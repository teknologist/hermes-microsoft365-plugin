"""Conditional schemas and native Hermes registration."""
from __future__ import annotations

import inspect
import json
from typing import Any, Callable

from .client import create_graph_client
from .contract import (
    OPERATION_REGISTRY,
    OPERATIONS,
    ConfigurationError,
    Settings,
    WRITE_OPERATIONS,
    operation_status,
)
from .execution import ExecutionError, run_async
from .errors import GraphError
from .handlers import HandlerContext, load_handlers
from .preflight import build_preflight
from .validation import check as check_operation_arguments, operation_schema


def _handler_context(settings: Settings | None) -> HandlerContext:
    """The context a registered handler runs with: the configuration of the running dispatch.

    The client factory is a *factory*: no secret is read, no credential is constructed and no
    client exists until a handler asks for one, which is after every check it makes itself.
    """
    if not isinstance(settings, Settings):
        raise ExecutionError("configuration_error", "Microsoft 365 dispatch configuration is ambiguous")
    return HandlerContext(
        settings=settings, client_factory=lambda: create_graph_client(settings)
    )


def _host_facing(record, settings: Settings | None) -> Callable[[dict], Any]:
    """One registered operation handler, in the call shape the host registers.

    ``handlers`` modules register ``(arguments, context)`` callables, which is what keeps the
    credential behind the injected context; the host calls a one-argument handler. The adapter
    is that boundary and nothing else: it builds the context and hands the arguments over
    untouched, so the payload a handler sees is exactly the payload the dispatch path gated.
    """

    def invoke(args):
        return record.function(args, _handler_context(settings))

    return invoke


def _dispatch_table(settings: Settings | None = None) -> dict[str, Callable[[dict], Any]]:
    """The dispatch mapping, filled from the self-populating handler registry (WP6).

    ``handlers.load_handlers()`` imports every module in ``microsoft365/handlers/`` and returns
    the mapping each of them registered at import time, so adding ``handlers/<service>.py``
    requires no edit here and there is no per-service import list to keep in step. The table
    holds **every** implemented operation, writes included: the operations this plugin may
    actually execute are decided by :func:`operation_is_executable`, never by membership here.
    """
    return {key: _host_facing(record, settings) for key, record in load_handlers().items()}


#: Every implemented operation handler, keyed ``"<service>.<operation>"``, filled from the
#: handler registry at import time. Every entry is executed through :func:`invoke_handler`,
#: which is the only place this plugin crosses the sync/async execution seam.
HANDLER_TABLE: dict[str, Callable[[dict], Any]] = _dispatch_table(None)
_DEFAULT_HANDLER_TABLE = dict(HANDLER_TABLE)


def invoke_handler(handler, args):
    """Run one registered handler, awaiting async handlers through the execution seam."""
    result = handler(args)
    if inspect.isawaitable(result):
        return run_async(result)
    return result


def tool_result(value) -> str:
    """Serialize a dispatch result into the string the host's tool contract accepts.

    The host registry (``tools/registry.py``) normalizes a handler result to a string or the
    multimodal envelope and turns anything else into a ``tool_result_contract`` error, so a
    handler payload that never crossed this boundary was reported as a failed call even when
    the Graph read succeeded. Handlers keep returning the structured ``dict`` they are
    contract-pinned on -- it is the *registration* boundary that owes the host a string, the
    same way ``microsoft365_preflight`` already does.

    ``json.dumps`` is the only transformation: the payload has already been through
    :func:`microsoft365.results.normalize_result` (redaction, string bounds, depth and cycle
    detection, binary artifacts outside string truncation), so nothing is re-bounded,
    re-redacted or dropped here. A value that is already a string (an error payload a
    dispatch path serialized itself) is handed over untouched.
    """
    if isinstance(value, str):
        return value
    return json.dumps(value, default=str)


def unavailable_handler(args, *, service: str) -> str:
    """Fallback for every operation that has no handler yet."""
    del args
    return json.dumps({"error": "operation_not_implemented", "service": service})


def not_executable_payload(service: str, operation: str) -> dict:
    """The refusal of an implemented operation that is not executable in this configuration.

    ``operation_not_implemented`` is the only taxonomy category that says "this plugin does not
    execute this operation" (the taxonomy is closed; ``errors.CATEGORIES`` is pinned by test),
    and the message names the operation so the two states stay distinguishable in a log: an
    operation with no handler at all, and an implemented operation withheld from the model.
    The message is the same one the ``pre_tool_call`` hook returns for the same call.
    """
    return {
        "error": "operation_not_implemented",
        "service": service,
        "operation": operation,
        "message": f"Microsoft 365 operation is not executable: {service}.{operation}",
        "retryable": False,
    }


def operation_is_executable(settings: Settings | None, service: str, operation) -> bool:
    """Re-evaluate the executability of one operation on the arguments about to run (invariant 15).

    ``HANDLER_TABLE`` holds the write handlers too, because WP6 implemented and contract-pinned
    them, so "a handler exists" must never be read as "this call may run". The verdict comes
    from the same gate the ``pre_tool_call`` hook uses -- ``active_actions`` over
    ``operation_status`` for the configured mode -- and is recomputed here, at dispatch time,
    from the *final* arguments. A modification that rewrites ``action`` from a read into a write
    after a hook approved the read payload (the CORE-1 defect class, and the scenario WP12
    exists for) therefore cannot reach a write handler: the write is not executable, so it is
    refused before any handler, secret, credential, client or request work.

    Without settings -- a direct dispatch, for instance from a test -- the registry's own
    declaration is the verdict. The configuration dimension is administrative (which operations
    an administrator selected); executability is not, so the weaker fallback is still
    fail-closed for every write.
    """
    if not isinstance(operation, str) or not operation:
        return False
    key = f"{service}.{operation}"
    definition = OPERATION_REGISTRY.get(key)
    if definition is None:
        return False
    if settings is None:
        return definition.executable
    return operation in active_actions(settings, service)


def service_tool_handler(
    service: str,
    args,
    settings: Settings | None = None,
    *,
    dispatch_table: dict[str, Callable[[dict], Any]] | None = None,
):
    """Dispatch one service tool call: registered handler, else unavailable (or refused).

    Order matters and is part of the contract: the operation is looked up first (so an
    operation this plugin does not implement keeps its explicit payload), then the
    executability gate re-evaluates the *final* arguments, and only then is a handler invoked.
    Nothing else happens before the gate -- no secret, no credential, no client, no request.
    """
    operation = args.get("action") if isinstance(args, dict) else None
    if dispatch_table is None:
        table = _dispatch_table(settings) if settings is not None else HANDLER_TABLE
    else:
        table = dispatch_table
        key = f"{service}.{operation}"
        # Keep the existing test/host seam: an explicit table entry replaced after
        # registration is an intentional override, never ambient dispatch state.
        if key in HANDLER_TABLE and HANDLER_TABLE[key] is not _DEFAULT_HANDLER_TABLE.get(key):
            table = {**table, key: HANDLER_TABLE[key]}
    handler = table.get(f"{service}.{operation}") if isinstance(operation, str) else None
    if handler is None:
        return unavailable_handler(args, service=service)
    if not operation_is_executable(settings, service, operation):
        return json.dumps(not_executable_payload(service, operation))
    if settings is not None:
        # Direct dispatch must enforce the same final-payload boundary as registration.
        rejection = check_operation_arguments(settings, service=service, arguments=args)
        if rejection is not None:
            return json.dumps(rejection.to_payload())
    try:
        return invoke_handler(handler, args)
    except ExecutionError as exc:
        return json.dumps({"error": exc.category, "message": str(exc)})
    except GraphError as exc:
        return json.dumps(exc.to_result())


def active_actions(settings: Settings, service: str) -> tuple[str, ...]:
    return tuple(
        operation
        for operation in OPERATIONS[service]
        if operation in settings.selected(service)
        and operation_status(settings.authentication_mode, service, operation).executable
    )


def schema_for(service: str, actions: tuple[str, ...]):
    """Build the registered model surface from the runtime argument contracts.

    Each branch is discriminated by a single ``action`` value.  Keeping the branches separate
    is important: a schema that unions all properties permits arguments from one operation to
    be smuggled into another, while the validator correctly rejects them.
    """
    if not actions:
        return None
    branches = []
    for action in actions:
        if service not in OPERATIONS or action not in OPERATIONS[service]:
            raise ValueError(f"{service}.{action} is not a declared operation")
        try:
            branches.append(operation_schema(f"{service}.{action}"))
        except KeyError as exc:
            raise ValueError(f"{service}.{action} has no argument contract") from exc
    outer_properties = {"action": {"type": "string", "enum": list(actions)}}
    # ``additionalProperties: false`` applies at this object level in JSON Schema;
    # declare every branch key here as well, while the oneOf branches remain the
    # authoritative validators for each action-specific value and requirement.
    for branch in branches:
        for property_name in branch["properties"]:
            outer_properties.setdefault(property_name, {})
    return {
        "name": f"microsoft365_{service}",
        "description": f"Microsoft 365 {service} operations with explicit host approval for writes.",
        "parameters": {
            "type": "object",
            "additionalProperties": False,
            # Keep the aggregate discriminator visible to hosts that inspect the outer object;
            # the oneOf branches remain authoritative for operation-specific properties.
            "properties": outer_properties,
            "required": ["action"],
            "oneOf": branches,
        },
    }


def _read_settings(ctx):
    raw = {
        key: ctx.get_config(key, default)
        for key, default in (
            ("tenant_id", ""),
            ("client_id", ""),
            ("user_id", ""),
            ("allowed_mailbox_ids", []),
            ("authentication_mode", "application"),
            ("capabilities", {}),
        )
    }
    try:
        return Settings.from_mapping(raw), []
    except ConfigurationError as exc:
        return None, [str(exc)]


def register_plugin(ctx) -> None:
    settings, errors = _read_settings(ctx)

    def preflight_handler(args, **kwargs):
        del args, kwargs
        return json.dumps(build_preflight(settings, configuration_errors=errors), default=str)

    ctx.register_tool(
        name="microsoft365_preflight",
        toolset="microsoft365",
        schema={
            "name": "microsoft365_preflight",
            "description": "Report local Microsoft 365 configuration, operation, and SDK status without contacting Graph.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        handler=preflight_handler,
        check_fn=lambda: True,
        is_async=False,
        emoji="🧩",
    )

    def pre_tool_call(*, tool_name="", args=None, **kwargs):
        del kwargs
        prefix = "microsoft365_"
        if settings is None or not isinstance(tool_name, str) or not tool_name.startswith(prefix):
            return None
        service = tool_name[len(prefix):]
        if service not in OPERATIONS:
            return None
        if not isinstance(args, dict):
            return {"action": "block", "message": "Microsoft 365 arguments must be an object"}
        # WP13: the operation-specific validator runs before anything else this hook can do --
        # before the approval directive, and before any handler could be reached. It is the
        # same function and the same payload the dispatch path re-checks (defence in depth),
        # so an invalid, disabled or unsupported call can never reach a secret, a credential
        # or a Graph client, and can never be approved.
        rejection = check_operation_arguments(settings, service=service, arguments=args)
        if rejection is not None:
            return {"action": "block", "message": rejection.message}
        operation = args.get("action")
        if operation not in active_actions(settings, service):
            return {"action": "block", "message": f"Microsoft 365 operation is not executable: {service}.{operation}"}
        if operation in WRITE_OPERATIONS:
            return {
                "action": "approve",
                "message": f"Microsoft 365 {operation}: external side effect",
                "rule_key": f"microsoft365.{service}.{operation}",
            }
        return None

    ctx.register_hook("pre_tool_call", pre_tool_call)
    if settings is None:
        return
    dispatch_table = _dispatch_table(settings)
    for service in OPERATIONS:
        actions = active_actions(settings, service)
        schema = schema_for(service, actions)
        if schema is None:
            continue

        def service_handler(args, _service=service, **kwargs):
            del kwargs
            # WP13 defence in depth: the dispatch path validates the very payload it is about
            # to execute with the same validator the hook used, so a call that did not pass the
            # hook (another dispatch path, a host that skips hooks, a payload changed after
            # approval) is refused here instead of reaching a handler, a secret or a client.
            rejection = check_operation_arguments(settings, service=_service, arguments=args)
            if rejection is not None:
                return json.dumps(rejection.to_payload())
            # Invariant 15: the settings travel with the call, so the executability gate inside
            # service_tool_handler is re-evaluated against this configuration and these final
            # arguments -- not against the payload an earlier hook saw.
            # The host's tool contract takes a string: a handler payload is serialized
            # here, at the registration boundary, so a successful read is reported as a
            # result instead of a ``tool_result_contract`` failure.
            return tool_result(
                service_tool_handler(_service, args, settings=settings, dispatch_table=dispatch_table)
            )

        ctx.register_tool(
            name=f"microsoft365_{service}",
            toolset="microsoft365",
            schema=schema,
            handler=service_handler,
            check_fn=lambda: True,
            is_async=False,
            emoji="📎",
        )
