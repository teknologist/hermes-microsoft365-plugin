"""Complete administrative operation and configuration contract."""
from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import Any, Mapping

OPERATIONS = {
    "outlook": ("search", "read", "list_attachments", "download_attachment", "create_draft", "send"),
    "sharepoint": ("search", "read", "download_files", "upload_files"),
    "onedrive": ("search", "read", "download_files", "upload_files"),
    "calendar": ("search", "create_events", "update_events"),
    "teams": ("list_teams", "list_channels", "search_messages", "send_messages"),
    "todo": ("list_task_lists", "search", "read", "create_tasks", "update_tasks"),
    "planner": ("list_plans", "list_buckets", "list_tasks", "read", "create_tasks", "update_tasks"),
}
WRITE_OPERATIONS = frozenset(
    {
        "create_draft", "send", "upload_files", "create_events", "update_events",
        "send_messages", "create_tasks", "update_tasks",
    }
)

# --------------------------------------------------------------------------------------
# Planner container semantics (WP2)
# --------------------------------------------------------------------------------------


class PlannerContractError(ValueError):
    """A Planner call would have been sent with an incomplete or malformed container."""


#: Claim states this repository is allowed to attach to an endpoint permission.
#:
#: ``documented_not_verified`` is the only state reachable offline: the role is the one
#: recorded for that endpoint (inherited from the pre-existing permission matrix, or from
#: the page named in ``documentation_page``), and no endpoint page was re-read and no
#: tenant was contacted while this milestone was built (rule R10; remote verification is
#: WP16). ``verified`` is reserved for a claim recomputed against the endpoint reference
#: **and** confirmed on a test tenant.
PERMISSION_CLAIM_STATUSES = frozenset({"documented_not_verified", "verified"})


# --------------------------------------------------------------------------------------
# Shared endpoint-table helpers
# --------------------------------------------------------------------------------------
#
# Every service whose endpoint contract is pinned declares a table of typed endpoint rows
# and derives its registry entry (permissions, container, identifiers, required headers,
# contract cases, claim statuses) from that table. The derivation is shared here so a
# service cannot quietly invent a different notion of "container" or "claim".


def _endpoint_container(endpoints) -> str:
    """The single container of an endpoint set, or the joined union when it has several."""
    containers = tuple(dict.fromkeys(endpoint.container for endpoint in endpoints))
    return containers[0] if len(containers) == 1 else "+".join(containers)


def _endpoint_permissions(endpoints) -> tuple[str, ...]:
    """Least-privileged application roles claimed by the endpoints themselves."""
    return tuple(
        dict.fromkeys(role for endpoint in endpoints for role in endpoint.application_permissions)
    )


def _required_endpoint_identifiers(endpoints) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    per_endpoint = [set(endpoint.required_identifiers) for endpoint in endpoints]
    shared = set.intersection(*per_endpoint) if per_endpoint else set()
    ordered = dict.fromkeys(
        name for endpoint in endpoints for name in endpoint.required_identifiers
    )
    return tuple(name for name in ordered if name in shared)


def _optional_endpoint_identifiers(endpoints) -> tuple[str, ...]:
    """Identifiers an endpoint accepts or requires without the operation requiring them."""
    required = set(_required_endpoint_identifiers(endpoints))
    ordered = dict.fromkeys(
        name for endpoint in endpoints for name in endpoint.required_identifiers
    )
    return tuple(name for name in ordered if name not in required)


def _missing_identifier_message(key: str, name: str, endpoints) -> str:
    where = "; ".join(sorted({endpoint.endpoint for endpoint in endpoints}))
    return f"{name} is required for {key} ({where})"


def _resolve_endpoint_identifiers(*, error, key, arguments, endpoints) -> dict[str, str]:
    """Normalize and validate the identifiers one endpoint set indexes by.

    Returns only the identifiers that will be sent, so a blank or absent required
    identifier -- and a blank optional one -- is refused with the endpoint it belongs to
    instead of being rendered into a request against the wrong container. ``None`` for an
    identifier the operation does not require means "not supplied" and selects the
    endpoint that has no such identifier; a blank string is an error and never a silent
    fallback to it. The caller supplies the typed error so each service keeps its own.
    """
    if not isinstance(arguments, Mapping):
        raise error(f"{key} arguments must be a mapping")
    resolved: dict[str, str] = {}

    def value(name: str, *, required: bool) -> str | None:
        raw = arguments.get(name)
        if raw is None:
            if required:
                raise error(_missing_identifier_message(key, name, endpoints))
            return None
        if not isinstance(raw, str):
            raise error(f"{name} for {key} must be a non-blank string")
        normalized = raw.strip()
        if normalized:
            return normalized
        if required:
            raise error(_missing_identifier_message(key, name, endpoints))
        raise error(f"{name} for {key} must be a non-blank string or be omitted")

    for name in _required_endpoint_identifiers(endpoints):
        resolved[name] = value(name, required=True)
    for name in _optional_endpoint_identifiers(endpoints):
        optional_value = value(name, required=False)
        if optional_value is not None:
            resolved[name] = optional_value
    return resolved


def _endpoint_backed_definition(
    *,
    service: str,
    key: str,
    endpoints,
    implementation_status: str | None = None,
    executable: bool = False,
) -> "OperationDefinition":
    """Registry entry derived from a service's own endpoint table.

    ``permissions``, ``container``, the identifier sets, the required headers, the contract
    cases and the claim statuses are all read from the endpoints, so a claim cannot drift
    from the endpoint it belongs to. The implementation status defaults to ``contract_verified``
    when every endpoint names the ``sdk_contract`` dispatcher and case a strict offline test
    pins, and to ``contract_foundation`` otherwise; a service with a handler behind every
    endpoint passes ``implementation_status="implemented"``. ``executable`` is never inferred --
    it is the caller's explicit declaration, because only a work package that owns the handler
    may expose an operation (R5, invariant 5).
    """
    operation = key.split(".", 1)[1]
    pinned = all(endpoint.contract_call and endpoint.contract_case for endpoint in endpoints)
    return OperationDefinition(
        service=service,
        operation=operation,
        permissions=_endpoint_permissions(endpoints),
        write=operation in WRITE_OPERATIONS,
        endpoints=tuple(endpoint.endpoint for endpoint in endpoints),
        container=_endpoint_container(endpoints),
        required_identifiers=_required_endpoint_identifiers(endpoints),
        optional_identifiers=_optional_endpoint_identifiers(endpoints),
        required_headers=tuple(
            dict.fromkeys(name for endpoint in endpoints for name in endpoint.required_headers)
        ),
        contract_cases=tuple(
            (endpoint.contract_call, endpoint.contract_case) for endpoint in endpoints
        ),
        permission_claim_statuses=tuple(
            dict.fromkeys(endpoint.claim_status for endpoint in endpoints)
        ),
        documentation_pages=tuple(
            dict.fromkeys(
                endpoint.documentation_page for endpoint in endpoints if endpoint.documentation_page
            )
        ),
        implementation_status=(
            implementation_status
            if implementation_status is not None
            else ("contract_verified" if pinned else "contract_foundation")
        ),
        executable=executable,
    )


@dataclass(frozen=True)
class PlannerEndpoint:
    """One verified request target of a Planner operation, with the container it indexes.

    Container semantics, recorded decision (WP2): a Planner operation indexes by exactly
    one kind of container, and the identifiers that identify that container must be
    non-blank before anything is sent. ``path_identifiers`` are rendered into
    ``path_template``; ``body_identifiers`` are identifiers the endpoint reads from the
    request body instead -- Planner *creation* carries ``plan_id`` and ``bucket_id`` in
    the body (``planId``/``bucketId``), never in the path, while listing buckets and
    tasks and reading or updating a task carry them in the path.
    """

    operation: str
    method: str
    path_template: str
    container: str
    path_identifiers: tuple[str, ...] = ()
    body_identifiers: tuple[str, ...] = ()
    application_permissions: tuple[str, ...] = ()
    claim_status: str = "documented_not_verified"
    contract_call: str = ""
    contract_case: str = ""
    required_headers: tuple[str, ...] = ()
    documentation_page: str = ""

    @property
    def endpoint(self) -> str:
        """The endpoint this row claims its application role from."""
        return f"{self.method} {self.path_template}"

    @property
    def required_identifiers(self) -> tuple[str, ...]:
        """Identifiers this endpoint cannot be called without."""
        return self.path_identifiers + self.body_identifiers


#: The complete Planner endpoint table. ``planner.list_plans`` is the only operation here
#: the SDK exposes twice: without ``group_id`` it reads the tenant-wide route, with it the
#: group-owned route.
PLANNER_ENDPOINTS: tuple[PlannerEndpoint, ...] = (
    PlannerEndpoint(
        operation="list_plans",
        method="GET",
        path_template="/planner/plans",
        container="tenant_or_group",
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_plans",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/planner-list-plans",
    ),
    PlannerEndpoint(
        operation="list_plans",
        method="GET",
        path_template="/groups/{group_id}/planner/plans",
        container="tenant_or_group",
        path_identifiers=("group_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_plans",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/planner-list-plans",
    ),
    PlannerEndpoint(
        operation="list_buckets",
        method="GET",
        path_template="/planner/plans/{plan_id}/buckets",
        container="plan",
        path_identifiers=("plan_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_buckets",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/planner-list-buckets",
    ),
    PlannerEndpoint(
        operation="list_tasks",
        method="GET",
        path_template="/planner/plans/{plan_id}/tasks",
        container="plan",
        path_identifiers=("plan_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="planner_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks",
    ),
    PlannerEndpoint(
        operation="read",
        method="GET",
        path_template="/planner/tasks/{planner_task_id}",
        container="task",
        path_identifiers=("planner_task_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_item_request_information",
        contract_case="planner_read",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/plannertask-get",
    ),
    PlannerEndpoint(
        operation="create_tasks",
        method="POST",
        path_template="/planner/tasks",
        container="plan_and_bucket",
        body_identifiers=("plan_id", "bucket_id"),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="planner_create_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/planner-post-tasks",
    ),
    PlannerEndpoint(
        operation="update_tasks",
        method="PATCH",
        path_template="/planner/tasks/{planner_task_id}",
        container="task",
        path_identifiers=("planner_task_id",),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="planner_update_tasks",
        required_headers=("If-Match",),
        documentation_page="https://learn.microsoft.com/en-us/graph/api/plannertask-update",
    ),
)

PLANNER_OPERATIONS: tuple[str, ...] = (
    "planner.list_plans",
    "planner.list_buckets",
    "planner.list_tasks",
    "planner.read",
    "planner.create_tasks",
    "planner.update_tasks",
)


def planner_endpoints(key: str) -> tuple[PlannerEndpoint, ...]:
    """Every verified endpoint of a Planner operation, in declaration order."""
    if key not in PLANNER_OPERATIONS:
        raise PlannerContractError(f"{key} is not a declared Planner operation")
    return tuple(
        endpoint for endpoint in PLANNER_ENDPOINTS if f"planner.{endpoint.operation}" == key
    )


def required_planner_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    return _required_endpoint_identifiers(planner_endpoints(key))


def optional_planner_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers an endpoint accepts or requires without the operation requiring them.

    ``group_id`` is the only one today: supplying it selects the group-owned plan route,
    omitting it (or passing ``None``) selects the tenant-wide route.
    """
    return _optional_endpoint_identifiers(planner_endpoints(key))


def validate_planner_identifiers(key: str, arguments: Mapping[str, Any]) -> dict[str, str]:
    """Normalize and validate the container identifiers a Planner call must carry.

    Returns only the identifiers that will be sent, so a blank or absent required
    identifier -- and a blank optional one -- is refused here, with the endpoint it
    belongs to, instead of being rendered into a request against the wrong container.
    ``None`` for an identifier the operation does not require means "not supplied": for
    ``planner.list_plans`` that selects the tenant-wide route, while a blank string is an
    error and never a silent fallback to it.
    """
    return _resolve_endpoint_identifiers(
        error=PlannerContractError,
        key=key,
        arguments=arguments,
        endpoints=planner_endpoints(key),
    )


# --------------------------------------------------------------------------------------
# Microsoft To Do semantics (WP3)
# --------------------------------------------------------------------------------------


class TodoContractError(ValueError):
    """A To Do call would have carried a field or filter the endpoint does not accept."""


@dataclass(frozen=True)
class TodoEndpoint:
    """One verified request target of a To Do operation, with the container it indexes.

    To Do indexes by **list** and by **task**: ``todo.list_task_lists`` reads the task lists
    of one user, and every other operation reads or writes a task inside one list. To Do
    addresses its container exclusively by path -- there is no body identifier -- so
    ``path_identifiers`` is the complete identifier set and every one of them must be
    non-blank before anything is sent.

    A claim may only call itself ``verified`` with recorded evidence: ``verified`` means the
    endpoint reference was recomputed **and** confirmed on a test tenant (WP16), so an
    offline milestone can only record ``documented_not_verified``.
    """

    operation: str
    method: str
    path_template: str
    container: str
    path_identifiers: tuple[str, ...] = ()
    application_permissions: tuple[str, ...] = ()
    claim_status: str = "documented_not_verified"
    claim_evidence: tuple[str, ...] = ()
    contract_call: str = ""
    contract_case: str = ""
    required_headers: tuple[str, ...] = ()
    documentation_page: str = ""

    def __post_init__(self) -> None:
        if self.claim_status not in PERMISSION_CLAIM_STATUSES:
            raise TodoContractError(
                f"{self.endpoint} declares unknown permission claim status {self.claim_status!r}"
            )
        if self.claim_status == "verified" and not self.claim_evidence:
            raise TodoContractError(
                f"{self.endpoint} claims a verified application role with no recorded evidence"
            )

    @property
    def endpoint(self) -> str:
        """The endpoint this row claims its application role from."""
        return f"{self.method} {self.path_template}"

    @property
    def required_identifiers(self) -> tuple[str, ...]:
        """Identifiers this endpoint cannot be called without."""
        return self.path_identifiers


#: The complete To Do endpoint table. Every role below is attributed to the endpoint it
#: belongs to instead of to a blanket "To Do" family, and no row carries a recorded endpoint
#: reference unless the repository already recorded one: only the task listing has one, so
#: the other four report ``not recorded`` rather than an invented URL.
TODO_ENDPOINTS: tuple[TodoEndpoint, ...] = (
    TodoEndpoint(
        operation="list_task_lists",
        method="GET",
        path_template="/users/{user_id}/todo/lists",
        container="user",
        path_identifiers=("user_id",),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="todo_lists",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/todo-list-lists",
    ),
    TodoEndpoint(
        operation="search",
        method="GET",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks",
        container="list",
        path_identifiers=("user_id", "todo_list_id"),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_collection_request_information",
        contract_case="todo_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks",
    ),
    TodoEndpoint(
        operation="read",
        method="GET",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        container="task",
        path_identifiers=("user_id", "todo_list_id", "todo_task_id"),
        application_permissions=("Tasks.Read.All",),
        contract_call="build_item_request_information",
        contract_case="todo_read",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/todotask-get",
    ),
    TodoEndpoint(
        operation="create_tasks",
        method="POST",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks",
        container="list",
        path_identifiers=("user_id", "todo_list_id"),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="todo_create_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks",
    ),
    TodoEndpoint(
        operation="update_tasks",
        method="PATCH",
        path_template="/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        container="task",
        path_identifiers=("user_id", "todo_list_id", "todo_task_id"),
        application_permissions=("Tasks.ReadWrite.All",),
        contract_call="build_write_request_information",
        contract_case="todo_update_tasks",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/todotask-update",
    ),
)

TODO_OPERATIONS: tuple[str, ...] = (
    "todo.list_task_lists",
    "todo.search",
    "todo.read",
    "todo.create_tasks",
    "todo.update_tasks",
)

#: The To Do operations that write. Their application-mode support is derived from the
#: endpoint table above by :func:`todo_write_support`, never asserted here.
TODO_WRITE_OPERATIONS: frozenset[str] = frozenset({"todo.create_tasks", "todo.update_tasks"})

#: The ``sdk_contract`` case names of the To Do writes, read from the endpoint table so the
#: field policy below is keyed by the same name the request builders use.
TODO_TASK_WRITE_CASES: frozenset[str] = frozenset(
    endpoint.contract_case
    for endpoint in TODO_ENDPOINTS
    if f"todo.{endpoint.operation}" in TODO_WRITE_OPERATIONS
)


def todo_endpoints(key: str) -> tuple[TodoEndpoint, ...]:
    """Every verified endpoint of a To Do operation, in declaration order."""
    if key not in TODO_OPERATIONS:
        raise TodoContractError(f"{key} is not a declared To Do operation")
    return tuple(endpoint for endpoint in TODO_ENDPOINTS if f"todo.{endpoint.operation}" == key)


def required_todo_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a call must supply whichever endpoint of the operation it uses."""
    return _required_endpoint_identifiers(todo_endpoints(key))


def optional_todo_identifiers(key: str) -> tuple[str, ...]:
    """Identifiers a To Do endpoint accepts without the operation requiring them.

    Empty today: unlike ``planner.list_plans``, no To Do operation has a second endpoint
    selectable by supplying an extra identifier.
    """
    return _optional_endpoint_identifiers(todo_endpoints(key))


def validate_todo_identifiers(key: str, arguments: Mapping[str, Any]) -> dict[str, str]:
    """Normalize and validate the container identifiers a To Do call must carry.

    Refuses a blank or absent required identifier -- naming the endpoint it belongs to --
    before any approval, client or token work happens, so a To Do request can never be
    rendered against the wrong list or task.
    """
    return _resolve_endpoint_identifiers(
        error=TodoContractError,
        key=key,
        arguments=arguments,
        endpoints=todo_endpoints(key),
    )


#: The only To Do task fields a write may carry, and where each accepted field comes from.
#:
#: The generated ``TodoTask`` model declares 22 Graph properties; this milestone cannot
#: re-read the endpoint reference offline (R10), so the writable set stays deliberately
#: minimal and fail-closed. A field outside this mapping is refused with the case name --
#: including ``subject``, which no To Do model has, ``id`` and the ``*DateTime`` stamps the
#: service maintains itself, and every collection property -- instead of being sent and
#: silently dropped by Graph. This is a strict subset of the model on purpose: refusing a
#: field the endpoint might accept is honest, accepting a field it ignores is not.
TODO_TASK_WRITE_FIELDS: dict[str, str] = {
    "title": "the task text: `TodoTask.title`, the field the To Do request contract pins",
    "body": "the generated `TodoTask.body` (`ItemBody`)",
    "due_date_time": "the generated `TodoTask.due_date_time` (`DateTimeTimeZone`)",
    "status": "the generated `TodoTask.status` (`TaskStatus`), serialized only as an enum member",
}


def validate_todo_task_fields(case: str, fields: Mapping[str, Any]) -> dict[str, Any]:
    """Refuse any To Do task field the endpoint contract does not accept.

    ``case`` is the ``sdk_contract`` case name of the write (``todo_update_tasks``), the same
    spelling the request builder uses. Returns the accepted mapping; the SDK layer maps the
    enum-valued entry to its generated member and type-checks the model-valued ones. ``None``
    is refused for the same reason a field outside the policy is: it carries no value, so it
    would vanish from the request body instead of changing anything.
    """
    if case not in TODO_TASK_WRITE_CASES:
        raise TodoContractError(f"{case} is not a declared To Do write case")
    if not isinstance(fields, Mapping) or not fields:
        raise TodoContractError(f"{case} requires a non-empty mapping of task fields")
    accepted: dict[str, Any] = {}
    for name, value in fields.items():
        if name not in TODO_TASK_WRITE_FIELDS:
            raise TodoContractError(
                f"unknown field for {case}: {name}; writable To Do task fields are "
                f"{', '.join(TODO_TASK_WRITE_FIELDS)}"
            )
        if value is None:
            raise TodoContractError(f"{name} of {case} must carry a value, not None")
        accepted[name] = value
    return accepted


#: What an application-mode write support status is allowed to mean.
WRITE_SUPPORT_STATUSES = frozenset({"not_verified", "supported"})


@dataclass(frozen=True)
class ApplicationWriteSupport:
    """How much is known about one write endpoint in application mode.

    ``supported`` requires recorded evidence -- at least one endpoint of the write whose
    permission claim is ``verified``, which needs the endpoint reference recomputed *and*
    confirmed on a test tenant (WP16). Without that evidence the status is ``not_verified``,
    the operation stays non-executable, and no other field of this record can change that.
    """

    operation: str
    status: str
    reason: str
    evidence: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if self.status not in WRITE_SUPPORT_STATUSES:
            raise TodoContractError(f"unknown application write support status: {self.status!r}")
        if self.status == "supported" and not self.evidence:
            raise TodoContractError(
                f"{self.operation} cannot be reported as supported in application mode "
                "without recorded endpoint evidence"
            )

    @property
    def supported(self) -> bool:
        return self.status == "supported"


#: The one reason a To Do write is reported ``not_verified``. Quoted verbatim in
#: ``references/graph-permissions.md`` so the document and the registry cannot disagree.
TODO_WRITE_NOT_VERIFIED_REASON = (
    "the application permission of this write is unverified: no endpoint reference was "
    "re-read offline and no test tenant was contacted (R10; remote verification is WP16)"
)


def todo_write_support(key: str) -> ApplicationWriteSupport:
    """Application-mode support of a To Do write, derived from its own endpoints.

    Derived, never asserted: the status flips to ``supported`` only when every endpoint of
    the write carries a ``verified`` application claim with recorded evidence, and stays
    ``not_verified`` until then. ``TodoEndpoint`` refuses a ``verified`` claim without
    evidence, so a promotion cannot be faked by editing one string.
    """
    if key not in TODO_WRITE_OPERATIONS:
        raise TodoContractError(f"{key} is not a declared To Do write operation")
    endpoints = todo_endpoints(key)
    verified = tuple(
        endpoint.endpoint
        for endpoint in endpoints
        if endpoint.claim_status == "verified" and endpoint.claim_evidence
    )
    if endpoints and len(verified) == len(endpoints):
        return ApplicationWriteSupport(
            operation=key,
            status="supported",
            reason="every endpoint of this write carries a verified application claim",
            evidence=verified,
        )
    return ApplicationWriteSupport(
        operation=key,
        status="not_verified",
        reason=TODO_WRITE_NOT_VERIFIED_REASON,
    )


#: Permission claims for the services whose endpoint contract the registry does not declare
#: yet. Planner, To Do, Outlook and Calendar rows are deliberately absent: each of their
#: operations derives its claim from its own endpoint table, so a blanket per-service claim
#: cannot drift from the endpoint it belongs to.
_PERMISSION_MAP = {
    "teams.list_teams": ("Team.ReadBasic.All",), "teams.list_channels": ("Channel.ReadBasic.All",),
    "teams.search_messages": (), "teams.send_messages": (),
}
_UNSUPPORTED_APPLICATION = frozenset({"teams.search_messages", "teams.send_messages"})


# --------------------------------------------------------------------------------------
# Outlook and Calendar endpoint contracts (WP6)
# --------------------------------------------------------------------------------------
# Outlook and Calendar address one user's mailbox and one user's calendar. Their endpoint
# table is declared here -- instead of staying in the blanket per-service permission map --
# so the role, the container, the identifier set and the strict offline contract case of
# every operation are attributed to the endpoint they belong to, exactly as Planner and To Do
# already do. No ``path_template`` below was assembled by hand: each one was rendered with the
# generated builder of the ``sdk_contract`` case named next to it, and
# ``tests/test_handlers_outlook_calendar.py`` rebuilds every row and compares the rendered
# method and path with the declaration, so an invented endpoint fails a test.
#
# No documentation page is recorded for these rows: no endpoint reference was re-read offline
# (R10), so the claim stays ``documented_not_verified`` and the page is reported as
# "not recorded" rather than guessed.


@dataclass(frozen=True)
class MessagingEndpoint:
    """One verified request target of an Outlook or Calendar operation.

    ``outlook.send`` is the only operation here with two endpoints: a composed message goes to
    ``/users/{user_id}/sendMail``, while sending a draft that already exists goes to
    ``/users/{user_id}/messages/{message_id}/send``. Each row therefore carries its own
    ``path_identifiers``: the identifier set of the *operation* is the intersection of its
    rows, so the operation requires only ``user_id`` while ``message_id`` stays optional, and
    the union keeps ``message_id`` visible instead of silently unlisted.
    """

    service: str
    operation: str
    method: str
    path_template: str
    container: str
    path_identifiers: tuple[str, ...] = ()
    application_permissions: tuple[str, ...] = ()
    claim_status: str = "documented_not_verified"
    contract_call: str = ""
    contract_case: str = ""
    required_headers: tuple[str, ...] = ()
    documentation_page: str = ""

    @property
    def key(self) -> str:
        return f"{self.service}.{self.operation}"

    @property
    def endpoint(self) -> str:
        """The endpoint this row claims its application role from."""
        return f"{self.method} {self.path_template}"

    @property
    def required_identifiers(self) -> tuple[str, ...]:
        """Identifiers this endpoint cannot be called without."""
        return self.path_identifiers


#: The single container kind of both services: one user's mailbox or calendar, addressed by
#: ``user_id``. Neither service has a tenant-wide route here.
MESSAGING_CONTAINER = "user"

OUTLOOK_ENDPOINTS: tuple[MessagingEndpoint, ...] = (
    MessagingEndpoint(
        service="outlook", operation="download_attachment", method="GET",
        path_template="/users/{user_id}/messages/{message_id}/attachments/{attachment_id}",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id", "message_id", "attachment_id"),
        application_permissions=("Mail.Read",),
        contract_call="build_item_request_information", contract_case="outlook_attachment",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/attachment-get",
    ),
    MessagingEndpoint(
        service="outlook",
        operation="list_attachments",
        method="GET",
        path_template="/users/{user_id}/messages/{message_id}/attachments",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id", "message_id"),
        application_permissions=("Mail.Read",),
        contract_call="build_collection_request_information",
        contract_case="outlook_attachments",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/message-list-attachments",
    ),
    MessagingEndpoint(
        service="outlook",
        operation="search",
        method="GET",
        path_template="/users/{user_id}/messages",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id",),
        application_permissions=("Mail.Read",),
        contract_call="build_collection_request_information",
        contract_case="outlook_messages",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/user-list-messages",
    ),
    MessagingEndpoint(
        service="outlook",
        operation="read",
        method="GET",
        path_template="/users/{user_id}/messages/{message_id}",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id", "message_id"),
        application_permissions=("Mail.Read",),
        contract_call="build_item_request_information",
        contract_case="outlook_read",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/message-get",
    ),
    MessagingEndpoint(
        service="outlook",
        operation="create_draft",
        method="POST",
        path_template="/users/{user_id}/messages",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id",),
        application_permissions=("Mail.ReadWrite",),
        contract_call="build_write_request_information",
        contract_case="outlook_create_draft",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/user-post-messages",
    ),
    MessagingEndpoint(
        service="outlook",
        operation="send",
        method="POST",
        path_template="/users/{user_id}/sendMail",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id",),
        application_permissions=("Mail.Send",),
        contract_call="build_write_request_information",
        contract_case="outlook_send",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/user-sendmail",
    ),
    MessagingEndpoint(
        service="outlook",
        operation="send",
        method="POST",
        path_template="/users/{user_id}/messages/{message_id}/send",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id", "message_id"),
        application_permissions=("Mail.Send",),
        contract_call="build_write_request_information",
        contract_case="outlook_send_existing",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/message-send",
    ),
)

CALENDAR_ENDPOINTS: tuple[MessagingEndpoint, ...] = (
    MessagingEndpoint(
        service="calendar",
        operation="search",
        method="GET",
        path_template="/users/{user_id}/calendar/events",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id",),
        application_permissions=("Calendars.Read",),
        contract_call="build_collection_request_information",
        contract_case="calendar_events",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/user-list-events",
    ),
    MessagingEndpoint(
        service="calendar",
        operation="create_events",
        method="POST",
        path_template="/users/{user_id}/calendar/events",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id",),
        application_permissions=("Calendars.ReadWrite",),
        contract_call="build_write_request_information",
        contract_case="calendar_create_events",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/user-post-events",
    ),
    MessagingEndpoint(
        service="calendar",
        operation="update_events",
        method="PATCH",
        path_template="/users/{user_id}/calendar/events/{event_id}",
        container=MESSAGING_CONTAINER,
        path_identifiers=("user_id", "event_id"),
        application_permissions=("Calendars.ReadWrite",),
        contract_call="build_write_request_information",
        contract_case="calendar_update_events",
        # A conditional update: the handler refuses a call without an ETag, so the endpoint
        # itself declares the header it cannot run without.
        required_headers=("If-Match",),
        documentation_page="https://learn.microsoft.com/en-us/graph/api/event-update",
    ),
)

#: The complete endpoint table of both services, in declaration order.
MESSAGING_ENDPOINTS: tuple[MessagingEndpoint, ...] = OUTLOOK_ENDPOINTS + CALENDAR_ENDPOINTS

MESSAGING_OPERATIONS: tuple[str, ...] = (
    "outlook.download_attachment",
    "outlook.list_attachments",
    "outlook.search",
    "outlook.read",
    "outlook.create_draft",
    "outlook.send",
    "calendar.search",
    "calendar.create_events",
    "calendar.update_events",
)

#: The operations this milestone exposes to the model: the three verified reads whose handler,
#: contract case and endpoint row all exist. It is deliberately a literal, not a derivation:
#: exposing an operation is a decision, and every write stays non-executable until the generic
#: host approval fix (CORE-1/CORE-2) is available in a supported Hermes release (R5, D1).
EXECUTABLE_OPERATIONS: tuple[str, ...] = (
    "outlook.download_attachment",
    "outlook.list_attachments",
    "outlook.search",
    "outlook.read",
    "calendar.search",
    "sharepoint.search",
    "sharepoint.read",
    "sharepoint.download_files",
    "onedrive.search",
    "onedrive.read",
    "onedrive.download_files",
    "teams.list_teams",
    "teams.list_channels",
    "todo.list_task_lists",
    "todo.search",
    "todo.read",
    "planner.list_plans",
    "planner.list_buckets",
    "planner.list_tasks",
    "planner.read",
)


def messaging_endpoints(key: str) -> tuple[MessagingEndpoint, ...]:
    """Every verified endpoint of an Outlook or Calendar operation, in declaration order."""
    if key not in MESSAGING_OPERATIONS:
        raise ValueError(f"{key} is not a declared outlook or calendar operation")
    return tuple(endpoint for endpoint in MESSAGING_ENDPOINTS if endpoint.key == key)


# --------------------------------------------------------------------------------------
# SharePoint and OneDrive endpoint contracts (WP7)
# --------------------------------------------------------------------------------------
# SharePoint and OneDrive both address files through a drive; SharePoint adds sites, whose
# default document library is a drive. The endpoint table below is declared here -- instead
# of staying in the blanket per-service permission map -- so the role, the container, the
# identifier set and the strict offline contract case of every operation are attributed to
# the endpoint they belong to, exactly as Planner, To Do and Outlook/Calendar do. No
# ``path_template`` below was assembled by hand: each one was rendered with the generated
# builder of the ``sdk_contract`` case named next to it, and ``tests/test_handlers_files.py``
# rebuilds every row and compares the rendered method and path with the declaration, so an
# invented endpoint fails a test.
#
# No documentation page is recorded for these rows: no endpoint reference was re-read offline
# (R10), so the claim stays ``documented_not_verified`` and the page is reported as
# "not recorded" rather than guessed.


@dataclass(frozen=True)
class DriveEndpoint:
    """One verified request target of a SharePoint or OneDrive operation.

    ``sharepoint.search`` is the only operation here with several request targets: the
    tenant-wide site search, a drive-scoped search, and the site→drive resolution step that
    precedes a site-scoped search. The identifier set of the *operation* is the intersection
    of its rows, so search requires no identifier while ``drive_id`` and ``site_id`` stay
    optional (selecting the row), and the union keeps both visible instead of silently
    unlisted. The ``drive_item_id`` path identifier of ``read``/``download_files``/
    ``upload_files`` is the generated addressing string ``"<root|item-id>:/<relative/path>"``
    that ``microsoft365.files.drive_item_address`` produces from the caller's ``item_id`` or
    ``item_path``.
    """

    service: str
    operation: str
    method: str
    path_template: str
    container: str
    path_identifiers: tuple[str, ...] = ()
    application_permissions: tuple[str, ...] = ()
    claim_status: str = "documented_not_verified"
    contract_call: str = ""
    contract_case: str = ""
    required_headers: tuple[str, ...] = ()
    documentation_page: str = ""

    @property
    def key(self) -> str:
        return f"{self.service}.{self.operation}"

    @property
    def endpoint(self) -> str:
        return f"{self.method} {self.path_template}"

    @property
    def required_identifiers(self) -> tuple[str, ...]:
        return self.path_identifiers


DRIVE_ENDPOINTS: tuple[DriveEndpoint, ...] = (
    DriveEndpoint(
        service="sharepoint",
        operation="search",
        method="GET",
        path_template="/sites",
        container="tenant",
        application_permissions=("Sites.Read.All",),
        contract_call="build_search_request_information",
        contract_case="sites_search",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/site-list",
    ),
    DriveEndpoint(
        service="sharepoint",
        operation="search",
        method="GET",
        path_template="/drives/{drive_id}/search(q='{query}')",
        container="drive",
        path_identifiers=("drive_id",),
        application_permissions=("Sites.Read.All",),
        contract_call="build_search_request_information",
        contract_case="drive_search",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-search",
    ),
    DriveEndpoint(
        service="sharepoint",
        operation="search",
        method="GET",
        path_template="/sites/{site_id}/drive",
        container="site",
        path_identifiers=("site_id",),
        application_permissions=("Sites.Read.All",),
        documentation_page="https://learn.microsoft.com/en-us/graph/api/site-get",
    ),
    DriveEndpoint(
        service="sharepoint",
        operation="read",
        method="GET",
        path_template="/drives/{drive_id}/items/{drive_item_id}",
        container="drive",
        path_identifiers=("drive_id", "drive_item_id"),
        application_permissions=("Sites.Read.All",),
        contract_call="build_item_request_information",
        contract_case="drive_item_read",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-get",
    ),
    DriveEndpoint(
        service="sharepoint",
        operation="download_files",
        method="GET",
        path_template="/drives/{drive_id}/items/{drive_item_id}/content",
        container="drive",
        path_identifiers=("drive_id", "drive_item_id"),
        application_permissions=("Files.Read.All",),
        contract_call="build_content_request_information",
        contract_case="download_files",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-get-content",
    ),
    DriveEndpoint(
        service="sharepoint",
        operation="upload_files",
        method="PUT",
        path_template="/drives/{drive_id}/items/{drive_item_id}/content",
        container="drive",
        path_identifiers=("drive_id", "drive_item_id"),
        application_permissions=("Files.ReadWrite.All",),
        contract_call="build_content_request_information",
        contract_case="upload_files",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-put-content",
    ),
    DriveEndpoint(
        service="onedrive",
        operation="search",
        method="GET",
        path_template="/drives/{drive_id}/search(q='{query}')",
        container="drive",
        path_identifiers=("drive_id",),
        application_permissions=("Files.Read.All",),
        contract_call="build_search_request_information",
        contract_case="drive_search",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-search",
    ),
    DriveEndpoint(
        service="onedrive",
        operation="read",
        method="GET",
        path_template="/drives/{drive_id}/items/{drive_item_id}",
        container="drive",
        path_identifiers=("drive_id", "drive_item_id"),
        application_permissions=("Files.Read.All",),
        contract_call="build_item_request_information",
        contract_case="drive_item_read",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-get",
    ),
    DriveEndpoint(
        service="onedrive",
        operation="download_files",
        method="GET",
        path_template="/drives/{drive_id}/items/{drive_item_id}/content",
        container="drive",
        path_identifiers=("drive_id", "drive_item_id"),
        application_permissions=("Files.Read.All",),
        contract_call="build_content_request_information",
        contract_case="download_files",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-get-content",
    ),
    DriveEndpoint(
        service="onedrive",
        operation="upload_files",
        method="PUT",
        path_template="/drives/{drive_id}/items/{drive_item_id}/content",
        container="drive",
        path_identifiers=("drive_id", "drive_item_id"),
        application_permissions=("Files.ReadWrite.All",),
        contract_call="build_content_request_information",
        contract_case="upload_files",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/driveitem-put-content",
    ),
)

DRIVE_OPERATIONS: tuple[str, ...] = (
    "sharepoint.search",
    "sharepoint.read",
    "sharepoint.download_files",
    "sharepoint.upload_files",
    "onedrive.search",
    "onedrive.read",
    "onedrive.download_files",
    "onedrive.upload_files",
)


def drive_endpoints(key: str) -> tuple[DriveEndpoint, ...]:
    """Every verified endpoint of a SharePoint or OneDrive operation, in declaration order."""
    if key not in DRIVE_OPERATIONS:
        raise ValueError(f"{key} is not a declared sharepoint or onedrive operation")
    return tuple(endpoint for endpoint in DRIVE_ENDPOINTS if endpoint.key == key)
# Teams Graph endpoint contracts (WP8)
# --------------------------------------------------------------------------------------
# Teams Graph has two application-readable collections and two delegated-only operations.
# The endpoint rows keep the path, permission role, and strict sdk_contract case together.
# ``joinedTeams`` and ``channels`` deliberately do not expose ``$top``: the Graph endpoint
# does not support it, so handlers bound the number of processed response items instead.
TEAMS_ENDPOINTS: tuple[MessagingEndpoint, ...] = (
    MessagingEndpoint(
        service="teams",
        operation="list_teams",
        method="GET",
        path_template="/users/{user_id}/joinedTeams",
        container="user",
        path_identifiers=("user_id",),
        application_permissions=("Team.ReadBasic.All",),
        contract_call="build_collection_request_information",
        contract_case="teams_joined",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams",
    ),
    MessagingEndpoint(
        service="teams",
        operation="list_channels",
        method="GET",
        path_template="/teams/{team_id}/channels",
        container="team",
        path_identifiers=("team_id",),
        application_permissions=("Channel.ReadBasic.All",),
        contract_call="build_collection_request_information",
        contract_case="teams_channels",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/channel-list",
    ),
    MessagingEndpoint(
        service="teams",
        operation="search_messages",
        method="POST",
        path_template="/search/query",
        container="search",
        contract_case="teams_search_messages",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/search-query",
        contract_call="build_write_request_information",
    ),
    MessagingEndpoint(
        service="teams",
        operation="send_messages",
        method="POST",
        path_template="/teams/{team_id}/channels/{channel_id}/messages",
        container="channel",
        path_identifiers=("team_id", "channel_id"),
        contract_case="teams_send_messages",
        documentation_page="https://learn.microsoft.com/en-us/graph/api/channel-post-messages",
        contract_call="build_write_request_information",
    ),
)
TEAMS_OPERATIONS: tuple[str, ...] = tuple(f"teams.{row.operation}" for row in TEAMS_ENDPOINTS)


def teams_endpoints(key: str) -> tuple[MessagingEndpoint, ...]:
    """Every endpoint row declared for one Teams Graph operation."""
    if key not in TEAMS_OPERATIONS:
        raise ValueError(f"{key} is not a declared Teams operation")
    return tuple(row for row in TEAMS_ENDPOINTS if row.key == key)


# --------------------------------------------------------------------------------------
# Per-authentication-mode operation matrix (WP13)
# --------------------------------------------------------------------------------------
#
# One operation, two authentication modes, two independent records. Application support and
# delegated support are stored separately -- each with its own status, its own permissions
# (application **roles** vs delegated **scopes**), the endpoint it belongs to, the write
# classification, the implementation status and the remote verification status -- because
# the two are different permission systems and one must never be derived from the other.
# The delegated scope table below is a literal record, not a mapping of the application
# role map: ``planner.read`` holds ``Tasks.Read.All`` as an app role and ``Tasks.Read`` as a
# delegated scope, and ``teams.search_messages``/``teams.send_messages`` claim no
# application role at all while still declaring the scopes delegated mode would need.

#: The authentication modes this plugin records support for.
AUTH_MODES = frozenset({"application", "delegated"})

#: The modes this plugin actually implements. Delegated is fully described in the matrix
#: (status, scopes, admin consent) but no delegated authentication flow exists yet (WP14), so
#: every delegated record resolves to ``not_implemented``: authentication is not implemented
#: for that mode, which is a different statement from an operation being unavailable in it.
IMPLEMENTED_AUTH_MODES = frozenset({"application"})

#: What a mode-level support status is allowed to mean.
MODE_SUPPORT_STATUSES = frozenset(
    {"supported", "not_verified", "unsupported_auth_mode", "not_implemented"}
)

#: What the endpoint permission claim of one mode is allowed to be. ``no_permission_claimed``
#: is the only status an unsupported mode may carry: nothing is claimed, so nothing is used.
PERMISSION_STATUSES = frozenset({"no_permission_claimed", "documented_not_verified", "verified"})

#: Remote verification is ``not_tested`` until WP16 runs against a disposable tenant, and
#: ``verified`` additionally requires recorded evidence.
REMOTE_VERIFICATION_STATUSES = frozenset({"not_tested", "verified"})

#: Evidence for a ``verified`` remote status. Empty in this repository on purpose: no tenant
#: has been contacted (R10; WP16 is the only work package allowed to fill it).
REMOTE_VERIFICATION_EVIDENCE: tuple[str, ...] = ()

#: Why an application permission always requires a tenant administrator.
APPLICATION_CONSENT_BASIS = (
    "an application permission is an app role: only a tenant administrator can grant it, "
    "so nothing an application-mode call needs can be consented to by a user"
)

#: Why a delegated scope is reported as requiring consent while delegated is not implemented.
UNVERIFIED_DELEGATED_CONSENT_BASIS = (
    "the delegated consent record was not re-read offline and delegated authentication is "
    "not implemented (R10; WP14), so consent is reported as required until it is recorded"
)

#: Why an operation with no claimed permission has nothing to consent to.
NO_PERMISSION_CONSENT_BASIS = (
    "no permission is claimed for this operation in this mode, so there is nothing to consent to"
)

#: The delegated scopes of every operation, recorded independently of ``_PERMISSION_MAP``.
#: These are the scopes a delegated implementation (WP14) would request; recording them here
#: does not make delegated authentication available, and no row of the application map is
#: used to build them.
_DELEGATED_SCOPES: dict[str, tuple[str, ...]] = {
    "outlook.download_attachment": ("Mail.Read",),
    "outlook.list_attachments": ("Mail.Read",),
    "outlook.search": ("Mail.Read",),
    "outlook.read": ("Mail.Read",),
    "outlook.create_draft": ("Mail.ReadWrite",),
    "outlook.send": ("Mail.Send",),
    "sharepoint.search": ("Sites.Read.All",),
    "sharepoint.read": ("Sites.Read.All",),
    "sharepoint.download_files": ("Files.Read.All",),
    "sharepoint.upload_files": ("Files.ReadWrite.All",),
    "onedrive.search": ("Files.Read.All",),
    "onedrive.read": ("Files.Read.All",),
    "onedrive.download_files": ("Files.Read.All",),
    "onedrive.upload_files": ("Files.ReadWrite.All",),
    "calendar.search": ("Calendars.Read",),
    "calendar.create_events": ("Calendars.ReadWrite",),
    "calendar.update_events": ("Calendars.ReadWrite",),
    "teams.list_teams": ("Team.ReadBasic.All",),
    "teams.list_channels": ("Channel.ReadBasic.All",),
    "teams.search_messages": ("Chat.Read", "ChannelMessage.Read.All"),
    "teams.send_messages": ("ChannelMessage.Send",),
    "todo.list_task_lists": ("Tasks.Read",),
    "todo.search": ("Tasks.Read",),
    "todo.read": ("Tasks.Read",),
    "todo.create_tasks": ("Tasks.ReadWrite",),
    "todo.update_tasks": ("Tasks.ReadWrite",),
    "planner.list_plans": ("Tasks.Read",),
    "planner.list_buckets": ("Tasks.Read",),
    "planner.list_tasks": ("Tasks.Read",),
    "planner.read": ("Tasks.Read",),
    "planner.create_tasks": ("Tasks.ReadWrite",),
    "planner.update_tasks": ("Tasks.ReadWrite",),
}


@dataclass(frozen=True)
class ModeSupport:
    """What one authentication mode supports for one operation, with its own claim.

    ``permissions`` are the application **roles** when ``mode`` is ``application`` and the
    delegated **scopes** when it is ``delegated``; the two are never derived from each other.
    Every field is validated here and fails closed, so a record cannot be promoted by editing
    one string:

    * a ``verified`` permission claim needs recorded evidence (WP16 is the only work that can
      record it);
    * a supported/not-verified mode must claim the permission it needs, and an unsupported
      mode may claim none at all;
    * admin consent must be stated with its basis, and a mode that claims no permission has
      nothing to consent to -- so the honest answer for an unverified delegated record is
      ``admin_consent = True`` with an explicit basis rather than a silent "not required".
    """

    mode: str
    status: str
    permissions: tuple[str, ...] = ()
    permission_status: str = "no_permission_claimed"
    permission_evidence: tuple[str, ...] = ()
    admin_consent: bool = False
    admin_consent_basis: str = ""
    reason: str = ""

    def __post_init__(self) -> None:
        if self.mode not in AUTH_MODES:
            raise ValueError(f"unknown authentication mode: {self.mode!r}")
        if self.status not in MODE_SUPPORT_STATUSES:
            raise ValueError(f"unknown authentication mode support status: {self.status!r}")
        if self.permission_status not in PERMISSION_STATUSES:
            raise ValueError(f"unknown permission status: {self.permission_status!r}")
        if self.permission_status == "verified" and not self.permission_evidence:
            raise ValueError(
                f"{self.mode} support cannot make a verified permission claim without "
                "recorded evidence"
            )
        if self.permissions and self.permission_status == "no_permission_claimed":
            raise ValueError(
                f"{self.mode} support declares permissions with no recorded permission claim"
            )
        if self.status == "unsupported_auth_mode" and self.permission_status != "no_permission_claimed":
            raise ValueError(
                f"{self.mode} support cannot claim a permission on an unsupported mode"
            )
        if not self.permissions and self.permission_status != "no_permission_claimed":
            raise ValueError(
                f"{self.mode} support claims a permission it does not declare"
            )
        if self.status in {"supported", "not_verified"} and not self.permissions:
            raise ValueError(
                f"{self.mode} support must claim the permission it needs"
            )
        if self.admin_consent and not self.permissions:
            raise ValueError(
                f"{self.mode} support cannot require admin consent without claiming a permission"
            )
        if self.admin_consent and not self.admin_consent_basis.strip():
            raise ValueError(
                f"{self.mode} support that requires admin consent must record the basis"
            )
        if not self.admin_consent and not self.admin_consent_basis.strip():
            raise ValueError(
                f"{self.mode} support that does not require admin consent must state why "
                "consent is not required"
            )

    @property
    def permission_kind(self) -> str:
        """What the entries of :attr:`permissions` are in this mode."""
        return "application_role" if self.mode == "application" else "delegated_scope"


def _application_permissions(key: str) -> tuple[str, ...]:
    """The application roles of an operation, from the endpoint table it owns."""
    service = key.split(".", 1)[0]
    if service == "planner":
        return _endpoint_permissions(planner_endpoints(key))
    if service == "todo":
        return _endpoint_permissions(todo_endpoints(key))
    if key in MESSAGING_OPERATIONS:
        return _endpoint_permissions(messaging_endpoints(key))
    if service in ("sharepoint", "onedrive"):
        return _endpoint_permissions(drive_endpoints(key))
    if key in TEAMS_OPERATIONS:
        return _endpoint_permissions(teams_endpoints(key))
    return tuple(_PERMISSION_MAP[key])


def _application_mode_support(key: str) -> ModeSupport:
    """Application support of one operation, derived from the honest sources it already has.

    Derived on every call -- not baked in at import -- so a To Do write keeps following
    :func:`todo_write_support` (its status is ``not_verified`` until every endpoint of the
    write carries a verified application claim with recorded evidence) and cannot be promoted
    by editing this function's inputs.
    """
    roles = _application_permissions(key)
    if key in _UNSUPPORTED_APPLICATION:
        return ModeSupport(
            mode="application",
            status="unsupported_auth_mode",
            admin_consent_basis=NO_PERMISSION_CONSENT_BASIS,
            reason=(
                f"{key} has no application permission: this endpoint only exists for "
                "delegated authentication"
            ),
        )
    if key in TODO_WRITE_OPERATIONS:
        support = todo_write_support(key)
        if not support.supported:
            return ModeSupport(
                mode="application",
                status="not_verified",
                permissions=roles,
                permission_status="documented_not_verified",
                admin_consent=True,
                admin_consent_basis=APPLICATION_CONSENT_BASIS,
                reason=support.reason,
            )
    return ModeSupport(
        mode="application",
        status="supported",
        permissions=roles,
        permission_status="documented_not_verified",
        admin_consent=True,
        admin_consent_basis=APPLICATION_CONSENT_BASIS,
        reason=(
            "the operation is available to application authentication; the role is the one "
            "recorded for its endpoint, not verified against a tenant (R10; WP16)"
        ),
    )


def _delegated_mode_support(key: str) -> ModeSupport:
    """Delegated support of one operation: scopes recorded, mode not implemented yet.

    The status is ``not_implemented`` while ``delegated`` is absent from
    :data:`IMPLEMENTED_AUTH_MODES`, so recording scopes cannot advertise a mode the plugin
    cannot perform. Only WP14 may add the mode, and it does so in one place.
    """
    scopes = _DELEGATED_SCOPES[key]
    implemented = "delegated" in IMPLEMENTED_AUTH_MODES
    unavailable = key in _UNSUPPORTED_APPLICATION
    return ModeSupport(
        mode="delegated",
        status="supported" if implemented else "not_implemented",
        permissions=scopes,
        permission_status="documented_not_verified",
        admin_consent=True,
        admin_consent_basis=UNVERIFIED_DELEGATED_CONSENT_BASIS,
        reason=(
            "delegated authentication is not implemented (WP14); this operation is only "
            "available in delegated mode, the scopes below are the ones it will need"
            if unavailable
            else (
                "delegated authentication is not implemented (WP14); the scopes below are the "
                "ones this operation will need"
            )
        ),
    )


#: What each implementation status is allowed to mean. ``contract_verified`` requires the
#: registry to declare, for every endpoint of the operation, the ``sdk_contract`` dispatcher
#: and case that a strict offline test pins (``tests/test_planner_contracts.py`` does this for
#: Planner). ``implemented`` additionally requires a registered handler -- and it says nothing
#: about exposure: an operation is reachable from the model only while it is ``executable``,
#: which is why the four Outlook/Calendar writes are ``implemented`` and withheld.
IMPLEMENTATION_STATUS_LABELS = {
    "contract_foundation":
        "Listed with a permission claim. The registry does not declare this operation's "
        "verified endpoint, container or identifier contract yet, so nothing here asserts "
        "that its request shape is correct.",
    "contract_verified":
        "Every endpoint the operation uses is pinned by a strict offline request-contract "
        "test. No handler exists for it, so the operation is not executable.",
    "implemented":
        "A handler and a strict offline request-contract test both exist. The operation is "
        "exposed to the model only while it is executable; a write stays implemented and "
        "non-executable until the generic host approval fix (CORE-1/CORE-2) is available "
        "in a supported Hermes release (R5).",
}


@dataclass(frozen=True)
class OperationDefinition:
    """One operation of the catalogue, with its per-mode support matrix (WP13).

    ``permissions`` are the application roles, and ``app``/``delegated`` carry the support of
    each authentication mode (status, roles/scopes, permission claim, admin consent). Both
    mode records are derived by :func:`_application_mode_support` / :func:`_delegated_mode_support`
    and validated in :meth:`__post_init__`, so the registry cannot declare a mode status that
    its own permissions do not back.
    """

    service: str
    operation: str
    permissions: tuple[str, ...]
    write: bool
    app: ModeSupport | None = None
    delegated: ModeSupport | None = None
    endpoints: tuple[str, ...] = ()
    container: str = ""
    required_identifiers: tuple[str, ...] = ()
    optional_identifiers: tuple[str, ...] = ()
    required_headers: tuple[str, ...] = ()
    contract_cases: tuple[tuple[str, str], ...] = ()
    permission_claim_statuses: tuple[str, ...] = ()
    documentation_pages: tuple[str, ...] = ()
    implementation_status: str = "contract_foundation"
    remote_verification: str = "not_tested"
    remote_verification_evidence: tuple[str, ...] = ()
    admin_consent: bool | None = None
    executable: bool = False

    def __post_init__(self) -> None:
        key = f"{self.service}.{self.operation}"
        app = self.app if self.app is not None else _application_mode_support(key)
        delegated = self.delegated if self.delegated is not None else _delegated_mode_support(key)
        object.__setattr__(self, "app", app)
        object.__setattr__(self, "delegated", delegated)

        if app.mode != "application":
            raise ValueError(f"{key} must record its application support as mode 'application'")
        if delegated.mode != "delegated":
            raise ValueError(f"{key} must record its delegated support as mode 'delegated'")
        if app.permissions != self.permissions:
            raise ValueError(
                f"{key} records application roles {app.permissions!r} but declares "
                f"permissions {self.permissions!r}"
            )
        if self.remote_verification not in REMOTE_VERIFICATION_STATUSES:
            raise ValueError(f"unknown remote verification status: {self.remote_verification!r}")
        if self.remote_verification == "verified" and not self.remote_verification_evidence:
            raise ValueError(
                f"{key} cannot claim remote verification without recorded evidence (WP16)"
            )
        claimed = [mode for mode in (app, delegated) if mode.permissions]
        expected_consent = any(mode.admin_consent for mode in claimed)
        if self.admin_consent is None:
            # Derived, never asserted: an operation needs an administrator when any mode it
            # can actually use claims a permission.
            object.__setattr__(self, "admin_consent", expected_consent)
        elif self.admin_consent is not expected_consent:
            raise ValueError(
                f"{key} records admin_consent={self.admin_consent!r} but its mode support "
                f"says {expected_consent!r}"
            )

    @property
    def key(self) -> str:
        return f"{self.service}.{self.operation}"

    @property
    def endpoint(self) -> str:
        """The endpoints this operation uses, or an empty string while none is recorded.

        Reported as unrecorded rather than invented: the services whose endpoint table is not
        declared yet (WP1/WP6-WP9) have no verified endpoint to name here.
        """
        return "; ".join(self.endpoints)

    def support_for(self, auth_mode: Any) -> ModeSupport | None:
        """The mode support record of ``auth_mode``, or ``None`` for an unknown mode.

        Derived from the same two builders the registry fields are built from, so a status
        that has to follow the endpoint tables (a To Do write) keeps following them instead of
        freezing at import time.
        """
        if auth_mode == "application":
            return _application_mode_support(self.key)
        if auth_mode == "delegated":
            return _delegated_mode_support(self.key)
        return None


def _planner_definition(key: str) -> OperationDefinition:
    return _endpoint_backed_definition(
        service="planner", key=key, endpoints=planner_endpoints(key),
        implementation_status="implemented", executable=key in EXECUTABLE_OPERATIONS,
    )


def _todo_definition(key: str) -> OperationDefinition:
    return _endpoint_backed_definition(
        service="todo", key=key, endpoints=todo_endpoints(key),
        implementation_status="implemented", executable=key in EXECUTABLE_OPERATIONS,
    )


def _teams_definition(key: str) -> OperationDefinition:
    """Registry entry for Teams Graph, with application reads only executable.

    Search and send are retained in the administrative registry with their delegated scopes,
    but application mode has no permission claim for either endpoint and therefore refuses them
    before client construction. Delegated authentication remains not implemented until WP14.
    """
    return _endpoint_backed_definition(
        service="teams",
        key=key,
        endpoints=teams_endpoints(key),
        implementation_status="implemented",
        executable=key in EXECUTABLE_OPERATIONS,
    )


def _messaging_definition(key: str) -> OperationDefinition:
    """Registry entry of an Outlook or Calendar operation (WP6).

    ``implemented`` (a handler exists for every endpoint, and every one is contract-pinned) is
    the honest label, and ``executable`` is taken from the single literal set
    :data:`EXECUTABLE_OPERATIONS`: the three verified reads are exposed to the model, the four
    writes are implemented and withheld (R5).
    """
    return _endpoint_backed_definition(
        service=key.split(".", 1)[0],
        key=key,
        endpoints=messaging_endpoints(key),
        implementation_status="implemented",
        executable=key in EXECUTABLE_OPERATIONS,
    )


def _drive_definition(key: str) -> OperationDefinition:
    """Registry entry of a SharePoint or OneDrive operation (WP7).

    ``implemented`` (a handler exists for every endpoint, and every endpoint is
    contract-pinned) is the honest label, and ``executable`` is taken from the single literal
    set :data:`EXECUTABLE_OPERATIONS`: search/read/download_files are exposed to the model for
    both services, the two uploads are implemented and withheld (R5).
    """
    return _endpoint_backed_definition(
        service=key.split(".", 1)[0],
        key=key,
        endpoints=drive_endpoints(key),
        implementation_status="implemented",
        executable=key in EXECUTABLE_OPERATIONS,
    )


OPERATION_REGISTRY = {
    **{
        key: OperationDefinition(
            service=key.split(".", 1)[0],
            operation=key.split(".", 1)[1],
            permissions=permissions,
            write=key.split(".", 1)[1] in WRITE_OPERATIONS,
        )
        for key, permissions in _PERMISSION_MAP.items()
    },
    **{key: _planner_definition(key) for key in PLANNER_OPERATIONS},
    **{key: _todo_definition(key) for key in TODO_OPERATIONS},
    **{key: _messaging_definition(key) for key in MESSAGING_OPERATIONS},
    **{key: _drive_definition(key) for key in DRIVE_OPERATIONS},
    **{key: _teams_definition(key) for key in TEAMS_OPERATIONS},
}


@dataclass(frozen=True)
class OperationStatus:
    auth_status: str
    implementation_status: str
    executable: bool
    reason: str


class ConfigurationError(ValueError):
    pass


#: The strict textual shape of a mailbox object ID: a 36-character GUID of 8-4-4-4-12
#: hex digits. Matching is case-insensitive on the hex characters but the entry is
#: retained verbatim (no stripping, no case-folding), so a whitespace-padded ID is
#: rejected rather than silently repaired, and wildcard or UPN forms never match.
_MAILBOX_ID_SHAPE = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)


def operation_status(auth_mode: str, service: str, operation: str) -> OperationStatus:
    """The honest status of one operation in one authentication mode, read from the matrix.

    The verdict is the operation's own mode support record (WP13): ``unsupported_auth_mode``
    when the mode cannot reach the endpoint at all, ``not_implemented`` when the authentication
    mode itself is not implemented (delegated, WP14), ``not_verified`` when the mode's
    permission for the endpoint is unverified, ``supported`` when the mode is available. A
    supported mode still never implies execution: ``executable`` stays whatever the registry
    declares, so support and executability cannot be conflated.
    """
    key = f"{service}.{operation}"
    definition = OPERATION_REGISTRY.get(key)
    if definition is None:
        return OperationStatus("unknown_operation", "unknown", False, f"Unknown operation: {key}")
    support = definition.support_for(auth_mode)
    if support is None:
        return OperationStatus(
            "not_implemented",
            definition.implementation_status,
            False,
            "Authentication mode is not implemented",
        )
    if support.status in {"unsupported_auth_mode", "not_implemented", "not_verified"}:
        return OperationStatus(support.status, definition.implementation_status, False, support.reason)
    return OperationStatus(
        support.status,
        definition.implementation_status,
        definition.executable,
        support.reason,
    )


@dataclass(frozen=True)
class Settings:
    tenant_id: str = ""
    client_id: str = ""
    user_id: str = ""
    authentication_mode: str = "application"
    capabilities: Mapping[str, Mapping[str, bool]] = field(default_factory=dict)
    allowed_mailbox_ids: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        """Reject a malformed mailbox policy without disclosing configured identifiers.

        Every entry must be an explicit mailbox object ID in GUID textual form
        (8-4-4-4-12 hex digits). Wildcard forms, UPN/other identifier forms, and
        whitespace-padded or otherwise malformed tokens are refused instead of
        silently repaired; a valid entry is retained verbatim, so the boundary
        comparison downstream stays exact. Duplicates are detected case-insensitively
        because a mailbox object ID is a UUID: the same ID spelled with different
        hex case is still one mailbox.
        """
        if type(self.allowed_mailbox_ids) is not list:
            raise ConfigurationError("allowed_mailbox_ids must be a list of explicit mailbox object IDs")
        seen = set()
        for value in self.allowed_mailbox_ids:
            if type(value) is not str or not _MAILBOX_ID_SHAPE.fullmatch(value):
                raise ConfigurationError(
                    "allowed_mailbox_ids must contain only explicit mailbox object IDs in "
                    "GUID textual form (8-4-4-4-12 hex digits); wildcard and other "
                    "identifier forms are not accepted"
                )
            key = value.lower()
            if key in seen:
                raise ConfigurationError("allowed_mailbox_ids must not contain duplicate mailbox IDs")
            seen.add(key)

    @classmethod
    def from_mapping(cls, raw: Mapping[str, Any] | None) -> "Settings":
        raw = raw or {}
        capabilities = raw.get("capabilities", {})
        if not isinstance(capabilities, Mapping):
            raise ConfigurationError(f"capabilities must be a mapping (got {type(capabilities).__name__})")
        unknown_services = sorted(set(capabilities) - set(OPERATIONS))
        if unknown_services:
            raise ConfigurationError(f"capabilities.{unknown_services[0]} is an unknown capability")
        normalized: dict[str, dict[str, bool]] = {}
        for service, service_value in capabilities.items():
            if service_value is True:
                normalized[service] = {operation: True for operation in OPERATIONS[service]}
                continue
            if service_value is False:
                normalized[service] = {operation: False for operation in OPERATIONS[service]}
                continue
            if not isinstance(service_value, Mapping):
                raise ConfigurationError(
                    f"capabilities.{service} must be a boolean or operation mapping "
                    f"(got {type(service_value).__name__})"
                )
            unknown_operations = sorted(set(service_value) - set(OPERATIONS[service]))
            if unknown_operations:
                raise ConfigurationError(
                    f"capabilities.{service}.{unknown_operations[0]} is an unknown operation"
                )
            normalized[service] = {}
            for operation, enabled in service_value.items():
                if type(enabled) is not bool:
                    raise ConfigurationError(
                        f"capabilities.{service}.{operation} must be a boolean "
                        f"(got {type(enabled).__name__})"
                    )
                normalized[service][operation] = enabled
        mode = raw.get("authentication_mode", "application")
        if mode not in {"application", "delegated"}:
            raise ConfigurationError("authentication_mode must be 'application' or 'delegated'")
        return cls(
            tenant_id=str(raw.get("tenant_id") or ""),
            client_id=str(raw.get("client_id") or ""),
            user_id=str(raw.get("user_id") or ""),
            authentication_mode=str(mode),
            capabilities=normalized,
            allowed_mailbox_ids=raw.get("allowed_mailbox_ids", []),
        )

    def selected(self, service: str) -> set[str]:
        return {name for name, enabled in self.capabilities.get(service, {}).items() if enabled}
