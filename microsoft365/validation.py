"""Operation-specific argument validation, run before approval *and* before dispatch (WP13).

One validator, two call sites
----------------------------
:func:`check` is the only validator in this plugin. The ``pre_tool_call`` hook
(``microsoft365.registration``) calls it before it can return an approval directive, and the
tool dispatch path calls it again immediately before a handler is looked up, so the same
payload gets the same answer before approval and before any handler, secret, credential or
Graph client is touched. It is deliberately the same function with the same arguments — not a
second, weaker check — and it never mutates or normalizes the payload: the arguments that were
approved are the arguments that are executed (R5).

What it rejects
---------------
* a call whose arguments are not an object; an unknown operation; an operation disabled by
  configuration; an authentication mode the operation does not support (or that is not
  implemented at all);
* any property the operation does not declare, any missing required property, and ``None``
  where a value is required;
* coercion: ``True``/``1``/``"false"`` are not booleans, a number or a list is not an
  identifier, a string is not an integer;
* empty, unbounded, unstripped or control-character-carrying strings; collections beyond their
  bounded size; recipients that are not addresses; paths that are absolute, traversing,
  backslashed or addressed; timestamps that are not ISO-8601 date-times (a naive one is
  accepted only when the operation also carries ``time_zone``); time zones that are not a
  plausible zone name; ETags that are blank, wildcards or header-injection attempts;
* base64 that is not padded base64, is larger than the encoded ceiling, or would decode larger
  than the decoded ceiling — both size gates run **before** any decode of the payload;
* a PATCH that is empty, carries a field outside the operation's allowlist, or carries ``None``
  or a nested object as a value;
* mutually exclusive argument forms (exactly one, or at most one, of the declared forms);
* ``/me`` as a user id in application mode, and a user-scoped operation with no user id at all.

What it may never leak
----------------------
A rejection message is assembled from static templates plus a **bounded, sanitized name** of
the offending argument. No value from the payload is ever interpolated: a rejected value cannot
appear in the message, in the tool result or in a log line built from it. Values are always
described by their kind and their bound, never by their content.

No credential surface
---------------------
This module never reads a secret, never builds a credential and never builds a Graph client —
it imports neither the host's scoped secret lookup, the credential library, nor the Graph client
library (pinned by test), so the guard sits strictly in front of that whole surface. It imports
only the operation registry, the argument-free settings record and the error taxonomy.
"""
from __future__ import annotations

import base64
import binascii
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Callable, Mapping

from .contract import OPERATIONS, OPERATION_REGISTRY, TODO_TASK_WRITE_FIELDS, Settings
from .errors import CATEGORIES, GraphError

#: Ceiling for an opaque Graph identifier (``user_id``, ``message_id``, ``drive_id`` ...).
MAX_IDENTIFIER_LENGTH = 256
#: Ceiling for a long free-text value (a mail or message body).
MAX_TEXT_LENGTH = 100_000
#: Ceiling for a short free-text value (a subject, a query, a filter, a title, a location).
MAX_SHORT_TEXT_LENGTH = 512
#: Ceiling for a drive-relative item path.
MAX_PATH_LENGTH = 400
#: Ceiling for the number of entries of a bounded collection.
MAX_COLLECTION_ITEMS = 100
#: Ceiling for the number of fields of a ``$select`` list.
MAX_SELECT_FIELDS = 20
#: Ceiling for a time zone name.
MAX_TIME_ZONE_LENGTH = 64
#: Ceiling for an ETag value, header-injection shaped input included.
MAX_ETAG_LENGTH = 200
#: Ceiling for a media type.
MAX_MEDIA_TYPE_LENGTH = 128
#: Ceiling for the decoded size of an encoded upload payload (10 MiB; the upload-session path
#: above it is not implemented by this plugin and is declared as such).
MAX_UPLOAD_DECODED_BYTES = 10 * 1024 * 1024
#: Ceiling for the *encoded* representation, computed from the decoded ceiling so the encoded
#: gate is the exact size of a maximum payload (a test pins the arithmetic).
MAX_UPLOAD_ENCODED_CHARS = 4 * ((MAX_UPLOAD_DECODED_BYTES + 2) // 3)

#: Every argument kind this validator knows.
KINDS = frozenset(
    {
        "identifier",
        "text",
        "boolean",
        "integer",
        "timestamp",
        "time_zone",
        "etag",
        "base64",
        "recipients",
        "path",
        "select",
        "enum",
        "media_type",
        "mapping",
    }
)

#: The bounds of every integer PATCH field, keyed by field name. A mapping field of kind
#: ``integer`` without an entry here is refused while the contract table is built, so an
#: unbounded integer can never be accepted.
PATCH_INTEGER_BOUNDS: Mapping[str, tuple[int, int]] = {
    "percent_complete": (0, 100),
    "priority": (0, 10),
}

#: The two exclusion forms: exactly one of the forms, or at most one of them.
EXCLUSION_MODES = frozenset({"exactly_one", "at_most_one"})

_TEXT_KIND_DEFAULTS: Mapping[str, Mapping[str, Any]] = {
    "identifier": {"max_length": MAX_IDENTIFIER_LENGTH},
    "text": {"max_length": MAX_SHORT_TEXT_LENGTH},
    "path": {"max_length": MAX_PATH_LENGTH},
    "etag": {"max_length": MAX_ETAG_LENGTH},
    "time_zone": {"max_length": MAX_TIME_ZONE_LENGTH},
    "base64": {"max_length": MAX_UPLOAD_ENCODED_CHARS},
    "select": {"max_items": MAX_SELECT_FIELDS},
    "recipients": {"max_items": MAX_COLLECTION_ITEMS},
}

_LABEL_PATTERN = re.compile(r"^[A-Za-z0-9_.\-]{1,48}$")
#: Control characters a free-text value may carry (none of the fatal ones, plus no DEL).
_CONTROL_PATTERN = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
#: Control characters no identifier, path or name may carry at all -- line breaks included,
#: because an identifier ends up in a URL path and a name ends up in a request header.
_ANY_CONTROL_PATTERN = re.compile(r"[\x00-\x1f\x7f]")
_EMAIL_PATTERN = re.compile(
    r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.\-]{1,64}@[A-Za-z0-9](?:[A-Za-z0-9\-.]{0,251}[A-Za-z0-9])?$"
)
_MEDIA_TYPE_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9!#$&^_.+\-]{0,63}/[A-Za-z0-9][A-Za-z0-9!#$&^_.+\-]{0,63}$")
_TIME_ZONE_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9 ._+\-/]{0,63}$")
_ETAG_PATTERN = re.compile(r'^(?:W/)?"[\x21\x23-\x7e]{1,200}"$|^[A-Za-z0-9._~\-]{1,200}$')
_BASE64_PATTERN = re.compile(r"^[A-Za-z0-9+/]*={0,2}$")
_SELECT_PATTERN = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]{0,63}$")
_TIMESTAMP_MAX_LENGTH = 64

#: The identifiers that name a user, and therefore the ones ``/me`` may not stand in for when
#: the caller authenticates as an application.
USER_IDENTIFIERS = ("user_id",)

#: Operations whose Graph endpoint indexes by ``/users/{user_id}``. Derived from the endpoint
#: tables where one is declared (Planner has none, To Do declares ``user_id`` on every
#: operation) and declared here for the services whose endpoint table WP1/WP6-WP9 will add.
#: ``validate`` requires ``user_id`` for exactly these operations; a test pins the two
#: sources against each other.
USER_SCOPED_OPERATIONS = frozenset(
    {
        "outlook.search",
        "outlook.list_attachments",
        "outlook.download_attachment",
        "outlook.read",
        "outlook.create_draft",
        "outlook.send",
        "calendar.search",
        "calendar.create_events",
        "calendar.update_events",
        "teams.list_teams",
        "todo.list_task_lists",
        "todo.search",
        "todo.read",
        "todo.create_tasks",
        "todo.update_tasks",
    }
)

#: The kinds of the To Do task fields a write may carry. The *allowlist* comes from
#: ``contract.TODO_TASK_WRITE_FIELDS`` (WP3 owns it); this only says how to check a value.
TODO_FIELD_KINDS: Mapping[str, str] = {
    "title": "text",
    "body": "text",
    "due_date_time": "timestamp",
    "status": "text",
}


class OperationArgumentsError(ValueError):
    """An argument contract that is itself malformed. Raised while the table is built."""


@dataclass(frozen=True)
class MutualExclusion:
    """Argument forms of which only one (or at most one) may be supplied."""

    forms: tuple[tuple[str, ...], ...]
    mode: str

    def __post_init__(self) -> None:
        if self.mode not in EXCLUSION_MODES:
            raise OperationArgumentsError(f"unknown exclusion mode: {self.mode!r}")
        if not self.forms or any(not form for form in self.forms):
            raise OperationArgumentsError("an exclusion needs at least one non-empty form")

    @property
    def description(self) -> str:
        return " OR ".join("+".join(form) for form in self.forms)


@dataclass(frozen=True)
class ArgumentSpec:
    """One argument an operation accepts, with the bounds it must respect."""

    name: str
    kind: str
    required: bool = False
    max_length: int = MAX_SHORT_TEXT_LENGTH
    max_items: int = MAX_COLLECTION_ITEMS
    minimum: int | None = None
    maximum: int | None = None
    values: tuple[str, ...] = ()
    fields: tuple[tuple[str, str], ...] = ()
    require_fields: tuple[str, ...] = ()

    @property
    def field_kinds(self) -> dict[str, str]:
        return dict(self.fields)


@dataclass(frozen=True)
class OperationArguments:
    """The complete argument contract of one operation (``"<service>.<operation>"``).

    Validated on construction, so a contract that would silently accept anything (an integer
    without bounds, an enum without values, a PATCH without an allowlist, a duplicated or
    structural property name, an exclusion naming a property that does not exist) cannot be
    added to the table.
    """

    key: str
    properties: tuple[ArgumentSpec, ...]
    exclusions: tuple[MutualExclusion, ...] = ()
    implementation_note: str = ""

    def __post_init__(self) -> None:
        if not self.key or "." not in self.key:
            raise OperationArgumentsError(f"{self.key!r} is not a <service>.<operation> key")
        seen: set[str] = set()
        for spec in self.properties:
            if spec.kind not in KINDS:
                raise OperationArgumentsError(
                    f"{self.key}.{spec.name} declares unknown argument kind {spec.kind!r}"
                )
            if spec.name in {"action", "operation"}:
                raise OperationArgumentsError(
                    f"{self.key}.{spec.name} is a structural argument and cannot be a property"
                )
            if spec.name in seen:
                raise OperationArgumentsError(
                    f"{self.key} declares the property {spec.name!r} more than once: property "
                    "names must be unique"
                )
            seen.add(spec.name)
            if spec.kind == "integer" and (spec.minimum is None or spec.maximum is None):
                raise OperationArgumentsError(
                    f"{self.key}.{spec.name} is an integer and must declare its bounds"
                )
            if spec.minimum is not None and spec.maximum is not None and spec.minimum > spec.maximum:
                raise OperationArgumentsError(f"{self.key}.{spec.name} has inverted bounds")
            if spec.kind == "enum" and not spec.values:
                raise OperationArgumentsError(
                    f"{self.key}.{spec.name} is an enum and must declare its allowed values"
                )
            if spec.kind == "mapping" and not spec.fields:
                raise OperationArgumentsError(
                    f"{self.key}.{spec.name} is a mapping and must declare its field allowlist"
                )
            if spec.kind == "mapping":
                names = [name for name, _ in spec.fields]
                if len(names) != len(set(names)):
                    raise OperationArgumentsError(
                        f"{self.key}.{spec.name} declares a field twice"
                    )
                for name, kind in spec.fields:
                    if kind not in KINDS or kind == "mapping":
                        raise OperationArgumentsError(
                            f"{self.key}.{spec.name}.{name} declares unknown field kind {kind!r}"
                        )
                    if kind == "integer" and name not in PATCH_INTEGER_BOUNDS:
                        raise OperationArgumentsError(
                            f"{self.key}.{spec.name}.{name} is an integer and must declare its "
                            "bounds in PATCH_INTEGER_BOUNDS"
                        )
                missing = [name for name in spec.require_fields if name not in set(names)]
                if missing:
                    raise OperationArgumentsError(
                        f"{self.key}.{spec.name} requires undeclared field {missing[0]!r}"
                    )
            if spec.kind in {"select", "recipients"} and spec.max_items < 1:
                raise OperationArgumentsError(
                    f"{self.key}.{spec.name} is a collection and needs a positive bound"
                )
        mappings = [spec for spec in self.properties if spec.kind == "mapping"]
        if len(mappings) > 1:
            raise OperationArgumentsError(f"{self.key} declares more than one mapping property")
        for exclusion in self.exclusions:
            for form in exclusion.forms:
                for name in form:
                    if name not in seen:
                        raise OperationArgumentsError(
                            f"{self.key} excludes undeclared property {name!r}"
                        )

    def spec(self, name: str) -> ArgumentSpec | None:
        return next((spec for spec in self.properties if spec.name == name), None)

    @property
    def property_names(self) -> tuple[str, ...]:
        return tuple(spec.name for spec in self.properties)

    @property
    def patch_allowlist(self) -> tuple[str, ...]:
        """The fields a PATCH of this operation may carry (empty when it has none)."""
        mapping = next((spec for spec in self.properties if spec.kind == "mapping"), None)
        return tuple(name for name, _ in mapping.fields) if mapping is not None else ()


def _spec(name: str, kind: str, *, required: bool = False, **overrides: Any) -> ArgumentSpec:
    values: dict[str, Any] = dict(_TEXT_KIND_DEFAULTS.get(kind, {}))
    values.update(overrides)
    return ArgumentSpec(name=name, kind=kind, required=required, **values)


def _integer(name: str, minimum: int, maximum: int, *, required: bool = False) -> ArgumentSpec:
    return _spec(name, "integer", required=required, minimum=minimum, maximum=maximum)


def _fields(allowlist: Mapping[str, str], *, required: bool = False, require: tuple[str, ...] = ()) -> ArgumentSpec:
    return _spec("fields", "mapping", required=required, fields=tuple(allowlist.items()), require_fields=require)


def _exclusive(*forms: tuple[str, ...]) -> MutualExclusion:
    return MutualExclusion(forms=tuple(forms), mode="exactly_one")


def _at_most_one(*forms: tuple[str, ...]) -> MutualExclusion:
    return MutualExclusion(forms=tuple(forms), mode="at_most_one")


#: The PATCH allowlist of ``calendar.update_events``. Everything else an ``event`` model
#: declares is either service-maintained or outside what this plugin writes.
EVENT_PATCH_FIELDS: Mapping[str, str] = {
    "subject": "text",
    "body": "text",
    "start_date_time": "timestamp",
    "end_date_time": "timestamp",
    "time_zone": "time_zone",
    "location": "text",
    "attendees": "recipients",
    "is_all_day": "boolean",
}

#: The PATCH allowlist of ``planner.update_tasks``: the fields of a ``PlannerTask`` this
#: plugin is willing to send, and how each value is checked.
PLANNER_PATCH_FIELDS: Mapping[str, str] = {
    "title": "text",
    "bucket_id": "identifier",
    "due_date_time": "timestamp",
    "start_date_time": "timestamp",
    "percent_complete": "integer",
    "priority": "integer",
}

#: The optional mailbox-kind override of every Outlook read.
#:
#: ``user`` is the ``/users/{id}`` route a personal **or shared** mailbox answers on; ``group``
#: is the ``/groups/{id}`` conversation route a Microsoft 365 group mailbox answers on. Left
#: out, the kind is resolved once from the directory (``microsoft365.mailbox``). It is an
#: explicit escape hatch, never a widening: ``user_id`` still has to pass the mailbox
#: allowlist, and naming a kind cannot reach a mailbox the allowlist does not contain.
_MAILBOX_TYPE = _spec("mailbox_type", "enum", values=("user", "group"))

_RECIPIENTS = _spec("to_recipients", "recipients", required=True)
_ITEM_FORMS = (("item_id",), ("item_path",))


def _file_item_arguments() -> tuple[ArgumentSpec, ...]:
    return (
        _spec("item_id", "identifier"),
        _spec("item_path", "path"),
    )


def _build_contracts() -> dict[str, OperationArguments]:
    """Every operation's argument contract, built once at import.

    The To Do write allowlist is read from :data:`microsoft365.contract.TODO_TASK_WRITE_FIELDS`
    rather than re-declared, so the field policy has exactly one source; a test pins the two
    against each other.
    """
    todo_allowlist = {
        name: TODO_FIELD_KINDS[name] for name in TODO_TASK_WRITE_FIELDS if name in TODO_FIELD_KINDS
    }
    if set(todo_allowlist) != set(TODO_TASK_WRITE_FIELDS):
        raise OperationArgumentsError(
            "the To Do write field kinds and contract.TODO_TASK_WRITE_FIELDS disagree"
        )

    contracts = {
        "outlook.search": OperationArguments(
            key="outlook.search",
            properties=(
                _spec("user_id", "identifier", required=True),
                _MAILBOX_TYPE,
                _spec("query", "text"),
                _spec("filter", "text"),
                _spec("has_attachments", "boolean"),
                _spec("sender", "text", max_length=318),
                _spec("received_after", "timestamp"),
                _spec("received_before", "timestamp"),
                _spec("select", "select"),
                _integer("top", 1, 100),
            ),
        ),
        "outlook.list_attachments": OperationArguments(
            key="outlook.list_attachments",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("message_id", "identifier", required=True),
                _MAILBOX_TYPE,
                _integer("top", 1, 100),
            ),
        ),
        "outlook.download_attachment": OperationArguments(
            key="outlook.download_attachment",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("message_id", "identifier", required=True),
                _spec("attachment_id", "identifier", required=True),
                _MAILBOX_TYPE,
                _spec("save_to_disk", "boolean"),
            ),
        ),
        "outlook.read": OperationArguments(
            key="outlook.read",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("message_id", "identifier", required=True),
                _MAILBOX_TYPE,
                _spec("select", "select"),
            ),
        ),
        "outlook.create_draft": OperationArguments(
            key="outlook.create_draft",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("subject", "text", required=True),
                _spec("body", "text", max_length=MAX_TEXT_LENGTH),
                _RECIPIENTS,
                _spec("cc_recipients", "recipients"),
                _spec("bcc_recipients", "recipients"),
            ),
        ),
        "outlook.send": OperationArguments(
            key="outlook.send",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("message_id", "identifier"),
                _spec("subject", "text"),
                _spec("body", "text", max_length=MAX_TEXT_LENGTH),
                _spec("to_recipients", "recipients"),
                _spec("cc_recipients", "recipients"),
                _spec("bcc_recipients", "recipients"),
                _spec("save_to_sent_items", "boolean"),
            ),
            # Sending an existing draft and composing a new message are two different calls;
            # mixing them would send one of them silently.
            exclusions=(_exclusive(("message_id",), ("subject", "to_recipients")),),
        ),
        "calendar.search": OperationArguments(
            key="calendar.search",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("filter", "text"),
                _integer("top", 1, 100),
                _spec("start_date_time", "timestamp"),
                _spec("end_date_time", "timestamp"),
                _spec("time_zone", "time_zone"),
            ),
        ),
        "calendar.create_events": OperationArguments(
            key="calendar.create_events",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("subject", "text", required=True),
                _spec("start_date_time", "timestamp", required=True),
                _spec("end_date_time", "timestamp", required=True),
                _spec("time_zone", "time_zone"),
                _spec("attendees", "recipients"),
                _spec("body", "text", max_length=MAX_TEXT_LENGTH),
                _spec("location", "text"),
                _spec("is_all_day", "boolean"),
            ),
        ),
        "calendar.update_events": OperationArguments(
            key="calendar.update_events",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("event_id", "identifier", required=True),
                # A calendar update is conditional: without the ETag the service cannot tell
                # whether the item changed since it was read (WP1/WP2 semantics).
                _spec("etag", "etag", required=True),
                _fields(EVENT_PATCH_FIELDS, required=True),
            ),
        ),
        "teams.list_teams": OperationArguments(
            key="teams.list_teams",
            properties=(_spec("user_id", "identifier", required=True),),
        ),
        "teams.list_channels": OperationArguments(
            key="teams.list_channels",
            properties=(
                _spec("team_id", "identifier", required=True),
                # ``$top`` is deliberately absent: Graph does not support it on this endpoint
                # (WP1 pins that), so ``top`` is an unknown property here.
                _spec("select", "select"),
            ),
        ),
        "teams.search_messages": OperationArguments(
            key="teams.search_messages",
            properties=(
                _spec("query", "text", required=True),
                _integer("from_", 0, 1000),
                _integer("size", 1, 100),
            ),
        ),
        "teams.send_messages": OperationArguments(
            key="teams.send_messages",
            properties=(
                _spec("team_id", "identifier", required=True),
                _spec("channel_id", "identifier", required=True),
                _spec("body", "text", max_length=MAX_TEXT_LENGTH, required=True),
                _spec("content_type", "enum", values=("text", "html")),
            ),
        ),
        "todo.list_task_lists": OperationArguments(
            key="todo.list_task_lists",
            properties=(
                _spec("user_id", "identifier", required=True),
                _integer("top", 1, 100),
            ),
        ),
        "todo.search": OperationArguments(
            key="todo.search",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("todo_list_id", "identifier", required=True),
                # The To Do task field is ``title``; there is no ``subject`` (WP3).
                _spec("title", "text"),
                _integer("top", 1, 100),
                _spec("select", "select"),
            ),
        ),
        "todo.read": OperationArguments(
            key="todo.read",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("todo_list_id", "identifier", required=True),
                _spec("todo_task_id", "identifier", required=True),
                _spec("select", "select"),
            ),
        ),
        "todo.create_tasks": OperationArguments(
            key="todo.create_tasks",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("todo_list_id", "identifier", required=True),
                _fields(todo_allowlist, required=True, require=("title",)),
            ),
        ),
        "todo.update_tasks": OperationArguments(
            key="todo.update_tasks",
            properties=(
                _spec("user_id", "identifier", required=True),
                _spec("todo_list_id", "identifier", required=True),
                _spec("todo_task_id", "identifier", required=True),
                _fields(todo_allowlist, required=True),
            ),
        ),
        "planner.list_plans": OperationArguments(
            key="planner.list_plans",
            properties=(_spec("group_id", "identifier"),),
        ),
        "planner.list_buckets": OperationArguments(
            key="planner.list_buckets",
            properties=(_spec("plan_id", "identifier", required=True),),
        ),
        "planner.list_tasks": OperationArguments(
            key="planner.list_tasks",
            properties=(
                _spec("plan_id", "identifier", required=True),
                _integer("top", 1, 100),
            ),
        ),
        "planner.read": OperationArguments(
            key="planner.read",
            properties=(_spec("planner_task_id", "identifier", required=True),),
        ),
        "planner.create_tasks": OperationArguments(
            key="planner.create_tasks",
            properties=(
                _spec("plan_id", "identifier", required=True),
                _spec("bucket_id", "identifier", required=True),
                _spec("title", "text", max_length=MAX_IDENTIFIER_LENGTH, required=True),
                _spec("due_date_time", "timestamp"),
                _spec("start_date_time", "timestamp"),
                _integer("percent_complete", 0, 100),
                _integer("priority", 0, 10),
            ),
        ),
        "planner.update_tasks": OperationArguments(
            key="planner.update_tasks",
            properties=(
                _spec("planner_task_id", "identifier", required=True),
                _spec("etag", "etag", required=True),
                _fields(PLANNER_PATCH_FIELDS, required=True),
            ),
        ),
    }

    for service in ("sharepoint", "onedrive"):
        download = f"{service}.download_files"
        contracts[download] = OperationArguments(
            key=download,
            properties=(_spec("drive_id", "identifier", required=True), *_file_item_arguments()),
            exclusions=(_exclusive(*_ITEM_FORMS),),
        )
        upload = f"{service}.upload_files"
        contracts[upload] = OperationArguments(
            key=upload,
            properties=(
                _spec("drive_id", "identifier", required=True),
                *_file_item_arguments(),
                _spec("content_base64", "base64", required=True),
                _spec("content_type", "media_type", max_length=MAX_MEDIA_TYPE_LENGTH),
                _spec("overwrite", "boolean"),
            ),
            exclusions=(_exclusive(*_ITEM_FORMS),),
        )
        read = f"{service}.read"
        contracts[read] = OperationArguments(
            key=read,
            properties=(
                _spec("drive_id", "identifier", required=True),
                *_file_item_arguments(),
                _spec("select", "select"),
            ),
            exclusions=(_exclusive(*_ITEM_FORMS),),
        )
        search = f"{service}.search"
        contracts[search] = OperationArguments(
            key=search,
            properties=(
                _spec("query", "text", required=True),
                _spec("drive_id", "identifier"),
                _spec("site_id", "identifier"),
                _integer("top", 1, 100),
            ),
            # A search is scoped by one container at most: both would mean two different
            # searches and only one of them would be sent.
            exclusions=(_at_most_one(("drive_id",), ("site_id",)),),
        )

    return contracts


#: Every operation's argument contract. ``tests/test_validation.py`` pins that it covers the
#: registry exactly, and that each contract rejects a malformed specification.
ARGUMENT_CONTRACTS: Mapping[str, OperationArguments] = _build_contracts()


@dataclass(frozen=True)
class Rejection:
    """Why a call must not proceed, in a form both call sites can report identically.

    ``message`` is a sanitized reason (static template plus a bounded argument name). The
    ``pre_tool_call`` hook returns it as the block message and the dispatch path returns it in
    the tool payload, so the same payload yields the same answer wherever it is checked.
    """

    category: str
    message: str

    def __post_init__(self) -> None:
        if self.category not in CATEGORIES:
            raise ValueError(f"unknown microsoft 365 error category: {self.category!r}")

    def to_payload(self) -> dict[str, Any]:
        """The tool payload for this rejection."""
        return {"error": self.category, "message": self.message, "retryable": False}

    def to_error(self) -> GraphError:
        """The same rejection as a raisable taxonomy error (for handlers)."""
        return GraphError(self.category, self.message)


def _label(value: Any) -> str:
    """A bounded, sanitized argument name: never the value, never an unbounded string."""
    if isinstance(value, str) and _LABEL_PATTERN.match(value):
        return value
    return "<unnamed>"


def _named(value: Any) -> str:
    """A bounded, printable argument name for a message: unusable names are truncated."""
    if isinstance(value, str):
        cleaned = "".join(character if character.isprintable() else "?" for character in value)
        return cleaned[:48]
    return type(value).__name__


class _Failure(Exception):
    """Internal control flow: one argument rule failed, with a sanitized detail."""

    def __init__(self, detail: str):
        super().__init__(detail)
        self.detail = detail


@dataclass(frozen=True)
class _Context:
    key: str
    mode: str
    arguments: Mapping[str, Any]


def _reject(detail: str) -> None:
    raise _Failure(detail)


def _mismatch(value: Any, expected: str) -> str:
    return f"must be {expected} (got {type(value).__name__})"


def _validate_identifier(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'a string identifier')}")
    if value != value.strip():
        _reject(f"{name} must not carry surrounding whitespace")
    if not value:
        _reject(f"{name} must not be blank")
    if len(value) > spec.max_length:
        _reject(f"{name} is longer than the {spec.max_length} characters this call accepts")
    if _ANY_CONTROL_PATTERN.search(value):
        _reject(f"{name} must not carry control characters")
    if name in USER_IDENTIFIERS and context.mode == "application" and value.lower() == "/me":
        _reject(
            f"{name} must be a real user id in application mode, not /me: application "
            "authentication has no signed-in user to resolve it to"
        )


def _validate_text(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'a string')}")
    if not value:
        _reject(f"{name} must not be blank")
    if len(value) > spec.max_length:
        _reject(f"{name} is longer than the {spec.max_length} characters this call accepts")
    if _CONTROL_PATTERN.search(value):
        _reject(f"{name} must not carry control characters")


def _validate_boolean(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context, spec
    if type(value) is not bool:
        _reject(
            f"{name} {_mismatch(value, 'a real boolean')}: this call does not coerce strings "
            "or numbers into a boolean"
        )


def _validate_integer(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if type(value) is not int:
        _reject(f"{name} {_mismatch(value, 'an integer')}: this call does not coerce it")
    if not spec.minimum <= value <= spec.maximum:
        _reject(f"{name} must be between {spec.minimum} and {spec.maximum}")


def _validate_timestamp(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del spec
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'an ISO-8601 date and time string')}")
    if len(value) > _TIMESTAMP_MAX_LENGTH:
        _reject(f"{name} is longer than the {_TIMESTAMP_MAX_LENGTH} characters a date and time accepts")
    if "T" not in value:
        _reject(f"{name} must be a date and time, not a date")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        _reject(f"{name} must be a valid ISO-8601 date and time")
    if parsed.tzinfo is None and "time_zone" not in context.arguments:
        _reject(
            f"{name} must carry a UTC offset or be accompanied by time_zone: an offsetless "
            "date and time would be applied in the wrong zone"
        )


def _validate_time_zone(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'a time zone name')}")
    if value != value.strip():
        _reject(f"{name} must not carry surrounding whitespace")
    if not value:
        _reject(f"{name} must not be blank")
    if len(value) > spec.max_length:
        _reject(f"{name} is longer than the {spec.max_length} characters a time zone name accepts")
    if not _TIME_ZONE_PATTERN.match(value):
        _reject(
            f"{name} must be a time zone name such as America/Sao_Paulo or "
            "'Pacific Standard Time'"
        )


def _validate_etag(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'a string etag')}")
    if not value:
        _reject(f"{name} must not be blank")
    if len(value) > spec.max_length:
        _reject(f"{name} is longer than the {spec.max_length} characters an etag accepts")
    if not _ETAG_PATTERN.match(value):
        _reject(
            f"{name} must be a quoted or token etag: a wildcard or a value carrying control "
            "characters cannot be used as a precondition"
        )


def _validate_base64(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'a base64 encoded string')}")
    if len(value) > spec.max_length:
        _reject(
            f"{name} is larger than the {spec.max_length} encoded characters this call accepts"
        )
    if len(value) % 4 or not _BASE64_PATTERN.match(value):
        _reject(f"{name} must be padded base64 with the standard alphabet")
    padding = len(value) - len(value.rstrip("="))
    decoded_size = (len(value) // 4) * 3 - padding
    if decoded_size > MAX_UPLOAD_DECODED_BYTES:
        _reject(
            f"{name} decodes to more than the {MAX_UPLOAD_DECODED_BYTES} bytes this call "
            "accepts (a larger transfer needs an upload session, which is not implemented)"
        )
    try:
        decoded = base64.b64decode(value, validate=True)
    except (binascii.Error, ValueError):
        _reject(f"{name} must be valid base64")
    if len(decoded) != decoded_size:
        _reject(f"{name} is not a well formed base64 payload")


def _validate_recipient(context: _Context, name: str, index: int, value: Any) -> None:
    del context
    label = f"{name}[{index}]"
    if isinstance(value, str):
        address = value
        if not _EMAIL_PATTERN.match(address):
            _reject(f"{label} must be an email address")
        return
    if not isinstance(value, Mapping):
        _reject(f"{label} {_mismatch(value, 'an email address or an address/name object')}")
    unknown = sorted(str(key) for key in value if key not in {"address", "name"})
    if unknown:
        _reject(
            f"{label} carries the unknown field {_named(unknown[0])}: a recipient accepts "
            "address and name"
        )
    address = value.get("address")
    if not isinstance(address, str) or not address:
        _reject(f"{label}.address must be an email address")
    if not _EMAIL_PATTERN.match(address):
        _reject(f"{label}.address must be an email address")
    if "name" in value:
        display = value["name"]
        if not isinstance(display, str):
            _reject(f"{label}.name {_mismatch(display, 'a string')}")
        if display != display.strip() or not display:
            _reject(f"{label}.name must be a non-blank, trimmed display name")
        if len(display) > MAX_IDENTIFIER_LENGTH:
            _reject(f"{label}.name is longer than {MAX_IDENTIFIER_LENGTH} characters")
        if _CONTROL_PATTERN.search(display):
            _reject(f"{label}.name must not carry control characters")


def _validate_recipients(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    if type(value) is not list:
        _reject(f"{name} {_mismatch(value, 'a list of recipients')}")
    if not value:
        _reject(f"{name} must not be empty")
    if len(value) > spec.max_items:
        _reject(f"{name} accepts at most {spec.max_items} recipients")
    for index, recipient in enumerate(value):
        _validate_recipient(context, name, index, recipient)


def _validate_select(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if type(value) is not list:
        _reject(f"{name} {_mismatch(value, 'a list of field names')}")
    if not value:
        _reject(f"{name} must not be empty")
    if len(value) > spec.max_items:
        _reject(f"{name} accepts at most {spec.max_items} field names")
    for index, entry in enumerate(value):
        if not isinstance(entry, str) or not _SELECT_PATTERN.match(entry):
            _reject(f"{name}[{index}] must be a Graph field name")


def _validate_path(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if not isinstance(value, str):
        _reject(f"{name} {_mismatch(value, 'a relative path string')}")
    if value != value.strip():
        _reject(f"{name} must not carry surrounding whitespace")
    if not value:
        _reject(f"{name} must not be blank")
    if len(value) > spec.max_length:
        _reject(f"{name} is longer than the {spec.max_length} characters a path accepts")
    if _CONTROL_PATTERN.search(value):
        _reject(f"{name} must not carry control characters")
    if "\\" in value:
        _reject(f"{name} must use forward slashes")
    if value.startswith("/"):
        _reject(f"{name} must be relative to the drive, not absolute")
    if ":" in value:
        _reject(f"{name} must not carry an addressing prefix")
    if any(segment in {"", ".", ".."} for segment in value.split("/")):
        _reject(f"{name} must not contain empty, current or parent path segments")


def _validate_enum(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    del context
    if not isinstance(value, str) or value not in spec.values:
        _reject(f"{name} must be one of: {', '.join(spec.values)}")


def _validate_media_type(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    if not isinstance(value, str) or not _MEDIA_TYPE_PATTERN.match(value):
        _reject(f"{name} must be a media type such as text/plain")
    if len(value) > spec.max_length:
        _reject(f"{name} is longer than the {spec.max_length} characters a media type accepts")


def _validate_mapping(context: _Context, name: str, value: Any, spec: ArgumentSpec) -> None:
    if not isinstance(value, Mapping):
        _reject(f"{name} {_mismatch(value, 'an object of fields to change')}")
    if not value:
        _reject(f"{name} must not be empty: an empty update changes nothing")
    allowed = [field_name for field_name, _ in spec.fields]
    accepted = set(allowed)
    for key in value:
        if not isinstance(key, str) or key not in accepted:
            _reject(
                f"{name}.{_named(key)} is not a field this call may change; the fields it "
                f"accepts are {', '.join(allowed)}"
            )
    for required_name in spec.require_fields:
        if required_name not in value:
            _reject(f"{name}.{required_name} is required when this call creates the item")
    for key, entry in value.items():
        if entry is None:
            _reject(f"{name}.{key} must carry a value, not None")
        if isinstance(entry, Mapping):
            _reject(f"{name}.{key} must be a flat value, not a nested object")
        kind = spec.field_kinds[key]
        _KIND_VALIDATORS[kind](
            context, f"{name}.{key}", entry, _field_spec(key, kind)
        )


def _check_outlook_search_filters(context: _Context) -> None:
    """Cross-argument rules of ``outlook.search`` no single property check can express.

    Runs only after every declared property has validated, so a payload rejected here has
    already passed every property-level rule, and the two ways of narrowing a search are
    refused when combined instead of one of them being dropped silently.
    """
    supplied = context.arguments
    if "filter" in supplied and (
        "query" in supplied
        or any(
            name in supplied
            for name in ("has_attachments", "sender", "received_after", "received_before")
        )
    ):
        _reject(
            "the raw OData filter cannot be combined with the subject query or the "
            "structured filters: supply the raw filter alone or the query and "
            "structured filters alone, never both"
        )
    sender = supplied.get("sender")
    if isinstance(sender, str) and not _EMAIL_PATTERN.match(sender):
        _reject(
            "sender must be an email address: the structured sender filter is built from "
            "an address, not free text"
        )
    after = supplied.get("received_after")
    before = supplied.get("received_before")
    if isinstance(after, str) and isinstance(before, str):
        try:
            lower = datetime.fromisoformat(after)
            upper = datetime.fromisoformat(before)
        except ValueError:
            return  # the property check already rejects what it cannot parse
        if lower.tzinfo is not None and upper.tzinfo is not None and lower >= upper:
            _reject(
                "received_after must be strictly earlier than received_before: "
                "received_after is inclusive and received_before exclusive, so equal "
                "or reversed bounds make the range empty"
            )


#: Cross-argument rules a single property check cannot express, keyed by operation. A
#: payload that fails one is refused in :func:`check` before approval, secrets, a
#: credential or a Graph client.
_CROSS_ARGUMENT_CHECKS: Mapping[str, Callable[[_Context], None]] = {
    "outlook.search": _check_outlook_search_filters,
}


#: One validator per kind. Adding a kind to :data:`KINDS` without adding it here is a
#: programming error, not a silently skipped check.
_KIND_VALIDATORS: Mapping[str, Any] = {
    "identifier": _validate_identifier,
    "text": _validate_text,
    "boolean": _validate_boolean,
    "integer": _validate_integer,
    "timestamp": _validate_timestamp,
    "time_zone": _validate_time_zone,
    "etag": _validate_etag,
    "base64": _validate_base64,
    "recipients": _validate_recipients,
    "path": _validate_path,
    "select": _validate_select,
    "enum": _validate_enum,
    "media_type": _validate_media_type,
    "mapping": _validate_mapping,
}


def _field_spec(field_name: str, kind: str) -> ArgumentSpec:
    """The spec of a bounded PATCH field: its kind's defaults plus its declared integer bounds."""
    bounds = PATCH_INTEGER_BOUNDS.get(field_name)
    if kind == "integer":
        if bounds is None:  # pragma: no cover - refused while the contract table is built
            raise OperationArgumentsError(
                f"{field_name} is an integer PATCH field with no declared bounds"
            )
        return _integer(field_name, bounds[0], bounds[1])
    return _spec(field_name, kind)


def _validate_arguments(context: _Context, contract: OperationArguments) -> None:
    supplied = {name: value for name, value in context.arguments.items() if name != "action"}
    for name in supplied:
        if not isinstance(name, str):
            _reject(
                f"{context.key} arguments must be named: a {type(name).__name__} key cannot "
                "be an argument name"
            )
    allowed = set(contract.property_names)
    unknown = sorted(name for name in supplied if name not in allowed)
    if unknown:
        _reject(
            f"the argument {_named(unknown[0])} is not accepted by {context.key}; accepted "
            f"arguments are {', '.join(contract.property_names)}"
        )

    for spec in contract.properties:
        if spec.name not in supplied:
            if spec.required:
                _reject(f"{spec.name} is required by {context.key}")
            continue
        value = supplied[spec.name]
        if value is None:
            _reject(f"{spec.name} must carry a value, not None")
        _KIND_VALIDATORS[spec.kind](context, spec.name, value, spec)

    for exclusion in contract.exclusions:
        present = [form for form in exclusion.forms if all(name in supplied for name in form)]
        if exclusion.mode == "exactly_one" and len(present) != 1:
            _reject(
                f"{context.key} requires exactly one of these argument forms: "
                f"{exclusion.description}"
            )
        if exclusion.mode == "at_most_one" and len(present) > 1:
            _reject(
                f"{context.key} accepts at most one of these argument forms: "
                f"{exclusion.description}"
            )
    cross_check = _CROSS_ARGUMENT_CHECKS.get(context.key)
    if cross_check is not None:
        cross_check(context)


def _check_bound_user(settings: Settings, key: str, arguments: Mapping[str, Any]) -> Rejection | None:
    """Enforce the Outlook allowlist or the legacy single-user resource boundary."""
    definition = OPERATION_REGISTRY[key]
    if "user_id" not in definition.required_identifiers:
        return None
    if definition.service == "outlook" and settings.allowed_mailbox_ids:
        supplied = arguments.get("user_id")
        if not isinstance(supplied, str) or supplied not in settings.allowed_mailbox_ids:
            return Rejection(
                "validation_error",
                f"microsoft365.{key}: user_id is outside the configured Outlook mailbox allowlist",
            )
        return None
    configured = settings.user_id.strip()
    if not configured:
        return None
    supplied = arguments.get("user_id")
    if not isinstance(supplied, str) or not supplied.strip() or supplied.strip() != configured:
        return Rejection(
            "validation_error",
            f"microsoft365.{key}: user_id is outside the configured single-user resource boundary",
        )
    return None


def check(settings: Settings | None, *, service: Any, arguments: Any) -> Rejection | None:
    """Reject a call that must not reach approval, a handler, a secret or a Graph client.

    Returns ``None`` when the call may proceed, otherwise a :class:`Rejection` naming the
    category and a sanitized reason. Both call sites use this function with the same payload,
    so the hook's block message and the dispatch path's error payload cannot disagree. The
    payload is never modified: the arguments that pass this check are the arguments that are
    approved and executed.
    """
    if not isinstance(service, str) or service not in OPERATIONS:
        return Rejection(
            "validation_error", f"microsoft365: {_named(service)} is not a declared service"
        )
    if settings is None:
        return Rejection(
            "configuration_error",
            f"microsoft365.{service}: the plugin configuration is unusable, so no operation "
            "may be run",
        )
    if not isinstance(arguments, Mapping):
        # The exact literal the hook already answered with before WP13, so the hook's early
        # guard and this validator cannot report two different things for the same payload.
        return Rejection("validation_error", "Microsoft 365 arguments must be an object")
    operation = arguments.get("action")
    if not isinstance(operation, str) or not operation:
        return Rejection(
            "validation_error",
            f"microsoft365.{service}: a string operation name is required as 'action'",
        )
    key = f"{service}.{operation}"
    definition = OPERATION_REGISTRY.get(key)
    if definition is None:
        return Rejection(
            "validation_error",
            f"microsoft365.{service}: {_named(operation)} is not a declared operation; "
            f"declared operations are {', '.join(OPERATIONS[service])}",
        )
    # The mode is a property of the whole call: an operation this mode cannot run is refused
    # as unsupported before its enablement is even considered, so an unavailable mode is never
    # reported as a configuration problem.
    support = definition.support_for(settings.authentication_mode)
    if support is None or support.status in {"unsupported_auth_mode", "not_implemented"}:
        reason = (
            support.reason
            if support is not None
            else (
                f"the configured authentication mode {_named(settings.authentication_mode)} is "
                "not supported"
            )
        )
        return Rejection("unsupported_auth_mode", f"microsoft365.{key}: {reason}")
    if operation not in settings.selected(service):
        return Rejection(
            "validation_error",
            f"microsoft365.{key} is disabled by the plugin configuration; enable "
            f"capabilities.{service}.{operation} to use it",
        )
    boundary_rejection = _check_bound_user(settings, key, arguments)
    if boundary_rejection is not None:
        return boundary_rejection
    contract = ARGUMENT_CONTRACTS.get(key)
    if contract is None:
        # Unreachable while the contract table covers the registry (a test pins that); if it
        # ever happens, refusing is the only safe answer.
        return Rejection(
            "configuration_error",
            f"microsoft365.{key} has no argument contract, so it cannot be validated",
        )
    context = _Context(
        key=key, mode=settings.authentication_mode, arguments=arguments
    )
    try:
        _validate_arguments(context, contract)
    except _Failure as failure:
        return Rejection("validation_error", f"microsoft365.{key}: {failure.detail}")
    return None


def operation_schema(key: str) -> dict[str, Any]:
    """The JSON schema of one operation's arguments, derived from the same contract.

    This is the per-action schema WP13 owes: the model-facing surface of one operation, with
    ``additionalProperties`` off so an undeclared argument cannot be smuggled past the schema
    that the validator would reject anyway.
    """
    contract = ARGUMENT_CONTRACTS.get(key)
    if contract is None:
        raise KeyError(f"{key} has no argument contract")
    operation = key.split(".", 1)[1]
    properties: dict[str, Any] = {"action": {"type": "string", "enum": [operation]}}
    for spec in contract.properties:
        properties[spec.name] = _schema_for(spec)
    required = ["action"] + sorted(spec.name for spec in contract.properties if spec.required)
    return {
        "type": "object",
        "properties": properties,
        "required": required,
        "additionalProperties": False,
    }


def _schema_for(spec: ArgumentSpec) -> dict[str, Any]:
    if spec.kind == "integer":
        return {"type": "integer", "minimum": spec.minimum, "maximum": spec.maximum}
    if spec.kind == "boolean":
        return {"type": "boolean"}
    if spec.kind in {"select", "recipients"}:
        items: dict[str, Any] = {"type": "string", "maxLength": 64}
        if spec.kind == "recipients":
            items = {
                "oneOf": [
                    {"type": "string", "maxLength": 318, "description": "an email address"},
                    {
                        "type": "object",
                        "properties": {"address": {"type": "string", "maxLength": 318}, "name": {"type": "string", "maxLength": MAX_IDENTIFIER_LENGTH}},
                        "required": ["address"],
                        "additionalProperties": False,
                    },
                ]
            }
        return {"type": "array", "items": items, "minItems": 1, "maxItems": spec.max_items}
    if spec.kind == "enum":
        return {"type": "string", "enum": list(spec.values)}
    if spec.kind == "mapping":
        return {
            "type": "object",
            "additionalProperties": False,
            "minProperties": 1,
            "properties": {
                name: _schema_for(_field_spec(name, kind)) for name, kind in spec.fields
            },
            "required": sorted(spec.require_fields),
        }
    schema: dict[str, Any] = {"type": "string", "maxLength": spec.max_length}
    if spec.kind == "timestamp":
        schema["maxLength"] = _TIMESTAMP_MAX_LENGTH
        schema["description"] = "an ISO-8601 date and time; add time_zone when it has no offset"
    return schema
