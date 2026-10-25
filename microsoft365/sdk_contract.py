"""Typed msgraph-sdk 1.62.0 request-information contracts."""
from __future__ import annotations

from kiota_abstractions.base_request_configuration import RequestConfiguration
from kiota_abstractions.headers_collection import HeadersCollection
from kiota_abstractions.serialization.parsable import Parsable

# Generated models are imported inside the builders below rather than at module import
# time: the module stays cheap to import and mirrors the repo's existing lazy-SDK style
# (``client.py`` resolves ``azure.identity``/``msgraph`` at execution time only).
_FRAMEWORK_FIELDS = frozenset({"backing_store", "additional_data"})

#: The compact candidate-metadata ``$select`` the outlook messages collection read sends when
#: the caller names no field of its own. Pinned here as a literal -- deliberately *not*
#: imported from the handler -- so the request pin in ``tests/test_handlers_outlook_calendar.py``
#: (which asserts the same literal, not the handler constant) catches a drift of the default
#: select on either side instead of computing the expected result from it.
OUTLOOK_MESSAGES_DEFAULT_SELECT = ["id", "subject", "from", "receivedDateTime", "hasAttachments", "webLink"]


def _required(value: str, name: str) -> str:
    value = str(value or "").strip()
    if not value:
        raise ValueError(f"{name} is required")
    return value


def _limit(value: int) -> int:
    if type(value) is not int or not 1 <= value <= 100:
        raise ValueError("limit must be an integer between 1 and 100")
    return value


def _odata_literal(value: str) -> str:
    return _required(value, "query").replace("'", "''")


def _select(select):
    """Normalize a caller-supplied ``$select`` list, or ``None`` for "no $select"."""
    if select is None:
        return None
    values = [str(value or "").strip() for value in select]
    if not values or any(not value for value in values):
        raise ValueError("select must be a non-empty list of field names")
    return values


def _request(builder, **query_values):
    query_type = getattr(type(builder), f"{type(builder).__name__}GetQueryParameters")
    query = query_type(**{name: value for name, value in query_values.items() if value is not None})
    configuration = RequestConfiguration(query_parameters=query)
    return builder.to_get_request_information(configuration)


def build_collection_request_information(
    client,
    case: str,
    *,
    user_id: str = "",
    todo_list_id: str = "",
    message_id: str = "",
    drive_id: str = "",
    drive_item_id: str = "",
    team_id: str = "",
    plan_id: str = "",
    owner_id: str = "",
    group_id: str | None = None,
    query: str = "",
    limit: int = 50,
):
    """Build request info with each generated collection's exact query type.

    The function stops at Kiota's transport boundary. It neither authenticates
    nor sends a request, making exact SDK shape testable offline.

    ``group_id`` selects the container for ``planner_plans``: ``None`` (default) lists
    every plan the caller can see through ``/planner/plans``; a non-blank value switches
    to the group-owned route ``/groups/{g}/planner/plans``. A blank string is rejected
    instead of silently falling back to the global route.
    """
    bounded = _limit(limit)
    if case == "outlook_attachments":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages.by_message_id(
            _required(message_id, "message_id")
        ).attachments
        return _request(builder, top=bounded, select=["id", "name", "size", "contentType", "isInline"])
    if case == "outlook_messages":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages
        return _request(
            builder,
            top=bounded,
            select=list(OUTLOOK_MESSAGES_DEFAULT_SELECT),
            filter=f"contains(subject,'{_odata_literal(query)}')",
        )
    if case == "calendar_events":
        builder = client.users.by_user_id(_required(user_id, "user_id")).calendar.events
        return _request(builder, top=bounded, filter=f"contains(subject,'{_odata_literal(query)}')")
    if case == "todo_lists":
        builder = client.users.by_user_id(_required(user_id, "user_id")).todo.lists
        return _request(builder, top=bounded)
    if case == "todo_tasks":
        builder = client.users.by_user_id(_required(user_id, "user_id")).todo.lists.by_todo_task_list_id(
            _required(todo_list_id, "todo_list_id")
        ).tasks
        return _request(builder, top=bounded, filter=f"contains(title,'{_odata_literal(query)}')")
    if case == "drive_children":
        builder = client.drives.by_drive_id(_required(drive_id, "drive_id")).items.by_drive_item_id(
            _required(drive_item_id, "drive_item_id")
        ).children
        return _request(builder, top=bounded)
    if case == "teams_joined":
        builder = client.users.by_user_id(_required(user_id, "user_id")).joined_teams
        return _request(builder)
    if case == "teams_channels":
        builder = client.teams.by_team_id(_required(team_id, "team_id")).channels
        return _request(builder, select=["id", "displayName"])
    if case == "planner_plans":
        if group_id is None:
            builder = client.planner.plans
        else:
            builder = client.groups.by_group_id(_required(group_id, "group_id")).planner.plans
        return _request(builder, top=bounded, filter=f"owner eq '{_odata_literal(owner_id)}'")
    if case == "planner_buckets":
        builder = client.planner.plans.by_planner_plan_id(_required(plan_id, "plan_id")).buckets
        return _request(builder, top=bounded)
    if case == "planner_tasks":
        builder = client.planner.plans.by_planner_plan_id(_required(plan_id, "plan_id")).tasks
        return _request(builder, top=bounded)
    raise ValueError(f"unknown collection contract: {case}")


ITEM_READ_CASES: tuple[str, ...] = (
    "outlook_attachment",
    "outlook_read",
    "calendar_read",
    "todo_read",
    "planner_read",
    "drive_item_read",
)


def build_item_request_information(
    client,
    case: str,
    *,
    user_id: str = "",
    message_id: str = "",
    event_id: str = "",
    todo_list_id: str = "",
    todo_task_id: str = "",
    planner_task_id: str = "",
    drive_id: str = "",
    drive_item_id: str = "",
    select=None,
    attachment_id: str = "",
):
    """Build request info for a *single item* read (Apêndice A "item" rows).

    Item reads are GETs against the generated ``*ItemRequestBuilder`` that the SDK
    derives from the collection builder, so the path is never assembled here.
    ``select`` is the caller's ``$select``; ``None`` means "send no ``$select``".
    """
    fields = _select(select)
    if case == "outlook_read":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages.by_message_id(
            _required(message_id, "message_id")
        )
        return _request(builder, select=fields)
    if case == "outlook_attachment":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages.by_message_id(
            _required(message_id, "message_id")
        ).attachments.by_attachment_id(_required(attachment_id, "attachment_id"))
        return _request(builder, select=fields)
    if case == "calendar_read":
        builder = client.users.by_user_id(_required(user_id, "user_id")).calendar.events.by_event_id(
            _required(event_id, "event_id")
        )
        return _request(builder, select=fields)
    if case == "todo_read":
        builder = (
            client.users.by_user_id(_required(user_id, "user_id"))
            .todo.lists.by_todo_task_list_id(_required(todo_list_id, "todo_list_id"))
            .tasks.by_todo_task_id(_required(todo_task_id, "todo_task_id"))
        )
        return _request(builder, select=_todo_task_select(fields))
    if case == "planner_read":
        builder = client.planner.tasks.by_planner_task_id(_required(planner_task_id, "planner_task_id"))
        return _request(builder, select=fields)
    if case == "drive_item_read":
        builder = client.drives.by_drive_id(_required(drive_id, "drive_id")).items.by_drive_item_id(
            _required(drive_item_id, "drive_item_id")
        )
        return _request(builder, select=fields)
    raise ValueError(f"unknown item contract: {case}")


SEARCH_CASES: tuple[str, ...] = ("sites_search", "drive_search", "drive_item_search")


def build_search_request_information(
    client,
    case: str,
    *,
    drive_id: str = "",
    drive_item_id: str = "",
    query: str = "",
):
    """Build request info for the search endpoints verified in Apêndice A.

    ``sites_search`` is a plain ``GET /sites`` whose ``$search`` value the generated
    builder URL-encodes, so the value is passed through untouched. ``drive_search`` and
    ``drive_item_search`` go through the generated ``search_with_q`` **method** whose URL
    template is ``.../search(q='{q}')``: the SDK encodes the path segment but does **not**
    escape single quotes as OData requires, so ``query`` is quote-doubled here.
    """
    if case == "sites_search":
        return _request(client.sites, search=_required(query, "query"))
    if case == "drive_item_search":
        builder = (
            client.drives.by_drive_id(_required(drive_id, "drive_id"))
            .items.by_drive_item_id(_required(drive_item_id, "drive_item_id"))
            .search_with_q(_odata_literal(query))
        )
        return builder.to_get_request_information()
    if case == "drive_search":
        builder = client.drives.by_drive_id(_required(drive_id, "drive_id")).search_with_q(
            _odata_literal(query)
        )
        return builder.to_get_request_information()
    raise ValueError(f"unknown search contract: {case}")


def _headers(**headers: str) -> HeadersCollection:
    """Build a real Kiota ``HeadersCollection`` -- a plain dict fails in ``configure()``."""
    collection = HeadersCollection()
    for name, value in headers.items():
        if value:
            collection.add(name, value)
    return collection


def _item_body(content: str, content_type: str):
    from msgraph.generated.models.body_type import BodyType
    from msgraph.generated.models.item_body import ItemBody

    try:
        kind = BodyType(str(content_type or "").strip())
    except ValueError:
        raise ValueError("content_type must be one of text, html") from None
    return ItemBody(content_type=kind, content=_required(content, "content"))


def _recipients(recipients):
    from msgraph.generated.models.email_address import EmailAddress
    from msgraph.generated.models.recipient import Recipient

    if not isinstance(recipients, (list, tuple)) or not recipients:
        raise ValueError("recipients must be a non-empty list of email addresses")
    addresses = [str(address or "").strip() for address in recipients]
    if any(not address for address in addresses):
        raise ValueError("recipients must be a non-empty list of email addresses")
    return [Recipient(email_address=EmailAddress(address=address)) for address in addresses]


def _draft_message(subject: str, content: str, recipients, content_type: str):
    from msgraph.generated.models.message import Message

    return Message(
        subject=_required(subject, "subject"),
        body=_item_body(content, content_type),
        to_recipients=_recipients(recipients),
    )


def _partial_model(model_class, fields, case: str):
    """Build a real generated model from an explicit, validated update mapping.

    ``fields`` keys must be real attributes of the generated model, so an attribute the
    endpoint ignores (``subject`` on ``TodoTask``) is refused here instead of being sent
    and silently dropped by Graph. Framework internals cannot be handed in.
    """
    if not isinstance(fields, dict) or not fields:
        raise ValueError(f"{case} requires a non-empty update_fields mapping")
    declared = {
        name for name in getattr(model_class, "__dataclass_fields__", {}) if name not in _FRAMEWORK_FIELDS
    }
    for name, value in fields.items():
        if name not in declared:
            raise ValueError(f"unknown field for {case}: {name}")
        if not (isinstance(value, (str, int, float, bool)) or isinstance(value, Parsable)):
            raise ValueError(f"field {name} of {case} must be a scalar or a generated model")
    return model_class(**fields)


# --------------------------------------------------------------------------------------
# Microsoft To Do semantics (WP3)
# --------------------------------------------------------------------------------------
#
# Two verified offline facts drive this section:
#
# * the generated ``TodoTask`` model has no ``subject`` property (To Do uses ``title``), and
#   its deserializers are the exact set of Graph wire names the model declares, so a
#   ``$select`` naming anything else is refused instead of being sent and ignored;
# * Kiota's JSON writer emits **nothing** for an enum-typed field whose value is not one of
#   the generated ``TaskStatus`` members -- a plain ``"inProgress"`` string serializes to
#   ``{}`` -- so a status has to be mapped to the generated member before the body is built.


def todo_task_field_names() -> frozenset[str]:
    """The Graph wire names the generated ``TodoTask`` model actually declares.

    Read from the generated model's own deserializers, never hand-written: ``subject`` is
    absent from the set because no To Do model has that property, so it cannot be selected
    or written through this plugin.
    """
    from msgraph.generated.models.todo_task import TodoTask

    return frozenset(TodoTask().get_field_deserializers())


def todo_task_status(value):
    """Map a caller-supplied status to the generated ``TaskStatus`` enum member.

    A generated member passes through unchanged and a string is mapped to the member it
    names. A value the generated enum does not declare is refused, listing the members the
    SDK actually has, rather than being handed to the writer -- which would drop the field
    from the request body without a word.
    """
    from msgraph.generated.models.task_status import TaskStatus

    if isinstance(value, TaskStatus):
        return value
    if isinstance(value, str):
        normalized = value.strip()
        if normalized:
            try:
                return TaskStatus(normalized)
            except ValueError:
                pass
    members = ", ".join(member.value for member in TaskStatus)
    raise ValueError(f"status must be one of the generated TaskStatus members: {members}")


def todo_task_status_matches(task_status, requested) -> bool:
    """Match one task's status client-side against a requested one.

    The endpoint's acceptance of a ``$filter`` on ``status`` is not recorded offline, so no
    such filter is sent and the comparison happens here, limited to the tasks the caller has
    already retrieved. A task with no status, or one whose value the generated enum does not
    declare, never matches.
    """
    from msgraph.generated.models.task_status import TaskStatus

    wanted = todo_task_status(requested)
    if isinstance(task_status, TaskStatus):
        return task_status is wanted
    if isinstance(task_status, str):
        normalized = task_status.strip()
        if normalized:
            try:
                return TaskStatus(normalized) is wanted
            except ValueError:
                return False
    return False


def _todo_task_select(fields):
    """Refuse a ``$select`` naming a field the generated ``TodoTask`` model does not have.

    A selection the model cannot return is silently ignored by Graph, so it is refused here
    with the case name instead of producing a request that drops the caller's intent.
    """
    if fields is None:
        return None
    declared = todo_task_field_names()
    for name in fields:
        if name not in declared:
            raise ValueError(f"unknown field for todo_read: {name}")
    return fields


def _todo_task_write_model(case: str, fields):
    """Build the real generated ``TodoTask`` a To Do write sends.

    The field policy lives in ``microsoft365/contract.py``, so a field the endpoint does not
    accept is refused before the SDK is reached, and the model is built from real generated
    values only: ``status`` is mapped to its enum member, and the model-valued fields are
    type-checked so a plain string cannot reach the serializer as an attribute error.
    """
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.todo_task import TodoTask

    from .contract import validate_todo_task_fields

    accepted = dict(validate_todo_task_fields(case, fields))
    if "title" in accepted:
        accepted["title"] = _required(accepted["title"], "title")
    if "body" in accepted and not isinstance(accepted["body"], ItemBody):
        raise ValueError(f"body of {case} must be an ItemBody")
    if "due_date_time" in accepted and not isinstance(accepted["due_date_time"], DateTimeTimeZone):
        raise ValueError(f"due_date_time of {case} must be a DateTimeTimeZone")
    if "status" in accepted:
        accepted["status"] = todo_task_status(accepted["status"])
    return TodoTask(**accepted)


WRITE_CASES: tuple[str, ...] = (
    "outlook_create_draft",
    "outlook_send",
    "outlook_send_existing",
    "calendar_create_events",
    "calendar_update_events",
    "todo_create_tasks",
    "todo_update_tasks",
    "planner_create_tasks",
    "planner_update_tasks",
    "teams_send_messages",
    "teams_search_messages",
    "upload_session",
)

# Rows whose request shape is fully verified here but which the product deliberately does
# not execute yet: >10 MiB transfers stay reported as not implemented (see WP4/D3).
UNIMPLEMENTED_CASES: frozenset[str] = frozenset({"upload_session"})


def build_write_request_information(
    client,
    case: str,
    *,
    user_id: str = "",
    message_id: str = "",
    event_id: str = "",
    todo_list_id: str = "",
    todo_task_id: str = "",
    plan_id: str = "",
    bucket_id: str = "",
    planner_task_id: str = "",
    team_id: str = "",
    channel_id: str = "",
    drive_id: str = "",
    drive_item_id: str = "",
    subject: str = "",
    content: str = "",
    content_type: str = "text",
    recipients=None,
    start=None,
    end=None,
    title: str = "",
    due_date_time=None,
    status=None,
    save_to_sent_items=None,
    etag: str = "",
    query: str = "",
    entity_types=None,
    from_: int = 0,
    size: int = 0,
    update_fields=None,
):
    """Build request info for every write row in Apêndice A.

    Bodies are **real generated models** (``Message``, ``SendMailPostRequestBody``,
    ``Event``, ``TodoTask``, ``PlannerTask``, ``ChatMessage``, ``QueryPostRequestBody``)
    so the serialized payload is exactly what Graph receives. ``If-Match`` is only added
    where an ETag was supplied, except ``planner_update_tasks`` where Graph requires it.
    """
    if case == "outlook_create_draft":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages
        return builder.to_post_request_information(
            _draft_message(subject, content, recipients, content_type),
            RequestConfiguration(headers=_headers()),
        )
    if case == "outlook_send":
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )

        if type(save_to_sent_items) is not bool:
            raise ValueError("save_to_sent_items must be a boolean")
        builder = client.users.by_user_id(_required(user_id, "user_id")).send_mail
        return builder.to_post_request_information(
            SendMailPostRequestBody(
                message=_draft_message(subject, content, recipients, content_type),
                save_to_sent_items=save_to_sent_items,
            ),
            RequestConfiguration(headers=_headers()),
        )
    if case == "outlook_send_existing":
        builder = client.users.by_user_id(_required(user_id, "user_id")).messages.by_message_id(
            _required(message_id, "message_id")
        ).send
        return builder.to_post_request_information(RequestConfiguration(headers=_headers()))
    if case == "calendar_create_events":
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.event import Event

        for name, value in (("start", start), ("end", end)):
            if not isinstance(value, DateTimeTimeZone):
                raise ValueError(f"{name} must be a DateTimeTimeZone")
        builder = client.users.by_user_id(_required(user_id, "user_id")).calendar.events
        return builder.to_post_request_information(
            Event(subject=_required(subject, "subject"), start=start, end=end),
            RequestConfiguration(headers=_headers()),
        )
    if case == "calendar_update_events":
        from msgraph.generated.models.event import Event

        builder = client.users.by_user_id(_required(user_id, "user_id")).calendar.events.by_event_id(
            _required(event_id, "event_id")
        )
        return builder.to_patch_request_information(
            _partial_model(Event, update_fields, case), RequestConfiguration(headers=_headers(**{"If-Match": etag}))
        )
    if case == "todo_create_tasks":
        builder = client.users.by_user_id(_required(user_id, "user_id")).todo.lists.by_todo_task_list_id(
            _required(todo_list_id, "todo_list_id")
        ).tasks
        fields: dict[str, object] = {"title": _required(title, "title")}
        if content:
            fields["body"] = _item_body(content, content_type)
        if due_date_time is not None:
            fields["due_date_time"] = due_date_time
        if status is not None:
            fields["status"] = status
        return builder.to_post_request_information(
            _todo_task_write_model(case, fields),
            RequestConfiguration(headers=_headers()),
        )
    if case == "todo_update_tasks":
        builder = (
            client.users.by_user_id(_required(user_id, "user_id"))
            .todo.lists.by_todo_task_list_id(_required(todo_list_id, "todo_list_id"))
            .tasks.by_todo_task_id(_required(todo_task_id, "todo_task_id"))
        )
        return builder.to_patch_request_information(
            _todo_task_write_model(case, update_fields),
            RequestConfiguration(headers=_headers(**{"If-Match": etag})),
        )
    if case == "planner_create_tasks":
        from msgraph.generated.models.planner_task import PlannerTask

        builder = client.planner.tasks
        return builder.to_post_request_information(
            PlannerTask(
                plan_id=_required(plan_id, "plan_id"),
                bucket_id=_required(bucket_id, "bucket_id"),
                title=_required(title, "title"),
            ),
            RequestConfiguration(headers=_headers()),
        )
    if case == "planner_update_tasks":
        from msgraph.generated.models.planner_task import PlannerTask

        builder = client.planner.tasks.by_planner_task_id(_required(planner_task_id, "planner_task_id"))
        if not str(etag or "").strip():
            raise ValueError("etag is required for planner_update_tasks")
        return builder.to_patch_request_information(
            _partial_model(PlannerTask, update_fields, case),
            RequestConfiguration(headers=_headers(**{"If-Match": etag})),
        )
    if case == "teams_send_messages":
        from msgraph.generated.models.chat_message import ChatMessage

        builder = (
            client.teams.by_team_id(_required(team_id, "team_id"))
            .channels.by_channel_id(_required(channel_id, "channel_id"))
            .messages
        )
        return builder.to_post_request_information(
            ChatMessage(body=_item_body(content, content_type)),
            RequestConfiguration(headers=_headers()),
        )
    if case == "teams_search_messages":
        from msgraph.generated.models.entity_type import EntityType
        from msgraph.generated.models.search_query import SearchQuery
        from msgraph.generated.models.search_request import SearchRequest
        from msgraph.generated.search.query.query_post_request_body import QueryPostRequestBody

        if not isinstance(entity_types, (list, tuple)) or not entity_types:
            raise ValueError("entity_types must be a non-empty list of EntityType members")
        if any(not isinstance(value, EntityType) for value in entity_types):
            raise ValueError("entity_types must be a non-empty list of EntityType members")
        if type(from_) is not int or from_ < 0:
            raise ValueError("from_ must be a non-negative integer")
        if type(size) is not int or size < 1:
            raise ValueError("size must be a positive integer")
        builder = client.search.query
        return builder.to_post_request_information(
            QueryPostRequestBody(
                requests=[
                    SearchRequest(
                        entity_types=[EntityType(value) for value in entity_types],
                        query=SearchQuery(query_string=_required(query, "query")),
                        from_=from_,
                        size=size,
                    )
                ]
            ),
            RequestConfiguration(headers=_headers()),
        )
    if case == "upload_session":
        from msgraph.generated.drives.item.items.item.create_upload_session.create_upload_session_post_request_body import (
            CreateUploadSessionPostRequestBody,
        )

        builder = client.drives.by_drive_id(_required(drive_id, "drive_id")).items.by_drive_item_id(
            _required(drive_item_id, "drive_item_id")
        ).create_upload_session
        return builder.to_post_request_information(
            CreateUploadSessionPostRequestBody(), RequestConfiguration(headers=_headers())
        )
    raise ValueError(f"unknown write contract: {case}")


CONTENT_CASES: tuple[str, ...] = ("download_files", "upload_files")


def build_content_request_information(
    client,
    case: str,
    *,
    drive_id: str = "",
    drive_item_id: str = "",
    content=None,
    content_type: str = "",
):
    """Build request info for the drive ``/content`` rows.

    Path addressing goes through ``items.by_drive_item_id("<root|item-id>:/<caminho>")``
    so the rendered URL always ends in ``/content``; no URL is concatenated by hand.
    """
    builder = (
        client.drives.by_drive_id(_required(drive_id, "drive_id"))
        .items.by_drive_item_id(_required(drive_item_id, "drive_item_id"))
        .content
    )
    if case == "download_files":
        return builder.to_get_request_information(RequestConfiguration(headers=_headers()))
    if case == "upload_files":
        if not isinstance(content, (bytes, bytearray)):
            raise ValueError("content must be bytes")
        return builder.to_put_request_information(
            bytes(content),
            RequestConfiguration(headers=_headers(**{"Content-Type": _required(content_type, "content_type")})),
        )
    raise ValueError(f"unknown content contract: {case}")
