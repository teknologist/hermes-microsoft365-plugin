"""Strict write-side msgraph-sdk contracts: request bodies, headers and /content.

Body assertions need a **real** Kiota serialization writer factory: the GET-side double
in ``tests/test_sdk_contract.py`` raises from ``get_serialization_writer_factory`` on
purpose. :class:`StrictSerializingTransportAdapter` therefore injects the genuine
``JsonSerializationWriterFactory`` from the installed
``microsoft-kiota-serialization-json`` distribution while keeping every ``send_*`` method
raising, so a serialized model body is observable and the network stays unreachable.

Nothing here uses ``SimpleNamespace``, ``MagicMock`` or a permissive ``__getattr__``: every
body is a real generated model imported from its own module.
"""
from __future__ import annotations

import json
from typing import Any
from urllib.parse import urlsplit

import pytest
from kiota_abstractions.request_adapter import RequestAdapter


class StrictSerializingTransportAdapter(RequestAdapter):
    """Serializes real models, never sends.

    ``get_serialization_writer_factory`` returns the installed JSON writer factory so
    ``to_*_request_information`` can serialize a generated model into
    ``RequestInformation.content``. All six ``send_*`` entry points still raise, so no
    code path in this module can reach a network.
    """

    def __init__(self):
        from kiota_serialization_json.json_serialization_writer_factory import (
            JsonSerializationWriterFactory,
        )

        self.base_url = "https://graph.microsoft.com/v1.0"
        self.writer_factory = JsonSerializationWriterFactory()

    def enable_backing_store(self, backing_store_factory=None):
        return None

    def get_serialization_writer_factory(self):
        return self.writer_factory

    async def send_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_collection_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_collection_of_primitive_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_primitive_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def send_no_response_content_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("network transport must not run")

    async def convert_to_native_async(self, *args: Any, **kwargs: Any):
        raise AssertionError("native conversion must not run")


@pytest.fixture
def graph_client():
    from msgraph import GraphServiceClient

    return GraphServiceClient(request_adapter=StrictSerializingTransportAdapter())


def serialized_body(request) -> dict:
    """Decode the real Kiota-serialized request body."""
    assert request.content, "expected a serialized request body"
    return json.loads(request.content.decode("utf-8"))


def test_writer_factory_is_the_real_installed_json_factory():
    from kiota_abstractions.serialization.serialization_writer_factory import (
        SerializationWriterFactory,
    )

    adapter = StrictSerializingTransportAdapter()

    assert isinstance(adapter.get_serialization_writer_factory(), SerializationWriterFactory)
    assert adapter.get_serialization_writer_factory().get_valid_content_type() == "application/json"


@pytest.mark.parametrize(
    "method",
    [
        "send_async",
        "send_collection_async",
        "send_collection_of_primitive_async",
        "send_primitive_async",
        "send_no_response_content_async",
        "convert_to_native_async",
    ],
)
def test_transport_remains_impossible_to_reach(method):
    import asyncio

    adapter = StrictSerializingTransportAdapter()

    with pytest.raises(AssertionError, match="must not run"):
        asyncio.run(getattr(adapter, method)())


OUTLOOK_DRAFT_ARGUMENTS = {
    "user_id": "user",
    "subject": "Weekly report",
    "content": "Body text",
    "recipients": ["someone@example.com", "else@example.com"],
}


def test_outlook_create_draft_posts_a_real_message_model(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client, "outlook_create_draft", **OUTLOOK_DRAFT_ARGUMENTS
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/users/user/messages"

    body = serialized_body(request)
    assert body["subject"] == "Weekly report"
    assert body["body"] == {"content": "Body text", "contentType": "text"}
    assert body["toRecipients"] == [
        {"emailAddress": {"address": "someone@example.com"}},
        {"emailAddress": {"address": "else@example.com"}},
    ]
    assert body["@odata.type"] == "#microsoft.graph.message"


@pytest.mark.parametrize(
    ("override", "message"),
    [
        ({"recipients": []}, "recipients must be a non-empty list"),
        ({"recipients": "someone@example.com"}, "recipients must be a non-empty list"),
        ({"recipients": [""]}, "recipients must be a non-empty list"),
        ({"content": "   "}, "content is required"),
        ({"subject": "  "}, "subject is required"),
    ],
)
def test_outlook_create_draft_refuses_an_incomplete_draft(graph_client, override, message):
    from microsoft365.sdk_contract import build_write_request_information

    arguments = dict(OUTLOOK_DRAFT_ARGUMENTS)
    arguments.update(override)

    with pytest.raises(ValueError, match=message):
        build_write_request_information(graph_client, "outlook_create_draft", **arguments)


def test_outlook_create_draft_requires_a_user_id(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    arguments = dict(OUTLOOK_DRAFT_ARGUMENTS)
    arguments["user_id"] = ""

    with pytest.raises(ValueError, match="user_id is required"):
        build_write_request_information(graph_client, "outlook_create_draft", **arguments)


def test_outlook_send_posts_the_generated_send_mail_body(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "outlook_send",
        save_to_sent_items=False,
        **OUTLOOK_DRAFT_ARGUMENTS,
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/users/user/sendMail"

    body = serialized_body(request)
    # the generated SendMailPostRequestBody serializes with its PascalCase Graph names
    assert set(body) == {"Message", "SaveToSentItems"}
    assert body["SaveToSentItems"] is False
    assert body["Message"]["subject"] == "Weekly report"
    assert body["Message"]["toRecipients"][0]["emailAddress"]["address"] == "someone@example.com"


def test_outlook_send_requires_the_save_to_sent_items_flag_to_be_a_bool(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="save_to_sent_items must be a boolean"):
        build_write_request_information(
            graph_client, "outlook_send", save_to_sent_items="false", **OUTLOOK_DRAFT_ARGUMENTS
        )


def test_outlook_send_existing_posts_without_any_body(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client, "outlook_send_existing", user_id="user", message_id="message"
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/users/user/messages/message/send"
    assert request.content in (None, b"")


def test_outlook_send_existing_requires_a_message_id(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="message_id is required"):
        build_write_request_information(
            graph_client, "outlook_send_existing", user_id="user", message_id=""
        )


def if_match(request) -> set:
    """The If-Match values actually configured on the request, or an empty set."""
    return request.headers.get("if-match") or set()


def test_calendar_create_events_posts_a_real_event_with_date_time_time_zone(graph_client):
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "calendar_create_events",
        user_id="user",
        subject="Standup",
        start=DateTimeTimeZone(date_time="2026-01-01T09:00:00", time_zone="UTC"),
        end=DateTimeTimeZone(date_time="2026-01-01T09:30:00", time_zone="UTC"),
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/users/user/calendar/events"

    body = serialized_body(request)
    assert body["subject"] == "Standup"
    assert body["start"] == {"dateTime": "2026-01-01T09:00:00", "timeZone": "UTC"}
    assert body["end"] == {"dateTime": "2026-01-01T09:30:00", "timeZone": "UTC"}
    assert body["@odata.type"] == "#microsoft.graph.event"
    assert if_match(request) == set()


@pytest.mark.parametrize("missing", ["start", "end"])
def test_calendar_create_events_requires_real_date_time_time_zone_models(graph_client, missing):
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from microsoft365.sdk_contract import build_write_request_information

    arguments = {
        "user_id": "user",
        "subject": "Standup",
        "start": DateTimeTimeZone(date_time="2026-01-01T09:00:00", time_zone="UTC"),
        "end": DateTimeTimeZone(date_time="2026-01-01T09:30:00", time_zone="UTC"),
    }
    arguments[missing] = {"dateTime": "2026-01-01T09:00:00", "timeZone": "UTC"}

    with pytest.raises(ValueError, match=f"{missing} must be a DateTimeTimeZone"):
        build_write_request_information(graph_client, "calendar_create_events", **arguments)


def test_calendar_update_events_patches_a_partial_body_with_if_match(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "calendar_update_events",
        user_id="user",
        event_id="event",
        etag='"etag-1"',
        update_fields={"subject": "Renamed"},
    )

    assert request.http_method.value == "PATCH"
    assert urlsplit(request.url).path == "/users/user/calendar/events/event"
    assert if_match(request) == {'"etag-1"'}

    body = serialized_body(request)
    # a partial PATCH: nothing the caller did not ask to change may appear
    assert body == {"@odata.type": "#microsoft.graph.event", "subject": "Renamed"}


def test_calendar_update_events_omits_if_match_when_no_etag_is_given(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "calendar_update_events",
        user_id="user",
        event_id="event",
        update_fields={"subject": "Renamed"},
    )

    assert if_match(request) == set()
    assert serialized_body(request) == {"@odata.type": "#microsoft.graph.event", "subject": "Renamed"}


def test_calendar_update_events_refuses_an_unknown_field(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="unknown field for calendar_update_events: subjet"):
        build_write_request_information(
            graph_client,
            "calendar_update_events",
            user_id="user",
            event_id="event",
            update_fields={"subjet": "typo"},
        )


def test_calendar_update_events_refuses_an_empty_patch(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="calendar_update_events requires a non-empty update_fields"):
        build_write_request_information(
            graph_client, "calendar_update_events", user_id="user", event_id="event", update_fields={}
        )


def test_todo_create_tasks_posts_a_real_todo_task_using_title(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "todo_create_tasks",
        user_id="user",
        todo_list_id="list",
        title="Buy milk",
        content="2 litres",
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/users/user/todo/lists/list/tasks"

    body = serialized_body(request)
    assert body["title"] == "Buy milk"
    assert body["body"] == {"content": "2 litres", "contentType": "text"}
    assert "subject" not in body


def test_todo_create_tasks_requires_a_title(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="title is required"):
        build_write_request_information(
            graph_client, "todo_create_tasks", user_id="user", todo_list_id="list", title="  "
        )


def test_todo_update_tasks_patches_a_partial_body_with_if_match(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "todo_update_tasks",
        user_id="user",
        todo_list_id="list",
        todo_task_id="task",
        etag='"etag-3"',
        update_fields={"title": "Buy oat milk"},
    )

    assert request.http_method.value == "PATCH"
    assert urlsplit(request.url).path == "/users/user/todo/lists/list/tasks/task"
    assert if_match(request) == {'"etag-3"'}
    assert serialized_body(request) == {"title": "Buy oat milk"}


def test_todo_update_tasks_refuses_subject_because_the_model_has_no_such_field(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="unknown field for todo_update_tasks: subject"):
        build_write_request_information(
            graph_client,
            "todo_update_tasks",
            user_id="user",
            todo_list_id="list",
            todo_task_id="task",
            update_fields={"subject": "To Do uses title"},
        )


def test_planner_create_tasks_posts_plan_bucket_and_title(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client, "planner_create_tasks", plan_id="plan", bucket_id="bucket", title="Draft"
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/planner/tasks"
    assert serialized_body(request) == {"bucketId": "bucket", "planId": "plan", "title": "Draft"}


@pytest.mark.parametrize(
    ("missing", "arguments"),
    [
        ("plan_id", {"plan_id": "", "bucket_id": "bucket", "title": "Draft"}),
        ("bucket_id", {"plan_id": "plan", "bucket_id": "", "title": "Draft"}),
        ("title", {"plan_id": "plan", "bucket_id": "bucket", "title": ""}),
    ],
)
def test_planner_create_tasks_requires_plan_bucket_and_title(graph_client, missing, arguments):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match=f"{missing} is required"):
        build_write_request_information(graph_client, "planner_create_tasks", **arguments)


def test_planner_update_tasks_requires_an_etag_and_sends_if_match(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="etag is required for planner_update_tasks"):
        build_write_request_information(
            graph_client,
            "planner_update_tasks",
            planner_task_id="task",
            update_fields={"title": "Renamed"},
        )

    request = build_write_request_information(
        graph_client,
        "planner_update_tasks",
        planner_task_id="task",
        etag='"etag-2"',
        update_fields={"title": "Renamed"},
    )

    assert request.http_method.value == "PATCH"
    assert urlsplit(request.url).path == "/planner/tasks/task"
    assert if_match(request) == {'"etag-2"'}
    assert serialized_body(request) == {"title": "Renamed"}


def test_teams_send_messages_posts_a_real_chat_message_body(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "teams_send_messages",
        team_id="team",
        channel_id="channel",
        content="hello",
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/teams/team/channels/channel/messages"
    assert serialized_body(request) == {"body": {"content": "hello", "contentType": "text"}}


def test_teams_send_messages_requires_team_channel_and_content(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="team_id is required"):
        build_write_request_information(
            graph_client, "teams_send_messages", team_id="", channel_id="channel", content="hi"
        )
    with pytest.raises(ValueError, match="channel_id is required"):
        build_write_request_information(
            graph_client, "teams_send_messages", team_id="team", channel_id="", content="hi"
        )
    with pytest.raises(ValueError, match="content is required"):
        build_write_request_information(
            graph_client, "teams_send_messages", team_id="team", channel_id="channel", content=""
        )


def test_teams_search_messages_posts_the_generated_search_query_body(graph_client):
    from msgraph.generated.models.entity_type import EntityType
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client,
        "teams_search_messages",
        query="hello",
        entity_types=[EntityType.ChatMessage],
        from_=0,
        size=25,
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/search/query"
    assert serialized_body(request) == {
        "requests": [
            {
                "entityTypes": ["chatMessage"],
                "from": 0,
                "query": {"queryString": "hello"},
                "size": 25,
            }
        ]
    }


def test_teams_search_messages_requires_real_entity_type_members(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="entity_types must be a non-empty list of EntityType members"):
        build_write_request_information(
            graph_client, "teams_search_messages", query="hello", entity_types=["chatMessage"], size=25
        )


def test_teams_search_messages_requires_a_query_and_a_bounded_page(graph_client):
    from msgraph.generated.models.entity_type import EntityType
    from microsoft365.sdk_contract import build_write_request_information

    arguments = {"query": "hello", "entity_types": [EntityType.ChatMessage], "from_": 0, "size": 25}

    with pytest.raises(ValueError, match="query is required"):
        build_write_request_information(
            graph_client, "teams_search_messages", **{**arguments, "query": " "}
        )
    with pytest.raises(ValueError, match="size must be a positive integer"):
        build_write_request_information(
            graph_client, "teams_search_messages", **{**arguments, "size": 0}
        )
    with pytest.raises(ValueError, match="from_ must be a non-negative integer"):
        build_write_request_information(
            graph_client, "teams_search_messages", **{**arguments, "from_": -1}
        )


def test_upload_session_builds_the_request_but_is_declared_not_implemented(graph_client):
    from microsoft365.sdk_contract import UNIMPLEMENTED_CASES, build_write_request_information

    request = build_write_request_information(
        graph_client, "upload_session", drive_id="drive", drive_item_id="item"
    )

    assert request.http_method.value == "POST"
    assert urlsplit(request.url).path == "/drives/drive/items/item/createUploadSession"
    assert serialized_body(request) == {}
    assert "upload_session" in UNIMPLEMENTED_CASES


def test_write_contract_rejects_an_unknown_case(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match="unknown write contract"):
        build_write_request_information(graph_client, "outlook_delete_message")


CONTENT_ARGUMENTS = {"drive_id": "drive", "drive_item_id": "root:/a/b.txt:"}


def test_content_contract_exposes_exactly_the_verified_cases():
    import microsoft365.sdk_contract as sdk_contract

    assert sdk_contract.CONTENT_CASES == ("download_files", "upload_files")


def test_download_files_url_ends_in_content(graph_client):
    from microsoft365.sdk_contract import build_content_request_information

    request = build_content_request_information(graph_client, "download_files", **CONTENT_ARGUMENTS)

    assert request.http_method.value == "GET"
    assert request.url.endswith("/content")
    assert urlsplit(request.url).path == "/drives/drive/items/root%3A%2Fa%2Fb.txt%3A/content"


def test_upload_files_puts_exact_bytes_with_content_type_to_the_content_url(graph_client):
    from microsoft365.sdk_contract import build_content_request_information

    payload = bytes(range(256))
    request = build_content_request_information(
        graph_client,
        "upload_files",
        content=payload,
        content_type="application/octet-stream",
        **CONTENT_ARGUMENTS,
    )

    assert request.http_method.value == "PUT"
    assert request.url.endswith("/content")
    assert urlsplit(request.url).path == "/drives/drive/items/root%3A%2Fa%2Fb.txt%3A/content"
    assert request.content == payload
    assert request.headers.get("content-type") == {"application/octet-stream"}


@pytest.mark.parametrize(
    ("override", "message"),
    [
        ({"content": "not bytes"}, "content must be bytes"),
        ({"content": {"a": 1}}, "content must be bytes"),
        ({"content": None}, "content must be bytes"),
        ({"content_type": "   "}, "content_type is required"),
    ],
)
def test_upload_files_refuses_a_non_binary_or_untyped_upload(graph_client, override, message):
    from microsoft365.sdk_contract import build_content_request_information

    arguments = {
        "content": b"payload",
        "content_type": "text/plain",
        **CONTENT_ARGUMENTS,
        **override,
    }

    with pytest.raises(ValueError, match=message):
        build_content_request_information(graph_client, "upload_files", **arguments)


@pytest.mark.parametrize(
    ("case", "override", "message"),
    [
        ("download_files", {"drive_id": ""}, "drive_id is required"),
        ("download_files", {"drive_item_id": ""}, "drive_item_id is required"),
        ("upload_files", {"drive_id": "  "}, "drive_id is required"),
    ],
)
def test_content_contract_requires_explicit_drive_and_item_ids(graph_client, case, override, message):
    from microsoft365.sdk_contract import build_content_request_information

    arguments = {"content": b"payload", "content_type": "text/plain", **CONTENT_ARGUMENTS, **override}

    with pytest.raises(ValueError, match=message):
        build_content_request_information(graph_client, case, **arguments)


def test_content_contract_rejects_an_unknown_case(graph_client):
    from microsoft365.sdk_contract import build_content_request_information

    with pytest.raises(ValueError, match="unknown content contract"):
        build_content_request_information(graph_client, "delete_files", **CONTENT_ARGUMENTS)


def test_every_row_in_apendice_a_is_bound_to_an_endpoint_or_a_declared_gap():
    """Anti-drift: the four dispatchers cover exactly the documented case names."""
    import microsoft365.sdk_contract as sdk_contract

    assert sdk_contract.ITEM_READ_CASES == (
        "outlook_attachment",
        "outlook_read",
        "calendar_read",
        "todo_read",
        "planner_read",
        "drive_item_read",
    )
    assert sdk_contract.SEARCH_CASES == ("sites_search", "drive_search", "drive_item_search")
    assert sdk_contract.CONTENT_CASES == ("download_files", "upload_files")
    assert sdk_contract.UNIMPLEMENTED_CASES <= set(sdk_contract.WRITE_CASES)
    assert len(set(sdk_contract.WRITE_CASES)) == len(sdk_contract.WRITE_CASES)
