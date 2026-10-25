from __future__ import annotations

from typing import Any
from urllib.parse import urlsplit

import pytest
from kiota_abstractions.request_adapter import RequestAdapter


class StrictTransportAdapter(RequestAdapter):
    """Concrete Kiota boundary double: no generated builder/member fakery."""

    def __init__(self):
        self.base_url = "https://graph.microsoft.com/v1.0"

    def enable_backing_store(self, backing_store_factory=None):
        return None

    def get_serialization_writer_factory(self):
        raise AssertionError("serialization is not expected for GET request-info tests")

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

    return GraphServiceClient(request_adapter=StrictTransportAdapter())


@pytest.mark.parametrize(
    ("case", "expected_path", "expected_query"),
    [
        ("outlook_messages", "/users/user/messages", {"%24filter": "contains(subject,'hello')", "%24top": 7, "%24select": ["id", "subject", "from", "receivedDateTime", "hasAttachments", "webLink"]}),
        ("calendar_events", "/users/user/calendar/events", {"%24filter": "contains(subject,'hello')", "%24top": 7}),
        ("todo_lists", "/users/user/todo/lists", {"%24top": 7}),
        ("todo_tasks", "/users/user/todo/lists/list/tasks", {"%24filter": "contains(title,'hello')", "%24top": 7}),
        ("drive_children", "/drives/drive/items/root/children", {"%24top": 7}),
        ("teams_joined", "/users/user/joinedTeams", {}),
        ("teams_channels", "/teams/team/channels", {"%24select": ["id", "displayName"]}),
        ("planner_plans", "/planner/plans", {"%24filter": "owner eq 'group'", "%24top": 7}),
        ("planner_buckets", "/planner/plans/plan/buckets", {"%24top": 7}),
        ("planner_tasks", "/planner/plans/plan/tasks", {"%24top": 7}),
    ],
)
def test_real_generated_collection_builder_accepts_typed_configuration(
    graph_client, case, expected_path, expected_query
):
    from microsoft365.sdk_contract import build_collection_request_information

    request = build_collection_request_information(
        graph_client,
        case,
        user_id="user",
        todo_list_id="list",
        drive_id="drive",
        drive_item_id="root",
        team_id="team",
        plan_id="plan",
        owner_id="group",
        query="hello",
        limit=7,
    )

    assert request.http_method.value == "GET"
    assert urlsplit(request.url).path == expected_path
    assert request.query_parameters == expected_query
    assert type(request).__name__ == "RequestInformation"


def test_real_builder_rejects_old_untyped_request_configuration(graph_client):
    from types import SimpleNamespace

    builder = graph_client.users.by_user_id("user").messages
    with pytest.raises(AttributeError, match="headers"):
        builder.to_get_request_information(
            SimpleNamespace(query_parameters={"$top": 7})
        )


def test_real_collection_response_normalizes_value_and_continuation():
    from msgraph.generated.models.message import Message
    from msgraph.generated.models.message_collection_response import MessageCollectionResponse
    from microsoft365.results import normalize_result

    response = MessageCollectionResponse(
        value=[Message(id="m1", subject="Subject")],
        odata_count=1,
        odata_next_link="https://graph.microsoft.com/v1.0/users/u/messages?$skiptoken=opaque",
    )

    normalized = normalize_result(response)

    assert normalized["value"] == [{"id": "m1", "subject": "Subject", "additional_data": {}}]
    assert normalized["@odata.count"] == 1
    assert normalized["@odata.nextLink"].endswith("$skiptoken=opaque")


def test_binary_result_bypasses_generic_string_truncation():
    import base64
    from microsoft365.results import BinaryResult, normalize_result

    payload = b"x" * 8192
    encoded = base64.b64encode(payload).decode("ascii")
    normalized = normalize_result(
        BinaryResult(content_base64=encoded, content_type="application/octet-stream", size=len(payload))
    )

    assert normalized["content_base64"] == encoded
    assert base64.b64decode(normalized["content_base64"], validate=True) == payload
    assert "TRUNCATED" not in normalized["content_base64"]
