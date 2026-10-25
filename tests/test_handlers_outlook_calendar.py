"""WP6 — Outlook and Calendar handlers, driven offline through the execution seam.

What this module proves
-----------------------
* the **self-populating handler registry**: every module dropped into
  ``microsoft365/handlers/`` is imported and its handlers registered without editing
  ``registration.py``, a module that cannot be imported fails loudly, and an unknown
  operation or a duplicate registration is refused;
* every handler builds its request from the **generated SDK builders** and is pinned
  byte-for-byte against the verified request-information contract of the same operation
  (``microsoft365/sdk_contract.py``, WP1): same method, URL, typed query, serialized body
  and ``If-Match``;
* the read handlers go through ``microsoft365/paging.py`` (validated next links, bounded
  traversal, truthful continuation) and the write handlers through the execution seam;
* an argument the verified contract cannot express, and a PATCH field that does not change
  the serialized request, are **refused** — never silently dropped;
* nothing reaches a network, and no raw SDK/Graph/Azure text reaches a tool payload.

Transport
---------
``HandlerGraphAdapter`` extends ``StrictTransportAdapter`` from ``tests/test_sdk_contract.py``:
every ``send_*`` the test did not supply a response for still raises
``AssertionError("network transport must not run")``. The writer factory is the real Kiota
JSON factory (module ``kiota_serialization_json.json_serialization_writer_factory``), so a
serialized body is assertable while the transport stays unreachable. Every request the
plugin builds is recorded, which is what makes the handler-versus-contract pin possible.

Nothing here uses ``SimpleNamespace``, ``MagicMock`` or a permissive ``__getattr__``.
"""
from __future__ import annotations

import importlib
import inspect
import json
import sys
import types
from pathlib import Path
from typing import Any
from urllib.parse import parse_qsl, unquote, urlsplit

import pytest
from kiota_abstractions.headers_collection import HeadersCollection
from kiota_abstractions.request_adapter import RequestAdapter
from kiota_serialization_json.json_serialization_writer_factory import JsonSerializationWriterFactory
from msgraph import GraphServiceClient
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.event import Event
from msgraph.generated.models.event_collection_response import EventCollectionResponse
from msgraph.generated.models.message import Message
from msgraph.generated.models.message_collection_response import MessageCollectionResponse

from microsoft365 import handlers as handler_package
from microsoft365.contract import OPERATIONS, OPERATION_REGISTRY, WRITE_OPERATIONS, Settings
from microsoft365.errors import CATEGORIES, GraphError
from microsoft365.handlers import HandlerContext, HandlerRegistrationError, handler, load_handlers
from microsoft365.sdk_contract import (
    build_collection_request_information,
    build_item_request_information,
    build_write_request_information,
)
from tests.test_sdk_contract import StrictTransportAdapter

REPO_ROOT = Path(__file__).resolve().parent.parent
HANDLER_MODULES = ("__init__.py", "outlook.py", "calendar.py")

OUTLOOK_SEARCH = "outlook.search"
OUTLOOK_READ = "outlook.read"
OUTLOOK_CREATE_DRAFT = "outlook.create_draft"
OUTLOOK_SEND = "outlook.send"
CALENDAR_SEARCH = "calendar.search"
CALENDAR_CREATE_EVENTS = "calendar.create_events"
CALENDAR_UPDATE_EVENTS = "calendar.update_events"

#: The eight operations WP6 owns.
WP6_HANDLERS = frozenset(
    {
        "outlook.list_attachments", "outlook.download_attachment",
        OUTLOOK_SEARCH,
        OUTLOOK_READ,
        OUTLOOK_CREATE_DRAFT,
        OUTLOOK_SEND,
        CALENDAR_SEARCH,
        CALENDAR_CREATE_EVENTS,
        CALENDAR_UPDATE_EVENTS,
    }
)

#: The four writes. They stay non-executable until the generic host approval fix (CORE-1 /
#: CORE-2) is available, so none of them may appear in a model-facing action enum.
WP6_WRITES = frozenset(
    {OUTLOOK_CREATE_DRAFT, OUTLOOK_SEND, CALENDAR_CREATE_EVENTS, CALENDAR_UPDATE_EVENTS}
)

USER = "user"
MESSAGE = "message"
EVENT = "event"
ETAG = '"etag-1"'
TIME_ZONE = "America/Sao_Paulo"
START = "2026-01-05T09:00:00"
END = "2026-01-05T10:00:00"


def settings() -> Settings:
    return Settings.from_mapping(
        {
            "tenant_id": "tenant",
            "client_id": "client",
            "user_id": USER,
            "capabilities": {service: True for service in OPERATIONS},
        }
    )


def request_identity(request_information: Any) -> tuple:
    """Canonical identity of a built request: method, path and typed query, sorted.

    The generated builders keep a typed configuration's query in ``query_parameters`` (with
    percent-encoded OData names) and a followed next link's query inside the raw URL, so the
    two sources are unioned -- never concatenated -- exactly as the paginator does. A
    ``/v1.0`` prefix is stripped so a first page and a followed link of the same collection
    are identified the same way.
    """
    parts = urlsplit(request_information.url)
    path = parts.path[len("/v1.0"):] if parts.path.startswith("/v1.0/") else parts.path
    query = list(parse_qsl(parts.query, keep_blank_values=True))
    for name, value in (getattr(request_information, "query_parameters", None) or {}).items():
        query.append((unquote(str(name)), tuple(value) if isinstance(value, list) else str(value)))
    return (
        request_information.http_method.value,
        path,
        tuple(sorted(set(query), key=repr)),
    )


class HandlerGraphAdapter(StrictTransportAdapter):
    """Strict transport double that records every request and answers only named requests.

    A supplied response may be a real generated model (returned as the SDK's typed response)
    or an exception instance (raised, so the canonical taxonomy has a real upstream failure to
    classify). Anything the test did not supply still goes to ``StrictTransportAdapter``,
    which raises: no code path in this module can reach a network.
    """

    def __init__(self, responses: dict | None = None):
        super().__init__()
        self._serialization = JsonSerializationWriterFactory()
        self._responses = dict(responses or {})
        self.requests: list = []

    def get_serialization_writer_factory(self):
        return self._serialization

    def _answer(self, request_information: Any) -> Any:
        self.requests.append(request_information)
        identity = request_identity(request_information)
        # A response may be keyed by the full request identity (a followed page: method, path
        # and the query the service put in the link) or by method and path alone (the first
        # page of a traversal). The specific key wins, so a rewritten follow-up request cannot
        # collect the first page a second time.
        for key in (identity, identity[:2]):
            if key in self._responses:
                response = self._responses[key]
                break
        else:
            raise AssertionError("network transport must not run")
        if isinstance(response, BaseException):
            raise response
        return response

    async def send_async(self, request_information, parsable_factory=None, error_map=None, **kwargs):
        return self._answer(request_information)

    async def send_no_response_content_async(self, request_information, error_map=None, **kwargs):
        return self._answer(request_information)

    async def send_collection_async(self, request_information, parsable_factory=None, error_map=None, **kwargs):
        return self._answer(request_information)

    async def send_primitive_async(self, request_information, *args, **kwargs):
        return self._answer(request_information)


def graph_client(adapter: RequestAdapter) -> GraphServiceClient:
    return GraphServiceClient(request_adapter=adapter)


def open_context(adapter: RequestAdapter) -> HandlerContext:
    """A handler context whose client is built from the injected strict double."""
    client = graph_client(adapter)
    return HandlerContext(settings=settings(), client_factory=lambda: client)


def invocation(key: str):
    """The registered handler of an operation, straight out of the registry."""
    record = load_handlers().get(key)
    assert record is not None, f"{key} is not registered"
    return record.function


def run(key: str, arguments: dict, adapter: RequestAdapter) -> dict:
    return invocation(key)(arguments, open_context(adapter))


def message_page(items: list, *, next_link: str | None = None, count: int | None = None):
    return MessageCollectionResponse(
        value=[Message(id=item_id) for item_id in items],
        odata_count=count,
        odata_next_link=next_link,
    )


def event_page(items: list, *, next_link: str | None = None):
    return EventCollectionResponse(
        value=[Event(id=item_id, subject=f"Event {item_id}") for item_id in items],
        odata_next_link=next_link,
    )


def search_query(request_information: Any) -> dict:
    """The typed query of a built request with OData names unfolded."""
    query = {}
    for name, value in (getattr(request_information, "query_parameters", None) or {}).items():
        query[unquote(str(name)).lstrip("$")] = value
    for name, value in parse_qsl(urlsplit(request_information.url).query, keep_blank_values=True):
        query.setdefault(name.lstrip("$"), value)
    return query


def serialized(request_information: Any) -> dict:
    assert request_information.content, "expected a serialized request body"
    return json.loads(request_information.content.decode("utf-8"))


def keys_in(payload: Any) -> set:
    """Every JSON key of a nested payload, so a sub-field of a nested object is findable."""
    found: set = set()
    if isinstance(payload, dict):
        for key, value in payload.items():
            found.add(key)
            found |= keys_in(value)
    elif isinstance(payload, list):
        for item in payload:
            found |= keys_in(item)
    return found


# ======================================================================================
# The registry mechanism every later handler work package depends on
# ======================================================================================


def test_load_handlers_imports_every_module_the_package_exposes():
    """No per-service import list exists: discovery is what populates the mapping."""
    import pkgutil

    names = {module.name for module in pkgutil.iter_modules(handler_package.__path__)}

    assert names == {"outlook", "outlook_attachments", "calendar", "files", "teams", "todo", "planner"}
    load_handlers()
    for name in names:
        assert f"{handler_package.__name__}.{name}" in sys.modules


def test_the_registry_holds_every_wp6_operation_with_its_write_classification():
    loaded = load_handlers()

    assert {key for key in loaded if key.startswith(("outlook.", "calendar."))} == set(WP6_HANDLERS)
    for key, record in loaded.items():
        assert record.key == key
        assert record.key == f"{record.service}.{record.operation}"
        assert record.write is (record.operation in WRITE_OPERATIONS), key
        assert callable(record.function), key


def test_a_new_service_module_is_discovered_without_editing_registration(monkeypatch):
    """The WP7/WP8/WP9 seam: drop a module in, register, nothing else changes."""
    import pkgutil

    module_name = f"{handler_package.__name__}.sample_service"
    registry = dict(handler_package._REGISTRY)
    registry.pop("teams.list_teams", None)
    monkeypatch.setattr(handler_package, "_REGISTRY", registry)
    module = types.ModuleType(module_name)
    exec(
        compile(
            "from microsoft365.handlers import handler\n"
            "\n"
            "@handler('teams', 'list_teams')\n"
            "def list_teams(arguments, context):\n"
            "    return {'operation': 'teams.list_teams'}\n",
            module_name,
            "exec",
        ),
        module.__dict__,
    )
    real_iter_modules = pkgutil.iter_modules
    package_path = list(handler_package.__path__)

    def discovering_iter_modules(path=None, prefix=""):
        found = list(real_iter_modules(path, prefix))
        if path is not None and list(path) == package_path:
            found.append(pkgutil.ModuleInfo(None, "sample_service", False))
        return iter(found)

    # The injected registration lands in the copied registry, so this test cannot leak
    # a handler into the mapping the other tests assert on.
    monkeypatch.setitem(sys.modules, module_name, module)
    monkeypatch.setattr(handler_package.pkgutil, "iter_modules", discovering_iter_modules)

    loaded = load_handlers()

    assert loaded["teams.list_teams"].service == "teams"
    assert loaded["teams.list_teams"].write is False
    assert loaded["outlook.search"].operation == "search"


def test_a_handler_module_that_cannot_be_imported_fails_loudly(monkeypatch):
    """A malformed module is never skipped: the discovery either imports it or raises."""
    import pkgutil

    monkeypatch.setattr(
        handler_package.pkgutil,
        "iter_modules",
        lambda path=None, prefix="": iter([pkgutil.ModuleInfo(None, "absent_handler_module", False)]),
    )

    with pytest.raises(ModuleNotFoundError):
        load_handlers()


def test_registering_an_unknown_operation_is_refused():
    with pytest.raises(HandlerRegistrationError) as caught:
        handler("outlook", "delete")

    assert "outlook.delete" in str(caught.value)


@pytest.mark.parametrize("service, operation", [("outlook", "search"), ("calendar", "search")])
def test_registering_a_duplicate_operation_is_refused(service, operation):
    with pytest.raises(HandlerRegistrationError) as caught:
        handler(service, operation)(lambda arguments, context: {})

    assert f"{service}.{operation}" in str(caught.value)


def test_every_registered_handler_is_synchronous():
    """Handlers cross the sync/async boundary only through the execution seam."""
    for key, record in load_handlers().items():
        assert not inspect.iscoroutinefunction(record.function), key


def test_handler_modules_never_hold_a_secret_credential_or_client():
    """The credential surface is the host's secret scope and ``client.py``, never a handler."""
    for name in HANDLER_MODULES:
        source = (REPO_ROOT / "microsoft365" / "handlers" / name).read_text(encoding="utf-8")
        for forbidden in (
            "get_secret",
            "secret_scope",
            "ClientSecretCredential",
            "GraphServiceClient",
            "azure.identity",
        ):
            assert forbidden not in source, (name, forbidden)


def test_handler_modules_never_assemble_a_url_or_import_a_transport():
    """Endpoints come from the generated builders; no URL is concatenated by hand."""
    for name in HANDLER_MODULES:
        source = (REPO_ROOT / "microsoft365" / "handlers" / name).read_text(encoding="utf-8")
        for forbidden in ("https://", "http://", "urlsplit", "urljoin", ".url =", "RequestsClient"):
            assert forbidden not in source, (name, forbidden)


def test_handler_modules_use_the_seam_and_the_paginator():
    package = (REPO_ROOT / "microsoft365" / "handlers" / "__init__.py").read_text(encoding="utf-8")
    outlook = (REPO_ROOT / "microsoft365" / "handlers" / "outlook.py").read_text(encoding="utf-8")
    calendar = (REPO_ROOT / "microsoft365" / "handlers" / "calendar.py").read_text(encoding="utf-8")

    assert "resolve_request_adapter" in package
    assert "execute_request" in package or "execute_request" in outlook or "execute_request" in calendar
    assert "paginate(" in outlook and "paginate(" in calendar


# ======================================================================================
# Reads — pinned to the verified request contracts, bounded by the paginator
# ======================================================================================


def test_outlook_search_builds_the_verified_request_and_reports_the_page():
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/messages"): message_page(["m1"], count=1)}
    )

    payload = run(OUTLOOK_SEARCH, {"action": "search", "user_id": USER, "query": "hello", "top": 7}, adapter)

    client = graph_client(adapter)
    # A search that named no select of its own carries the compact candidate metadata; the pin
    # is the shared ``outlook_messages`` request contract, which now declares that select as an
    # independent literal (not the handler constant), so a drift of the default select fails.
    verified = build_collection_request_information(
        client, "outlook_messages", user_id=USER, query="hello", limit=7
    )
    assert request_identity(adapter.requests[0]) == request_identity(verified)
    # the default select is the compact candidate metadata, asserted literally
    assert search_query(adapter.requests[0])["select"] == [
        "id", "subject", "from", "receivedDateTime", "hasAttachments", "webLink",
    ]
    assert payload["operation"] == OUTLOOK_SEARCH
    assert payload["status"] == "succeeded"
    # only the field ``m1`` carries is projected; the absent metadata is not invented
    assert payload["result"]["value"] == [{"id": "m1"}]
    assert payload["continuation"] == {
        "truncated": False,
        "next_link_present": False,
        "pages_fetched": 1,
        "items_fetched": 1,
        "stop_reason": "complete",
    }


def test_outlook_search_passes_a_caller_filter_and_select_as_typed_query_parameters():
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/messages"): message_page([])}
    )

    run(
        OUTLOOK_SEARCH,
        {
            "action": "search",
            "user_id": USER,
            "filter": "isRead eq false",
            "select": ["id", "subject"],
            "top": 5,
        },
        adapter,
    )

    query = search_query(adapter.requests[0])
    assert query == {"filter": "isRead eq false", "select": ["id", "subject"], "top": 5}


def test_outlook_search_follows_a_validated_next_link_and_stops_at_the_caller_limit():
    first = "https://graph.microsoft.com/v1.0/users/user/messages?$skiptoken=page-2"
    adapter = HandlerGraphAdapter(
        {
            ("GET", f"/users/{USER}/messages"): message_page(["m1", "m2"], next_link=first, count=4),
            ("GET", f"/users/{USER}/messages", (("$skiptoken", "page-2"),)): message_page(["m3", "m4"]),
        }
    )

    payload = run(OUTLOOK_SEARCH, {"action": "search", "user_id": USER, "query": "hello", "top": 4}, adapter)

    assert [item["id"] for item in payload["result"]["value"]] == ["m1", "m2", "m3", "m4"]
    assert payload["continuation"] == {
        "truncated": False,
        "next_link_present": False,
        "pages_fetched": 2,
        "items_fetched": 4,
        "stop_reason": "complete",
    }
    assert len(adapter.requests) == 2


def test_outlook_search_reports_a_truthful_continuation_when_the_budget_cuts_a_page():
    first = "https://graph.microsoft.com/v1.0/users/user/messages?$skiptoken=page-2"
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/messages"): message_page(["m1", "m2"], next_link=first, count=4)}
    )

    payload = run(OUTLOOK_SEARCH, {"action": "search", "user_id": USER, "query": "hello", "top": 2}, adapter)

    assert payload["continuation"]["stop_reason"] == "caller_limit"
    assert payload["continuation"]["truncated"] is True
    assert payload["continuation"]["items_fetched"] == 2
    assert payload["result"]["@odata.nextLink"].endswith("$skiptoken=page-2")
    assert len(adapter.requests) == 1


def test_outlook_search_refuses_query_and_filter_together_before_any_client_work(monkeypatch):
    calls: list = []
    adapter = HandlerGraphAdapter()

    def factory():
        calls.append("client")
        return graph_client(adapter)

    context = HandlerContext(settings=settings(), client_factory=factory)
    arguments = {"action": "search", "user_id": USER, "query": "hello", "filter": "isRead eq false"}

    with pytest.raises(GraphError) as caught:
        invocation(OUTLOOK_SEARCH)(arguments, context)

    assert caught.value.category == "validation_error"
    assert "query" in caught.value.message and "filter" in caught.value.message
    assert calls == []
    assert adapter.requests == []


def test_outlook_search_never_reaches_the_network_with_the_strict_double():
    adapter = HandlerGraphAdapter()

    with pytest.raises(GraphError) as caught:
        run(OUTLOOK_SEARCH, {"action": "search", "user_id": USER, "query": "hello"}, adapter)

    assert caught.value.category == "transport_error"
    assert "network transport must not run" not in json.dumps(caught.value.to_result())
    assert len(adapter.requests) == 3


def test_outlook_read_is_pinned_to_the_verified_item_contract():
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/messages/{MESSAGE}"): Message(id=MESSAGE, subject="Weekly")}
    )

    payload = run(
        OUTLOOK_READ,
        {"action": "read", "user_id": USER, "message_id": MESSAGE, "select": ["id", "subject"]},
        adapter,
    )

    client = graph_client(adapter)
    verified = build_item_request_information(
        client, "outlook_read", user_id=USER, message_id=MESSAGE, select=["id", "subject"]
    )
    assert request_identity(adapter.requests[0]) == request_identity(verified)
    assert search_query(adapter.requests[0])["select"] == ["id", "subject"]
    assert payload["operation"] == OUTLOOK_READ
    assert payload["result"] == {
        "id": MESSAGE,
        "subject": "Weekly",
        "additional_data": {},
    }


def test_outlook_read_sends_no_select_when_the_caller_did_not_ask_for_one():
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/messages/{MESSAGE}"): Message(id=MESSAGE)}
    )

    run(OUTLOOK_READ, {"action": "read", "user_id": USER, "message_id": MESSAGE}, adapter)

    client = graph_client(adapter)
    verified = build_item_request_information(
        client, "outlook_read", user_id=USER, message_id=MESSAGE
    )
    assert request_identity(adapter.requests[0]) == request_identity(verified)
    assert "select" not in search_query(adapter.requests[0])


def test_calendar_search_matches_the_verified_calendar_contract():
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/calendar/events"): event_page(["e1"])}
    )

    payload = run(
        CALENDAR_SEARCH,
        {"action": "search", "user_id": USER, "filter": "contains(subject,'hello')", "top": 7},
        adapter,
    )

    client = graph_client(adapter)
    verified = build_collection_request_information(
        client, "calendar_events", user_id=USER, query="hello", limit=7
    )
    assert request_identity(adapter.requests[0]) == request_identity(verified)
    assert payload["result"]["value"][0]["id"] == "e1"
    assert payload["continuation"]["stop_reason"] == "complete"


def test_calendar_search_sends_only_the_typed_top_when_no_filter_is_given():
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/calendar/events"): event_page([])}
    )

    run(CALENDAR_SEARCH, {"action": "search", "user_id": USER, "top": 3}, adapter)

    assert search_query(adapter.requests[0]) == {"top": 3}


def test_calendar_search_refuses_a_date_range_the_verified_contract_cannot_express():
    adapter = HandlerGraphAdapter()
    context = HandlerContext(settings=settings(), client_factory=lambda: graph_client(adapter))
    arguments = {
        "action": "search",
        "user_id": USER,
        "start_date_time": START,
        "end_date_time": END,
        "time_zone": TIME_ZONE,
    }

    with pytest.raises(GraphError) as caught:
        invocation(CALENDAR_SEARCH)(arguments, context)

    assert caught.value.category == "validation_error"
    assert "start_date_time" in caught.value.message
    assert adapter.requests == []


def test_a_read_failure_is_reported_through_the_sanitized_taxonomy():
    from kiota_abstractions.api_error import APIError

    failure = APIError("raw graph text must never surface", 404, {})
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/messages/{MESSAGE}"): failure}
    )

    with pytest.raises(GraphError) as caught:
        run(OUTLOOK_READ, {"action": "read", "user_id": USER, "message_id": MESSAGE}, adapter)

    assert caught.value.category == "not_found"
    assert caught.value.category in CATEGORIES
    assert "raw graph text" not in json.dumps(caught.value.to_result())


def test_calendar_search_bounds_the_traversal_by_the_caller_limit():
    first = "https://graph.microsoft.com/v1.0/users/user/calendar/events?$skiptoken=page-2"
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER}/calendar/events"): event_page(["e1", "e2"], next_link=first)}
    )

    payload = run(CALENDAR_SEARCH, {"action": "search", "user_id": USER, "top": 2}, adapter)

    assert payload["continuation"]["stop_reason"] == "caller_limit"
    assert payload["continuation"]["next_link_present"] is True
    assert len(adapter.requests) == 1


# ======================================================================================
# Writes — implemented and offline-testable through the seam, deliberately not executable
# ======================================================================================


def test_outlook_create_draft_is_pinned_to_the_verified_publish_request():
    adapter = HandlerGraphAdapter(
        {("POST", f"/users/{USER}/messages"): Message(id="draft-1", subject="Weekly report")}
    )
    arguments = {
        "action": "create_draft",
        "user_id": USER,
        "subject": "Weekly report",
        "body": "Body text",
        "to_recipients": ["a@example.com", "b@example.com"],
    }

    payload = run(OUTLOOK_CREATE_DRAFT, arguments, adapter)

    client = graph_client(adapter)
    verified = build_write_request_information(
        client,
        "outlook_create_draft",
        user_id=USER,
        subject="Weekly report",
        content="Body text",
        recipients=["a@example.com", "b@example.com"],
    )
    executed = adapter.requests[0]
    assert executed.http_method.value == "POST"
    assert executed.url == verified.url
    # The body is the verified contract's own body, byte for byte.
    assert executed.content == verified.content
    assert serialized(executed)["toRecipients"] == [
        {"emailAddress": {"address": "a@example.com"}},
        {"emailAddress": {"address": "b@example.com"}},
    ]
    assert payload["operation"] == OUTLOOK_CREATE_DRAFT
    assert payload["result"]["id"] == "draft-1"


@pytest.mark.parametrize(
    "override, missing",
    [
        ({"body": None}, "body"),
        ({"body": "   "}, "body"),
        ({"to_recipients": None}, "to_recipients"),
        ({"to_recipients": []}, "to_recipients"),
    ],
)
def test_outlook_create_draft_requires_a_body_and_recipients_before_any_client_work(
    override, missing
):
    adapter = HandlerGraphAdapter()
    calls: list = []

    def factory():
        calls.append("client")
        return graph_client(adapter)

    arguments = {
        "action": "create_draft",
        "user_id": USER,
        "subject": "Weekly report",
        "body": "Body text",
        "to_recipients": ["a@example.com"],
    }
    arguments.update(override)

    with pytest.raises(GraphError) as caught:
        invocation(OUTLOOK_CREATE_DRAFT)(
            arguments, HandlerContext(settings=settings(), client_factory=factory)
        )

    assert caught.value.category == "validation_error"
    assert missing in caught.value.message
    assert calls == []
    assert adapter.requests == []


def test_outlook_create_draft_refuses_a_message_id_as_its_content():
    """A message id is the alternate *send* path; it cannot stand in for a draft body.

    The payload below is otherwise complete (subject, body and recipients), so refusing it is
    the message-id rule and nothing else: without that rule the request would be sent.
    """
    adapter = HandlerGraphAdapter({("POST", f"/users/{USER}/messages"): Message(id="draft-3")})
    arguments = {
        "action": "create_draft",
        "user_id": USER,
        "message_id": MESSAGE,
        "subject": "Weekly report",
        "body": "Body text",
        "to_recipients": ["a@example.com"],
    }

    with pytest.raises(GraphError) as caught:
        run(OUTLOOK_CREATE_DRAFT, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "message_id" in caught.value.message
    assert adapter.requests == []


def test_outlook_create_draft_carries_declared_cc_and_bcc_recipients():
    adapter = HandlerGraphAdapter({("POST", f"/users/{USER}/messages"): Message(id="draft-2")})
    arguments = {
        "action": "create_draft",
        "user_id": USER,
        "subject": "Weekly report",
        "body": "Body text",
        "to_recipients": ["a@example.com"],
        "cc_recipients": [{"address": "c@example.com", "name": "C"}],
        "bcc_recipients": ["d@example.com"],
    }

    run(OUTLOOK_CREATE_DRAFT, arguments, adapter)

    body = serialized(adapter.requests[0])
    assert body["ccRecipients"] == [{"emailAddress": {"address": "c@example.com", "name": "C"}}]
    assert body["bccRecipients"] == [{"emailAddress": {"address": "d@example.com"}}]


def test_outlook_send_is_pinned_to_the_verified_send_request():
    adapter = HandlerGraphAdapter({("POST", f"/users/{USER}/sendMail"): None})
    arguments = {
        "action": "send",
        "user_id": USER,
        "subject": "Weekly report",
        "body": "Body text",
        "to_recipients": ["a@example.com"],
        "save_to_sent_items": True,
    }

    payload = run(OUTLOOK_SEND, arguments, adapter)

    client = graph_client(adapter)
    verified = build_write_request_information(
        client,
        "outlook_send",
        user_id=USER,
        subject="Weekly report",
        content="Body text",
        recipients=["a@example.com"],
        save_to_sent_items=True,
    )
    executed = adapter.requests[0]
    assert executed.http_method.value == "POST"
    assert executed.url == verified.url
    assert executed.content == verified.content
    # The generated ``SendMailPostRequestBody`` serializes with its PascalCase Graph names;
    # WP1's contract pins exactly that shape (tests/test_sdk_contract_writes.py) and the
    # handler reproduces it byte for byte.
    body = serialized(executed)
    assert set(body) == {"Message", "SaveToSentItems"}
    assert body["SaveToSentItems"] is True
    assert payload["result"] == {"accepted": True, "form": "composed"}


def test_outlook_send_does_not_assert_a_save_to_sent_items_default():
    adapter = HandlerGraphAdapter({("POST", f"/users/{USER}/sendMail"): None})
    arguments = {
        "action": "send",
        "user_id": USER,
        "subject": "Weekly report",
        "to_recipients": ["a@example.com"],
    }

    run(OUTLOOK_SEND, arguments, adapter)

    body = serialized(adapter.requests[0])
    assert "SaveToSentItems" not in body
    assert "body" not in body["Message"]


def test_outlook_send_reports_the_draft_path_of_an_existing_message():
    adapter = HandlerGraphAdapter({("POST", f"/users/{USER}/messages/{MESSAGE}/send"): None})

    payload = run(OUTLOOK_SEND, {"action": "send", "user_id": USER, "message_id": MESSAGE}, adapter)

    client = graph_client(adapter)
    verified = build_write_request_information(
        client, "outlook_send_existing", user_id=USER, message_id=MESSAGE
    )
    executed = adapter.requests[0]
    assert executed.http_method.value == "POST"
    assert executed.url == verified.url
    assert executed.content == verified.content
    assert payload["result"] == {"accepted": True, "form": "existing"}


def test_outlook_send_refuses_mixing_a_draft_id_with_a_composed_message():
    adapter = HandlerGraphAdapter()
    arguments = {
        "action": "send",
        "user_id": USER,
        "message_id": MESSAGE,
        "subject": "Weekly report",
        "to_recipients": ["a@example.com"],
    }

    with pytest.raises(GraphError) as caught:
        run(OUTLOOK_SEND, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "message_id" in caught.value.message
    assert adapter.requests == []


def test_outlook_send_refuses_a_call_that_sends_nothing():
    adapter = HandlerGraphAdapter()

    with pytest.raises(GraphError) as caught:
        run(OUTLOOK_SEND, {"action": "send", "user_id": USER}, adapter)

    assert caught.value.category == "validation_error"
    assert adapter.requests == []


def test_calendar_create_events_is_pinned_to_the_verified_event_request():
    adapter = HandlerGraphAdapter(
        {("POST", f"/users/{USER}/calendar/events"): Event(id="event-1", subject="Planning")}
    )
    arguments = {
        "action": "create_events",
        "user_id": USER,
        "subject": "Planning",
        "start_date_time": START,
        "end_date_time": END,
        "time_zone": TIME_ZONE,
    }

    payload = run(CALENDAR_CREATE_EVENTS, arguments, adapter)

    client = graph_client(adapter)
    verified = build_write_request_information(
        client,
        "calendar_create_events",
        user_id=USER,
        subject="Planning",
        start=DateTimeTimeZone(date_time=START, time_zone=TIME_ZONE),
        end=DateTimeTimeZone(date_time=END, time_zone=TIME_ZONE),
    )
    executed = adapter.requests[0]
    assert executed.http_method.value == "POST"
    assert executed.url == verified.url
    assert executed.content == verified.content
    assert serialized(executed)["start"] == {"dateTime": START, "timeZone": TIME_ZONE}
    assert payload["result"]["id"] == "event-1"


def test_calendar_create_events_carries_every_declared_optional_field():
    adapter = HandlerGraphAdapter({("POST", f"/users/{USER}/calendar/events"): Event(id="event-2")})
    arguments = {
        "action": "create_events",
        "user_id": USER,
        "subject": "Planning",
        "start_date_time": "2026-01-05T09:00:00-03:00",
        "end_date_time": "2026-01-05T10:00:00-03:00",
        "body": "Agenda",
        "location": "Room 1",
        "attendees": ["a@example.com", {"address": "b@example.com", "name": "B"}],
        "is_all_day": False,
    }

    run(CALENDAR_CREATE_EVENTS, arguments, adapter)

    body = serialized(adapter.requests[0])
    assert keys_in(body) >= {"subject", "start", "end", "body", "location", "attendees", "isAllDay"}
    assert body["location"] == {"displayName": "Room 1"}
    assert body["isAllDay"] is False
    assert body["attendees"] == [
        {"emailAddress": {"address": "a@example.com"}, "@odata.type": "#microsoft.graph.attendee"},
        {
            "emailAddress": {"address": "b@example.com", "name": "B"},
            "@odata.type": "#microsoft.graph.attendee",
        },
    ]


def test_calendar_create_events_refuses_an_offsetless_date_without_a_time_zone():
    adapter = HandlerGraphAdapter()
    arguments = {
        "action": "create_events",
        "user_id": USER,
        "subject": "Planning",
        "start_date_time": START,
        "end_date_time": END,
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_CREATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "time_zone" in caught.value.message
    assert adapter.requests == []


def test_calendar_update_events_is_pinned_to_the_verified_patch_request():
    adapter = HandlerGraphAdapter(
        {("PATCH", f"/users/{USER}/calendar/events/{EVENT}"): Event(id=EVENT, subject="Renamed")}
    )
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": {"subject": "Renamed"},
    }

    payload = run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    client = graph_client(adapter)
    verified = build_write_request_information(
        client,
        "calendar_update_events",
        user_id=USER,
        event_id=EVENT,
        etag=ETAG,
        update_fields={"subject": "Renamed"},
    )
    executed = adapter.requests[0]
    assert executed.http_method.value == "PATCH"
    assert executed.url == verified.url
    assert executed.content == verified.content
    assert list(executed.headers.get("If-Match")) == [ETAG]
    assert serialized(executed) == {"@odata.type": "#microsoft.graph.event", "subject": "Renamed"}
    assert payload["result"]["subject"] == "Renamed"


#: One payload per declared update field, and the wire key that field must produce.
UPDATE_FIELD_CASES = {
    "subject": ({"subject": "Renamed"}, "subject"),
    "body": ({"body": "New agenda"}, "body"),
    "start_date_time": ({"start_date_time": START, "time_zone": TIME_ZONE}, "start"),
    "end_date_time": ({"end_date_time": END, "time_zone": TIME_ZONE}, "end"),
    "time_zone": ({"start_date_time": START, "time_zone": TIME_ZONE}, "timeZone"),
    "location": ({"location": "Room 2"}, "location"),
    "attendees": ({"attendees": ["a@example.com"]}, "attendees"),
    "is_all_day": ({"is_all_day": True}, "isAllDay"),
}


def test_the_update_field_cases_cover_exactly_the_declared_allowlist():
    from microsoft365.validation import EVENT_PATCH_FIELDS

    assert set(UPDATE_FIELD_CASES) == set(EVENT_PATCH_FIELDS)


@pytest.mark.parametrize("name", sorted(UPDATE_FIELD_CASES))
def test_every_accepted_update_field_changes_the_serialized_request(name):
    """A field the endpoint would ignore is refused, never sent as a no-op."""
    fields, wire_key = UPDATE_FIELD_CASES[name]
    adapter = HandlerGraphAdapter(
        {("PATCH", f"/users/{USER}/calendar/events/{EVENT}"): Event(id=EVENT)}
    )
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": dict(fields),
    }

    run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    assert wire_key in keys_in(serialized(adapter.requests[0])), name


def test_calendar_update_events_refuses_a_field_the_endpoint_ignores():
    adapter = HandlerGraphAdapter()
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": {"id": "other"},
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "id" in caught.value.message
    assert adapter.requests == []


def test_calendar_update_events_refuses_an_empty_patch():
    adapter = HandlerGraphAdapter()
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": {},
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert adapter.requests == []


def test_calendar_update_events_refuses_a_time_zone_that_changes_nothing():
    """``time_zone`` alone is a sub-field of a date: on its own it changes no request."""
    adapter = HandlerGraphAdapter()
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": {"time_zone": TIME_ZONE},
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "time_zone only applies to a start or an end date and time" in caught.value.message
    assert adapter.requests == []


def test_calendar_update_events_requires_an_etag():
    adapter = HandlerGraphAdapter()
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "fields": {"subject": "Renamed"},
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "etag" in caught.value.message
    assert adapter.requests == []


WRITE_CALLS = [
    (
        OUTLOOK_CREATE_DRAFT,
        {
            "action": "create_draft",
            "user_id": USER,
            "subject": "Weekly report",
            "body": "Body text",
            "to_recipients": ["a@example.com"],
        },
        Message(id="draft-1"),
        "POST",
        f"/users/{USER}/messages",
    ),
    (
        OUTLOOK_SEND,
        {
            "action": "send",
            "user_id": USER,
            "subject": "Weekly report",
            "body": "Body text",
            "to_recipients": ["a@example.com"],
        },
        None,
        "POST",
        f"/users/{USER}/sendMail",
    ),
    (
        CALENDAR_CREATE_EVENTS,
        {
            "action": "create_events",
            "user_id": USER,
            "subject": "Planning",
            "start_date_time": START,
            "end_date_time": END,
            "time_zone": TIME_ZONE,
        },
        Event(id="event-1"),
        "POST",
        f"/users/{USER}/calendar/events",
    ),
    (
        CALENDAR_UPDATE_EVENTS,
        {
            "action": "update_events",
            "user_id": USER,
            "event_id": EVENT,
            "etag": ETAG,
            "fields": {"subject": "Renamed"},
        },
        Event(id=EVENT),
        "PATCH",
        f"/users/{USER}/calendar/events/{EVENT}",
    ),
]


@pytest.mark.parametrize("key, arguments, response, method, path", WRITE_CALLS)
def test_every_write_reaches_the_seam_with_the_declared_request(
    key, arguments, response, method, path
):
    adapter = HandlerGraphAdapter({(method, path): response})

    run(key, arguments, adapter)

    assert len(adapter.requests) == 1
    assert adapter.requests[0].http_method.value == method


@pytest.mark.parametrize("key, arguments, response, method, path", WRITE_CALLS)
def test_every_write_fails_closed_when_the_transport_refuses_to_send(
    key, arguments, response, method, path
):
    """The strict double raises on every send: a write cannot reach a network."""
    del response, method, path
    adapter = HandlerGraphAdapter()

    with pytest.raises(GraphError) as caught:
        run(key, arguments, adapter)

    assert caught.value.category == "transport_error"
    assert "network transport must not run" not in json.dumps(caught.value.to_result())
    assert len(adapter.requests) == 1


# ======================================================================================
# Guard pins: every bound and refusal below fails a test when it is removed
# ======================================================================================


def test_outlook_search_applies_the_declared_item_budget_when_none_was_asked_for():
    from microsoft365.handlers import DEFAULT_LIMIT

    assert DEFAULT_LIMIT == 50
    adapter = HandlerGraphAdapter({("GET", f"/users/{USER}/messages"): message_page([])})

    run(OUTLOOK_SEARCH, {"action": "search", "user_id": USER}, adapter)

    # No query and no filter: no ``$filter`` is sent at all, the budget is the default, and a
    # select-less search still carries the compact candidate metadata (never the body).
    assert search_query(adapter.requests[0]) == {
        "top": 50,
        "select": ["id", "subject", "from", "receivedDateTime", "hasAttachments", "webLink"],
    }


@pytest.mark.parametrize("top", [0, 101, "5", True])
def test_a_read_refuses_an_item_budget_outside_the_contract_range(top):
    adapter = HandlerGraphAdapter()

    with pytest.raises(GraphError) as caught:
        run(OUTLOOK_SEARCH, {"action": "search", "user_id": USER, "top": top}, adapter)

    assert caught.value.category == "validation_error"
    assert "top" in caught.value.message
    assert adapter.requests == []


@pytest.mark.parametrize(
    "service, operation, arguments, expected",
    [
        ("outlook", "search", {"user_id": ""}, "user_id"),
        ("outlook", "search", {"user_id": USER, "query": "   "}, "query"),
        ("outlook", "search", {"user_id": USER, "filter": "  "}, "filter"),
        ("outlook", "search", {"user_id": USER, "select": []}, "select"),
        ("outlook", "search", {"user_id": USER, "select": "id"}, "select"),
        ("outlook", "read", {"user_id": USER}, "message_id"),
        ("outlook", "read", {"user_id": USER, "message_id": " "}, "message_id"),
        ("calendar", "search", {"user_id": USER, "filter": ""}, "filter"),
        ("outlook", "create_draft", {"user_id": USER, "to_recipients": ["a@example.com"]}, "subject"),
        (
            "outlook",
            "create_draft",
            {
                "user_id": USER,
                "subject": "Weekly report",
                "body": "Body text",
                "to_recipients": [42],
            },
            "to_recipients",
        ),
        (
            "outlook",
            "send",
            {
                "user_id": USER,
                "subject": "Weekly report",
                "to_recipients": ["a@example.com"],
                "save_to_sent_items": "true",
            },
            "save_to_sent_items",
        ),
        (
            "calendar",
            "create_events",
            {
                "user_id": USER,
                "subject": "Planning",
                "start_date_time": START,
                "end_date_time": END,
                "time_zone": " ",
            },
            "time_zone",
        ),
        (
            "calendar",
            "create_events",
            {
                "user_id": USER,
                "subject": "Planning",
                "start_date_time": START,
                "end_date_time": END,
                "time_zone": TIME_ZONE,
                "is_all_day": "false",
            },
            "is_all_day",
        ),
        (
            "calendar",
            "update_events",
            {
                "user_id": USER,
                "event_id": EVENT,
                "etag": ETAG,
                "fields": {"start_date_time": START},
            },
            "time_zone",
        ),
        (
            "calendar",
            "update_events",
            {
                "user_id": USER,
                "event_id": EVENT,
                "etag": ETAG,
                "fields": {"is_all_day": "yes"},
            },
            "is_all_day",
        ),
        (
            "calendar",
            "update_events",
            {
                "user_id": USER,
                "event_id": EVENT,
                "etag": " ",
                "fields": {"subject": "Renamed"},
            },
            "etag",
        ),
    ],
)
def test_an_argument_the_handler_cannot_express_is_refused_before_any_client(service, operation, arguments, expected):
    handler_key = f"{service}.{operation}"
    payload = {"action": operation, **arguments}
    adapter = HandlerGraphAdapter()
    calls: list = []

    def factory():
        calls.append("client")
        return graph_client(adapter)

    with pytest.raises(GraphError) as caught:
        invocation(handler_key)(payload, HandlerContext(settings=settings(), client_factory=factory))

    assert caught.value.category == "validation_error"
    assert expected in caught.value.message
    assert calls == []
    assert adapter.requests == []


def test_a_read_reports_a_typed_failure_when_the_client_cannot_be_built():
    from microsoft365.execution import ExecutionError

    def factory():
        raise RuntimeError("raw credential text must never surface")

    context = HandlerContext(settings=settings(), client_factory=factory)

    with pytest.raises(ExecutionError) as caught:
        invocation(OUTLOOK_READ)({"action": "read", "user_id": USER, "message_id": MESSAGE}, context)

    assert caught.value.category == "authentication_required"
    assert "raw credential text" not in str(caught.value)


def test_a_context_without_a_usable_client_factory_fails_closed():
    from microsoft365.execution import ExecutionError

    # A context whose factory is not callable cannot even be built.
    with pytest.raises(ExecutionError) as caught:
        HandlerContext(settings=settings(), client_factory=None)

    assert caught.value.category == "configuration_error"


def test_a_handler_refuses_a_context_of_the_wrong_shape():
    from microsoft365.execution import ExecutionError

    with pytest.raises(ExecutionError) as caught:
        invocation(OUTLOOK_READ)(
            {"action": "read", "user_id": USER, "message_id": MESSAGE}, {"settings": settings()}
        )

    assert caught.value.category == "configuration_error"


def test_a_handler_context_requires_real_settings():
    from microsoft365.execution import ExecutionError

    with pytest.raises(ExecutionError) as caught:
        HandlerContext(settings={"tenant_id": "tenant"}, client_factory=lambda: None)

    assert caught.value.category == "configuration_error"


def test_update_events_refuses_when_a_field_did_not_reach_the_request(monkeypatch):
    """The change-guard itself: a declared field that produces no wire key is refused."""
    from microsoft365.handlers import calendar

    adapter = HandlerGraphAdapter(
        {("PATCH", f"/users/{USER}/calendar/events/{EVENT}"): Event(id=EVENT)}
    )
    monkeypatch.setitem(calendar.EVENT_WIRE_KEYS, "subject", ("no-such-key",))
    arguments = {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": {"subject": "Renamed"},
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_UPDATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "would not change the request" in caught.value.message
    assert adapter.requests == []


def test_create_events_refuses_when_a_field_did_not_reach_the_request(monkeypatch):
    """The same change-guard on the create path."""
    from microsoft365.handlers import calendar

    adapter = HandlerGraphAdapter(
        {("POST", f"/users/{USER}/calendar/events"): Event(id="event-1")}
    )
    monkeypatch.setitem(calendar.EVENT_WIRE_KEYS, "location", ("no-such-key",))
    arguments = {
        "action": "create_events",
        "user_id": USER,
        "subject": "Planning",
        "start_date_time": START,
        "end_date_time": END,
        "time_zone": TIME_ZONE,
        "location": "Room 1",
    }

    with pytest.raises(GraphError) as caught:
        run(CALENDAR_CREATE_EVENTS, arguments, adapter)

    assert caught.value.category == "validation_error"
    assert "would not change the request" in caught.value.message
    assert adapter.requests == []


# ======================================================================================
# The milestone surface: three verified reads executable, four writes implemented and withheld
# ======================================================================================

#: The operations this milestone exposes to the model: the eleven verified reads. Spelled out
#: literally, so a reverted flag -- or one flag too many -- fails here.
EXECUTABLE_OPERATIONS = frozenset(
    {
        OUTLOOK_SEARCH, OUTLOOK_READ, "outlook.list_attachments", "outlook.download_attachment", CALENDAR_SEARCH,
        "sharepoint.search", "sharepoint.read", "sharepoint.download_files",
        "onedrive.search", "onedrive.read", "onedrive.download_files",
        "teams.list_teams", "teams.list_channels",
        "todo.list_task_lists", "todo.search", "todo.read",
        "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read",
    }
)

#: The six writes: implemented, contract-pinned, exercised offline through the seam, and
#: non-executable until the generic host approval fix (CORE-1/CORE-2) is available (R5).
NON_EXECUTABLE_WRITES = frozenset(
    {
        OUTLOOK_CREATE_DRAFT, OUTLOOK_SEND, CALENDAR_CREATE_EVENTS, CALENDAR_UPDATE_EVENTS,
        "sharepoint.upload_files", "onedrive.upload_files",
        "todo.create_tasks", "todo.update_tasks", "planner.create_tasks", "planner.update_tasks",
    }
)

#: Every operation a handler exists for today: WP6 owns the Outlook/Calendar handlers and WP7
#: owns the SharePoint/OneDrive handlers; WP8's delegated-only handlers are tested separately.
IMPLEMENTED_OPERATIONS = EXECUTABLE_OPERATIONS | NON_EXECUTABLE_WRITES

#: Concrete identifiers for the declared endpoint rows, and the extra arguments each contract
#: case needs. Used to rebuild every declared endpoint for real (see
#: ``test_the_declared_messaging_endpoints_are_the_ones_the_sdk_contract_builds``).
ENDPOINT_IDENTIFIERS = {"attachment_id": "attachment", "user_id": USER, "message_id": MESSAGE, "event_id": EVENT}

ENDPOINT_CASE_ARGUMENTS: dict[tuple[str, str], dict] = {
    ("build_collection_request_information", "outlook_attachments"): {},
    ("build_item_request_information", "outlook_attachment"): {},
    ("build_collection_request_information", "outlook_messages"): {"query": "invoice"},
    ("build_collection_request_information", "calendar_events"): {"query": "planning"},
    ("build_item_request_information", "outlook_read"): {},
    ("build_collection_request_information", "outlook_group_threads"): {},
    ("build_collection_request_information", "outlook_group_posts"): {},
    ("build_collection_request_information", "outlook_group_post_attachments"): {},
    ("build_write_request_information", "outlook_create_draft"): {
        "subject": "Weekly report",
        "content": "Body text",
        "recipients": ["a@example.com"],
    },
    ("build_write_request_information", "outlook_send"): {
        "subject": "Weekly report",
        "content": "Body text",
        "recipients": ["a@example.com"],
        "save_to_sent_items": True,
    },
    ("build_write_request_information", "outlook_send_existing"): {},
    ("build_write_request_information", "calendar_create_events"): {
        "subject": "Planning",
        "start": DateTimeTimeZone(date_time=START, time_zone=TIME_ZONE),
        "end": DateTimeTimeZone(date_time=END, time_zone=TIME_ZONE),
    },
    ("build_write_request_information", "calendar_update_events"): {
        "etag": ETAG,
        "update_fields": {"subject": "Renamed"},
    },
}


def test_the_registry_declares_exactly_the_three_verified_reads_executable():
    """The milestone flip itself, with every write still withheld."""
    from microsoft365.contract import IMPLEMENTATION_STATUS_LABELS, OPERATION_REGISTRY

    assert {
        key for key, definition in OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_OPERATIONS)

    for key in sorted(IMPLEMENTED_OPERATIONS):
        definition = OPERATION_REGISTRY[key]
        # exactly the same evidence for both halves: an endpoint row and a contract case
        assert definition.endpoints, key
        assert definition.contract_cases, key
        assert definition.implementation_status == "implemented", key
        assert definition.implementation_status in IMPLEMENTATION_STATUS_LABELS, key
        assert definition.write is (key in NON_EXECUTABLE_WRITES), key
        assert definition.executable is (key in EXECUTABLE_OPERATIONS), key

    assert {
        key for key in OPERATION_REGISTRY if key.startswith(("outlook.", "calendar."))
    } == set(WP6_HANDLERS)


def test_operation_status_exposes_the_reads_and_withholds_every_write():
    """``operation_status`` is the gate the hook and the dispatch path both read."""
    from microsoft365.contract import operation_status

    for key in sorted(IMPLEMENTED_OPERATIONS):
        service, operation = key.split(".", 1)
        application = operation_status("application", service, operation)
        assert application.auth_status == ("not_verified" if key in {"todo.create_tasks", "todo.update_tasks"} else "supported"), key
        assert application.executable is (key in EXECUTABLE_OPERATIONS), key
        # delegated authentication is not implemented (WP14), so nothing runs there
        assert operation_status("delegated", service, operation).executable is False, key


def test_the_dispatch_table_is_exactly_the_implemented_handler_set():
    """One source of dispatch truth: ``HANDLER_TABLE`` is the self-populating registry."""
    from microsoft365.registration import HANDLER_TABLE

    assert set(HANDLER_TABLE) == set(IMPLEMENTED_OPERATIONS) | {"teams.search_messages", "teams.send_messages", "todo.list_task_lists", "todo.search", "todo.read", "todo.create_tasks", "todo.update_tasks", "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read", "planner.create_tasks", "planner.update_tasks"}
    for key in sorted(IMPLEMENTED_OPERATIONS):
        assert callable(HANDLER_TABLE[key]), key


def test_the_declared_messaging_endpoints_are_the_ones_the_sdk_contract_builds():
    """Never invent an endpoint: each declared row is rebuilt by its own contract case."""
    from microsoft365 import sdk_contract
    from microsoft365.contract import messaging_endpoints

    for key in sorted(WP6_HANDLERS):
        endpoints = messaging_endpoints(key)
        assert endpoints, key
        for endpoint in endpoints:
            arguments = dict(
                ENDPOINT_CASE_ARGUMENTS[(endpoint.contract_call, endpoint.contract_case)]
            )
            # Only the path placeholders are rendered: an identifier the endpoint needs
            # without putting it in the URL (the group download's attachment_id) is not a
            # segment and has nothing to substitute.
            for name in endpoint.path_identifiers:
                arguments[name] = ENDPOINT_IDENTIFIERS[name]
            request = getattr(sdk_contract, endpoint.contract_call)(
                graph_client(HandlerGraphAdapter()), endpoint.contract_case, **arguments
            )

            assert request.http_method.value == endpoint.method, endpoint.endpoint
            expected = endpoint.path_template
            for name, value in ENDPOINT_IDENTIFIERS.items():
                expected = expected.replace(f"{{{name}}}", value)
            assert urlsplit(request.url).path == expected, endpoint.endpoint


def test_write_operations_stay_out_of_the_model_facing_schema():
    """R5 / CORE-1: a write must not be reachable from the model surface at all."""
    from microsoft365.registration import active_actions, schema_for

    configuration = settings()
    for key in sorted(WP6_WRITES):
        service, operation = key.split(".", 1)
        actions = active_actions(configuration, service)
        assert operation not in actions, key
        schema = schema_for(service, actions)
        enum = [] if schema is None else schema["parameters"]["properties"]["action"]["enum"]
        assert operation not in enum, key


def test_the_reads_are_model_facing_and_the_withheld_writes_keep_their_handler():
    """The dispatch table holds the writes; the registry flag is what withholds them."""
    from microsoft365.registration import HANDLER_TABLE, active_actions, schema_for

    configuration = settings()
    for key in sorted(EXECUTABLE_OPERATIONS):
        service, operation = key.split(".", 1)
        actions = active_actions(configuration, service)
        assert operation in actions, key
        assert key in HANDLER_TABLE, key
        schema = schema_for(service, actions)
        assert operation in schema["parameters"]["properties"]["action"]["enum"], key

    for key in sorted(NON_EXECUTABLE_WRITES):
        service, operation = key.split(".", 1)
        actions = active_actions(configuration, service)
        assert operation not in actions, key
        # the handler exists -- so it is the executability flag, not a missing handler, that
        # keeps the write off the model surface
        assert key in HANDLER_TABLE, key
        schema = schema_for(service, actions)
        enum = [] if schema is None else schema["parameters"]["properties"]["action"]["enum"]
        assert operation not in enum, key

    # Not dead code: the withheld writes are exactly the ones the parametrized seam tests
    # above drive to a real request through the strict double. The two SharePoint/OneDrive
    # uploads are driven to a real request in tests/test_handlers_files.py instead.
    assert {key for key, *_ in WRITE_CALLS} == set(WP6_WRITES)


def test_known_limitations_records_the_milestone_state():
    """Invariant 6: the document moves in the same commit as the capacity it describes."""
    text = (REPO_ROOT / "docs" / "known-limitations.md").read_text(encoding="utf-8")

    for key in sorted(IMPLEMENTED_OPERATIONS):
        assert key.split(".", 1)[1] in text or key in text, key
    assert "Exactly twenty reads are executable" in text
    assert "`active_actions()` never returns them" in text
    assert "final arguments" in text
    # the pre-WP6 wording claimed nothing consumed the paginator
    assert "The four executable search reads consume it" in text
    # and that a list-valued PATCH field was unimplemented everywhere
    assert "so affected updates stay unimplemented" not in text
