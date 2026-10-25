"""WP7 — SharePoint and OneDrive handlers, driven offline through the execution seam.

What this module proves
-----------------------
* the self-populating handler registry picks up ``handlers/files.py`` without editing
  ``registration.py`` (the same discovery seam the Outlook/Calendar work package added);
* every handler builds its request from the **generated SDK builders** and is pinned against
  the verified request-information contract of the same operation
  (``microsoft365/sdk_contract.py``, WP1): same method, URL and typed query, the search
  ``search(q='…')`` path, the site→drive resolution path, the item read path and the
  ``/content`` transfer path;
* the search handlers go through ``microsoft365/paging.py`` (validated next links, bounded
  traversal, truthful continuation) and the read/download/upload handlers through the
  execution seam and the bounded ``microsoft365/files.py`` transfer;
* exactly ``search``, ``read`` and ``download_files`` are executable and model-facing for
  both services, while ``upload_files`` is implemented, contract-pinned and exercised offline
  through the seam, and deliberately **not executable** (R5 / CORE-1 / CORE-2);
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

import json
from pathlib import Path
from typing import Any
from urllib.parse import parse_qsl, unquote, urlsplit

import pytest
from kiota_abstractions.request_adapter import RequestAdapter
from kiota_serialization_json.json_serialization_writer_factory import JsonSerializationWriterFactory
from msgraph import GraphServiceClient
from msgraph.generated.drives.item.search_with_q.search_with_q_get_response import (
    SearchWithQGetResponse,
)
from msgraph.generated.models.drive import Drive
from msgraph.generated.models.drive_item import DriveItem
from msgraph.generated.models.file import File
from msgraph.generated.models.site import Site
from msgraph.generated.models.site_collection_response import SiteCollectionResponse

from microsoft365 import handlers as handler_package
from microsoft365.contract import OPERATIONS, OPERATION_REGISTRY, WRITE_OPERATIONS, Settings
from microsoft365.errors import CATEGORIES, GraphError
from microsoft365.handlers import HandlerContext, load_handlers
from microsoft365.sdk_contract import (
    build_content_request_information,
    build_item_request_information,
    build_search_request_information,
)
from tests.test_sdk_contract import StrictTransportAdapter

REPO_ROOT = Path(__file__).resolve().parent.parent

SHAREPOINT_SEARCH = "sharepoint.search"
SHAREPOINT_READ = "sharepoint.read"
SHAREPOINT_DOWNLOAD = "sharepoint.download_files"
SHAREPOINT_UPLOAD = "sharepoint.upload_files"
ONEDRIVE_SEARCH = "onedrive.search"
ONEDRIVE_READ = "onedrive.read"
ONEDRIVE_DOWNLOAD = "onedrive.download_files"
ONEDRIVE_UPLOAD = "onedrive.upload_files"

#: The eight operations WP7 owns.
WP7_HANDLERS = frozenset(
    {
        SHAREPOINT_SEARCH, SHAREPOINT_READ, SHAREPOINT_DOWNLOAD, SHAREPOINT_UPLOAD,
        ONEDRIVE_SEARCH, ONEDRIVE_READ, ONEDRIVE_DOWNLOAD, ONEDRIVE_UPLOAD,
    }
)

#: The six reads this milestone exposes to the model. The uploads stay implemented and
#: non-executable until the generic host approval fix (CORE-1/CORE-2) is available.
WP7_EXECUTABLE = frozenset(
    {
        SHAREPOINT_SEARCH, SHAREPOINT_READ, SHAREPOINT_DOWNLOAD,
        ONEDRIVE_SEARCH, ONEDRIVE_READ, ONEDRIVE_DOWNLOAD,
    }
)

#: The two uploads: implemented, contract-pinned, exercised offline and withheld.
WP7_WRITES = frozenset({SHAREPOINT_UPLOAD, ONEDRIVE_UPLOAD})

DRIVE = "drive"
SITE = "site"
ITEM = "item"
ITEM_ADDRESS = "root:/a/b.txt:"
ITEM_ENCODED_PATH = "/drives/drive/items/root%3A%2Fa%2Fb.txt%3A"
MIME = "application/pdf"


def settings() -> Settings:
    return Settings.from_mapping(
        {
            "tenant_id": "tenant",
            "client_id": "client",
            "user_id": "user",
            "capabilities": {service: True for service in OPERATIONS},
        }
    )


def request_identity(request_information: Any) -> tuple:
    """Canonical identity of a built request: method, path and typed query, sorted.

    Mirrors the paginator's own identity: a typed configuration's query values live in
    ``query_parameters`` (with percent-encoded OData names) while a followed next link's query
    lives inside the raw URL, so the two sources are unioned — never concatenated.
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
    """Strict transport double that records every request and answers only named requests."""

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
    client = graph_client(adapter)
    return HandlerContext(settings=settings(), client_factory=lambda: client)


def invocation(key: str):
    record = load_handlers().get(key)
    assert record is not None, f"{key} is not registered"
    return record.function


def run(key: str, arguments: dict, adapter: RequestAdapter) -> dict:
    return invocation(key)(arguments, open_context(adapter))


def search_query(request_information: Any) -> dict:
    """The typed query of a built request with OData names unfolded."""
    query = {}
    for name, value in (getattr(request_information, "query_parameters", None) or {}).items():
        query[unquote(str(name)).lstrip("$")] = value
    for name, value in parse_qsl(urlsplit(request_information.url).query, keep_blank_values=True):
        query.setdefault(name.lstrip("$"), value)
    return query


def drive_item(*, size: int, mime_type: str = MIME, name: str = "report.pdf", item_id: str = ITEM) -> DriveItem:
    return DriveItem(id=item_id, name=name, size=size, file=File(mime_type=mime_type))


def search_page(items: list, *, next_link: str | None = None) -> SearchWithQGetResponse:
    return SearchWithQGetResponse(value=items, odata_next_link=next_link)


def site_page(items: list, *, next_link: str | None = None) -> SiteCollectionResponse:
    return SiteCollectionResponse(value=items, odata_next_link=next_link)


# ======================================================================================
# Registry mechanism: files.py is discovered, holds every WP7 operation, writes withheld
# ======================================================================================


def test_the_handler_registry_holds_every_wp7_operation_with_its_write_classification():
    loaded = load_handlers()

    assert {key for key in loaded if key.startswith(("sharepoint.", "onedrive."))} == set(WP7_HANDLERS)
    for key, record in loaded.items():
        assert record.key == key
        assert record.key == f"{record.service}.{record.operation}"
        assert record.write is (record.operation in WRITE_OPERATIONS), key
        assert callable(record.function), key


def test_the_dispatch_table_is_exactly_the_implemented_handler_set():
    """One source of dispatch truth: HANDLER_TABLE is the self-populating registry."""
    from microsoft365.registration import HANDLER_TABLE

    implemented = {
        "outlook.search", "outlook.read", "outlook.list_attachments", "outlook.download_attachment", "outlook.create_draft", "outlook.send",
        "calendar.search", "calendar.create_events", "calendar.update_events",
        "teams.list_teams", "teams.list_channels", "teams.search_messages", "teams.send_messages",
    } | set(WP7_HANDLERS)
    assert set(HANDLER_TABLE) == implemented | {"todo.list_task_lists", "todo.search", "todo.read", "todo.create_tasks", "todo.update_tasks", "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read", "planner.create_tasks", "planner.update_tasks"}
    for key in sorted(WP7_HANDLERS):
        assert callable(HANDLER_TABLE[key]), key


# ======================================================================================
# Search — sites search, drive search and site→drive resolution, pinned to the contracts
# ======================================================================================


def test_sharepoint_search_without_a_container_searches_sites():
    adapter = HandlerGraphAdapter(
        {("GET", "/sites"): site_page([Site(id="s1", display_name="Marketing", web_url="https://x")])}
    )

    payload = run(SHAREPOINT_SEARCH, {"action": "search", "query": "projeto", "top": 7}, adapter)

    client = graph_client(adapter)
    verified = build_search_request_information(client, "sites_search", query="projeto")
    assert urlsplit(adapter.requests[0].url).path == urlsplit(verified.url).path
    assert search_query(adapter.requests[0]) == {"search": "projeto", "top": 7}
    assert payload["operation"] == SHAREPOINT_SEARCH
    assert payload["result"]["value"][0]["display_name"] == "Marketing"
    assert payload["continuation"]["stop_reason"] == "complete"


def test_sharepoint_search_by_drive_targets_the_verified_search_endpoint():
    adapter = HandlerGraphAdapter(
        {("GET", "/drives/drive/search(q='hello')"): search_page([DriveItem(id="f1", name="a.txt")])}
    )

    payload = run(SHAREPOINT_SEARCH, {"action": "search", "query": "hello", "drive_id": DRIVE, "top": 7}, adapter)

    client = graph_client(adapter)
    verified = build_search_request_information(client, "drive_search", drive_id=DRIVE, query="hello")
    assert urlsplit(adapter.requests[0].url).path == urlsplit(verified.url).path
    assert search_query(adapter.requests[0]) == {"top": 7}
    assert payload["result"]["value"][0]["name"] == "a.txt"


def test_sharepoint_search_doubles_odata_quotes_in_the_drive_search_path():
    adapter = HandlerGraphAdapter(
        {("GET", "/drives/drive/search(q='it%27%27s')"): search_page([])}
    )

    run(SHAREPOINT_SEARCH, {"action": "search", "query": "it's", "drive_id": DRIVE}, adapter)

    client = graph_client(adapter)
    verified = build_search_request_information(client, "drive_search", drive_id=DRIVE, query="it's")
    assert urlsplit(adapter.requests[0].url).path == urlsplit(verified.url).path


def test_sharepoint_search_resolves_a_site_to_its_drive_before_searching():
    adapter = HandlerGraphAdapter(
        {
            ("GET", "/sites/site/drive"): Drive(id="resolved-drive", drive_type="documentLibrary"),
            ("GET", "/drives/resolved-drive/search(q='hello')"): search_page([DriveItem(id="f1", name="a.txt")]),
        }
    )

    payload = run(SHAREPOINT_SEARCH, {"action": "search", "query": "hello", "site_id": SITE}, adapter)

    assert [r.http_method.value for r in adapter.requests] == ["GET", "GET"]
    assert urlsplit(adapter.requests[0].url).path == "/sites/site/drive"
    assert urlsplit(adapter.requests[1].url).path == "/drives/resolved-drive/search(q='hello')"
    assert payload["result"]["value"][0]["name"] == "a.txt"


def test_onedrive_search_by_drive_targets_the_verified_search_endpoint():
    adapter = HandlerGraphAdapter(
        {("GET", "/drives/drive/search(q='hello')"): search_page([DriveItem(id="f1", name="a.txt")])}
    )

    run(ONEDRIVE_SEARCH, {"action": "search", "query": "hello", "drive_id": DRIVE}, adapter)

    client = graph_client(adapter)
    verified = build_search_request_information(client, "drive_search", drive_id=DRIVE, query="hello")
    assert urlsplit(adapter.requests[0].url).path == urlsplit(verified.url).path


def test_onedrive_search_refuses_a_site_container_before_any_client_work():
    adapter = HandlerGraphAdapter()
    calls: list = []

    def factory():
        calls.append("client")
        return graph_client(adapter)

    context = HandlerContext(settings=settings(), client_factory=factory)
    arguments = {"action": "search", "query": "hello", "site_id": SITE}

    with pytest.raises(GraphError) as caught:
        invocation(ONEDRIVE_SEARCH)(arguments, context)

    assert caught.value.category == "validation_error"
    assert calls == []
    assert adapter.requests == []


def test_onedrive_search_requires_a_drive_before_any_client_work():
    adapter = HandlerGraphAdapter()
    calls: list = []

    def factory():
        calls.append("client")
        return graph_client(adapter)

    context = HandlerContext(settings=settings(), client_factory=factory)
    arguments = {"action": "search", "query": "hello"}

    with pytest.raises(GraphError) as caught:
        invocation(ONEDRIVE_SEARCH)(arguments, context)

    assert caught.value.category == "validation_error"
    assert calls == []
    assert adapter.requests == []


def test_search_follows_a_validated_next_link_and_stops_at_the_caller_limit():
    first = "https://graph.microsoft.com/v1.0/drives/drive/search(q='hello')?$skiptoken=page-2"
    adapter = HandlerGraphAdapter(
        {
            ("GET", "/drives/drive/search(q='hello')"): search_page([DriveItem(id="f1"), DriveItem(id="f2")], next_link=first),
            (
                "GET",
                "/drives/drive/search(q='hello')",
                (("$skiptoken", "page-2"),),
            ): search_page([DriveItem(id="f3"), DriveItem(id="f4")]),
        }
    )

    payload = run(SHAREPOINT_SEARCH, {"action": "search", "query": "hello", "drive_id": DRIVE, "top": 4}, adapter)

    assert [item["id"] for item in payload["result"]["value"]] == ["f1", "f2", "f3", "f4"]
    assert payload["continuation"]["pages_fetched"] == 2
    assert payload["continuation"]["stop_reason"] == "complete"
    assert len(adapter.requests) == 2


def test_search_never_reaches_the_network_with_the_strict_double():
    adapter = HandlerGraphAdapter()

    with pytest.raises(GraphError) as caught:
        run(SHAREPOINT_SEARCH, {"action": "search", "query": "hello", "drive_id": DRIVE}, adapter)

    assert caught.value.category == "transport_error"
    assert "network transport must not run" not in json.dumps(caught.value.to_result())
    assert len(adapter.requests) == 3


# ======================================================================================
# Read — a single drive item, pinned to the verified item contract
# ======================================================================================


def test_read_by_path_is_pinned_to_the_verified_item_contract():
    adapter = HandlerGraphAdapter(
        {("GET", ITEM_ENCODED_PATH): DriveItem(id="f1", name="a.txt", size=4)}
    )

    payload = run(SHAREPOINT_READ, {"action": "read", "drive_id": DRIVE, "item_path": "a/b.txt"}, adapter)

    client = graph_client(adapter)
    verified = build_item_request_information(client, "drive_item_read", drive_id=DRIVE, drive_item_id=ITEM_ADDRESS)
    assert request_identity(adapter.requests[0]) == request_identity(verified)
    assert payload["operation"] == SHAREPOINT_READ
    assert payload["result"]["name"] == "a.txt"


def test_read_by_id_is_pinned_to_the_verified_item_contract():
    adapter = HandlerGraphAdapter(
        {("GET", f"/drives/{DRIVE}/items/{ITEM}"): DriveItem(id=ITEM, name="a.txt")}
    )

    payload = run(ONEDRIVE_READ, {"action": "read", "drive_id": DRIVE, "item_id": ITEM, "select": ["name", "size"]}, adapter)

    client = graph_client(adapter)
    verified = build_item_request_information(client, "drive_item_read", drive_id=DRIVE, drive_item_id=ITEM, select=["name", "size"])
    assert request_identity(adapter.requests[0]) == request_identity(verified)
    assert search_query(adapter.requests[0])["select"] == ["name", "size"]
    assert payload["result"]["name"] == "a.txt"


def test_read_sends_no_select_when_the_caller_did_not_ask_for_one():
    adapter = HandlerGraphAdapter({("GET", ITEM_ENCODED_PATH): DriveItem(id="f1")})

    run(SHAREPOINT_READ, {"action": "read", "drive_id": DRIVE, "item_path": "a/b.txt"}, adapter)

    assert "select" not in search_query(adapter.requests[0])


def test_read_refuses_a_traversal_path_before_any_client_work():
    adapter = HandlerGraphAdapter()
    calls: list = []

    def factory():
        calls.append("client")
        return graph_client(adapter)

    context = HandlerContext(settings=settings(), client_factory=factory)

    with pytest.raises(GraphError) as caught:
        invocation(SHAREPOINT_READ)({"action": "read", "drive_id": DRIVE, "item_path": "../secret.txt"}, context)

    assert caught.value.category == "validation_error"
    assert calls == []
    assert adapter.requests == []


# ======================================================================================
# Download and upload — the bounded /content transfer through the execution seam
# ======================================================================================


def test_download_returns_exact_bytes_from_the_content_endpoint():
    payload = bytes(range(256)) * 4
    adapter = HandlerGraphAdapter(
        {
            ("GET", ITEM_ENCODED_PATH): drive_item(size=len(payload), name="b.txt"),
            ("GET", f"{ITEM_ENCODED_PATH}/content"): payload,
        }
    )

    result = run(SHAREPOINT_DOWNLOAD, {"action": "download_files", "drive_id": DRIVE, "item_path": "a/b.txt"}, adapter)

    paths = [urlsplit(request.url).path for request in adapter.requests]
    assert paths == [ITEM_ENCODED_PATH, f"{ITEM_ENCODED_PATH}/content"]
    assert result["operation"] == SHAREPOINT_DOWNLOAD
    assert result["result"]["content_type"] == MIME
    assert result["result"]["name"] == "b.txt"
    assert result["result"]["size"] == len(payload)


def test_download_by_id_targets_the_same_content_endpoint():
    payload = b"abc"
    adapter = HandlerGraphAdapter(
        {
            ("GET", f"/drives/{DRIVE}/items/{ITEM}"): drive_item(size=3, name="a.txt"),
            ("GET", f"/drives/{DRIVE}/items/{ITEM}/content"): payload,
        }
    )

    run(ONEDRIVE_DOWNLOAD, {"action": "download_files", "drive_id": DRIVE, "item_id": ITEM}, adapter)

    assert urlsplit(adapter.requests[1].url).path == f"/drives/{DRIVE}/items/{ITEM}/content"


def test_download_content_type_comes_from_authenticated_metadata_not_the_caller():
    payload = b"<h1>"
    adapter = HandlerGraphAdapter(
        {
            ("GET", ITEM_ENCODED_PATH): drive_item(size=4, mime_type="application/pdf", name="report.html"),
            ("GET", f"{ITEM_ENCODED_PATH}/content"): payload,
        }
    )

    result = run(SHAREPOINT_DOWNLOAD, {"action": "download_files", "drive_id": DRIVE, "item_path": "a/b.txt"}, adapter)

    # The content type is the one authenticated metadata declared, not anything the caller
    # supplied (the download_files argument contract has no content_type at all).
    assert result["result"]["content_type"] == "application/pdf"
    assert result["result"]["name"] == "report.html"


def test_upload_puts_exact_bytes_with_content_type_to_the_content_url():
    import base64

    payload = bytes(range(256)) * 16
    adapter = HandlerGraphAdapter(
        {("PUT", f"{ITEM_ENCODED_PATH}/content"): DriveItem(id="f1", name="a.txt")}
    )

    result = run(
        SHAREPOINT_UPLOAD,
        {
            "action": "upload_files",
            "drive_id": DRIVE,
            "item_path": "a/b.txt",
            "content_base64": base64.b64encode(payload).decode("ascii"),
            "content_type": "text/csv",
        },
        adapter,
    )

    request = adapter.requests[0]
    assert request.http_method.value == "PUT"
    assert urlsplit(request.url).path == f"{ITEM_ENCODED_PATH}/content"
    assert request.content == payload
    assert result["operation"] == SHAREPOINT_UPLOAD
    assert result["result"]["size_bytes"] == len(payload)
    assert result["result"]["content_type"] == "text/csv"


def test_upload_defaults_the_content_type_to_octet_stream():
    import base64

    payload = b"payload"
    adapter = HandlerGraphAdapter(
        {("PUT", f"{ITEM_ENCODED_PATH}/content"): DriveItem(id="f1")}
    )

    result = run(
        SHAREPOINT_UPLOAD,
        {
            "action": "upload_files",
            "drive_id": DRIVE,
            "item_path": "a/b.txt",
            "content_base64": base64.b64encode(payload).decode("ascii"),
        },
        adapter,
    )

    assert result["result"]["content_type"] == "application/octet-stream"


def test_upload_refuses_overwrite_false_which_a_simple_put_cannot_honor():
    import base64

    adapter = HandlerGraphAdapter()
    calls: list = []

    def factory():
        calls.append("client")
        return graph_client(adapter)

    context = HandlerContext(settings=settings(), client_factory=factory)
    arguments = {
        "action": "upload_files",
        "drive_id": DRIVE,
        "item_path": "a/b.txt",
        "content_base64": base64.b64encode(b"payload").decode("ascii"),
        "overwrite": False,
    }

    with pytest.raises(GraphError) as caught:
        invocation(SHAREPOINT_UPLOAD)(arguments, context)

    assert caught.value.category == "validation_error"
    assert calls == []
    assert adapter.requests == []


def test_upload_rejects_an_oversized_encoded_payload_before_decoding():
    import base64

    from microsoft365.files import SIMPLE_TRANSFER_LIMIT_BYTES

    oversized = base64.b64encode(b"x" * (SIMPLE_TRANSFER_LIMIT_BYTES + 4096)).decode("ascii")
    adapter = HandlerGraphAdapter()

    with pytest.raises(GraphError) as caught:
        run(
            SHAREPOINT_UPLOAD,
            {
                "action": "upload_files",
                "drive_id": DRIVE,
                "item_path": "a/b.txt",
                "content_base64": oversized,
                "content_type": "text/plain",
            },
            adapter,
        )

    assert caught.value.category == "operation_not_implemented"
    assert adapter.requests == []


# ======================================================================================
# The milestone surface: six reads executable, uploads implemented and withheld
# ======================================================================================


def test_the_registry_declares_exactly_the_six_file_reads_executable():
    """The WP7 flip: search/read/download executable for both services, upload withheld."""
    from microsoft365.contract import EXECUTABLE_OPERATIONS

    assert set(EXECUTABLE_OPERATIONS) == {
        "outlook.search", "outlook.read", "outlook.list_attachments", "outlook.download_attachment", "calendar.search",
        "teams.list_teams", "teams.list_channels",
    } | set(WP7_EXECUTABLE) | {"todo.list_task_lists", "todo.search", "todo.read", "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read"}

    assert {
        key for key, definition in OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_OPERATIONS)


def test_every_wp7_operation_is_implemented_and_contract_pinned():
    from microsoft365.contract import IMPLEMENTATION_STATUS_LABELS

    for key in sorted(WP7_HANDLERS):
        definition = OPERATION_REGISTRY[key]
        assert definition.endpoints, key
        assert definition.contract_cases, key
        assert definition.implementation_status == "implemented", key
        assert definition.implementation_status in IMPLEMENTATION_STATUS_LABELS, key
        assert definition.write is (key in WP7_WRITES), key
        assert definition.executable is (key in WP7_EXECUTABLE), key


def test_operation_status_exposes_the_reads_and_withholds_every_upload():
    from microsoft365.contract import operation_status

    for key in sorted(WP7_HANDLERS):
        service, operation = key.split(".", 1)
        application = operation_status("application", service, operation)
        assert application.auth_status == "supported", key
        assert application.executable is (key in WP7_EXECUTABLE), key
        assert operation_status("delegated", service, operation).executable is False, key


def test_uploads_stay_out_of_the_model_facing_schema():
    from microsoft365.registration import active_actions, schema_for

    configuration = settings()
    for key in sorted(WP7_WRITES):
        service, operation = key.split(".", 1)
        actions = active_actions(configuration, service)
        assert operation not in actions, key
        schema = schema_for(service, actions)
        enum = [] if schema is None else schema["parameters"]["properties"]["action"]["enum"]
        assert operation not in enum, key


def test_the_reads_are_model_facing_and_the_withheld_uploads_keep_their_handler():
    from microsoft365.registration import HANDLER_TABLE, active_actions, schema_for

    configuration = settings()
    for key in sorted(WP7_EXECUTABLE):
        service, operation = key.split(".", 1)
        actions = active_actions(configuration, service)
        assert operation in actions, key
        assert key in HANDLER_TABLE, key
        schema = schema_for(service, actions)
        assert operation in schema["parameters"]["properties"]["action"]["enum"], key

    for key in sorted(WP7_WRITES):
        service, operation = key.split(".", 1)
        actions = active_actions(configuration, service)
        assert operation not in actions, key
        assert key in HANDLER_TABLE, key


# ======================================================================================
# Endpoint reality: every declared row is the one the SDK builds
# ======================================================================================


def test_the_declared_drive_endpoints_are_the_ones_the_sdk_contract_builds():
    """Never invent an endpoint: each declared row is rebuilt by its own contract case."""
    from microsoft365 import sdk_contract
    from microsoft365.contract import drive_endpoints

    identifiers = {"drive_id": DRIVE, "drive_item_id": ITEM, "site_id": SITE}
    case_arguments = {
        ("build_search_request_information", "sites_search"): {"query": "projeto"},
        ("build_search_request_information", "drive_search"): {"query": "hello"},
        ("build_item_request_information", "drive_item_read"): {},
        ("build_content_request_information", "download_files"): {},
        ("build_content_request_information", "upload_files"): {
            "content": b"x",
            "content_type": "text/plain",
        },
    }

    for key in sorted(WP7_HANDLERS):
        endpoints = drive_endpoints(key)
        assert endpoints, key
        for endpoint in endpoints:
            if endpoint.contract_call and endpoint.contract_case:
                arguments = dict(case_arguments[(endpoint.contract_call, endpoint.contract_case)])
                for name in endpoint.required_identifiers:
                    arguments[name] = identifiers[name]
                request = getattr(sdk_contract, endpoint.contract_call)(
                    graph_client(HandlerGraphAdapter()), endpoint.contract_case, **arguments
                )
                assert request.http_method.value == endpoint.method, endpoint.endpoint
                expected = endpoint.path_template
                for name, value in identifiers.items():
                    expected = expected.replace(f"{{{name}}}", value)
                expected = expected.replace("{query}", "hello" if endpoint.contract_case == "drive_search" else "projeto")
                assert urlsplit(request.url).path == expected, endpoint.endpoint
            else:
                # The site→drive resolution row has no sdk_contract case: it is rebuilt from
                # the generated builder directly.
                request = graph_client(HandlerGraphAdapter()).sites.by_site_id(SITE).drive.to_get_request_information()
                assert request.http_method.value == endpoint.method
                assert urlsplit(request.url).path == "/sites/site/drive"


# ======================================================================================
# Guard pins: every bound and refusal below fails a test when it is removed
# ======================================================================================


@pytest.mark.parametrize(
    "service, operation, arguments, expected",
    [
        ("sharepoint", "search", {"query": "   "}, "query"),
        ("sharepoint", "search", {"query": "hello", "drive_id": ""}, "drive_id"),
        ("sharepoint", "read", {"drive_id": DRIVE}, "item"),
        ("sharepoint", "read", {"drive_id": DRIVE, "item_path": "/etc/passwd"}, "item_path"),
        ("onedrive", "read", {"drive_id": DRIVE, "item_id": " "}, "item_id"),
        ("onedrive", "search", {"query": "hello", "site_id": SITE, "drive_id": DRIVE}, "at most one"),
    ],
)
def test_an_argument_the_handler_cannot_express_is_refused_before_any_client(
    service, operation, arguments, expected
):
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
    assert calls == []
    assert adapter.requests == []


def test_a_handler_refuses_a_context_of_the_wrong_shape():
    from microsoft365.execution import ExecutionError

    with pytest.raises(ExecutionError) as caught:
        invocation(SHAREPOINT_READ)(
            {"action": "read", "drive_id": DRIVE, "item_id": ITEM}, {"settings": settings()}
        )

    assert caught.value.category == "configuration_error"


def test_the_handler_module_never_holds_a_secret_credential_or_client():
    source = (REPO_ROOT / "microsoft365" / "handlers" / "files.py").read_text(encoding="utf-8")
    for forbidden in (
        "get_secret",
        "secret_scope",
        "ClientSecretCredential",
        "GraphServiceClient",
        "azure.identity",
    ):
        assert forbidden not in source, forbidden


def test_the_handler_module_never_assembles_a_url_or_imports_a_transport():
    source = (REPO_ROOT / "microsoft365" / "handlers" / "files.py").read_text(encoding="utf-8")
    for forbidden in ("https://", "http://", "urlsplit", "urljoin", ".url =", "RequestsClient"):
        assert forbidden not in source, forbidden


def test_known_limitations_records_the_wp7_state():
    """Invariant 6: the document moves in the same commit as the capacity it describes."""
    text = (REPO_ROOT / "docs" / "known-limitations.md").read_text(encoding="utf-8")

    for key in sorted(WP7_HANDLERS):
        assert key in text, key
    assert "upload_files" in text
    assert "sharepoint.search" in text
    assert "not_tested" in text
