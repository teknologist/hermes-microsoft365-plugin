"""Strict Planner contracts: container, identifiers, body, ETag and permission claims.

WP2 pinned, for every Planner operation, *which container it indexes by*, which
identifiers must be non-blank before a request is built at all, which endpoint the request
targets, and which application role is claimed for that endpoint. The tables in this
module are literal: they are the specification, and the administrative registry in
``microsoft365/contract.py`` must match them, so removing or weakening a guard in the
registry fails a test here.

No ``SimpleNamespace``, no ``MagicMock``, no permissive ``__getattr__``: every request is
built by the real generated builder of the installed SDK. The concrete ``RequestAdapter``
double raises from every ``send_*`` method -- the network is unreachable -- while request
bodies still serialize through the installed Kiota JSON writer factory.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

import pytest
from kiota_abstractions.request_adapter import RequestAdapter

REPO_ROOT = Path(__file__).resolve().parents[1]
PERMISSION_MATRIX = REPO_ROOT / "microsoft365" / "references" / "graph-permissions.md"
KNOWN_LIMITATIONS = REPO_ROOT / "docs" / "known-limitations.md"

#: The exact executable set of this milestone, including WP9 reads. Spelled out literally, so neither a flipped write nor a lost
#: read flag passes.
EXECUTABLE_READS = frozenset(
    {
        "outlook.search", "outlook.read", "outlook.list_attachments", "outlook.download_attachment", "calendar.search",
        "sharepoint.search", "sharepoint.read", "sharepoint.download_files",
        "onedrive.search", "onedrive.read", "onedrive.download_files",
        "teams.list_teams", "teams.list_channels",
        "todo.list_task_lists", "todo.search", "todo.read",
        "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read",
    }
)

#: Writes implemented and this milestone keeps non-executable (R5).
WITHHELD_WRITES = frozenset(
    {
        "outlook.create_draft", "outlook.send", "calendar.create_events", "calendar.update_events",
        "sharepoint.upload_files", "onedrive.upload_files",
        "todo.create_tasks", "todo.update_tasks", "planner.create_tasks", "planner.update_tasks",
    }
)


class StrictPlannerTransportAdapter(RequestAdapter):
    """Serializes real generated models, never sends."""

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

    return GraphServiceClient(request_adapter=StrictPlannerTransportAdapter())


def serialized_body(request) -> dict:
    assert request.content, "expected a serialized request body"
    return json.loads(request.content.decode("utf-8"))


def if_match(request) -> set:
    """The If-Match values actually configured on the request, or an empty set."""
    return request.headers.get("if-match") or set()


def camel(identifier: str) -> str:
    """``plan_id`` -> ``planId``: the Graph body name of a declared body identifier."""
    head, *rest = identifier.split("_")
    return head + "".join(part.title() for part in rest)


# --- the literal specification -------------------------------------------------------

#: Sample values for the identifiers the registry declares.
SAMPLE_IDENTIFIERS = {"group_id": "group", "plan_id": "plan", "planner_task_id": "task"}

#: (operation, method, path_template, container, path identifiers, body identifiers,
#:  application roles, required headers, contract dispatcher, contract case, page)
EXPECTED_ENDPOINTS = (
    (
        "planner.list_plans", "GET", "/planner/plans", "tenant_or_group", (), (),
        ("Tasks.Read.All",), (), "build_collection_request_information", "planner_plans", "https://learn.microsoft.com/en-us/graph/api/planner-list-plans",
    ),
    (
        "planner.list_plans", "GET", "/groups/{group_id}/planner/plans", "tenant_or_group",
        ("group_id",), (), ("Tasks.Read.All",), (), "build_collection_request_information",
        "planner_plans", "https://learn.microsoft.com/en-us/graph/api/planner-list-plans",
    ),
    (
        "planner.list_buckets", "GET", "/planner/plans/{plan_id}/buckets", "plan",
        ("plan_id",), (), ("Tasks.Read.All",), (), "build_collection_request_information",
        "planner_buckets", "https://learn.microsoft.com/en-us/graph/api/planner-list-buckets",
    ),
    (
        "planner.list_tasks", "GET", "/planner/plans/{plan_id}/tasks", "plan", ("plan_id",), (),
        ("Tasks.Read.All",), (), "build_collection_request_information", "planner_tasks",
        "https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks",
    ),
    (
        "planner.read", "GET", "/planner/tasks/{planner_task_id}", "task",
        ("planner_task_id",), (), ("Tasks.Read.All",), (), "build_item_request_information",
        "planner_read", "https://learn.microsoft.com/en-us/graph/api/plannertask-get",
    ),
    (
        "planner.create_tasks", "POST", "/planner/tasks", "plan_and_bucket", (),
        ("plan_id", "bucket_id"), ("Tasks.ReadWrite.All",), (),
        "build_write_request_information", "planner_create_tasks",
        "https://learn.microsoft.com/en-us/graph/api/planner-post-tasks",
    ),
    (
        "planner.update_tasks", "PATCH", "/planner/tasks/{planner_task_id}", "task",
        ("planner_task_id",), (), ("Tasks.ReadWrite.All",), ("If-Match",),
        "build_write_request_information", "planner_update_tasks",
        "https://learn.microsoft.com/en-us/graph/api/plannertask-update",
    ),
)

#: operation -> (container, required identifiers, optional identifiers, required headers)
EXPECTED_OPERATION_CONTRACTS = {
    "planner.list_plans": ("tenant_or_group", (), ("group_id",), ()),
    "planner.list_buckets": ("plan", ("plan_id",), (), ()),
    "planner.list_tasks": ("plan", ("plan_id",), (), ()),
    "planner.read": ("task", ("planner_task_id",), (), ()),
    "planner.create_tasks": ("plan_and_bucket", ("plan_id", "bucket_id"), (), ()),
    "planner.update_tasks": ("task", ("planner_task_id",), (), ("If-Match",)),
}

#: operation -> the least-privileged application roles its endpoints claim
EXPECTED_CLAIMS = {
    "planner.list_plans": ("Tasks.Read.All",),
    "planner.list_buckets": ("Tasks.Read.All",),
    "planner.list_tasks": ("Tasks.Read.All",),
    "planner.read": ("Tasks.Read.All",),
    "planner.create_tasks": ("Tasks.ReadWrite.All",),
    "planner.update_tasks": ("Tasks.ReadWrite.All",),
}

#: sample values for the required identifiers of every operation that declares one
SAMPLE_REQUIRED = {
    "planner.list_buckets": {"plan_id": "plan"},
    "planner.list_tasks": {"plan_id": "plan"},
    "planner.read": {"planner_task_id": "task"},
    "planner.create_tasks": {"plan_id": "plan", "bucket_id": "bucket"},
    "planner.update_tasks": {"planner_task_id": "task"},
}


def _list_plans_tenant_wide(client):
    from microsoft365.sdk_contract import build_collection_request_information

    return build_collection_request_information(client, "planner_plans", owner_id="owner", limit=7)


def _list_plans_for_group(client):
    from microsoft365.sdk_contract import build_collection_request_information

    return build_collection_request_information(
        client, "planner_plans", group_id="group", owner_id="owner", limit=7
    )


def _list_buckets(client):
    from microsoft365.sdk_contract import build_collection_request_information

    return build_collection_request_information(client, "planner_buckets", plan_id="plan", limit=7)


def _list_tasks(client):
    from microsoft365.sdk_contract import build_collection_request_information

    return build_collection_request_information(client, "planner_tasks", plan_id="plan", limit=7)


def _read_task(client):
    from microsoft365.sdk_contract import build_item_request_information

    return build_item_request_information(client, "planner_read", planner_task_id="task", select=["id", "title"])


def _create_task(client):
    from microsoft365.sdk_contract import build_write_request_information

    return build_write_request_information(
        client, "planner_create_tasks", plan_id="plan", bucket_id="bucket", title="Draft"
    )


def _update_task(client):
    from microsoft365.sdk_contract import build_write_request_information

    return build_write_request_information(
        client, "planner_update_tasks", planner_task_id="task", etag='"etag"', update_fields={"title": "Renamed"}
    )


#: (method, path_template) -> the strict call that must render exactly that endpoint
STRICT_ENDPOINTS = {
    ("GET", "/planner/plans"): {
        "build": _list_plans_tenant_wide,
        "contract_call": "build_collection_request_information",
        "contract_case": "planner_plans",
        "query_parameters": {"%24filter": "owner eq 'owner'", "%24top": 7},
        "serialized_body": None,
        "if_match": set(),
    },
    ("GET", "/groups/{group_id}/planner/plans"): {
        "build": _list_plans_for_group,
        "contract_call": "build_collection_request_information",
        "contract_case": "planner_plans",
        "query_parameters": {"%24filter": "owner eq 'owner'", "%24top": 7},
        "serialized_body": None,
        "if_match": set(),
    },
    ("GET", "/planner/plans/{plan_id}/buckets"): {
        "build": _list_buckets,
        "contract_call": "build_collection_request_information",
        "contract_case": "planner_buckets",
        "query_parameters": {"%24top": 7},
        "serialized_body": None,
        "if_match": set(),
    },
    ("GET", "/planner/plans/{plan_id}/tasks"): {
        "build": _list_tasks,
        "contract_call": "build_collection_request_information",
        "contract_case": "planner_tasks",
        "query_parameters": {"%24top": 7},
        "serialized_body": None,
        "if_match": set(),
    },
    ("GET", "/planner/tasks/{planner_task_id}"): {
        "build": _read_task,
        "contract_call": "build_item_request_information",
        "contract_case": "planner_read",
        "query_parameters": {"%24select": ["id", "title"]},
        "serialized_body": None,
        "if_match": set(),
    },
    ("POST", "/planner/tasks"): {
        "build": _create_task,
        "contract_call": "build_write_request_information",
        "contract_case": "planner_create_tasks",
        "query_parameters": None,
        "serialized_body": {"bucketId": "bucket", "planId": "plan", "title": "Draft"},
        "if_match": set(),
    },
    ("PATCH", "/planner/tasks/{planner_task_id}"): {
        "build": _update_task,
        "contract_call": "build_write_request_information",
        "contract_case": "planner_update_tasks",
        "query_parameters": None,
        "serialized_body": {"title": "Renamed"},
        "if_match": {'"etag"'},
    },
}


# --- the registry must match the specification ---------------------------------------


def test_registry_declares_exactly_the_pinned_planner_endpoints():
    from microsoft365.contract import OPERATION_REGISTRY, PLANNER_ENDPOINTS, PLANNER_OPERATIONS

    declared = tuple(
        (
            f"planner.{endpoint.operation}",
            endpoint.method,
            endpoint.path_template,
            endpoint.container,
            endpoint.path_identifiers,
            endpoint.body_identifiers,
            endpoint.application_permissions,
            endpoint.required_headers,
            endpoint.contract_call,
            endpoint.contract_case,
            endpoint.documentation_page,
        )
        for endpoint in PLANNER_ENDPOINTS
    )

    assert declared == EXPECTED_ENDPOINTS
    assert PLANNER_OPERATIONS == tuple(EXPECTED_OPERATION_CONTRACTS)
    for key in PLANNER_OPERATIONS:
        expected = tuple(
            f"{method} {path}" for (operation, method, path, *_) in EXPECTED_ENDPOINTS if operation == key
        )
        assert OPERATION_REGISTRY[key].endpoints == expected
        assert OPERATION_REGISTRY[key].contract_cases == tuple(
            (call, case)
            for (operation, _, _, _, _, _, _, _, call, case, _) in EXPECTED_ENDPOINTS
            if operation == key
        )


def test_planner_container_and_identifier_semantics_are_pinned():
    from microsoft365.contract import OPERATION_REGISTRY

    declared = {
        key: (
            definition.container,
            definition.required_identifiers,
            definition.optional_identifiers,
            definition.required_headers,
        )
        for key, definition in OPERATION_REGISTRY.items()
        if key.startswith("planner.")
    }

    assert declared == EXPECTED_OPERATION_CONTRACTS


def test_every_path_placeholder_is_a_declared_path_identifier():
    from microsoft365.contract import PLANNER_ENDPOINTS

    for endpoint in PLANNER_ENDPOINTS:
        placeholders = set(re.findall(r"\{(\w+)\}", endpoint.path_template))
        assert placeholders == set(endpoint.path_identifiers), endpoint.endpoint
        assert not (set(endpoint.path_identifiers) & set(endpoint.body_identifiers))


def test_strict_table_covers_exactly_the_endpoints_the_registry_declares():
    from microsoft365.contract import PLANNER_ENDPOINTS

    declared = {(endpoint.method, endpoint.path_template) for endpoint in PLANNER_ENDPOINTS}

    assert len(declared) == len(PLANNER_ENDPOINTS), "a Planner endpoint is declared twice"
    assert declared == set(STRICT_ENDPOINTS)
    for (method, path_template), row in STRICT_ENDPOINTS.items():
        endpoint = next(
            item for item in PLANNER_ENDPOINTS if (item.method, item.path_template) == (method, path_template)
        )
        assert (endpoint.contract_call, endpoint.contract_case) == (
            row["contract_call"],
            row["contract_case"],
        )


@pytest.mark.parametrize("method,path_template", list(STRICT_ENDPOINTS))
def test_declared_planner_endpoint_is_rendered_exactly_by_the_real_builder(
    graph_client, method, path_template
):
    from microsoft365.contract import PLANNER_ENDPOINTS

    endpoint = next(
        item for item in PLANNER_ENDPOINTS if (item.method, item.path_template) == (method, path_template)
    )
    row = STRICT_ENDPOINTS[(method, path_template)]
    request = row["build"](graph_client)
    rendered = path_template.format(**SAMPLE_IDENTIFIERS)

    assert request.http_method.value == method
    assert urlsplit(request.url).path == rendered
    if row["query_parameters"] is not None:
        assert request.query_parameters == row["query_parameters"]
    assert if_match(request) == row["if_match"]
    if row["serialized_body"] is None:
        assert not request.content
    else:
        assert serialized_body(request) == row["serialized_body"]

    # the declared container identifiers are really the ones the request carries
    for name in endpoint.path_identifiers:
        assert SAMPLE_IDENTIFIERS[name] in rendered
    for name in endpoint.body_identifiers:
        assert camel(name) in serialized_body(request)


# --- blank containers are refused before anything is sent -----------------------------


@pytest.mark.parametrize(
    ("key", "arguments", "name", "endpoint"),
    [
        ("planner.list_buckets", {}, "plan_id", "GET /planner/plans/{plan_id}/buckets"),
        ("planner.list_buckets", {"plan_id": ""}, "plan_id", "GET /planner/plans/{plan_id}/buckets"),
        ("planner.list_buckets", {"plan_id": "   "}, "plan_id", "GET /planner/plans/{plan_id}/buckets"),
        ("planner.list_buckets", {"plan_id": None}, "plan_id", "GET /planner/plans/{plan_id}/buckets"),
        ("planner.list_tasks", {}, "plan_id", "GET /planner/plans/{plan_id}/tasks"),
        ("planner.list_tasks", {"plan_id": "\t"}, "plan_id", "GET /planner/plans/{plan_id}/tasks"),
        ("planner.read", {}, "planner_task_id", "GET /planner/tasks/{planner_task_id}"),
        ("planner.read", {"planner_task_id": " "}, "planner_task_id", "GET /planner/tasks/{planner_task_id}"),
        ("planner.create_tasks", {"plan_id": "plan", "bucket_id": "  "}, "bucket_id", "POST /planner/tasks"),
        ("planner.create_tasks", {"bucket_id": "bucket"}, "plan_id", "POST /planner/tasks"),
        ("planner.create_tasks", {}, "plan_id", "POST /planner/tasks"),
        ("planner.update_tasks", {}, "planner_task_id", "PATCH /planner/tasks/{planner_task_id}"),
        ("planner.update_tasks", {"planner_task_id": ""}, "planner_task_id", "PATCH /planner/tasks/{planner_task_id}"),
    ],
)
def test_absent_or_blank_required_identifiers_are_refused_with_the_endpoint(
    key, arguments, name, endpoint
):
    from microsoft365.contract import PlannerContractError, validate_planner_identifiers

    with pytest.raises(PlannerContractError) as excinfo:
        validate_planner_identifiers(key, arguments)

    message = str(excinfo.value)
    assert f"{name} is required for {key}" in message
    assert endpoint in message


@pytest.mark.parametrize(
    ("key", "arguments", "name"),
    [
        ("planner.list_tasks", {"plan_id": 7}, "plan_id"),
        ("planner.list_tasks", {"plan_id": True}, "plan_id"),
        ("planner.read", {"planner_task_id": ["task"]}, "planner_task_id"),
        ("planner.create_tasks", {"plan_id": "plan", "bucket_id": {"id": "bucket"}}, "bucket_id"),
        ("planner.list_plans", {"group_id": 7}, "group_id"),
        ("planner.list_plans", {"group_id": ""}, "group_id"),
        ("planner.list_plans", {"group_id": "  "}, "group_id"),
    ],
)
def test_non_string_or_blank_identifiers_are_refused_without_coercion(key, arguments, name):
    from microsoft365.contract import PlannerContractError, validate_planner_identifiers

    with pytest.raises(PlannerContractError, match=re.escape(name)) as excinfo:
        validate_planner_identifiers(key, arguments)

    assert key in str(excinfo.value)
    assert "must be a non-blank string" in str(excinfo.value)


def test_plan_listing_container_is_optional_but_never_a_silent_blank_fallback(graph_client):
    from microsoft365.contract import validate_planner_identifiers
    from microsoft365.sdk_contract import build_collection_request_information

    assert validate_planner_identifiers("planner.list_plans", {}) == {}
    assert validate_planner_identifiers("planner.list_plans", {"group_id": None}) == {}
    assert validate_planner_identifiers("planner.list_plans", {"group_id": " group "}) == {"group_id": "group"}

    tenant_wide = build_collection_request_information(graph_client, "planner_plans", owner_id="owner")
    group_owned = build_collection_request_information(
        graph_client, "planner_plans", group_id="group", owner_id="owner"
    )
    assert urlsplit(tenant_wide.url).path == "/planner/plans"
    assert urlsplit(group_owned.url).path == "/groups/group/planner/plans"


def test_registry_guard_and_sdk_guard_refuse_a_blank_group_container_independently(graph_client):
    from microsoft365.contract import PlannerContractError, validate_planner_identifiers
    from microsoft365.sdk_contract import build_collection_request_information

    with pytest.raises(PlannerContractError):
        validate_planner_identifiers("planner.list_plans", {"group_id": "   "})

    with pytest.raises(ValueError) as excinfo:
        build_collection_request_information(graph_client, "planner_plans", group_id="   ")
    # the SDK-level guard is a separate one, not the registry error bubbling through
    assert excinfo.type is ValueError
    assert "group_id is required" in str(excinfo.value)


def test_normalized_identifiers_render_the_declared_endpoint(graph_client):
    from microsoft365.contract import validate_planner_identifiers
    from microsoft365.sdk_contract import build_collection_request_information

    identifiers = validate_planner_identifiers("planner.list_tasks", {"plan_id": " plan "})

    assert identifiers == {"plan_id": "plan"}
    request = build_collection_request_information(graph_client, "planner_tasks", **identifiers)
    assert urlsplit(request.url).path == "/planner/plans/plan/tasks"


@pytest.mark.parametrize(
    ("key", "arguments"),
    [
        ("planner.list_plans", []),
        ("planner.list_tasks", "plan_id=plan"),
        ("planner.read", None),
        ("planner.read", 7),
    ],
)
def test_planner_arguments_must_be_a_mapping(key, arguments):
    from microsoft365.contract import PlannerContractError, validate_planner_identifiers

    with pytest.raises(PlannerContractError, match="arguments must be a mapping"):
        validate_planner_identifiers(key, arguments)


@pytest.mark.parametrize("key", ["planner.delete", "planner.list_buckets.tasks", "outlook.search", ""])
def test_only_declared_planner_operations_can_be_validated(key):
    from microsoft365.contract import PlannerContractError, validate_planner_identifiers

    with pytest.raises(PlannerContractError, match=re.escape(f"{key} is not a declared Planner operation")):
        validate_planner_identifiers(key, {})


# --- ETag and partial body boundaries -------------------------------------------------


def test_declared_required_headers_are_enforced_by_the_sdk_contract(graph_client):
    from microsoft365.contract import PLANNER_ENDPOINTS
    from microsoft365.sdk_contract import build_write_request_information

    assert {endpoint.contract_case: endpoint.required_headers for endpoint in PLANNER_ENDPOINTS} == {
        "planner_plans": (),
        "planner_buckets": (),
        "planner_tasks": (),
        "planner_read": (),
        "planner_create_tasks": (),
        "planner_update_tasks": ("If-Match",),
    }

    with pytest.raises(ValueError, match="etag is required for planner_update_tasks"):
        build_write_request_information(
            graph_client, "planner_update_tasks", planner_task_id="task", update_fields={"title": "Renamed"}
        )

    request = build_write_request_information(
        graph_client, "planner_update_tasks", planner_task_id="task", etag='"etag"', update_fields={"title": "Renamed"}
    )
    assert if_match(request) == {'"etag"'}
    assert serialized_body(request) == {"title": "Renamed"}


def test_planner_patch_accepts_the_real_assignment_model_but_refuses_a_bare_list(graph_client):
    from msgraph.generated.models.planner_assignments import PlannerAssignments
    from microsoft365.sdk_contract import build_write_request_information

    accepted = build_write_request_information(
        graph_client,
        "planner_update_tasks",
        planner_task_id="task",
        etag='"etag"',
        update_fields={"assignments": PlannerAssignments()},
    )
    assert serialized_body(accepted) == {"assignments": {}}

    # The shared partial-model helper deliberately cannot express a *list* of generated
    # models; it refuses instead of silently dropping the field.
    with pytest.raises(
        ValueError, match=re.escape("field assignments of planner_update_tasks must be a scalar or a generated model")
    ):
        build_write_request_information(
            graph_client,
            "planner_update_tasks",
            planner_task_id="task",
            etag='"etag"',
            update_fields={"assignments": [PlannerAssignments()]},
        )


# --- honest status labels -------------------------------------------------------------


def test_no_planner_operation_is_executable_while_its_handler_does_not_exist():
    from microsoft365.contract import OPERATION_REGISTRY, operation_status
    from microsoft365.registration import HANDLER_TABLE

    planner_keys = sorted(key for key in OPERATION_REGISTRY if key.startswith("planner."))

    assert planner_keys == sorted(EXPECTED_OPERATION_CONTRACTS)
    for key in planner_keys:
        definition = OPERATION_REGISTRY[key]
        assert definition.implementation_status == "implemented", key
        assert definition.executable is (key in EXECUTABLE_READS), key
        assert key in HANDLER_TABLE, key

        service, operation = key.split(".", 1)
        status = operation_status("application", service, operation)
        assert status.auth_status == "supported"
        assert status.implementation_status == "implemented"
        assert status.executable is (key in EXECUTABLE_READS)

    # The milestone's exact sets, spelled out: three verified Outlook/Calendar reads are
    # executable, and a handler exists for exactly those reads plus the four writes WP6
    # implemented and this milestone deliberately withholds (R5). Nothing else has a handler.
    assert {
        key for key, definition in OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_READS)
    assert set(HANDLER_TABLE) == set(EXECUTABLE_READS) | set(WITHHELD_WRITES) | {"teams.search_messages", "teams.send_messages"}


def test_implementation_status_labels_are_explicit_and_backed_by_declared_evidence():
    from microsoft365.contract import IMPLEMENTATION_STATUS_LABELS, OPERATION_REGISTRY
    from microsoft365.registration import HANDLER_TABLE

    for key, definition in OPERATION_REGISTRY.items():
        assert definition.implementation_status in IMPLEMENTATION_STATUS_LABELS, key
        if definition.implementation_status == "contract_verified":
            assert definition.endpoints, key
            assert definition.contract_cases, key
        if definition.implementation_status == "implemented" or definition.executable:
            assert definition.contract_cases, key
            assert key in HANDLER_TABLE, key
        # a handler is never registered for an operation the registry does not describe
        if key in HANDLER_TABLE:
            assert definition.implementation_status == "implemented", key


def test_planner_permission_claims_are_endpoint_specific_and_not_promoted():
    from microsoft365.contract import (
        OPERATION_REGISTRY,
        PERMISSION_CLAIM_STATUSES,
        PLANNER_ENDPOINTS,
    )

    for endpoint in PLANNER_ENDPOINTS:
        assert endpoint.application_permissions, endpoint.endpoint
        assert endpoint.claim_status in PERMISSION_CLAIM_STATUSES
        assert endpoint.claim_status == "documented_not_verified", endpoint.endpoint
        assert endpoint.documentation_page in ("",) or endpoint.documentation_page.startswith(
            "https://learn.microsoft.com/en-us/graph/api/"
        )

    assert {key: OPERATION_REGISTRY[key].permissions for key in EXPECTED_CLAIMS} == EXPECTED_CLAIMS
    assert {
        key: OPERATION_REGISTRY[key].permission_claim_statuses for key in EXPECTED_CLAIMS
    } == {key: ("documented_not_verified",) for key in EXPECTED_CLAIMS}

    for key, roles in EXPECTED_CLAIMS.items():
        endpoints = [item for item in PLANNER_ENDPOINTS if f"planner.{item.operation}" == key]
        assert endpoints
        for endpoint in endpoints:
            assert endpoint.application_permissions == roles, endpoint.endpoint


def test_preflight_reports_planner_claims_without_claiming_execution():
    from microsoft365.contract import Settings
    from microsoft365.preflight import build_preflight

    payload = build_preflight(
        Settings.from_mapping({"capabilities": {"planner": True}}), sdk_available=True
    )

    assert sorted(payload["operation_status"]) == sorted(EXPECTED_CLAIMS)
    for key, status in payload["operation_status"].items():
        assert status["permissions"] == list(EXPECTED_CLAIMS[key])
        assert status["implementation_status"] == "implemented"
        assert status["executable"] is (key in EXECUTABLE_READS)
        assert status["write"] is (key in {"planner.create_tasks", "planner.update_tasks"})
    assert payload["locally_ready"] is False


# --- the documents must match the registry -------------------------------------------


def planner_matrix_rows() -> list[tuple[str, ...]]:
    rows = []
    for line in PERMISSION_MATRIX.read_text(encoding="utf-8").splitlines():
        cells = tuple(cell.strip().strip("`") for cell in line.split("|")[1:-1])
        if len(cells) == 6 and cells[0].startswith("planner."):
            rows.append(cells)
    return rows


def test_permission_matrix_reports_one_row_per_planner_endpoint():
    from microsoft365.contract import OPERATION_REGISTRY, PLANNER_ENDPOINTS

    expected = sorted(
        (
            f"planner.{endpoint.operation}",
            endpoint.endpoint,
            " ".join(endpoint.application_permissions),
            endpoint.claim_status,
            OPERATION_REGISTRY[f"planner.{endpoint.operation}"].implementation_status,
            endpoint.documentation_page or "not recorded",
        )
        for endpoint in PLANNER_ENDPOINTS
    )

    assert sorted(planner_matrix_rows()) == expected


def test_permission_matrix_keeps_no_blanket_planner_claim():
    text = PERMISSION_MATRIX.read_text(encoding="utf-8")
    rows = [
        [cell.strip().strip("`") for cell in line.split("|")[1:-1]]
        for line in text.splitlines()
        if line.startswith("|")
    ]

    # no table row may cover Planner as one blanket operation family again
    assert [
        cells[0]
        for cells in rows
        if cells[:1] and cells[0] in {"Planner list/read", "Planner create/update"}
    ] == []
    assert "documented_not_verified" in text


def test_known_limitations_records_the_planner_contract_state():
    text = KNOWN_LIMITATIONS.read_text(encoding="utf-8")

    assert "strict offline request-contract tests" in text
    assert "The four Planner reads are executable" in text
    assert "documented_not_verified" in text
    # the pre-WP2 text claimed these were still unresolved
    assert "Planner ownership/container semantics" not in text
    assert "Planner write serialization/ETag behavior" not in text
