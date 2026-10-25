"""Strict Microsoft To Do contracts: field semantics, status enum and honest write support.

WP3 pins, for every To Do operation, the endpoint it targets, the identifiers that must be
non-blank before anything is sent, the task fields a write may carry, and how much is
actually known about application-mode write support. The tables here are literal: they are
the specification, and ``microsoft365/contract.py`` must match them, so removing or
weakening a guard in the registry fails a test in this module.

Two facts are asserted with real installed-SDK objects instead of prose:

* the generated ``TodoTask`` model has **no** ``subject`` property, so a To Do request that
  carries ``subject`` -- as a filter, a ``$select`` value or a body field -- is refused
  rather than being sent and silently ignored by Graph;
* Kiota's JSON writer emits nothing at all for an enum-typed field whose value is not one
  of the generated ``TaskStatus`` members (verified offline: a plain ``"inProgress"``
  string serializes to ``{}``), so ``status`` must be mapped to that member before the
  request body is built.

No ``SimpleNamespace``, no ``MagicMock``, no permissive ``__getattr__``: every request is
built by the real generated builder of the installed SDK and serialized through the
installed Kiota JSON writer factory. The concrete ``RequestAdapter`` double raises from
every ``send_*`` method, so no code path here can reach a network.
"""
from __future__ import annotations

import dataclasses
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


class StrictTodoTransportAdapter(RequestAdapter):
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

    return GraphServiceClient(request_adapter=StrictTodoTransportAdapter())


def serialized_body(request) -> dict:
    assert request.content, "expected a serialized request body"
    return json.loads(request.content.decode("utf-8"))


def if_match(request) -> set:
    """The If-Match values actually configured on the request, or an empty set."""
    return request.headers.get("if-match") or set()


# --- the literal specification -------------------------------------------------------

#: Sample values for the identifiers the registry declares.
SAMPLE_IDENTIFIERS = {"user_id": "user", "todo_list_id": "list", "todo_task_id": "task"}

#: (operation, method, path_template, container, path identifiers, application roles,
#:  required headers, contract dispatcher, contract case, recorded page)
EXPECTED_ENDPOINTS = (
    (
        "todo.list_task_lists", "GET", "/users/{user_id}/todo/lists", "user", ("user_id",),
        ("Tasks.Read.All",), (), "build_collection_request_information", "todo_lists", "https://learn.microsoft.com/en-us/graph/api/todo-list-lists",
    ),
    (
        "todo.search", "GET", "/users/{user_id}/todo/lists/{todo_list_id}/tasks", "list",
        ("user_id", "todo_list_id"), ("Tasks.Read.All",), (),
        "build_collection_request_information", "todo_tasks",
        "https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks",
    ),
    (
        "todo.read", "GET", "/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        "task", ("user_id", "todo_list_id", "todo_task_id"), ("Tasks.Read.All",), (),
        "build_item_request_information", "todo_read", "https://learn.microsoft.com/en-us/graph/api/todotask-get",
    ),
    (
        "todo.create_tasks", "POST", "/users/{user_id}/todo/lists/{todo_list_id}/tasks", "list",
        ("user_id", "todo_list_id"), ("Tasks.ReadWrite.All",), (),
        "build_write_request_information", "todo_create_tasks", "https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks",
    ),
    (
        "todo.update_tasks", "PATCH",
        "/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}", "task",
        ("user_id", "todo_list_id", "todo_task_id"), ("Tasks.ReadWrite.All",), (),
        "build_write_request_information", "todo_update_tasks", "https://learn.microsoft.com/en-us/graph/api/todotask-update",
    ),
)

#: operation -> (container, required identifiers, optional identifiers, required headers)
EXPECTED_OPERATION_CONTRACTS = {
    "todo.list_task_lists": ("user", ("user_id",), (), ()),
    "todo.search": ("list", ("user_id", "todo_list_id"), (), ()),
    "todo.read": ("task", ("user_id", "todo_list_id", "todo_task_id"), (), ()),
    "todo.create_tasks": ("list", ("user_id", "todo_list_id"), (), ()),
    "todo.update_tasks": ("task", ("user_id", "todo_list_id", "todo_task_id"), (), ()),
}

#: operation -> the least-privileged application roles its endpoint claims
EXPECTED_CLAIMS = {
    "todo.list_task_lists": ("Tasks.Read.All",),
    "todo.search": ("Tasks.Read.All",),
    "todo.read": ("Tasks.Read.All",),
    "todo.create_tasks": ("Tasks.ReadWrite.All",),
    "todo.update_tasks": ("Tasks.ReadWrite.All",),
}

TODO_WRITE_KEYS = ("todo.create_tasks", "todo.update_tasks")
TODO_READ_KEYS = ("todo.list_task_lists", "todo.search", "todo.read")

#: The only To Do task fields a write may carry (the recorded field policy).
EXPECTED_WRITABLE_FIELDS = ("title", "body", "due_date_time", "status")

#: Fields a To Do write must refuse: none of these is accepted by the endpoint contract.
REFUSED_WRITE_FIELDS = (
    "subject",  # no To Do model has this property; To Do uses title
    "id",
    "odata_type",
    "created_date_time",
    "last_modified_date_time",
    "body_last_modified_date_time",
    "completed_date_time",
    "has_attachments",
    "attachments",
    "attachment_sessions",
    "extensions",
    "importance",
    "is_reminder_on",
    "reminder_date_time",
    "start_date_time",
    "categories",
    "checklist_items",
    "linked_resources",
    "recurrence",
)


def _list_task_lists(client):
    from microsoft365.sdk_contract import build_collection_request_information

    return build_collection_request_information(client, "todo_lists", user_id="user", limit=7)


def _search_tasks(client):
    from microsoft365.sdk_contract import build_collection_request_information

    return build_collection_request_information(
        client, "todo_tasks", user_id="user", todo_list_id="list", query="milk", limit=7
    )


def _read_task(client):
    from microsoft365.sdk_contract import build_item_request_information

    return build_item_request_information(
        client, "todo_read", user_id="user", todo_list_id="list", todo_task_id="task",
        select=["id", "title", "status"],
    )


def _create_task(client):
    from microsoft365.sdk_contract import build_write_request_information

    return build_write_request_information(
        client, "todo_create_tasks", user_id="user", todo_list_id="list", title="Buy milk"
    )


def _update_task(client):
    from microsoft365.sdk_contract import build_write_request_information

    return build_write_request_information(
        client, "todo_update_tasks", user_id="user", todo_list_id="list", todo_task_id="task",
        update_fields={"title": "Buy oat milk"},
    )


#: (method, path_template) -> the strict call that must render exactly that endpoint
STRICT_ENDPOINTS = {
    ("GET", "/users/{user_id}/todo/lists"): {
        "build": _list_task_lists,
        "contract_call": "build_collection_request_information",
        "contract_case": "todo_lists",
        "query_parameters": {"%24top": 7},
        "serialized_body": None,
        "if_match": set(),
    },
    ("GET", "/users/{user_id}/todo/lists/{todo_list_id}/tasks"): {
        "build": _search_tasks,
        "contract_call": "build_collection_request_information",
        "contract_case": "todo_tasks",
        "query_parameters": {"%24filter": "contains(title,'milk')", "%24top": 7},
        "serialized_body": None,
        "if_match": set(),
    },
    ("GET", "/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}"): {
        "build": _read_task,
        "contract_call": "build_item_request_information",
        "contract_case": "todo_read",
        "query_parameters": {"%24select": ["id", "title", "status"]},
        "serialized_body": None,
        "if_match": set(),
    },
    ("POST", "/users/{user_id}/todo/lists/{todo_list_id}/tasks"): {
        "build": _create_task,
        "contract_call": "build_write_request_information",
        "contract_case": "todo_create_tasks",
        "query_parameters": None,
        "serialized_body": {"title": "Buy milk"},
        "if_match": set(),
    },
    ("PATCH", "/users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}"): {
        "build": _update_task,
        "contract_call": "build_write_request_information",
        "contract_case": "todo_update_tasks",
        "query_parameters": None,
        "serialized_body": {"title": "Buy oat milk"},
        "if_match": set(),
    },
}


# --- the registry must match the specification ---------------------------------------


def test_registry_declares_exactly_the_pinned_todo_endpoints():
    from microsoft365.contract import OPERATION_REGISTRY, TODO_ENDPOINTS, TODO_OPERATIONS

    declared = tuple(
        (
            f"todo.{endpoint.operation}",
            endpoint.method,
            endpoint.path_template,
            endpoint.container,
            endpoint.path_identifiers,
            endpoint.application_permissions,
            endpoint.required_headers,
            endpoint.contract_call,
            endpoint.contract_case,
            endpoint.documentation_page,
        )
        for endpoint in TODO_ENDPOINTS
    )

    assert declared == EXPECTED_ENDPOINTS
    assert TODO_OPERATIONS == tuple(EXPECTED_OPERATION_CONTRACTS)
    for key in TODO_OPERATIONS:
        expected_endpoints = tuple(
            f"{method} {path}"
            for (operation, method, path, *_) in EXPECTED_ENDPOINTS
            if operation == key
        )
        assert OPERATION_REGISTRY[key].endpoints == expected_endpoints
        assert OPERATION_REGISTRY[key].contract_cases == tuple(
            (call, case)
            for (operation, _, _, _, _, _, _, call, case, _) in EXPECTED_ENDPOINTS
            if operation == key
        )


def test_todo_container_and_identifier_semantics_are_pinned():
    from microsoft365.contract import OPERATION_REGISTRY

    declared = {
        key: (
            definition.container,
            definition.required_identifiers,
            definition.optional_identifiers,
            definition.required_headers,
        )
        for key, definition in OPERATION_REGISTRY.items()
        if key.startswith("todo.")
    }

    assert declared == EXPECTED_OPERATION_CONTRACTS


def test_every_path_placeholder_is_a_declared_path_identifier():
    from microsoft365.contract import TODO_ENDPOINTS

    for endpoint in TODO_ENDPOINTS:
        placeholders = set(re.findall(r"\{(\w+)\}", endpoint.path_template))
        assert placeholders == set(endpoint.path_identifiers), endpoint.endpoint
        assert endpoint.contract_call and endpoint.contract_case, endpoint.endpoint


def test_strict_table_covers_exactly_the_endpoints_the_registry_declares():
    from microsoft365.contract import TODO_ENDPOINTS

    declared = {(endpoint.method, endpoint.path_template) for endpoint in TODO_ENDPOINTS}

    assert len(declared) == len(TODO_ENDPOINTS), "a To Do endpoint is declared twice"
    assert declared == set(STRICT_ENDPOINTS)
    for (method, path_template), row in STRICT_ENDPOINTS.items():
        endpoint = next(
            item
            for item in TODO_ENDPOINTS
            if (item.method, item.path_template) == (method, path_template)
        )
        assert (endpoint.contract_call, endpoint.contract_case) == (
            row["contract_call"],
            row["contract_case"],
        )


@pytest.mark.parametrize("method,path_template", list(STRICT_ENDPOINTS))
def test_declared_todo_endpoint_is_rendered_exactly_by_the_real_builder(
    graph_client, method, path_template
):
    from microsoft365.contract import TODO_ENDPOINTS

    endpoint = next(
        item for item in TODO_ENDPOINTS if (item.method, item.path_template) == (method, path_template)
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

    for name in endpoint.path_identifiers:
        assert SAMPLE_IDENTIFIERS[name] in rendered


# --- the search and read path uses the real model fields -----------------------------


def _todo_task_request_builders():
    """Every To Do request the contract can build, as (label, builder) pairs."""
    from microsoft365.sdk_contract import (
        build_collection_request_information,
        build_item_request_information,
        build_write_request_information,
    )

    return (
        (
            "list_task_lists",
            lambda client: build_collection_request_information(
                client, "todo_lists", user_id="user", limit=7
            ),
        ),
        (
            "search",
            lambda client: build_collection_request_information(
                client, "todo_tasks", user_id="user", todo_list_id="list", query="milk", limit=7
            ),
        ),
        (
            "read",
            lambda client: build_item_request_information(
                client, "todo_read", user_id="user", todo_list_id="list", todo_task_id="task",
                select=["id", "title", "status"],
            ),
        ),
        (
            "create_tasks",
            lambda client: build_write_request_information(
                client, "todo_create_tasks", user_id="user", todo_list_id="list", title="Buy milk",
                status="notStarted",
            ),
        ),
        (
            "update_tasks",
            lambda client: build_write_request_information(
                client, "todo_update_tasks", user_id="user", todo_list_id="list",
                todo_task_id="task", update_fields={"status": "completed"},
            ),
        ),
    )


def test_no_todo_request_ever_carries_the_outlook_subject_field(graph_client):
    for label, build in _todo_task_request_builders():
        request = build(graph_client)
        assert "subject" not in request.url, label
        assert "subject" not in str(request.query_parameters), label
        if request.content:
            assert "subject" not in request.content.decode("utf-8"), label


def test_todo_task_search_sends_the_title_filter_and_no_status_filter(graph_client):
    from microsoft365.sdk_contract import build_collection_request_information

    request = build_collection_request_information(
        graph_client, "todo_tasks", user_id="user", todo_list_id="list", query="milk", limit=7
    )

    assert request.query_parameters["%24filter"] == "contains(title,'milk')"
    # the endpoint's acceptance of a $filter on `status` is not recorded offline, so no
    # status filter is sent; matching by status stays client-side (see the next test).
    assert "status" not in json.dumps(request.query_parameters)
    assert "subject" not in request.url


def test_todo_task_status_matches_client_side_with_the_generated_enum():
    from msgraph.generated.models.task_status import TaskStatus
    from microsoft365.sdk_contract import todo_task_status_matches

    assert todo_task_status_matches(TaskStatus.Completed, "completed") is True
    assert todo_task_status_matches("completed", TaskStatus.Completed) is True
    assert todo_task_status_matches(TaskStatus.Completed, TaskStatus.InProgress) is False
    # a task without a status, or one carrying a value the generated enum does not declare,
    # never matches and never raises
    assert todo_task_status_matches(None, "completed") is False
    assert todo_task_status_matches("somethingElse", "completed") is False
    assert todo_task_status_matches(7, "completed") is False


def test_todo_task_field_names_come_from_the_generated_model():
    from microsoft365.sdk_contract import todo_task_field_names

    names = todo_task_field_names()

    assert {"id", "title", "status", "dueDateTime", "importance", "isReminderOn"} <= names
    assert "subject" not in names
    assert len(names) >= 20


def test_todo_task_read_select_is_refused_for_a_field_the_model_does_not_declare(graph_client):
    from microsoft365.sdk_contract import build_item_request_information

    accepted = build_item_request_information(
        graph_client, "todo_read", user_id="user", todo_list_id="list", todo_task_id="task",
        select=["id", "title", "dueDateTime"],
    )
    assert accepted.query_parameters["%24select"] == ["id", "title", "dueDateTime"]

    with pytest.raises(ValueError, match="unknown field for todo_read: subject"):
        build_item_request_information(
            graph_client, "todo_read", user_id="user", todo_list_id="list", todo_task_id="task",
            select=["id", "subject"],
        )


# --- status is mapped to the generated enum before serialization ----------------------


def test_todo_task_status_maps_to_the_generated_enum():
    from msgraph.generated.models.task_status import TaskStatus
    from microsoft365.sdk_contract import todo_task_status

    assert todo_task_status("inProgress") is TaskStatus.InProgress
    assert todo_task_status(" inProgress ") is TaskStatus.InProgress
    assert todo_task_status(TaskStatus.NotStarted) is TaskStatus.NotStarted


@pytest.mark.parametrize("value", ["done", "DONE", "inprogress", "", "   ", 7, True, None, ["completed"]])
def test_todo_task_status_refuses_a_value_the_generated_enum_does_not_declare(value):
    from msgraph.generated.models.task_status import TaskStatus
    from microsoft365.sdk_contract import todo_task_status

    with pytest.raises(ValueError, match="TaskStatus members") as excinfo:
        todo_task_status(value)

    message = str(excinfo.value)
    for member in TaskStatus:
        assert member.value in message


def test_todo_update_serializes_status_as_the_generated_enum_member(graph_client):
    from msgraph.generated.models.task_status import TaskStatus
    from microsoft365.sdk_contract import build_write_request_information

    for supplied in ("inProgress", TaskStatus.InProgress):
        request = build_write_request_information(
            graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
            todo_task_id="task", update_fields={"title": "Buy oat milk", "status": supplied},
        )
        body = serialized_body(request)
        # a status the writer does not recognise as a generated member would be absent
        # from the body entirely, so equality here is the guard against a silent drop
        assert body == {"title": "Buy oat milk", "status": "inProgress"}


def test_todo_create_serializes_the_optional_status_as_the_generated_enum(graph_client):
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client, "todo_create_tasks", user_id="user", todo_list_id="list", title="Buy milk",
        status="notStarted",
    )

    assert serialized_body(request) == {"title": "Buy milk", "status": "notStarted"}

    with pytest.raises(ValueError, match="TaskStatus members"):
        build_write_request_information(
            graph_client, "todo_create_tasks", user_id="user", todo_list_id="list",
            title="Buy milk", status="done",
        )


# --- a field the endpoint ignores is refused, never dropped ---------------------------


def test_todo_task_write_field_policy_is_recorded_and_minimal():
    from msgraph.generated.models.todo_task import TodoTask
    from microsoft365.contract import TODO_TASK_WRITE_FIELDS

    assert tuple(TODO_TASK_WRITE_FIELDS) == EXPECTED_WRITABLE_FIELDS
    assert all(reason.strip() for reason in TODO_TASK_WRITE_FIELDS.values())

    model_fields = set(TodoTask.__dataclass_fields__)
    # every accepted field is a real generated property: nothing here is invented
    assert set(TODO_TASK_WRITE_FIELDS) <= model_fields
    # and the policy is a strict subset: the model declares more than the endpoint contract
    assert set(TODO_TASK_WRITE_FIELDS) < model_fields


def test_todo_write_refuses_an_empty_or_missing_field_mapping(graph_client):
    from microsoft365.contract import TodoContractError, validate_todo_task_fields
    from microsoft365.sdk_contract import build_write_request_information

    for arguments in ({}, None, "title=Buy milk", ["title"]):
        with pytest.raises(TodoContractError, match="non-empty mapping of task fields"):
            validate_todo_task_fields("todo_update_tasks", arguments)

    with pytest.raises(ValueError, match="non-empty mapping of task fields"):
        build_write_request_information(
            graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
            todo_task_id="task", update_fields={},
        )


def test_todo_write_field_policy_is_bound_to_the_declared_write_cases():
    from microsoft365.contract import (
        TODO_TASK_WRITE_CASES,
        TODO_WRITE_OPERATIONS,
        TodoContractError,
        validate_todo_task_fields,
    )

    assert tuple(sorted(TODO_WRITE_OPERATIONS)) == tuple(sorted(TODO_WRITE_KEYS))
    # the policy is keyed by the sdk_contract case name the request builder uses, and every
    # one of those names comes from the endpoint table
    assert tuple(sorted(TODO_TASK_WRITE_CASES)) == ("todo_create_tasks", "todo_update_tasks")
    for case in ("todo.read", "todo.search", "planner.create_tasks", "", "todo.update_tasks"):
        with pytest.raises(TodoContractError, match="not a declared To Do write case"):
            validate_todo_task_fields(case, {"title": "Buy milk"})


@pytest.mark.parametrize("field", REFUSED_WRITE_FIELDS)
def test_todo_update_refuses_a_field_the_endpoint_does_not_accept(graph_client, field):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match=re.escape(f"unknown field for todo_update_tasks: {field}")):
        build_write_request_information(
            graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
            todo_task_id="task", update_fields={field: "value"},
        )


@pytest.mark.parametrize("field", REFUSED_WRITE_FIELDS)
def test_todo_field_policy_refuses_the_same_fields_before_the_sdk_is_reached(field):
    from microsoft365.contract import (
        TODO_TASK_WRITE_CASES,
        TodoContractError,
        validate_todo_task_fields,
    )

    for case in sorted(TODO_TASK_WRITE_CASES):
        with pytest.raises(TodoContractError) as excinfo:
            validate_todo_task_fields(case, {field: "value"})
        message = str(excinfo.value)
        assert f"unknown field for {case}: {field}" in message
        assert "writable To Do task fields are" in message
        for accepted in EXPECTED_WRITABLE_FIELDS:
            assert accepted in message


@pytest.mark.parametrize(
    ("value", "message"),
    [(None, "title of todo_update_tasks must carry a value"), ("", "title is required")],
)
def test_todo_write_refuses_a_null_or_blank_title_instead_of_dropping_it(
    graph_client, value, message
):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match=message):
        build_write_request_information(
            graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
            todo_task_id="task", update_fields={"title": value},
        )


@pytest.mark.parametrize("field", ["body", "due_date_time", "status"])
def test_todo_write_refuses_a_null_value_instead_of_dropping_the_field(graph_client, field):
    from microsoft365.sdk_contract import build_write_request_information

    with pytest.raises(ValueError, match=f"{field} of todo_update_tasks must carry a value"):
        build_write_request_information(
            graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
            todo_task_id="task", update_fields={field: None},
        )


@pytest.mark.parametrize(
    ("case", "field", "value", "message"),
    [
        ("todo_update_tasks", "due_date_time", "2026-01-01T00:00:00Z", "must be a DateTimeTimeZone"),
        ("todo_update_tasks", "body", "just text", "must be an ItemBody"),
        ("todo_create_tasks", "due_date_time", {"dateTime": "2026-01-01T00:00:00Z"}, "must be a DateTimeTimeZone"),
    ],
)
def test_todo_write_refuses_a_typed_field_that_is_not_the_generated_model(
    graph_client, case, field, value, message
):
    from microsoft365.sdk_contract import build_write_request_information

    arguments = {"user_id": "user", "todo_list_id": "list", "title": "Buy milk"}
    if case == "todo_update_tasks":
        arguments["todo_task_id"] = "task"
        arguments["update_fields"] = {field: value}
    else:
        arguments[field] = value

    with pytest.raises(ValueError, match=message):
        build_write_request_information(graph_client, case, **arguments)


def test_todo_write_accepts_the_real_generated_models_for_body_and_due_date(graph_client):
    from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
    from msgraph.generated.models.item_body import ItemBody
    from msgraph.generated.models.body_type import BodyType
    from microsoft365.sdk_contract import build_write_request_information

    request = build_write_request_information(
        graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
        todo_task_id="task",
        update_fields={
            "body": ItemBody(content_type=BodyType.Html, content="<p>Buy oat milk</p>"),
            "due_date_time": DateTimeTimeZone(date_time="2026-01-01T00:00:00", time_zone="UTC"),
        },
    )

    assert serialized_body(request) == {
        "body": {"content": "<p>Buy oat milk</p>", "contentType": "html"},
        "dueDateTime": {"dateTime": "2026-01-01T00:00:00", "timeZone": "UTC"},
    }


def test_todo_update_keeps_if_match_optional_and_requires_no_header(graph_client):
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.sdk_contract import build_write_request_information

    assert OPERATION_REGISTRY["todo.update_tasks"].required_headers == ()

    without = build_write_request_information(
        graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
        todo_task_id="task", update_fields={"status": "completed"},
    )
    assert if_match(without) == set()

    with_etag = build_write_request_information(
        graph_client, "todo_update_tasks", user_id="user", todo_list_id="list",
        todo_task_id="task", etag='"etag-4"', update_fields={"status": "completed"},
    )
    assert if_match(with_etag) == {'"etag-4"'}
    assert serialized_body(with_etag) == {"status": "completed"}


# --- identifiers are refused before anything is sent ----------------------------------


@pytest.mark.parametrize(
    ("key", "arguments", "name", "endpoint"),
    [
        ("todo.list_task_lists", {}, "user_id", "GET /users/{user_id}/todo/lists"),
        ("todo.list_task_lists", {"user_id": ""}, "user_id", "GET /users/{user_id}/todo/lists"),
        ("todo.list_task_lists", {"user_id": "   "}, "user_id", "GET /users/{user_id}/todo/lists"),
        ("todo.search", {"user_id": "user"}, "todo_list_id", "GET /users/{user_id}/todo/lists/{todo_list_id}/tasks"),
        (
            "todo.search", {"user_id": "user", "todo_list_id": "\t"}, "todo_list_id",
            "GET /users/{user_id}/todo/lists/{todo_list_id}/tasks",
        ),
        (
            "todo.read", {"user_id": "user", "todo_list_id": "list"}, "todo_task_id",
            "GET /users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        ),
        (
            "todo.read", {"user_id": "user", "todo_list_id": "list", "todo_task_id": None},
            "todo_task_id", "GET /users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        ),
        ("todo.create_tasks", {"user_id": "user"}, "todo_list_id", "POST /users/{user_id}/todo/lists/{todo_list_id}/tasks"),
        (
            "todo.update_tasks", {"user_id": "user", "todo_list_id": " "}, "todo_list_id",
            "PATCH /users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}",
        ),
    ],
)
def test_absent_or_blank_required_identifiers_are_refused_with_the_endpoint(
    key, arguments, name, endpoint
):
    from microsoft365.contract import TodoContractError, validate_todo_identifiers

    with pytest.raises(TodoContractError) as excinfo:
        validate_todo_identifiers(key, arguments)

    message = str(excinfo.value)
    assert f"{name} is required for {key}" in message
    assert endpoint in message


@pytest.mark.parametrize(
    ("key", "arguments", "name"),
    [
        ("todo.search", {"user_id": 7, "todo_list_id": "list"}, "user_id"),
        ("todo.search", {"user_id": True, "todo_list_id": "list"}, "user_id"),
        ("todo.read", {"user_id": "user", "todo_list_id": "list", "todo_task_id": ["task"]}, "todo_task_id"),
        ("todo.create_tasks", {"user_id": "user", "todo_list_id": {"id": "list"}}, "todo_list_id"),
    ],
)
def test_non_string_identifiers_are_refused_without_coercion(key, arguments, name):
    from microsoft365.contract import TodoContractError, validate_todo_identifiers

    with pytest.raises(TodoContractError, match=re.escape(name)) as excinfo:
        validate_todo_identifiers(key, arguments)

    assert key in str(excinfo.value)
    assert "must be a non-blank string" in str(excinfo.value)


@pytest.mark.parametrize(
    ("key", "arguments"),
    [("todo.search", []), ("todo.read", "user_id=user"), ("todo.read", None), ("todo.list_task_lists", 7)],
)
def test_todo_arguments_must_be_a_mapping(key, arguments):
    from microsoft365.contract import TodoContractError, validate_todo_identifiers

    with pytest.raises(TodoContractError, match="arguments must be a mapping"):
        validate_todo_identifiers(key, arguments)


@pytest.mark.parametrize("key", ["todo.delete", "todo.list_task_lists.tasks", "planner.read", ""])
def test_only_declared_todo_operations_can_be_validated(key):
    from microsoft365.contract import TodoContractError, validate_todo_identifiers

    with pytest.raises(
        TodoContractError, match=re.escape(f"{key} is not a declared To Do operation")
    ):
        validate_todo_identifiers(key, {})


def test_normalized_identifiers_render_the_declared_endpoint(graph_client):
    from microsoft365.contract import validate_todo_identifiers
    from microsoft365.sdk_contract import build_collection_request_information

    identifiers = validate_todo_identifiers(
        "todo.search", {"user_id": " user ", "todo_list_id": " list "}
    )

    assert identifiers == {"user_id": "user", "todo_list_id": "list"}
    request = build_collection_request_information(graph_client, "todo_tasks", **identifiers, query="milk")
    assert urlsplit(request.url).path == "/users/user/todo/lists/list/tasks"


def test_registry_guard_and_sdk_guard_refuse_a_blank_list_container_independently(graph_client):
    from microsoft365.contract import TodoContractError, validate_todo_identifiers
    from microsoft365.sdk_contract import build_collection_request_information

    with pytest.raises(TodoContractError):
        validate_todo_identifiers("todo.search", {"user_id": "user", "todo_list_id": "   "})

    with pytest.raises(ValueError) as excinfo:
        build_collection_request_information(
            graph_client, "todo_tasks", user_id="user", todo_list_id="   ", query="milk"
        )
    # the SDK-level guard is a separate one, not the registry error bubbling through
    assert excinfo.type is ValueError
    assert "todo_list_id is required" in str(excinfo.value)


# --- honest status labels: nothing is promoted without evidence -----------------------


def test_no_todo_operation_is_executable_while_its_handler_does_not_exist():
    from microsoft365.contract import OPERATION_REGISTRY, operation_status
    from microsoft365.registration import HANDLER_TABLE

    todo_keys = sorted(key for key in OPERATION_REGISTRY if key.startswith("todo."))

    assert todo_keys == sorted(EXPECTED_OPERATION_CONTRACTS)
    for key in todo_keys:
        definition = OPERATION_REGISTRY[key]
        assert definition.implementation_status == "implemented", key
        assert definition.executable is (key in EXECUTABLE_READS), key
        assert key in HANDLER_TABLE, key

        service, operation = key.split(".", 1)
        status = operation_status("application", service, operation)
        assert status.executable is (key in EXECUTABLE_READS), key
        assert status.implementation_status == "implemented", key
        expected_auth = "not_verified" if key in TODO_WRITE_KEYS else "supported"
        assert status.auth_status == expected_auth, key

    # The milestone's exact sets: three Outlook/Calendar reads are the only executable
    # operations, and the only handlers registered are those reads plus the four writes WP6
    # implemented and withholds (R5). No To Do operation, read or write, has a handler.
    assert {
        key for key, definition in OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_READS)
    assert set(HANDLER_TABLE) == set(EXECUTABLE_READS) | set(WITHHELD_WRITES) | {"teams.search_messages", "teams.send_messages"}
    assert set(todo_keys).issubset(HANDLER_TABLE)


def test_todo_writes_keep_explicit_not_verified_status_until_endpoint_evidence_exists():
    from microsoft365.contract import (
        WRITE_SUPPORT_STATUSES,
        operation_status,
        todo_write_support,
    )

    assert WRITE_SUPPORT_STATUSES == frozenset({"not_verified", "supported"})

    for key in TODO_WRITE_KEYS:
        support = todo_write_support(key)

        assert support.operation == key
        assert support.status == "not_verified"
        assert support.supported is False
        assert support.evidence == ()
        assert "unverified" in support.reason
        # the endpoint's own application claim is what would have to be verified
        assert todo_write_support.__doc__.strip()

        service, operation = key.split(".", 1)
        status = operation_status("application", service, operation)
        assert status.auth_status == "not_verified"
        # a To Do write stays non-executable: the milestone exposes three reads and no write
        assert key not in EXECUTABLE_READS
        assert status.executable is False


def test_application_write_support_cannot_be_promoted_without_recorded_evidence():
    from microsoft365.contract import (
        ApplicationWriteSupport,
        TodoContractError,
        todo_write_support,
    )

    with pytest.raises(TodoContractError, match="without recorded endpoint evidence"):
        ApplicationWriteSupport(
            operation="todo.create_tasks", status="supported", reason="looks fine to me"
        )

    with pytest.raises(TodoContractError, match="unknown application write support status"):
        ApplicationWriteSupport(operation="todo.create_tasks", status="probably", reason="?")

    # the guard is not a blanket ban: a supported claim with recorded evidence is a real state
    promoted = ApplicationWriteSupport(
        operation="todo.create_tasks",
        status="supported",
        reason="endpoint reference re-read and confirmed on a test tenant",
        evidence=("POST /users/{user_id}/todo/lists/{todo_list_id}/tasks",),
    )
    assert promoted.supported is True
    assert promoted.evidence

    with pytest.raises(TodoContractError, match="not a declared To Do write operation"):
        todo_write_support("todo.read")
    with pytest.raises(TodoContractError, match="not a declared To Do write operation"):
        todo_write_support("planner.create_tasks")


def test_todo_endpoint_cannot_claim_verified_permissions_without_recorded_evidence():
    from microsoft365.contract import PERMISSION_CLAIM_STATUSES, TodoContractError, TodoEndpoint

    assert PERMISSION_CLAIM_STATUSES == frozenset({"documented_not_verified", "verified"})

    arguments = {
        "operation": "create_tasks",
        "method": "POST",
        "path_template": "/users/{user_id}/todo/lists/{todo_list_id}/tasks",
        "container": "list",
        "application_permissions": ("Tasks.ReadWrite.All",),
    }

    with pytest.raises(TodoContractError, match="with no recorded evidence"):
        TodoEndpoint(**arguments, claim_status="verified")

    verified = TodoEndpoint(
        **arguments,
        claim_status="verified",
        claim_evidence=("reference re-read + test tenant confirmation",),
    )
    assert verified.claim_status == "verified"

    with pytest.raises(TodoContractError, match="unknown permission claim status"):
        TodoEndpoint(**arguments, claim_status="sounds_right")


def test_write_support_is_derived_from_the_endpoint_claims_not_asserted(monkeypatch):
    """Promotion is possible only through evidence, and never silently changes execution."""
    from microsoft365 import contract

    assert contract.todo_write_support("todo.create_tasks").supported is False

    promoted = tuple(
        dataclasses.replace(
            endpoint,
            claim_status="verified",
            claim_evidence=("endpoint reference re-read and confirmed on a test tenant",),
        )
        if f"todo.{endpoint.operation}" == "todo.create_tasks"
        else endpoint
        for endpoint in contract.TODO_ENDPOINTS
    )
    monkeypatch.setattr(contract, "TODO_ENDPOINTS", promoted)

    support = contract.todo_write_support("todo.create_tasks")
    assert support.supported is True
    assert support.evidence == ("POST /users/{user_id}/todo/lists/{todo_list_id}/tasks",)

    # the endpoint is now verified for application mode, but the operation itself is still
    # not executable: support and executability are separate, and only WP9 may flip the flag
    status = contract.operation_status("application", "todo", "create_tasks")
    assert status.auth_status == "supported"
    assert "todo.create_tasks" not in EXECUTABLE_READS
    assert status.executable is False
    assert contract.OPERATION_REGISTRY["todo.create_tasks"].executable is False
    # the milestone's exact executable set is unchanged by the promotion
    assert {
        key for key, definition in contract.OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_READS)


def test_todo_permission_claims_are_endpoint_specific_and_not_promoted():
    from microsoft365.contract import (
        OPERATION_REGISTRY,
        PERMISSION_CLAIM_STATUSES,
        TODO_ENDPOINTS,
    )

    for endpoint in TODO_ENDPOINTS:
        assert endpoint.application_permissions, endpoint.endpoint
        assert endpoint.claim_status in PERMISSION_CLAIM_STATUSES
        assert endpoint.claim_status == "documented_not_verified", endpoint.endpoint
        assert endpoint.claim_evidence == ()
        assert endpoint.documentation_page == "" or endpoint.documentation_page.startswith(
            "https://learn.microsoft.com/en-us/graph/api/"
        )

    assert {key: OPERATION_REGISTRY[key].permissions for key in EXPECTED_CLAIMS} == EXPECTED_CLAIMS
    assert {
        key: OPERATION_REGISTRY[key].permission_claim_statuses for key in EXPECTED_CLAIMS
    } == {key: ("documented_not_verified",) for key in EXPECTED_CLAIMS}

    for key, roles in EXPECTED_CLAIMS.items():
        endpoints = [item for item in TODO_ENDPOINTS if f"todo.{item.operation}" == key]
        assert endpoints
        for endpoint in endpoints:
            assert endpoint.application_permissions == roles, endpoint.endpoint


def test_preflight_reports_todo_claims_without_claiming_execution():
    from microsoft365.contract import Settings
    from microsoft365.preflight import build_preflight

    payload = build_preflight(
        Settings.from_mapping({"capabilities": {"todo": True}}), sdk_available=True
    )

    assert sorted(payload["operation_status"]) == sorted(EXPECTED_CLAIMS)
    for key, status in payload["operation_status"].items():
        assert status["permissions"] == list(EXPECTED_CLAIMS[key])
        assert status["implementation_status"] == "implemented"
        assert status["executable"] is (key in EXECUTABLE_READS)
        assert status["write"] is (key in TODO_WRITE_KEYS)
    assert payload["locally_ready"] is False
    assert payload["remote_verification"] == "not_tested"


# --- the documents must match the registry -------------------------------------------


def todo_matrix_rows() -> list[tuple[str, ...]]:
    rows = []
    for line in PERMISSION_MATRIX.read_text(encoding="utf-8").splitlines():
        cells = tuple(cell.strip().strip("`") for cell in line.split("|")[1:-1])
        if len(cells) == 6 and cells[0].startswith("todo."):
            rows.append(cells)
    return rows


def test_permission_matrix_reports_one_row_per_todo_endpoint():
    from microsoft365.contract import OPERATION_REGISTRY, TODO_ENDPOINTS

    expected = sorted(
        (
            f"todo.{endpoint.operation}",
            endpoint.endpoint,
            " ".join(endpoint.application_permissions),
            endpoint.claim_status,
            OPERATION_REGISTRY[f"todo.{endpoint.operation}"].implementation_status,
            endpoint.documentation_page or "not recorded",
        )
        for endpoint in TODO_ENDPOINTS
    )

    assert sorted(todo_matrix_rows()) == expected


def test_permission_matrix_keeps_no_blanket_todo_claim():
    text = PERMISSION_MATRIX.read_text(encoding="utf-8")
    rows = [
        [cell.strip().strip("`") for cell in line.split("|")[1:-1]]
        for line in text.splitlines()
        if line.startswith("|")
    ]

    # no table row may cover To Do as one blanket operation family again
    assert [
        cells[0]
        for cells in rows
        if cells[:1] and cells[0] in {"To Do list/search/read", "To Do create/update"}
    ] == []
    assert "To Do endpoint claims" in text


def test_permission_matrix_states_the_todo_write_status_and_field_policy():
    from microsoft365.contract import TODO_TASK_WRITE_FIELDS, todo_write_support

    text = PERMISSION_MATRIX.read_text(encoding="utf-8")
    section = text.split("## Microsoft To Do endpoint claims", 1)[-1]

    assert "## Microsoft To Do endpoint claims" in text
    support = todo_write_support("todo.update_tasks")
    assert support.status in section
    assert support.reason in section
    for key in TODO_WRITE_KEYS:
        assert key in section
    for field in TODO_TASK_WRITE_FIELDS:
        assert field in section
    # the filter decision is documented instead of being silently absent
    assert "client-side" in section
    assert "not recorded" in section


def test_known_limitations_records_the_todo_contract_state():
    text = KNOWN_LIMITATIONS.read_text(encoding="utf-8")

    assert "tests/test_todo_contracts.py" in text
    assert "The three To Do reads are executable" in text
    assert "documented_not_verified" in text
    # the pre-WP3 wording claimed To Do application-write support was merely pending; the
    # state is now explicit and derived
    assert "To Do application-write support, Microsoft Search chat-message support" not in text
    assert "not_verified" in text
