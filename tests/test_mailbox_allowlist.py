"""Offline mailbox policy proofs through configured host registration."""
from __future__ import annotations

import json
from urllib.parse import urlsplit

import pytest
from msgraph import GraphServiceClient
from msgraph.generated.models.message_collection_response import MessageCollectionResponse

from microsoft365 import register, registration
from microsoft365.contract import ConfigurationError, Settings
from microsoft365.preflight import build_preflight
from microsoft365.validation import check
from tests.test_handlers_outlook_calendar import HandlerGraphAdapter
from tests.test_registration_secrets import RecordingContext

# Synthetic mailbox object IDs in GUID textual form; never real tenant identifiers.
MAIN_ID = "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d"
BILLINGS_ID = "d4c3b2a1-6f5e-47a6-9876-543210fedcba"
THIRD_ID = "deadbeef-0000-4000-8000-0000000000ff"


def test_registered_outlook_allowlist_overrides_single_user_before_client(monkeypatch):
    adapter = HandlerGraphAdapter({
        ("GET", f"/users/{MAIN_ID}/messages"): MessageCollectionResponse(value=[]),
        ("GET", f"/users/{BILLINGS_ID}/messages"): MessageCollectionResponse(value=[]),
    })
    clients = []
    monkeypatch.setattr(
        registration, "create_graph_client",
        lambda settings: clients.append(settings) or GraphServiceClient(request_adapter=adapter),
    )
    ctx = RecordingContext({
        "user_id": "legacy-user",
        "allowed_mailbox_ids": [MAIN_ID, BILLINGS_ID],
        "capabilities": {"outlook": {"search": True}},
    })
    register(ctx)
    hook = ctx.hooks["pre_tool_call"]
    handler = ctx.tools["microsoft365_outlook"]["handler"]

    for mailbox in (MAIN_ID, BILLINGS_ID):
        args = {"action": "search", "user_id": mailbox}
        assert hook(tool_name="microsoft365_outlook", args=args) is None
        result = json.loads(handler(args))
        assert result["status"] == "succeeded"
        assert result["result"]["value"] == []
    assert len(clients) == 2
    # The allowlisted IDs reach the Graph request with their exact spelling: the
    # handler normalizes the identifier in no way, so an authorized ID cannot be
    # rewritten into a different one in flight.
    assert [urlsplit(request.url).path for request in adapter.requests] == [
        f"/users/{MAIN_ID}/messages", f"/users/{BILLINGS_ID}/messages",
    ]

    # A third mailbox, the legacy user, a case-variant of an allowlisted ID and a
    # whitespace-padded ID are all rejected: the comparison is exact, not normalized.
    for denied in (THIRD_ID, "legacy-user", BILLINGS_ID.upper(), f" {BILLINGS_ID}"):
        args = {"action": "search", "user_id": denied}
        assert hook(tool_name="microsoft365_outlook", args=args)["action"] == "block"
        result = json.loads(handler(args))
        assert result["error"] == "validation_error"
        assert denied not in json.dumps(result)
    assert len(clients) == 2
    assert len(adapter.requests) == 2


@pytest.mark.parametrize("invalid", [
    # list shape, string entries, duplicates
    None, "main-mailbox", {}, (), 1, False, [None], [1], [True], [""], [" \t"], ["main", "main"],
    # unsupported wildcard forms and non-GUID identifiers
    ["*"], ["main-mailbox"], ["user@contoso.onmicrosoft.com"],
    # malformed tokens: whitespace padding, non-hex character, short final group,
    # and a duplicate that differs only by hex case
    [MAIN_ID + " "], [" " + MAIN_ID],
    ["a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5g"],
    ["a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5"],
    [MAIN_ID, MAIN_ID.upper()],
])
def test_malformed_allowlist_disables_service_registration(invalid):
    ctx = RecordingContext({"allowed_mailbox_ids": invalid, "capabilities": {"outlook": {"search": True}}})
    register(ctx)
    assert set(ctx.tools) == {"microsoft365_preflight"}
    result = json.loads(ctx.tools["microsoft365_preflight"]["handler"]({}))
    assert result["locally_ready"] is False
    assert "allowed_mailbox_ids" in result["configuration_errors"][0]
    assert "main-mailbox" not in json.dumps(result)
    with pytest.raises(ConfigurationError):
        Settings(allowed_mailbox_ids=invalid)


def test_allowlist_requires_guid_textual_shape_verbatim():
    """Wildcard, UPN and malformed tokens are refused at configuration time, and a
    valid entry is retained verbatim: no stripping, no case-folding, no repair."""
    for entry in (
        "*",
        "user@contoso.onmicrosoft.com",
        "main-mailbox",
        " " + MAIN_ID,
        MAIN_ID + " ",
        MAIN_ID.replace("-", ""),
        "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5",
        "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5g",
    ):
        with pytest.raises(ConfigurationError):
            Settings(allowed_mailbox_ids=[entry])
    upper = MAIN_ID.upper()
    assert Settings(allowed_mailbox_ids=[upper]).allowed_mailbox_ids == [upper]
    # Duplicates are detected case-insensitively: same ID, different spelling is one ID.
    with pytest.raises(ConfigurationError):
        Settings(allowed_mailbox_ids=[MAIN_ID, upper])


def test_allowlist_comparison_is_exact_not_normalized():
    """With a nonempty allowlist only the verbatim entry passes: no case-folding, no
    whitespace stripping, no other identifier repair can map an input onto it."""
    settings = Settings(
        user_id="ignored", allowed_mailbox_ids=[MAIN_ID],
        capabilities={"outlook": {"search": True}},
    )
    assert check(settings, service="outlook", arguments={"action": "search", "user_id": MAIN_ID}) is None
    for variant in (MAIN_ID.upper(), " " + MAIN_ID, MAIN_ID[:-1] + "e"):
        rejection = check(settings, service="outlook", arguments={"action": "search", "user_id": variant})
        assert rejection.category == "validation_error"


@pytest.mark.parametrize("service,action", [("calendar", "search"), ("todo", "list_task_lists"), ("teams", "list_teams")])
def test_allowlist_does_not_widen_other_user_scoped_services(service, action):
    settings = Settings.from_mapping({
        "user_id": "main", "allowed_mailbox_ids": [MAIN_ID, BILLINGS_ID],
        "capabilities": {service: {action: True}},
    })
    assert check(settings, service=service, arguments={"action": action, "user_id": "main"}) is None
    rejection = check(settings, service=service, arguments={"action": action, "user_id": "billings"})
    assert rejection is not None
    assert rejection.category == "validation_error"


@pytest.mark.parametrize("allowlist", [None, []])
def test_absent_or_empty_allowlist_retains_legacy_boundary(allowlist):
    raw = {"user_id": "main", "capabilities": {"outlook": {"search": True}}}
    if allowlist is not None:
        raw["allowed_mailbox_ids"] = allowlist
    settings = Settings.from_mapping(raw)
    assert check(settings, service="outlook", arguments={"action": "search", "user_id": "main"}) is None
    assert check(settings, service="outlook", arguments={"action": "search", "user_id": "billings"}).category == "validation_error"


@pytest.mark.parametrize("other_service", [None, "calendar", "todo", "teams"])
def test_preflight_allowlist_satisfies_only_outlook_user_requirement(monkeypatch, other_service):
    from microsoft365 import preflight

    monkeypatch.setattr(preflight, "_secret_presence", lambda: (True, None))
    capabilities = {"outlook": {"search": True}}
    if other_service:
        capabilities[other_service] = {"list_task_lists" if other_service == "todo" else "list_teams" if other_service == "teams" else "search": True}
    settings = Settings.from_mapping({
        "tenant_id": "tenant-sentinel", "client_id": "client-sentinel",
        "allowed_mailbox_ids": [MAIN_ID, BILLINGS_ID], "capabilities": capabilities,
    })
    result = build_preflight(settings, sdk_available=True)
    assert result["requirements"]["user_id"] is bool(other_service)
    assert ("user_id" in result["missing"]) is bool(other_service)
    assert result["locally_ready"] is (other_service is None)
    text = json.dumps(result)
    assert MAIN_ID not in text and BILLINGS_ID not in text


def test_direct_dispatch_rechecks_mailbox_before_client(monkeypatch):
    """Calling the dispatch entry point without the registration closure is still bounded."""
    settings = Settings(allowed_mailbox_ids=[MAIN_ID], capabilities={"outlook": {"search": True}})
    calls = []
    adapter = HandlerGraphAdapter({
        ("GET", f"/users/{MAIN_ID}/messages"): MessageCollectionResponse(value=[]),
        ("GET", f"/users/{THIRD_ID}/messages"): MessageCollectionResponse(value=[]),
    })
    monkeypatch.setattr(registration, "create_graph_client", lambda settings: calls.append(settings) or GraphServiceClient(request_adapter=adapter))
    allowed = registration.service_tool_handler("outlook", {"action": "search", "user_id": MAIN_ID}, settings)
    assert allowed["status"] == "succeeded"
    calls.clear()
    adapter.requests.clear()
    result = registration.service_tool_handler("outlook", {"action": "search", "user_id": THIRD_ID}, settings)
    result = json.loads(result) if isinstance(result, str) else result
    assert result.get("error") == "validation_error"
    assert calls == []
    assert adapter.requests == []


def test_post_hook_mailbox_rewrite_cannot_escape_allowlist(monkeypatch):
    calls = []
    monkeypatch.setattr(registration, "create_graph_client", lambda settings: calls.append(settings))
    ctx = RecordingContext({"allowed_mailbox_ids": [MAIN_ID], "capabilities": {"outlook": {"search": True}}})
    register(ctx)
    args = {"action": "search", "user_id": MAIN_ID}
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args=args) is None
    args["user_id"] = THIRD_ID
    assert json.loads(ctx.tools["microsoft365_outlook"]["handler"](args))["error"] == "validation_error"
    assert calls == []
