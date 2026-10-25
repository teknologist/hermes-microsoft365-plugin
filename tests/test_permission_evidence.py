from __future__ import annotations

from urllib.parse import urlparse

import pytest

from microsoft365.contract import (
    DRIVE_ENDPOINTS,
    MESSAGING_ENDPOINTS,
    PLANNER_ENDPOINTS,
    TODO_ENDPOINTS,
    TEAMS_ENDPOINTS,
    OPERATION_REGISTRY,
)

ALL_ENDPOINTS = PLANNER_ENDPOINTS + TODO_ENDPOINTS + MESSAGING_ENDPOINTS + DRIVE_ENDPOINTS + TEAMS_ENDPOINTS


def endpoint_key(endpoint):
    key = getattr(endpoint, "key", None)
    if key:
        return key
    service = "planner" if endpoint in PLANNER_ENDPOINTS else "todo"
    return f"{service}.{endpoint.operation}"


@pytest.mark.parametrize("endpoint", ALL_ENDPOINTS)
def test_every_endpoint_has_official_static_evidence(endpoint):
    """The matrix must cite the endpoint page it was recomputed against."""
    parsed = urlparse(endpoint.documentation_page)
    assert parsed.scheme == "https"
    assert parsed.netloc == "learn.microsoft.com"
    assert parsed.path.startswith("/en-us/graph/api/")
    assert "?" not in endpoint.documentation_page
    assert "#" not in endpoint.documentation_page


def test_registry_documentation_pages_cannot_drift_from_endpoint_rows():
    """Operation metadata is exactly the union of its endpoint citations."""
    for definition in OPERATION_REGISTRY.values():
        rows = [row for row in ALL_ENDPOINTS if endpoint_key(row) == definition.key]
        assert definition.documentation_pages == tuple(dict.fromkeys(row.documentation_page for row in rows))


def test_static_evidence_never_promotes_permission_or_remote_claims():
    """Official docs are static evidence only; tenant verification stays fail-closed."""
    for definition in OPERATION_REGISTRY.values():
        assert definition.remote_verification == "not_tested"
        assert not definition.remote_verification_evidence
        for mode in (definition.app, definition.delegated):
            assert mode.permission_status != "verified"
            assert not mode.permission_evidence


def test_unsupported_teams_application_operations_remain_claim_free():
    for key in ("teams.search_messages", "teams.send_messages"):
        definition = OPERATION_REGISTRY[key]
        assert definition.app.status == "unsupported_auth_mode"
        assert definition.app.permissions == ()
        assert definition.app.permission_status == "no_permission_claimed"


def test_all_thirty_two_operations_are_present_in_the_evidence_matrix():
    assert len(OPERATION_REGISTRY) == 32
    assert {endpoint_key(row) for row in ALL_ENDPOINTS} == set(OPERATION_REGISTRY)
