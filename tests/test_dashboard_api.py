import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from microsoft365.contract import OPERATION_REGISTRY, OPERATIONS
from microsoft365.dashboard import plugin_api


@pytest.fixture
def client(monkeypatch):
    monkeypatch.setattr(plugin_api, "load_config", lambda: {})
    monkeypatch.setattr(plugin_api, "save_config", lambda config, **kwargs: None)
    return TestClient(plugin_api.create_app())


def test_configuration_is_sanitized(client):
    plugin_api.load_config = lambda: {"plugins": {"entries": {"microsoft365": {"settings": {
        "tenant_id": "tenant-sentinel", "client_id": "client-sentinel", "user_id": "user-sentinel",
        "authentication_mode": "application", "capabilities": {"outlook": {"search": True}},
    }}}}}
    body = client.get("/configuration").json()
    text = json.dumps(body)
    assert "tenant-sentinel" not in text
    assert "client-sentinel" not in text
    assert "user-sentinel" not in text
    assert body["tenant_id_configured"] is True
    assert body["client_id_configured"] is True
    assert body["user_id_configured"] is True


def test_configuration_put_rejects_unknown_and_authority_fields(client):
    base = {"tenant_id": "t", "client_id": "c", "user_id": "u", "authentication_mode": "application", "capabilities": {}}
    for extra in ({"secret": "x"}, {"executable": True}, {"verified": True}, {"unexpected": "x"}):
        response = client.put("/configuration", json={**base, **extra})
        assert response.status_code in (400, 422)
        assert "x" not in response.text


def test_configuration_put_persists_only_plugin_settings(monkeypatch):
    saved = []
    monkeypatch.setattr(plugin_api, "load_config", lambda: {"other": {"keep": True}})
    monkeypatch.setattr(plugin_api, "save_config", lambda config, **kwargs: saved.append(config))
    client = TestClient(plugin_api.create_app())
    payload = {"tenant_id": "tenant-sentinel", "client_id": "client-sentinel", "user_id": "user-sentinel",
               "authentication_mode": "application", "capabilities": {"outlook": {"search": True}}}
    response = client.put("/configuration", json=payload)
    assert response.status_code == 200
    assert saved == [{"other": {"keep": True}, "plugins": {"entries": {"microsoft365": {"settings": payload}}}}]
    assert "tenant-sentinel" not in json.dumps(response.json())


def test_capabilities_contains_every_registry_operation(client):
    response = client.get("/capabilities")
    assert response.status_code == 200
    rows = response.json()["operations"]
    assert {row["key"] for row in rows} == set(OPERATION_REGISTRY)
    assert len(rows) == sum(map(len, OPERATIONS.values())) == 32


def test_preflight_is_local_and_sanitized(monkeypatch):
    sentinel = "token-sentinel"
    monkeypatch.setattr(plugin_api, "load_config", lambda: {"plugins": {"entries": {"microsoft365": {"settings": {
        "tenant_id": sentinel, "client_id": sentinel, "user_id": sentinel, "capabilities": {}
    }}}}})
    called = []
    monkeypatch.setattr(plugin_api, "build_preflight", lambda settings, **kwargs: called.append(settings) or {"locally_ready": False, "remote_verification": "not_tested"})
    response = TestClient(plugin_api.create_app()).get("/preflight")
    assert response.status_code == 200
    assert response.json()["remote_verification"] == "not_tested"
    assert sentinel not in response.text
    assert called


def test_configuration_update_preserves_omitted_mailbox_policy(monkeypatch):
    # Synthetic mailbox object IDs in GUID textual form; never real tenant identifiers.
    main_id = "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d"
    billings_id = "d4c3b2a1-6f5e-47a6-9876-543210fedcba"
    replacement_id = "deadbeef-0000-4000-8000-0000000000ff"
    payload = {"tenant_id": "t", "client_id": "c", "user_id": "u",
               "authentication_mode": "application", "capabilities": {"outlook": {"search": True}}}
    state = {"plugins": {"entries": {"microsoft365": {"settings": {
        **payload, "allowed_mailbox_ids": [main_id, billings_id],
    }}}}}
    monkeypatch.setattr(plugin_api, "load_config", lambda: state)
    monkeypatch.setattr(plugin_api, "save_config", lambda config, **kwargs: state.update(config))
    client = TestClient(plugin_api.create_app())
    assert main_id not in client.get("/configuration").text and billings_id not in client.get("/configuration").text
    response = client.put("/configuration", json=payload)
    assert response.status_code == 200
    assert main_id not in response.text and billings_id not in response.text
    assert state["plugins"]["entries"]["microsoft365"]["settings"]["allowed_mailbox_ids"] == [main_id, billings_id]
    response = client.put("/configuration", json={**payload, "allowed_mailbox_ids": [replacement_id]})
    assert response.status_code == 200
    assert replacement_id not in response.text
    assert state["plugins"]["entries"]["microsoft365"]["settings"]["allowed_mailbox_ids"] == [replacement_id]
    assert client.put("/configuration", json={**payload, "allowed_mailbox_ids": "invalid"}).status_code == 400
    assert client.put("/configuration", json={**payload, "allowed_mailbox_ids": ["not-a-guid"]}).status_code == 400
    assert client.put("/configuration", json={**payload, "allowed_mailbox_ids": []}).status_code == 200
    assert not state["plugins"]["entries"]["microsoft365"]["settings"].get("allowed_mailbox_ids")
