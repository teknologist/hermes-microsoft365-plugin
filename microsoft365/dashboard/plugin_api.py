"""Sanitized FastAPI surface for the Microsoft 365 native dashboard.

This module deliberately keeps configuration persistence and the capability catalogue behind
small functions so it can be imported in a Hermes-free test environment.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
from typing import Any

try:
    from fastapi import APIRouter, FastAPI, Request
    from fastapi.responses import JSONResponse
except ImportError:  # pragma: no cover - the host supplies FastAPI
    APIRouter = None  # type: ignore[assignment]
    FastAPI = None  # type: ignore[assignment]
    Request = Any  # type: ignore[misc,assignment]
    JSONResponse = None  # type: ignore[assignment]

from ..contract import ConfigurationError, OPERATION_REGISTRY, OPERATIONS, Settings, operation_status
from ..preflight import build_preflight

try:  # Hermes is optional when this package is tested outside the host.
    from hermes_cli.config import load_config, save_config
except ImportError:  # pragma: no cover - exercised only without Hermes installed
    _LOCAL_CONFIG: dict[str, Any] = {}

    def load_config() -> dict[str, Any]:
        return deepcopy(_LOCAL_CONFIG)

    def save_config(config: dict[str, Any], **_: Any) -> None:
        _LOCAL_CONFIG.clear()
        _LOCAL_CONFIG.update(deepcopy(config))


_ALLOWED_FIELDS = frozenset({"tenant_id", "client_id", "user_id", "authentication_mode", "capabilities", "allowed_mailbox_ids"})
_PLUGIN_PATH = ("plugins", "entries", "microsoft365", "settings")


def _settings_from_config() -> tuple[Settings | None, list[str]]:
    try:
        config = load_config() or {}
        current: Any = config
        for key in _PLUGIN_PATH:
            if not isinstance(current, dict):
                current = {}
                break
            current = current.get(key, {})
        if not isinstance(current, dict):
            return None, ["invalid Microsoft 365 configuration"]
        return Settings.from_mapping(current), []
    except (ConfigurationError, TypeError, ValueError, AttributeError):
        return None, ["invalid Microsoft 365 configuration"]
    except Exception:
        # A host/config backend exception must never become a traceback or leak its message.
        return None, ["configuration unavailable"]


def _configuration_response(settings: Settings | None) -> dict[str, Any]:
    if settings is None:
        return {
            "authentication_mode": "application",
            "capabilities": {},
            "tenant_id_configured": False,
            "client_id_configured": False,
            "user_id_configured": False,
        }
    return {
        "authentication_mode": settings.authentication_mode,
        "capabilities": {service: dict(operations) for service, operations in settings.capabilities.items()},
        "tenant_id_configured": bool(settings.tenant_id.strip()),
        "client_id_configured": bool(settings.client_id.strip()),
        "user_id_configured": bool(settings.user_id.strip()),
    }


def _error(status: int, message: str) -> JSONResponse:
    return JSONResponse(status_code=status, content={"error": message})


def _strict_settings(body: Any) -> Settings:
    if not isinstance(body, dict):
        raise ConfigurationError("request body must be a JSON object")
    unknown = set(body) - _ALLOWED_FIELDS
    if unknown:
        raise ConfigurationError("unknown configuration field")
    missing = (_ALLOWED_FIELDS - {"allowed_mailbox_ids"}) - set(body)
    if missing:
        raise ConfigurationError("configuration fields are required")
    for name in ("tenant_id", "client_id", "user_id", "authentication_mode"):
        if type(body[name]) is not str:
            raise ConfigurationError("configuration values have invalid types")
    if type(body["capabilities"]) is not dict:
        raise ConfigurationError("capabilities must be a mapping")
    if "allowed_mailbox_ids" not in body:
        current, errors = _settings_from_config()
        if errors or current is None:
            raise ConfigurationError("existing configuration is invalid")
        body = {**body, "allowed_mailbox_ids": current.allowed_mailbox_ids}
    return Settings.from_mapping(body)


def _save_settings(settings: Settings) -> None:
    config = deepcopy(load_config() or {})
    if not isinstance(config, dict):
        raise ConfigurationError("configuration is not an object")
    settings_dict = {
        "tenant_id": settings.tenant_id,
        "client_id": settings.client_id,
        "user_id": settings.user_id,
        "authentication_mode": settings.authentication_mode,
        "capabilities": {service: dict(ops) for service, ops in settings.capabilities.items()},
    }
    if settings.allowed_mailbox_ids:
        settings_dict["allowed_mailbox_ids"] = list(settings.allowed_mailbox_ids)
    node = config
    for key in _PLUGIN_PATH[:-1]:
        existing = node.get(key)
        if not isinstance(existing, dict):
            existing = {}
            node[key] = existing
        node = existing
    node[_PLUGIN_PATH[-1]] = settings_dict
    save_config(config)


router = APIRouter() if APIRouter is not None else None


if router is not None:
    @router.get("/configuration")
    async def get_configuration() -> dict[str, Any]:
        settings, _ = _settings_from_config()
        return _configuration_response(settings)

    @router.put("/configuration")
    async def put_configuration(request: Request) -> Any:
        try:
            body = await request.json()
            settings = _strict_settings(body)
            _save_settings(settings)
        except (ConfigurationError, TypeError, ValueError):
            return _error(400, "invalid configuration")
        except Exception:
            return _error(400, "configuration could not be saved")
        return _configuration_response(settings)

    @router.get("/capabilities")
    async def get_capabilities() -> dict[str, Any]:
        settings, _ = _settings_from_config()
        mode = settings.authentication_mode if settings else "application"
        rows = []
        for key, definition in OPERATION_REGISTRY.items():
            status = operation_status(mode, definition.service, definition.operation)
            rows.append({
                "key": key,
                "service": definition.service,
                "operation": definition.operation,
                "permissions": list(definition.permissions),
                "write": definition.write,
                "status": asdict(status),
            })
        return {"authentication_mode": mode, "operations": rows}

    @router.get("/preflight")
    async def get_preflight() -> Any:
        settings, errors = _settings_from_config()
        try:
            return build_preflight(settings, configuration_errors=errors)
        except Exception:
            return _error(400, "preflight unavailable")


def create_app() -> Any:
    if FastAPI is None or router is None:  # pragma: no cover
        raise RuntimeError("FastAPI is required to create the dashboard API")
    app = FastAPI(title="Microsoft 365 dashboard API", docs_url=None, redoc_url=None)
    app.include_router(router)
    return app


app = create_app() if FastAPI is not None else None
