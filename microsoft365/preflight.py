"""Side-effect-free local preflight reporting."""
from __future__ import annotations

from dataclasses import asdict
from typing import Iterable

from .contract import OPERATIONS, OPERATION_REGISTRY, Settings, operation_status

SECRET_ENV_NAME = "MICROSOFT365_CLIENT_SECRET"  # pragma: allowlist secret


def _secret_presence() -> tuple[bool, str | None]:
    try:
        from agent.secret_scope import get_secret

        return bool(get_secret(SECRET_ENV_NAME, None)), None
    except Exception as exc:
        # Unscoped access and backend failures fail closed; never return details
        # that may contain provider-specific secret material.
        return False, type(exc).__name__


def _requirements(settings: Settings, selected: list[str]) -> dict[str, bool]:
    """Derive identity requirements from the selected operation registry rows."""
    definitions = [OPERATION_REGISTRY[key] for key in selected]
    application = settings.authentication_mode == "application"
    return {
        "tenant_id": application and bool(selected),
        "client_id": application and bool(selected),
        "user_id": any(
            "user_id" in definition.required_identifiers
            and not (definition.service == "outlook" and settings.allowed_mailbox_ids)
            for definition in definitions
        ),
    }


def build_preflight(
    settings: Settings | None,
    *,
    sdk_available: bool | None = None,
    configuration_errors: Iterable[str] = (),
) -> dict:
    errors = list(configuration_errors)
    if sdk_available is None:
        try:
            import azure.identity  # noqa: F401
            import msgraph  # noqa: F401
            sdk_available = True
        except ImportError:
            sdk_available = False
    if settings is None:
        return {
            "locally_ready": False,
            "configuration_errors": errors,
            "secret_present": None,
            "secret_scope_error": None,
            "remote_verification": "not_tested",
            "selected_operations": [],
            "operation_status": {},
            "configuration": {},
            "requirements": {"tenant_id": False, "client_id": False, "user_id": False},
        }

    application_mode = settings.authentication_mode == "application"
    secret_present: bool | None = None
    secret_error: str | None = None
    if application_mode:
        secret_present, secret_error = _secret_presence()

    selected = sorted(
        f"{service}.{operation}"
        for service in OPERATIONS
        for operation in settings.selected(service)
    )
    statuses = {}
    for key in selected:
        service, operation = key.split(".", 1)
        status = operation_status(settings.authentication_mode, service, operation)
        definition = OPERATION_REGISTRY[key]
        remote_verification = getattr(definition, "remote_verification", "not_tested")
        evidence = getattr(definition, "remote_verification_evidence", ())
        if remote_verification != "verified" or not evidence:
            remote_verification = "not_tested"
            evidence = ()
        statuses[key] = {
            **asdict(status),
            "permissions": list(definition.permissions),
            "write": definition.write,
            "remote_verification": remote_verification,
            "remote_verification_evidence": list(evidence),
        }
    requirements = _requirements(settings, selected)
    configured = {
        "tenant_id": bool(settings.tenant_id.strip()),
        "client_id": bool(settings.client_id.strip()),
        "user_id": bool(settings.user_id.strip()),
    }
    missing = [name for name, required in requirements.items() if required and not configured[name]]
    if selected and application_mode and not secret_present:
        missing.append(SECRET_ENV_NAME)
    if not selected:
        missing.append("capabilities")
    all_executable = bool(selected) and all(item["executable"] for item in statuses.values())
    return {
        "locally_ready": bool(sdk_available and not errors and not missing and all_executable),
        "sdk_available": bool(sdk_available),
        "configuration_errors": errors,
        "missing": sorted(set(missing)),
        "secret_present": secret_present,
        "secret_scope_error": secret_error,
        "authentication_mode": settings.authentication_mode,
        "remote_verification": "not_tested",
        "selected_operations": selected,
        "operation_status": statuses,
        "configuration": {
            "tenant_id_configured": configured["tenant_id"],
            "client_id_configured": configured["client_id"],
            "user_id_configured": configured["user_id"],
            "authentication_mode": settings.authentication_mode,
            "capabilities": {key: dict(value) for key, value in settings.capabilities.items()},
        },
        "requirements": requirements,
    }
