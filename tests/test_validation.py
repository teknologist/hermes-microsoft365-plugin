"""WP13 — the per-authentication-mode operation matrix and the strict argument validators.

What this file proves
---------------------
1. The registry records application support **and** delegated support separately (status,
   roles/scopes, admin consent, endpoint, write classification, implementation status,
   remote verification), and the delegated scopes are an independent record — never derived
   from the application role map.
2. One validator, ``microsoft365.validation.check``, rejects every invalid, disabled and
   unsupported payload before anything can happen, and it is the *same* call in the
   ``pre_tool_call`` hook (before approval) and in the tool dispatch path (defence in
   depth): identical payload in, identical rejection out.
3. No invalid payload can reach a secret lookup, a credential construction or a Graph client
   construction. The proof monkeypatches ``agent.secret_scope.get_secret``,
   ``azure.identity.ClientSecretCredential`` and ``msgraph.GraphServiceClient`` with
   functions that record the attempt and raise, then runs every invalid payload through the
   registered hook and the registered tool; the positive control shows the same probe *is*
   reached by a valid payload, so the guard — not an accident — is what blocks the rest.

Dispatch-path note: no service tool is registered while every operation is
``executable = False`` (WP13 must not flip that flag; WP6-WP9 own it). The tests that need
the *dispatch* half therefore simulate only ``active_actions`` — a plain function, no mock
library and no permissive double — so the registered handler closure is exercised exactly as
it will be once a handler exists, while the registry stays honest.

Only real objects are used: no ``SimpleNamespace``, no ``MagicMock``, no permissive
``__getattr__``. Nothing here touches the network, a tenant or a credential.
"""
from __future__ import annotations

import base64
import json
from pathlib import Path

import pytest

#: A synthetic probe value with no credential shape (same convention as tests/test_errors.py).
VERIFICATION_PROBE = "SENTINEL-REDACTION-PROBE-8F3A2B"

USER = "user@example.invalid"
DRIVE = "b!drive-id"
ITEM = "01ABCDEFGHIJKLMN"
PLAN = "plan-id"
BUCKET = "bucket-id"
LIST = "list-id"
TASK = "task-id"
EVENT = "event-id"
TEAM = "team-id"
CHANNEL = "channel-id"
ETAG = 'W/"etag-value"'
BASE64_PAYLOAD = base64.b64encode(b"hello world").decode("ascii")

RECIPIENT = {"address": "ana@example.invalid", "name": "Ana"}
AWARE_START = "2026-09-16T10:00:00-03:00"
AWARE_END = "2026-09-16T11:00:00-03:00"
TIME_ZONE = "America/Sao_Paulo"

SERVICES = ("outlook", "sharepoint", "onedrive", "calendar", "teams", "todo", "planner")

#: The two Teams operations this plugin can only reach in delegated mode: in application mode
#: the endpoint does not exist for an app-only token. They are covered by their own test
#: instead of the happy-path table.
APPLICATION_UNREACHABLE = ("teams.search_messages", "teams.send_messages")

#: One known-good payload per operation that application mode can run today.
VALID_PAYLOADS: dict[str, dict] = {
    "outlook.list_attachments": {"action": "list_attachments", "user_id": USER, "message_id": "message"},
    "outlook.download_attachment": {"action": "download_attachment", "user_id": USER, "message_id": "message", "attachment_id": "attachment"},
    "outlook.search": {"action": "search", "user_id": USER, "query": "subject:invoice", "top": 10},
    "outlook.read": {"action": "read", "user_id": USER, "message_id": "AAMkAG"},
    "outlook.create_draft": {
        "action": "create_draft",
        "user_id": USER,
        "subject": "Kickoff",
        "body": "<p>hello</p>",
        "to_recipients": [RECIPIENT],
    },
    "outlook.send": {"action": "send", "user_id": USER, "message_id": "AAMkAG"},
    "sharepoint.search": {"action": "search", "query": "projeto", "top": 5},
    "sharepoint.read": {"action": "read", "drive_id": DRIVE, "item_id": ITEM},
    "sharepoint.download_files": {"action": "download_files", "drive_id": DRIVE, "item_path": "docs/a.txt"},
    "sharepoint.upload_files": {
        "action": "upload_files",
        "drive_id": DRIVE,
        "item_path": "docs/b.txt",
        "content_base64": BASE64_PAYLOAD,
        "content_type": "text/plain",
    },
    "onedrive.search": {"action": "search", "query": "projeto", "drive_id": DRIVE},
    "onedrive.read": {"action": "read", "drive_id": DRIVE, "item_path": "docs/a.txt", "select": ["name", "size"]},
    "onedrive.download_files": {"action": "download_files", "drive_id": DRIVE, "item_id": ITEM},
    "onedrive.upload_files": {
        "action": "upload_files",
        "drive_id": DRIVE,
        "item_id": ITEM,
        "content_base64": BASE64_PAYLOAD,
    },
    "calendar.search": {"action": "search", "user_id": USER, "time_zone": TIME_ZONE, "top": 10},
    "calendar.create_events": {
        "action": "create_events",
        "user_id": USER,
        "subject": "Kickoff",
        "start_date_time": AWARE_START,
        "end_date_time": AWARE_END,
        "time_zone": TIME_ZONE,
        "attendees": [RECIPIENT],
        "is_all_day": False,
    },
    "calendar.update_events": {
        "action": "update_events",
        "user_id": USER,
        "event_id": EVENT,
        "etag": ETAG,
        "fields": {"subject": "Renamed"},
    },
    "todo.list_task_lists": {"action": "list_task_lists", "user_id": USER},
    "todo.search": {"action": "search", "user_id": USER, "todo_list_id": LIST, "title": "Draft"},
    "todo.read": {"action": "read", "user_id": USER, "todo_list_id": LIST, "todo_task_id": TASK},
    "todo.create_tasks": {
        "action": "create_tasks",
        "user_id": USER,
        "todo_list_id": LIST,
        "fields": {"title": "Draft"},
    },
    "todo.update_tasks": {
        "action": "update_tasks",
        "user_id": USER,
        "todo_list_id": LIST,
        "todo_task_id": TASK,
        "fields": {"title": "Renamed"},
    },
    "teams.list_teams": {"action": "list_teams", "user_id": USER},
    "teams.list_channels": {"action": "list_channels", "team_id": TEAM, "select": ["displayName"]},
    "planner.list_plans": {"action": "list_plans"},
    "planner.list_buckets": {"action": "list_buckets", "plan_id": PLAN},
    "planner.list_tasks": {"action": "list_tasks", "plan_id": PLAN, "top": 20},
    "planner.read": {"action": "read", "planner_task_id": TASK},
    "planner.create_tasks": {
        "action": "create_tasks",
        "plan_id": PLAN,
        "bucket_id": BUCKET,
        "title": "Draft",
    },
    "planner.update_tasks": {
        "action": "update_tasks",
        "planner_task_id": TASK,
        "etag": ETAG,
        "fields": {"title": "Renamed"},
    },
}

#: Payloads that must be rejected in the mode they are declared for.
#: ``(label, service, payload, expected_category, mode)`` where ``mode`` selects the
#: settings used for the call (``application``, ``delegated``, ``disabled_read``, ``hybrid``).
INVALID_CALLS: tuple[tuple[str, str, object, str, str], ...] = (
    # -- structure ---------------------------------------------------------------------
    ("arguments are not an object", "outlook", None, "validation_error", "application"),
    ("arguments are a list", "outlook", ["search"], "validation_error", "application"),
    ("arguments are a string", "outlook", USER, "validation_error", "application"),
    ("unknown operation", "outlook", {"action": "delete", "user_id": USER}, "validation_error", "application"),
    ("missing action", "outlook", {"user_id": USER}, "validation_error", "application"),
    ("action is not a string", "outlook", {"action": 7}, "validation_error", "application"),
    ("unknown property", "outlook", {"action": "search", "user_id": USER, "nope": 1}, "validation_error", "application"),
    ("non-string property key", "outlook", {"action": "search", "user_id": USER, 7: 1}, "validation_error", "application"),
    ("missing required identifier", "outlook", {"action": "search"}, "validation_error", "application"),
    ("required identifier is None", "outlook", {"action": "read", "user_id": USER, "message_id": None}, "validation_error", "application"),
    ("missing planner task id", "planner", {"action": "read"}, "validation_error", "application"),
    ("missing todo task id", "todo", {"action": "read", "user_id": USER, "todo_list_id": LIST}, "validation_error", "application"),
    ("missing drive id", "onedrive", {"action": "read", "item_id": ITEM}, "validation_error", "application"),
    ("top is not declared for channels", "teams", {"action": "list_channels", "team_id": TEAM, "top": 5}, "validation_error", "application"),
    # -- coercion ----------------------------------------------------------------------
    ("boolean given as string", "outlook", {"action": "send", "user_id": USER, "message_id": "M", "save_to_sent_items": "false"}, "validation_error", "application"),
    ("boolean given as single-char string", "outlook", {"action": "send", "user_id": USER, "message_id": "M", "save_to_sent_items": "true"}, "validation_error", "application"),
    ("boolean given as integer", "calendar", {"action": "create_events", "user_id": USER, "subject": "s", "start_date_time": AWARE_START, "end_date_time": AWARE_END, "is_all_day": 1}, "validation_error", "application"),
    ("identifier given as integer", "outlook", {"action": "read", "user_id": USER, "message_id": 7}, "validation_error", "application"),
    ("identifier given as boolean", "outlook", {"action": "read", "user_id": USER, "message_id": True}, "validation_error", "application"),
    ("identifier given as a list", "outlook", {"action": "search", "user_id": ["u"]}, "validation_error", "application"),
    ("integer given as string", "calendar", {"action": "search", "user_id": USER, "top": "10"}, "validation_error", "application"),
    ("integer given as boolean", "calendar", {"action": "search", "user_id": USER, "top": True}, "validation_error", "application"),
    # -- strings -----------------------------------------------------------------------
    ("blank required text", "outlook", {"action": "create_draft", "user_id": USER, "subject": "", "to_recipients": [RECIPIENT]}, "validation_error", "application"),
    ("identifier is only whitespace", "outlook", {"action": "search", "user_id": "   "}, "validation_error", "application"),
    ("identifier is not stripped", "outlook", {"action": "read", "user_id": USER, "message_id": " AAMkAG "}, "validation_error", "application"),
    ("identifier is too long", "outlook", {"action": "read", "user_id": "u" * 300, "message_id": "M"}, "validation_error", "application"),
    ("identifier carries a control character", "outlook", {"action": "read", "user_id": USER, "message_id": "AA\nMk"}, "validation_error", "application"),
    ("slash-me user id in application mode", "outlook", {"action": "search", "user_id": "/me"}, "validation_error", "application"),
    ("query is too long", "outlook", {"action": "search", "user_id": USER, "query": "q" * 513}, "validation_error", "application"),
    ("text carries a null character", "outlook", {"action": "create_draft", "user_id": USER, "subject": "a\x00b", "to_recipients": [RECIPIENT]}, "validation_error", "application"),
    # -- collections and recipients -----------------------------------------------------
    ("recipients given as a string", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": "ana@example.invalid"}, "validation_error", "application"),
    ("recipients empty", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": []}, "validation_error", "application"),
    ("recipients too many", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": [RECIPIENT] * 101}, "validation_error", "application"),
    ("recipient is not an address", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": [7]}, "validation_error", "application"),
    ("recipient is blank", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": [{"address": ""}]}, "validation_error", "application"),
    ("recipient has no domain separator", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": ["ana.invalid"]}, "validation_error", "application"),
    ("recipient has an embedded space", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": ["ana @example.invalid"]}, "validation_error", "application"),
    ("recipient carries a header injection", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": ["ana@example.invalid\r\nBcc: x@example.invalid"]}, "validation_error", "application"),
    ("recipient mapping has an unknown key", "outlook", {"action": "create_draft", "user_id": USER, "subject": "s", "to_recipients": [{"address": "ana@example.invalid", "cc": True}]}, "validation_error", "application"),
    ("select too many fields", "onedrive", {"action": "read", "drive_id": DRIVE, "item_id": ITEM, "select": ["f"] * 21}, "validation_error", "application"),
    ("select holds a non-string", "onedrive", {"action": "read", "drive_id": DRIVE, "item_id": ITEM, "select": [7]}, "validation_error", "application"),
    # -- paths -------------------------------------------------------------------------
    ("path traversal", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": "../secret.txt"}, "validation_error", "application"),
    ("path traversal in the middle", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": "docs/../secret.txt"}, "validation_error", "application"),
    ("absolute path", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": "/etc/passwd"}, "validation_error", "application"),
    ("windows drive path", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": "C:/temp/a.txt"}, "validation_error", "application"),
    ("backslash path", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": "docs\\a.txt"}, "validation_error", "application"),
    ("path with a control character", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": "a\x00b"}, "validation_error", "application"),
    ("empty path", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": ""}, "validation_error", "application"),
    ("path is not stripped", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_path": " docs/a.txt"}, "validation_error", "application"),
    # -- timestamps and time zones -----------------------------------------------------
    ("timestamp is not a string", "calendar", {"action": "create_events", "user_id": USER, "subject": "s", "start_date_time": 20260916, "end_date_time": AWARE_END}, "validation_error", "application"),
    ("timestamp is not a date", "calendar", {"action": "create_events", "user_id": USER, "subject": "s", "start_date_time": "next tuesday", "end_date_time": AWARE_END}, "validation_error", "application"),
    ("timestamp has no time component", "calendar", {"action": "create_events", "user_id": USER, "subject": "s", "start_date_time": "2026-09-16", "end_date_time": AWARE_END}, "validation_error", "application"),
    ("naive timestamp without a time zone", "calendar", {"action": "create_events", "user_id": USER, "subject": "s", "start_date_time": "2026-09-16T10:00:00", "end_date_time": "2026-09-16T11:00:00"}, "validation_error", "application"),
    ("time zone is not a string", "calendar", {"action": "search", "user_id": USER, "time_zone": 7}, "validation_error", "application"),
    ("time zone is blank", "calendar", {"action": "search", "user_id": USER, "time_zone": "  "}, "validation_error", "application"),
    ("time zone has a tab", "calendar", {"action": "search", "user_id": USER, "time_zone": "America/\tSao_Paulo"}, "validation_error", "application"),
    # -- etags -------------------------------------------------------------------------
    ("etag is not a string", "calendar", {"action": "update_events", "user_id": USER, "event_id": EVENT, "etag": 7, "fields": {"subject": "s"}}, "validation_error", "application"),
    ("etag is blank", "calendar", {"action": "update_events", "user_id": USER, "event_id": EVENT, "etag": "", "fields": {"subject": "s"}}, "validation_error", "application"),
    ("etag wildcard", "calendar", {"action": "update_events", "user_id": USER, "event_id": EVENT, "etag": "*", "fields": {"subject": "s"}}, "validation_error", "application"),
    ("etag carries a header injection", "calendar", {"action": "update_events", "user_id": USER, "event_id": EVENT, "etag": 'W/"a"\r\nX-Injected: 1', "fields": {"subject": "s"}}, "validation_error", "application"),
    ("etag is too long", "calendar", {"action": "update_events", "user_id": USER, "event_id": EVENT, "etag": '"' + "a" * 300 + '"', "fields": {"subject": "s"}}, "validation_error", "application"),
    ("missing etag", "planner", {"action": "update_tasks", "planner_task_id": TASK, "fields": {"title": "x"}}, "validation_error", "application"),
    # -- base64 ------------------------------------------------------------------------
    ("base64 given as bytes", "sharepoint", {"action": "upload_files", "drive_id": DRIVE, "item_id": ITEM, "content_base64": b"aGVsbG8="}, "validation_error", "application"),
    ("base64 with an invalid alphabet", "sharepoint", {"action": "upload_files", "drive_id": DRIVE, "item_id": ITEM, "content_base64": "aGVsbG8*"}, "validation_error", "application"),
    ("base64 without padding", "sharepoint", {"action": "upload_files", "drive_id": DRIVE, "item_id": ITEM, "content_base64": "aGVsbG8"}, "validation_error", "application"),
    ("content type is not a media type", "sharepoint", {"action": "upload_files", "drive_id": DRIVE, "item_id": ITEM, "content_base64": BASE64_PAYLOAD, "content_type": "plain text"}, "validation_error", "application"),
    # -- patch fields ------------------------------------------------------------------
    ("empty patch", "todo", {"action": "update_tasks", "user_id": USER, "todo_list_id": LIST, "todo_task_id": TASK, "fields": {}}, "validation_error", "application"),
    ("patch with a field the endpoint does not accept", "todo", {"action": "update_tasks", "user_id": USER, "todo_list_id": LIST, "todo_task_id": TASK, "fields": {"subject": "x"}}, "validation_error", "application"),
    ("patch value is None", "todo", {"action": "update_tasks", "user_id": USER, "todo_list_id": LIST, "todo_task_id": TASK, "fields": {"title": None}}, "validation_error", "application"),
    ("patch value is nested", "todo", {"action": "update_tasks", "user_id": USER, "todo_list_id": LIST, "todo_task_id": TASK, "fields": {"title": {"x": 1}}}, "validation_error", "application"),
    ("patch is a list", "todo", {"action": "update_tasks", "user_id": USER, "todo_list_id": LIST, "todo_task_id": TASK, "fields": ["title"]}, "validation_error", "application"),
    ("create without a title", "todo", {"action": "create_tasks", "user_id": USER, "todo_list_id": LIST, "fields": {"body": "x"}}, "validation_error", "application"),
    ("patch with an unknown planner field", "planner", {"action": "update_tasks", "planner_task_id": TASK, "etag": ETAG, "fields": {"nope": "x"}}, "validation_error", "application"),
    ("planner patch value is None", "planner", {"action": "update_tasks", "planner_task_id": TASK, "etag": ETAG, "fields": {"title": None}}, "validation_error", "application"),
    ("planner patch timestamp without a time zone", "planner", {"action": "update_tasks", "planner_task_id": TASK, "etag": ETAG, "fields": {"due_date_time": "2026-09-16T10:00:00"}}, "validation_error", "application"),
    # -- mutually exclusive forms ------------------------------------------------------
    ("send with both forms", "outlook", {"action": "send", "user_id": USER, "message_id": "M", "subject": "s", "to_recipients": [RECIPIENT]}, "validation_error", "application"),
    ("send with neither form", "outlook", {"action": "send", "user_id": USER}, "validation_error", "application"),
    ("read by id and by path", "sharepoint", {"action": "read", "drive_id": DRIVE, "item_id": ITEM, "item_path": "docs/a.txt"}, "validation_error", "application"),
    ("read without an item", "sharepoint", {"action": "read", "drive_id": DRIVE}, "validation_error", "application"),
    ("search scoped to a drive and a site", "sharepoint", {"action": "search", "query": "q", "drive_id": DRIVE, "site_id": "site-id"}, "validation_error", "application"),
    # -- configuration and mode --------------------------------------------------------
    ("disabled operation", "outlook", {"action": "read", "user_id": USER, "message_id": "M"}, "validation_error", "disabled_read"),
    ("teams search in application mode", "teams", {"action": "search_messages", "query": "hello"}, "unsupported_auth_mode", "application"),
    ("teams channel message in application mode", "teams", {"action": "send_messages", "team_id": TEAM, "channel_id": CHANNEL, "body": "hi"}, "unsupported_auth_mode", "application"),
    ("valid payload in delegated mode", "outlook", {"action": "search", "user_id": USER}, "unsupported_auth_mode", "delegated"),
    ("unknown mode", "outlook", {"action": "search", "user_id": USER}, "unsupported_auth_mode", "hybrid"),
)


def _settings(mode: str = "application"):
    from microsoft365.contract import OPERATIONS, Settings

    capabilities = {service: True for service in OPERATIONS}
    if mode == "disabled_read":
        capabilities["outlook"] = {"search": True, "read": False}
        mode = "application"
    if mode == "hybrid":
        # ``Settings.from_mapping`` refuses this mode (and registration then registers
        # nothing at all); the record is built directly so the validator's own refusal of an
        # unsupported authentication mode is still exercised.
        return Settings(authentication_mode="hybrid", capabilities={})
    return Settings.from_mapping({"capabilities": capabilities, "authentication_mode": mode})


def _check(service: str, payload, *, mode: str = "application"):
    from microsoft365.validation import check

    return check(_settings(mode), service=service, arguments=payload)


class RecordingContext:
    """A concrete plugin context: no mock, no permissive attribute access."""

    def __init__(self, config=None):
        self.config = config or {}
        self.tools: dict[str, dict] = {}
        self.hooks: dict[str, object] = {}

    def get_config(self, key, default=None):
        return self.config.get(key, default)

    def register_tool(self, name, **kwargs):
        self.tools[name] = kwargs

    def register_hook(self, name, callback):
        self.hooks[name] = callback


def _config(mode: str = "application", *, capabilities=None):
    from microsoft365.contract import OPERATIONS

    return {
        "capabilities": capabilities or {service: True for service in OPERATIONS},
        "authentication_mode": "application" if mode == "disabled_read" else mode,
    }


def _registered(mode: str = "application", *, capabilities=None, monkeypatch=None):
    """Register the plugin, optionally with every capability-enabled operation executable.

    Without ``monkeypatch`` the registration is exactly the milestone's: preflight, the hook,
    and one service tool per service that has an executable operation (today Outlook and
    Calendar, whose three verified reads are executable). With it, ``active_actions`` is
    replaced by a plain function that reports every capability-enabled operation as executable,
    including the writes, which is what a post-CORE-1 host will make true; no registry flag
    changes.
    """
    from microsoft365 import register, registration

    if monkeypatch is not None:
        from microsoft365.contract import OPERATIONS

        def simulated_active_actions(settings, service):
            if settings.authentication_mode != "application":
                return ()
            selected = settings.selected(service)
            return tuple(operation for operation in OPERATIONS[service] if operation in selected)

        monkeypatch.setattr(registration, "active_actions", simulated_active_actions)

    ctx = RecordingContext(_config(mode, capabilities=capabilities))
    register(ctx)
    return ctx


def test_the_simulated_dispatch_seam_is_only_active_with_a_monkeypatch(monkeypatch):
    """Guard on the test seam itself: without it, only the executable services register."""
    plain = _registered()
    simulated = _registered(monkeypatch=monkeypatch)

    # Exactly the services whose operations the registry declares executable (the nine
    # verified reads of WP6/WP7) get a tool; the other services keep only preflight.
    assert set(plain.tools) == {
        "microsoft365_preflight",
        "microsoft365_outlook",
        "microsoft365_calendar",
        "microsoft365_sharepoint",
        "microsoft365_onedrive",
        "microsoft365_teams",
        "microsoft365_todo",
        "microsoft365_planner",
        }
    assert {name for name in SERVICES if f"microsoft365_{name}" in plain.tools} == {
        key.split(".", 1)[0] for key in EXECUTABLE_READS
    }
    assert "microsoft365_outlook" in simulated.tools
    assert set(simulated.tools) == {"microsoft365_preflight"} | {f"microsoft365_{name}" for name in SERVICES}


# ======================================================================================
# 1. the per-mode matrix
# ======================================================================================


def test_every_operation_records_application_and_delegated_support_separately():
    from microsoft365.contract import OPERATION_REGISTRY

    for key, definition in OPERATION_REGISTRY.items():
        assert definition.app.mode == "application", key
        assert definition.delegated.mode == "delegated", key
        assert definition.app.permission_kind == "application_role", key
        assert definition.delegated.permission_kind == "delegated_scope", key
        assert definition.support_for("application") == definition.app, key
        assert definition.support_for("delegated") == definition.delegated, key
        assert definition.support_for("hybrid") is None, key


def test_application_roles_are_pinned_per_operation():
    from microsoft365.contract import OPERATION_REGISTRY

    assert {key: definition.app.permissions for key, definition in OPERATION_REGISTRY.items()} == {
        "outlook.list_attachments": ("Mail.Read", "Group.Read.All"),
        "outlook.download_attachment": ("Mail.Read", "Group.Read.All"),
        "outlook.search": ("Mail.Read", "Group.Read.All"),
        "outlook.read": ("Mail.Read", "Group.Read.All"),
        "outlook.create_draft": ("Mail.ReadWrite",),
        "outlook.send": ("Mail.Send",),
        "sharepoint.search": ("Sites.Read.All",),
        "sharepoint.read": ("Sites.Read.All",),
        "sharepoint.download_files": ("Files.Read.All",),
        "sharepoint.upload_files": ("Files.ReadWrite.All",),
        "onedrive.search": ("Files.Read.All",),
        "onedrive.read": ("Files.Read.All",),
        "onedrive.download_files": ("Files.Read.All",),
        "onedrive.upload_files": ("Files.ReadWrite.All",),
        "calendar.search": ("Calendars.Read",),
        "calendar.create_events": ("Calendars.ReadWrite",),
        "calendar.update_events": ("Calendars.ReadWrite",),
        "teams.list_teams": ("Team.ReadBasic.All",),
        "teams.list_channels": ("Channel.ReadBasic.All",),
        "teams.search_messages": (),
        "teams.send_messages": (),
        "todo.list_task_lists": ("Tasks.Read.All",),
        "todo.search": ("Tasks.Read.All",),
        "todo.read": ("Tasks.Read.All",),
        "todo.create_tasks": ("Tasks.ReadWrite.All",),
        "todo.update_tasks": ("Tasks.ReadWrite.All",),
        "planner.list_plans": ("Tasks.Read.All",),
        "planner.list_buckets": ("Tasks.Read.All",),
        "planner.list_tasks": ("Tasks.Read.All",),
        "planner.read": ("Tasks.Read.All",),
        "planner.create_tasks": ("Tasks.ReadWrite.All",),
        "planner.update_tasks": ("Tasks.ReadWrite.All",),
    }


def test_delegated_scopes_are_recorded_independently_of_the_application_roles():
    from microsoft365.contract import OPERATION_REGISTRY

    expected = {
        "outlook.list_attachments": ("Mail.Read", "Group.Read.All"),
        "outlook.download_attachment": ("Mail.Read", "Group.Read.All"),
        "outlook.search": ("Mail.Read", "Group.Read.All"),
        "outlook.read": ("Mail.Read", "Group.Read.All"),
        "outlook.create_draft": ("Mail.ReadWrite",),
        "outlook.send": ("Mail.Send",),
        "sharepoint.search": ("Sites.Read.All",),
        "sharepoint.read": ("Sites.Read.All",),
        "sharepoint.download_files": ("Files.Read.All",),
        "sharepoint.upload_files": ("Files.ReadWrite.All",),
        "onedrive.search": ("Files.Read.All",),
        "onedrive.read": ("Files.Read.All",),
        "onedrive.download_files": ("Files.Read.All",),
        "onedrive.upload_files": ("Files.ReadWrite.All",),
        "calendar.search": ("Calendars.Read",),
        "calendar.create_events": ("Calendars.ReadWrite",),
        "calendar.update_events": ("Calendars.ReadWrite",),
        "teams.list_teams": ("Team.ReadBasic.All",),
        "teams.list_channels": ("Channel.ReadBasic.All",),
        "teams.search_messages": ("Chat.Read", "ChannelMessage.Read.All"),
        "teams.send_messages": ("ChannelMessage.Send",),
        "todo.list_task_lists": ("Tasks.Read",),
        "todo.search": ("Tasks.Read",),
        "todo.read": ("Tasks.Read",),
        "todo.create_tasks": ("Tasks.ReadWrite",),
        "todo.update_tasks": ("Tasks.ReadWrite",),
        "planner.list_plans": ("Tasks.Read",),
        "planner.list_buckets": ("Tasks.Read",),
        "planner.list_tasks": ("Tasks.Read",),
        "planner.read": ("Tasks.Read",),
        "planner.create_tasks": ("Tasks.ReadWrite",),
        "planner.update_tasks": ("Tasks.ReadWrite",),
    }

    assert {key: definition.delegated.permissions for key, definition in OPERATION_REGISTRY.items()} == expected

    # The two records are independent: Planner carries an app role that is not the delegated
    # scope, and the operations application mode cannot use at all still declare theirs.
    assert OPERATION_REGISTRY["planner.read"].app.permissions == ("Tasks.Read.All",)
    assert OPERATION_REGISTRY["planner.read"].delegated.permissions == ("Tasks.Read",)
    assert OPERATION_REGISTRY["teams.search_messages"].app.permissions == ()
    assert OPERATION_REGISTRY["teams.search_messages"].delegated.permissions == (
        "Chat.Read",
        "ChannelMessage.Read.All",
    )
    for key, definition in OPERATION_REGISTRY.items():
        assert definition.delegated.permissions == expected[key], key


def test_per_mode_support_requires_evidence_and_fails_closed():
    from microsoft365.contract import ModeSupport

    with pytest.raises(ValueError, match="unknown authentication mode"):
        ModeSupport(mode="hybrid", status="supported", permissions=("Mail.Read",))
    with pytest.raises(ValueError, match="unknown authentication mode support status"):
        ModeSupport(mode="application", status="probably", permissions=("Mail.Read",))
    with pytest.raises(ValueError, match="verified permission claim"):
        ModeSupport(
            mode="delegated",
            status="not_implemented",
            permissions=("Mail.Read",),
            permission_status="verified",
        )
    with pytest.raises(ValueError, match="must claim the permission it needs"):
        ModeSupport(mode="application", status="supported", permissions=())
    with pytest.raises(ValueError, match="cannot claim a permission on an unsupported mode"):
        ModeSupport(
            mode="application",
            status="unsupported_auth_mode",
            permissions=("Mail.Read",),
            permission_status="documented_not_verified",
        )
    with pytest.raises(ValueError, match="cannot claim a permission on an unsupported mode"):
        ModeSupport(
            mode="application",
            status="unsupported_auth_mode",
            permissions=(),
            permission_status="documented_not_verified",
        )
    with pytest.raises(ValueError, match="must record the basis"):
        ModeSupport(
            mode="application",
            status="supported",
            permissions=("Mail.Read",),
            permission_status="documented_not_verified",
            admin_consent=True,
        )
    with pytest.raises(ValueError, match="cannot require admin consent"):
        ModeSupport(
            mode="delegated",
            status="not_implemented",
            permissions=(),
            admin_consent=True,
            admin_consent_basis="impossible",
        )
    with pytest.raises(ValueError, match="must state why consent is not required"):
        ModeSupport(
            mode="delegated",
            status="not_implemented",
            permissions=(),
            admin_consent=False,
        )


def test_admin_consent_and_remote_verification_are_derived_and_fail_closed():
    from microsoft365.contract import (
        IMPLEMENTED_AUTH_MODES,
        OPERATION_REGISTRY,
        REMOTE_VERIFICATION_STATUSES,
        OperationDefinition,
    )

    assert IMPLEMENTED_AUTH_MODES == frozenset({"application"})
    assert REMOTE_VERIFICATION_STATUSES == frozenset({"not_tested", "verified"})

    for key, definition in OPERATION_REGISTRY.items():
        assert definition.remote_verification == "not_tested", key
        assert definition.remote_verification_evidence == (), key
        claimed = [mode for mode in (definition.app, definition.delegated) if mode.permissions]
        assert definition.admin_consent is any(mode.admin_consent for mode in claimed), key
        for mode in (definition.app, definition.delegated):
            if mode.permissions:
                assert mode.admin_consent is True, key
                assert mode.admin_consent_basis.strip(), key
            else:
                assert mode.admin_consent is False, key

    with pytest.raises(ValueError, match="unknown remote verification status"):
        OperationDefinition(
            service="outlook",
            operation="search",
            permissions=("Mail.Read", "Group.Read.All"),
            write=False,
            remote_verification="probably",
        )
    with pytest.raises(ValueError, match="cannot claim remote verification"):
        OperationDefinition(
            service="outlook",
            operation="search",
            permissions=("Mail.Read", "Group.Read.All"),
            write=False,
            remote_verification="verified",
        )


def test_operation_status_is_derived_from_the_mode_matrix():
    from microsoft365.contract import OPERATION_REGISTRY, operation_status

    for key, definition in OPERATION_REGISTRY.items():
        service, operation = key.split(".", 1)
        for mode in ("application", "delegated"):
            support = definition.support_for(mode)
            status = operation_status(mode, service, operation)
            assert status.auth_status == support.status, (key, mode)
            assert status.reason == support.reason, (key, mode)
            assert status.implementation_status == definition.implementation_status, (key, mode)
            # executability is the registry flag *and* an available mode: the three verified
            # reads run in application mode, nothing runs in the unimplemented delegated mode
            assert status.executable is (
                mode == "application" and (key in EXECUTABLE_READS or key in {"todo.list_task_lists", "todo.search", "todo.read", "planner.list_plans", "planner.list_buckets", "planner.list_tasks", "planner.read"})
            ), (key, mode)

    unknown = operation_status("application", "outlook", "delete")
    assert unknown.auth_status == "unknown_operation"
    assert unknown.executable is False
    assert operation_status("hybrid", "outlook", "search").auth_status == "not_implemented"
    assert operation_status("application", "teams", "search_messages").auth_status == "unsupported_auth_mode"
    assert operation_status("application", "todo", "create_tasks").auth_status == "not_verified"
    assert operation_status("delegated", "outlook", "search").auth_status == "not_implemented"


def test_the_two_teams_operations_are_unsupported_in_application_mode():
    from microsoft365.contract import OPERATION_REGISTRY, operation_status

    for key in APPLICATION_UNREACHABLE:
        definition = OPERATION_REGISTRY[key]
        assert definition.app.status == "unsupported_auth_mode", key
        assert definition.app.permissions == (), key
        assert definition.app.permission_status == "no_permission_claimed", key
        assert definition.delegated.permissions, key
        assert definition.delegated.status == "not_implemented", key
        assert "delegated" in definition.delegated.reason, key
        service, operation = key.split(".", 1)
        assert operation_status("application", service, operation).auth_status == "unsupported_auth_mode"
        assert operation_status("delegated", service, operation).auth_status == "not_implemented"


def test_only_the_three_verified_reads_are_executable_and_every_handler_is_registered():
    """The milestone pin, in its exact form: named executable set, exact dispatch mapping.

    This replaces the earlier "no operation is executable and nothing is registered" pin. It is
    strictly stronger: it fails if a write is flipped, if a read loses its flag, if a handler is
    removed from the registry, or if a handler is added without a contract case behind it.
    """
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.registration import HANDLER_TABLE

    assert len(OPERATION_REGISTRY) == 32
    assert {
        key for key, definition in OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_READS)
    assert set(HANDLER_TABLE) == set(EXECUTABLE_READS) | set(WITHHELD_WRITES) | {"teams.search_messages", "teams.send_messages"}

    for key, definition in OPERATION_REGISTRY.items():
        assert definition.executable is (key in EXECUTABLE_READS), key
        if definition.executable:
            assert key in HANDLER_TABLE, key
            assert definition.contract_cases, key
        elif key in WITHHELD_WRITES:
            # implemented and withheld, never removed (invariant 14)
            assert key in HANDLER_TABLE, key
            assert definition.write is True, key
        elif key in {"teams.search_messages", "teams.send_messages"}:
            assert key in HANDLER_TABLE, key
            assert definition.write is (key == "teams.send_messages"), key
        else:
            assert key not in HANDLER_TABLE, key
        assert definition.endpoint == "; ".join(definition.endpoints), key
        assert definition.write is (
            definition.operation
            in {
                "create_draft",
                "send",
                "upload_files",
                "create_events",
                "update_events",
                "send_messages",
                "create_tasks",
                "update_tasks",
            }
        ), key


# ======================================================================================
# 2. the argument contract table
# ======================================================================================


def test_the_argument_contract_covers_exactly_the_registry_and_every_operation_has_one():
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.validation import ARGUMENT_CONTRACTS

    assert set(ARGUMENT_CONTRACTS) == set(OPERATION_REGISTRY)
    for key, contract in ARGUMENT_CONTRACTS.items():
        assert contract.key == key
        assert contract.properties, key
        names = [spec.name for spec in contract.properties]
        assert len(names) == len(set(names)), key
        assert "action" not in names, key
        assert "operation" not in names, key


def test_every_valid_payload_is_accepted_and_the_table_covers_the_whole_catalogue():
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.validation import check

    assert set(VALID_PAYLOADS) | set(APPLICATION_UNREACHABLE) == set(OPERATION_REGISTRY)
    settings = _settings("application")
    for key, payload in VALID_PAYLOADS.items():
        service = key.split(".", 1)[0]
        assert check(settings, service=service, arguments=payload) is None, key


def test_user_scoped_operations_require_a_real_user_id():
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.validation import ARGUMENT_CONTRACTS, USER_SCOPED_OPERATIONS

    assert USER_SCOPED_OPERATIONS == frozenset(
        key
        for key in OPERATION_REGISTRY
        if key.split(".", 1)[0] in {"outlook", "calendar", "todo"} or key == "teams.list_teams"
    )
    for key, contract in ARGUMENT_CONTRACTS.items():
        user_id = next((spec for spec in contract.properties if spec.name == "user_id"), None)
        if key in USER_SCOPED_OPERATIONS:
            assert user_id is not None and user_id.required, key
        else:
            assert user_id is None, key


def test_todo_patch_allowlist_is_the_one_the_contract_defines():
    from microsoft365.contract import TODO_TASK_WRITE_FIELDS
    from microsoft365.validation import ARGUMENT_CONTRACTS

    for key in ("todo.create_tasks", "todo.update_tasks"):
        contract = ARGUMENT_CONTRACTS[key]
        fields = next(spec for spec in contract.properties if spec.name == "fields")
        assert {name for name, _ in fields.fields} == set(TODO_TASK_WRITE_FIELDS), key
        assert contract.patch_allowlist == tuple(name for name, _ in fields.fields), key


def test_argument_contract_rejects_a_malformed_specification():
    from microsoft365.validation import ArgumentSpec, MutualExclusion, OperationArguments

    with pytest.raises(ValueError, match="unknown argument kind"):
        OperationArguments(key="x.y", properties=(ArgumentSpec(name="a", kind="whatever"),))
    with pytest.raises(ValueError, match="must declare its allowed values"):
        OperationArguments(key="x.y", properties=(ArgumentSpec(name="a", kind="enum"),))
    with pytest.raises(ValueError, match="must declare its bounds"):
        OperationArguments(key="x.y", properties=(ArgumentSpec(name="a", kind="integer"),))
    with pytest.raises(ValueError, match="must declare its field allowlist"):
        OperationArguments(key="x.y", properties=(ArgumentSpec(name="a", kind="mapping"),))
    with pytest.raises(ValueError, match="must declare its bounds"):
        OperationArguments(
            key="x.y",
            properties=(
                ArgumentSpec(name="a", kind="mapping", fields=(("percent", "integer"),)),
            ),
        )
    with pytest.raises(ValueError, match="must be unique"):
        OperationArguments(
            key="x.y",
            properties=(ArgumentSpec(name="a", kind="text"), ArgumentSpec(name="a", kind="text")),
        )
    with pytest.raises(ValueError, match="structural argument"):
        OperationArguments(key="x.y", properties=(ArgumentSpec(name="action", kind="text"),))
    with pytest.raises(ValueError, match="unknown exclusion mode"):
        OperationArguments(
            key="x.y",
            properties=(ArgumentSpec(name="a", kind="text"), ArgumentSpec(name="b", kind="text")),
            exclusions=(MutualExclusion(forms=(("a",),), mode="one_or_two"),),
        )
    with pytest.raises(ValueError, match="undeclared property"):
        OperationArguments(
            key="x.y",
            properties=(ArgumentSpec(name="a", kind="text"),),
            exclusions=(MutualExclusion(forms=(("b",),), mode="exactly_one"),),
        )


def test_operation_schema_is_derived_from_the_same_spec_table():
    from microsoft365.validation import ARGUMENT_CONTRACTS, operation_schema

    for key, contract in ARGUMENT_CONTRACTS.items():
        schema = operation_schema(key)
        assert set(schema["properties"]) == {"action"} | {spec.name for spec in contract.properties}, key
        assert schema["required"] == ["action"] + sorted(
            spec.name for spec in contract.properties if spec.required
        ), key
        assert schema["additionalProperties"] is False, key
        assert schema["properties"]["action"]["enum"] == [key.split(".", 1)[1]], key


def test_operation_schema_types_follow_the_declared_kind():
    from microsoft365.validation import operation_schema

    assert operation_schema("outlook.search")["properties"]["top"]["type"] == "integer"
    assert operation_schema("outlook.search")["properties"]["user_id"]["type"] == "string"
    assert operation_schema("teams.list_channels")["properties"]["select"]["type"] == "array"
    assert operation_schema("calendar.create_events")["properties"]["is_all_day"]["type"] == "boolean"
    assert operation_schema("teams.send_messages")["properties"]["content_type"]["enum"] == ["text", "html"]


# ======================================================================================
# 3. every rejection rule, one guard at a time
# ======================================================================================


@pytest.mark.parametrize(
    "label,service,payload,expected_category,mode",
    INVALID_CALLS,
    ids=[case[0] for case in INVALID_CALLS],
)
def test_every_invalid_payload_is_rejected_with_the_expected_category(
    label, service, payload, expected_category, mode
):
    rejection = _check(service, payload, mode=mode)

    assert rejection is not None, label
    assert rejection.category == expected_category, label
    assert rejection.message, label
    assert rejection.to_payload() == {
        "error": expected_category,
        "message": rejection.message,
        "retryable": False,
    }, label


def test_rejection_names_the_operation_and_the_offending_argument():
    rejection = _check("outlook", {"action": "search", "user_id": USER, "nope": 1})

    assert "microsoft365.outlook.search" in rejection.message
    assert "nope" in rejection.message
    # the allowed arguments are reported back, so the model can correct itself
    assert "user_id" in rejection.message


def test_rejection_never_echoes_the_rejected_value_and_bounds_the_argument_name():
    from microsoft365.validation import check

    settings = _settings("application")
    secret_shaped = check(
        settings,
        service="outlook",
        arguments={"action": "send", "user_id": USER, "message_id": "M", "save_to_sent_items": VERIFICATION_PROBE},
    )
    assert secret_shaped is not None
    assert VERIFICATION_PROBE not in secret_shaped.message
    assert VERIFICATION_PROBE not in json.dumps(secret_shaped.to_payload())

    hostile_key = "k" * 300
    bounded = check(settings, service="outlook", arguments={"action": "search", "user_id": USER, hostile_key: 1})
    assert bounded is not None
    assert hostile_key not in bounded.message
    assert len(bounded.message) < 300


def test_oversized_base64_is_refused_before_decoding():
    from microsoft365.validation import MAX_UPLOAD_DECODED_BYTES, MAX_UPLOAD_ENCODED_CHARS, check

    # The encoded ceiling is exactly the encoded size of a maximum payload, so the two bounds
    # cannot drift apart.
    assert MAX_UPLOAD_ENCODED_CHARS == 4 * ((MAX_UPLOAD_DECODED_BYTES + 2) // 3)

    settings = _settings("application")
    accepted = check(
        settings,
        service="sharepoint",
        arguments={
            "action": "upload_files",
            "drive_id": DRIVE,
            "item_id": ITEM,
            "content_base64": base64.b64encode(b"A" * MAX_UPLOAD_DECODED_BYTES).decode("ascii"),
        },
    )
    assert accepted is None

    # Over the ceiling: refused as "too large", and with a payload that would also fail the
    # alphabet check, so the refusal provably happens on size before anything decodes it.
    refusal = check(
        settings,
        service="sharepoint",
        arguments={
            "action": "upload_files",
            "drive_id": DRIVE,
            "item_id": ITEM,
            "content_base64": "?" * (MAX_UPLOAD_ENCODED_CHARS + 4),
        },
    )
    assert refusal is not None
    assert refusal.category == "validation_error"
    assert "encoded" in refusal.message
    assert "alphabet" not in refusal.message


def test_bounded_strings_accept_the_declared_limit_and_refuse_one_more_character():
    from microsoft365.validation import MAX_IDENTIFIER_LENGTH, MAX_TEXT_LENGTH, check

    settings = _settings("application")

    at_limit = check(
        settings,
        service="outlook",
        arguments={"action": "read", "user_id": "u" * MAX_IDENTIFIER_LENGTH, "message_id": "M"},
    )
    assert at_limit is None
    over_limit = check(
        settings,
        service="outlook",
        arguments={"action": "read", "user_id": "u" * (MAX_IDENTIFIER_LENGTH + 1), "message_id": "M"},
    )
    assert over_limit is not None

    body_ok = check(
        settings,
        service="outlook",
        arguments={
            "action": "create_draft",
            "user_id": USER,
            "subject": "s",
            "body": "b" * MAX_TEXT_LENGTH,
            "to_recipients": [RECIPIENT],
        },
    )
    assert body_ok is None
    body_too_long = check(
        settings,
        service="outlook",
        arguments={
            "action": "create_draft",
            "user_id": USER,
            "subject": "s",
            "body": "b" * (MAX_TEXT_LENGTH + 1),
            "to_recipients": [RECIPIENT],
        },
    )
    assert body_too_long is not None


def test_collection_bounds_accept_the_limit_and_refuse_one_more():
    from microsoft365.validation import MAX_COLLECTION_ITEMS, check

    settings = _settings("application")
    at_limit = check(
        settings,
        service="outlook",
        arguments={
            "action": "create_draft",
            "user_id": USER,
            "subject": "s",
            "to_recipients": [RECIPIENT] * MAX_COLLECTION_ITEMS,
        },
    )
    assert at_limit is None
    over_limit = check(
        settings,
        service="outlook",
        arguments={
            "action": "create_draft",
            "user_id": USER,
            "subject": "s",
            "to_recipients": [RECIPIENT] * (MAX_COLLECTION_ITEMS + 1),
        },
    )
    assert over_limit is not None


def test_etag_accepts_both_strong_and_weak_forms_but_never_a_wildcard():
    from microsoft365.validation import check

    settings = _settings("application")
    # Sample ETag values. The first is a well-known published example value, not a credential;
    # the marker below keeps the secret scanner honest about that.
    for etag in ('"33a64df551425fcc55e4d42a148795d9f25f89d4"', ETAG, "unquoted-etag-token"):  # pragma: allowlist secret
        assert (
            check(
                settings,
                service="calendar",
                arguments={
                    "action": "update_events",
                    "user_id": USER,
                    "event_id": EVENT,
                    "etag": etag,
                    "fields": {"subject": "s"},
                },
            )
            is None
        ), etag
    for etag in ("*", '"', 'W/"a"W/"b"', ""):
        assert (
            check(
                settings,
                service="calendar",
                arguments={
                    "action": "update_events",
                    "user_id": USER,
                    "event_id": EVENT,
                    "etag": etag,
                    "fields": {"subject": "s"},
                },
            )
            is not None
        ), etag


def test_naive_timestamp_is_accepted_when_the_operation_carries_a_time_zone():
    from microsoft365.validation import check

    accepted = check(
        _settings("application"),
        service="calendar",
        arguments={
            "action": "create_events",
            "user_id": USER,
            "subject": "s",
            "start_date_time": "2026-09-16T10:00:00",
            "end_date_time": "2026-09-16T11:00:00",
            "time_zone": TIME_ZONE,
        },
    )

    assert accepted is None


def test_delegated_mode_is_refused_as_unsupported_until_it_exists():
    from microsoft365.validation import check

    rejection = check(_settings("delegated"), service="outlook", arguments={"action": "search", "user_id": USER})

    assert rejection is not None
    assert rejection.category == "unsupported_auth_mode"
    assert "delegated" in rejection.message
    assert check(_settings("application"), service="outlook", arguments={"action": "search", "user_id": USER}) is None


def test_disabled_operation_is_refused_even_when_its_arguments_are_valid():
    from microsoft365.validation import check

    settings = _settings("disabled_read")
    rejection = check(settings, service="outlook", arguments={"action": "read", "user_id": USER, "message_id": "M"})

    assert rejection is not None
    assert rejection.category == "validation_error"
    assert "read" in rejection.message
    assert check(settings, service="outlook", arguments={"action": "search", "user_id": USER}) is None


def test_missing_settings_is_refused_as_a_configuration_failure():
    from microsoft365.validation import check

    rejection = check(None, service="outlook", arguments={"action": "search", "user_id": USER})

    assert rejection is not None
    assert rejection.category == "configuration_error"


def test_unknown_service_is_refused():
    from microsoft365.validation import check

    rejection = check(_settings("application"), service="mail", arguments={"action": "search"})

    assert rejection is not None
    assert rejection.category == "validation_error"


# ======================================================================================
# 4. one validator, two call sites (the hook and the dispatch path agree)
# ======================================================================================


@pytest.mark.parametrize(
    "label,service,payload,expected_category,mode",
    [case for case in INVALID_CALLS if case[4] != "hybrid"],
    ids=[case[0] for case in INVALID_CALLS if case[4] != "hybrid"],
)
def test_hook_and_dispatch_reject_identically(monkeypatch, label, service, payload, expected_category, mode):
    from microsoft365.validation import check

    settings = _settings(mode)
    ctx = _registered(
        mode,
        capabilities=settings.capabilities,
        monkeypatch=monkeypatch,
    )
    tool_name = f"microsoft365_{service}"
    tool = ctx.tools.get(tool_name)
    expected = check(settings, service=service, arguments=payload)

    directive = ctx.hooks["pre_tool_call"](tool_name=tool_name, args=payload)

    assert directive is not None and directive["action"] == "block", label
    assert expected is not None and expected.category == expected_category, label
    assert directive["message"] == expected.message, label

    if tool is not None:
        dispatched = json.loads(tool["handler"](payload))
        assert dispatched["error"] == expected_category, label
        assert dispatched["message"] == expected.message, label


def test_dispatch_path_rejects_before_reaching_a_registered_handler(monkeypatch):
    from microsoft365 import registration

    invoked: list[str] = []

    def probe(args):
        invoked.append(str(args.get("action")))
        return json.dumps({"error": "unreachable"})

    for key in VALID_PAYLOADS:
        monkeypatch.setitem(registration.HANDLER_TABLE, key, probe)

    ctx = _registered(monkeypatch=monkeypatch)
    tool = ctx.tools["microsoft365_outlook"]["handler"]

    rejected = json.loads(tool({"action": "read", "user_id": USER, "message_id": 7}))
    assert rejected["error"] == "validation_error"
    # a valid payload for the same operation does reach dispatch (no handler effect here)
    reached = json.loads(tool({"action": "search", "user_id": USER}))
    assert reached == {"error": "unreachable"}

    assert invoked == ["search"]


# --------------------------------------------------------------------------------------
# Invariant 15: executability is re-evaluated on the final arguments, in the dispatch path
# --------------------------------------------------------------------------------------

#: The eleven reads this milestone exposes to the model.
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

#: The six writes implemented, contract-pinned and deliberately left non-executable
#: while the generic host approval fix (CORE-1/CORE-2) is unreleased (R5).
WITHHELD_WRITES = frozenset(
    {
        "outlook.create_draft", "outlook.send", "calendar.create_events", "calendar.update_events",
        "sharepoint.upload_files", "onedrive.upload_files",
        "todo.create_tasks", "todo.update_tasks", "planner.create_tasks", "planner.update_tasks",
    }
)


def test_a_read_mutated_into_a_write_is_refused_at_dispatch(monkeypatch):
    """CORE-1's defect class, closed inside the plugin.

    ``HANDLER_TABLE`` holds the write handlers -- WP6 implemented and pinned them -- so "a
    handler exists" must never mean "this call may run". A hook approves ``outlook.search``;
    a later modification rewrites ``action`` into a write. Every step a handler would take
    (secret lookup, credential, client, request) is armed to raise, so this test fails if the
    dispatch path lets the mutated payload through.
    """
    import agent.secret_scope
    import azure.identity
    import msgraph

    from microsoft365 import register, registration
    from microsoft365.validation import check

    def boom(name):
        """A runtime step that records being reached and then raises (the seam sanitizes it)."""

        def fail(*args, **kwargs):
            record.append(name)
            raise AssertionError(f"{name} must not be reached for a non-executable operation")

        return fail

    record: list = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", boom("get_secret"))
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", boom("ClientSecretCredential"))
    monkeypatch.setattr(msgraph, "GraphServiceClient", boom("GraphServiceClient"))

    ctx = RecordingContext(
        {
            "tenant_id": "tenant",
            "client_id": "client",
            "authentication_mode": "application",
            "capabilities": {service: True for service in SERVICES},
        }
    )
    register(ctx)
    hook = ctx.hooks["pre_tool_call"]

    for key in sorted(WITHHELD_WRITES):
        payload = dict(VALID_PAYLOADS[key])
        service, _ = key.split(".", 1)
        tool = ctx.tools[f"microsoft365_{service}"]["handler"]

        # Validation accepts this payload: only the executability gate can refuse it.
        assert check(_settings(), service=service, arguments=payload) is None, key

        # Both gates say the same thing about the same final arguments.
        directive = hook(tool_name=f"microsoft365_{service}", args=payload)
        assert directive == {
            "action": "block",
            "message": f"Microsoft 365 operation is not executable: {key}",
        }, key

        refused = json.loads(tool(payload))
        assert refused["error"] == "operation_not_implemented", key
        assert refused["message"] == directive["message"], key
        # the dispatch function the tool handler calls re-evaluates independently
        assert json.loads(
            registration.service_tool_handler(service, payload, settings=_settings())
        ) == refused, key

    # The handler is never invoked either: a recording probe is not reached.
    reached: list = []

    def write_probe(args):
        reached.append(str(args.get("action")))
        return json.dumps({"error": "write_handler_reached"})

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.send", write_probe)
    outlook_tool = ctx.tools["microsoft365_outlook"]["handler"]
    assert (
        json.loads(outlook_tool(dict(VALID_PAYLOADS["outlook.send"])))["error"]
        == "operation_not_implemented"
    )
    assert reached == []

    # Nothing a handler would do was reached for any of the four writes: no secret lookup, no
    # credential, no client, and therefore no request.
    assert record == []

    # Positive control: the same payload shape, but a read, dispatches as far as the credential
    # -- which raises, so the honest outcome is the taxonomy's ``authentication_required`` and
    # the armed runtime is proven live rather than inert.
    read = json.loads(outlook_tool(dict(VALID_PAYLOADS["outlook.search"])))
    assert record == ["get_secret"]
    assert read["error"] == "authentication_required"


def test_the_withheld_writes_keep_their_handler_and_the_reads_are_the_only_executable_set():
    """The gate, not a missing handler, is what withholds the writes (invariant 15's premise)."""
    from microsoft365.contract import OPERATION_REGISTRY
    from microsoft365.registration import HANDLER_TABLE, active_actions

    configuration = _settings()
    assert {
        key for key, definition in OPERATION_REGISTRY.items() if definition.executable
    } == set(EXECUTABLE_READS)
    assert set(HANDLER_TABLE) == set(EXECUTABLE_READS) | set(WITHHELD_WRITES) | {"teams.search_messages", "teams.send_messages"}

    for key in sorted(WITHHELD_WRITES):
        service, operation = key.split(".", 1)
        assert key in HANDLER_TABLE, key
        assert operation not in active_actions(configuration, service), key


def test_hook_lets_an_executable_read_through_and_blocks_a_withheld_write():
    """The hook's executability gate, on the milestone's real state.

    ``outlook.search`` is executable, so a valid call needs no block and no approval. The
    withheld write ``outlook.send`` is valid yet not executable, so the hook blocks it with the
    same message the dispatch path returns -- never an approval directive.
    """
    ctx = _registered(
        capabilities={"outlook": {"search": True, "send": True}, "teams": {"send_messages": True}}
    )

    assert (
        ctx.hooks["pre_tool_call"](
            tool_name="microsoft365_outlook", args={"action": "search", "user_id": USER}
        )
        is None
    )
    assert ctx.hooks["pre_tool_call"](
        tool_name="microsoft365_outlook", args={"action": "send", "user_id": USER, "message_id": "AAMkAG"}
    ) == {
        "action": "block",
        "message": "Microsoft 365 operation is not executable: outlook.send",
    }
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args=None) == {
        "action": "block",
        "message": "Microsoft 365 arguments must be an object",
    }
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_preflight", args={}) is None
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_nosuch", args={}) is None
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_teams", args={"action": "send_messages", "team_id": TEAM, "channel_id": CHANNEL, "body": "hi"})["action"] == "block"


def test_invalid_write_is_blocked_before_the_approval_directive(monkeypatch):
    ctx = _registered(
        capabilities={"outlook": {"search": True, "read": True, "send": True}},
        monkeypatch=monkeypatch,
    )
    hook = ctx.hooks["pre_tool_call"]

    approved = hook(
        tool_name="microsoft365_outlook",
        args={"action": "send", "user_id": USER, "message_id": "AAMkAG"},
    )
    assert approved == {
        "action": "approve",
        "message": "Microsoft 365 send: external side effect",
        "rule_key": "microsoft365.outlook.send",
    }

    blocked = hook(
        tool_name="microsoft365_outlook",
        args={"action": "send", "user_id": USER, "message_id": "AAMkAG", "save_to_sent_items": "false"},
    )
    assert blocked["action"] == "block"
    assert blocked["message"] != approved["message"]
    assert "save_to_sent_items" in blocked["message"]


# ======================================================================================
# 5. the central proof: nothing invalid reaches a secret, a credential or a client
# ======================================================================================


def _boom(record, name):
    def fail(*args, **kwargs):
        record.append(name)
        raise AssertionError(f"{name} must not be reached")

    return fail


def _raising_runtime(monkeypatch):
    import agent.secret_scope
    import azure.identity
    import msgraph

    record: list[str] = []
    monkeypatch.setattr(agent.secret_scope, "get_secret", _boom(record, "get_secret"))
    monkeypatch.setattr(azure.identity, "ClientSecretCredential", _boom(record, "ClientSecretCredential"))
    monkeypatch.setattr(msgraph, "GraphServiceClient", _boom(record, "GraphServiceClient"))
    return record


def test_invalid_arguments_never_touch_secret_credential_or_client(monkeypatch):
    record = _raising_runtime(monkeypatch)
    contexts = {
        mode: _registered(mode, capabilities=_settings(mode).capabilities, monkeypatch=monkeypatch)
        for mode in ("application", "delegated", "disabled_read")
    }

    for label, service, payload, expected_category, mode in INVALID_CALLS:
        if mode == "hybrid":
            # A malformed configuration registers no service tool at all; the validator's own
            # refusal is covered by test_unknown_authentication_mode_is_refused.
            continue
        ctx = contexts[mode]
        tool_name = f"microsoft365_{service}"
        directive = ctx.hooks["pre_tool_call"](tool_name=tool_name, args=payload)
        assert directive is not None and directive["action"] == "block", label

        tool = ctx.tools.get(tool_name)
        if tool is not None:
            dispatched = json.loads(tool["handler"](payload))
            assert dispatched["error"] == expected_category, label

    assert record == []


def test_unknown_authentication_mode_is_refused_without_registering_anything(monkeypatch):
    from microsoft365.contract import operation_status
    from microsoft365.validation import check

    record = _raising_runtime(monkeypatch)

    rejection = check(_settings("hybrid"), service="outlook", arguments={"action": "search", "user_id": USER})
    assert rejection is not None
    assert rejection.category == "unsupported_auth_mode"
    assert operation_status("hybrid", "outlook", "search").auth_status == "not_implemented"

    # The unknown mode registers nothing at all -- not even the service tools a valid
    # configuration now registers, because the executable reads do not exist in that mode.
    ctx = _registered("hybrid")
    assert set(ctx.tools) == {"microsoft365_preflight"}
    # Control: a valid configuration *does* register the service tool, so the difference above
    # is the authentication mode and not an empty capability set.
    valid = _registered(capabilities={"outlook": {"search": True}})
    assert "microsoft365_outlook" in valid.tools
    assert set(valid.tools) == {"microsoft365_preflight", "microsoft365_outlook"}
    assert record == []


def test_a_valid_payload_does_reach_the_probe_so_the_guard_is_what_blocks(monkeypatch):
    """Positive control: the probes do trip for a valid payload, so the guard is the blocker."""
    from microsoft365 import client, registration
    from microsoft365.contract import Settings

    record = _raising_runtime(monkeypatch)

    def probe(args):
        del args
        client.create_graph_client(Settings(tenant_id="tenant", client_id="client"))
        return json.dumps({"error": "unreachable"})

    monkeypatch.setitem(registration.HANDLER_TABLE, "outlook.search", probe)
    ctx = _registered(monkeypatch=monkeypatch)

    with pytest.raises(AssertionError, match="get_secret must not be reached"):
        json.loads(ctx.tools["microsoft365_outlook"]["handler"](dict(VALID_PAYLOADS["outlook.search"])))

    assert record == ["get_secret"]


def test_validation_module_keeps_no_secret_credential_or_client_access():
    from microsoft365 import validation

    source = Path(validation.__file__).read_text(encoding="utf-8")
    for forbidden in (
        "get_secret",
        "secret_scope",
        "ClientSecretCredential",
        "GraphServiceClient",
        "azure.identity",
        "msgraph",
    ):
        assert forbidden not in source

# -- outlook.search structured filters: cross-argument rules (foundation slice) ---------


def test_outlook_search_structured_filters_are_validated_before_any_client_work():
    """The same ``check`` the hook and the dispatch path both call refuses an invalid
    sender, a raw/structured filter conflict and a reversed date range, while a
    well-formed structured search (including the boolean ``false``) passes."""
    base = {"action": "search", "user_id": USER}
    valid = {
        **base,
        "query": "facture",
        "has_attachments": False,
        "sender": "o'hara@example.test",
        "received_after": "2026-10-01T00:00:00Z",
        "received_before": "2026-11-01T00:00:00Z",
        "top": 20,
    }
    assert _check("outlook", valid) is None

    for payload in (
        {**base, "sender": "not-an-address"},
        {**base, "filter": "isRead eq false", "has_attachments": True},
        {**base, "filter": "isRead eq false", "sender": "a@example.test"},
        {**base, "filter": "isRead eq false", "received_after": "2026-10-01T00:00:00Z"},
        {
            **base,
            "received_after": "2026-11-01T00:00:00Z",
            "received_before": "2026-10-01T00:00:00Z",
        },
        # reversed although the raw wall-clock ordering looks fine: the offsets differ
        # (2026-10-02T01:00:00+02:00 is 2026-10-01T23:00:00Z)
        {
            **base,
            "received_after": "2026-10-02T00:00:00Z",
            "received_before": "2026-10-02T01:00:00+02:00",
        },
        # equal bounds: received_after is inclusive and received_before exclusive,
        # so the interval is empty and must be refused, not silently searched
        {
            **base,
            "received_after": "2026-10-01T00:00:00Z",
            "received_before": "2026-10-01T00:00:00Z",
        },
        # equal instants with different zone offsets
        {
            **base,
            "received_after": "2026-10-02T00:00:00Z",
            "received_before": "2026-10-01T22:00:00-02:00",
        },
        # the raw filter and the subject query are two different filters:
        # combined, the query would be silently dropped
        {**base, "filter": "isRead eq false", "query": "facture"},
    ):
        rejection = _check("outlook", payload)
        assert rejection is not None, payload
        assert rejection.category == "validation_error", payload

    # The raw filter alone, and the structured filters alone, remain fully valid.
    assert _check("outlook", {**base, "filter": "isRead eq false"}) is None
    assert _check("outlook", {**base, "has_attachments": False}) is None
    assert _check("outlook", {**base, "query": "facture"}) is None
