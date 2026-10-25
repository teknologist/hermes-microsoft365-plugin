"""Which kind of Microsoft 365 mailbox an allowlisted identifier names.

Why a mailbox has a kind at all
-------------------------------
Three things are called "a mailbox" in Microsoft 365 and only two of them answer on the same
Graph route:

* a **user mailbox** -- ``/users/{id}/messages``;
* a **shared mailbox** -- an Exchange recipient that is still a user object in the directory,
  so it answers on exactly the same ``/users/{id}`` route under the same ``Mail.Read``
  application role, and needs no special handling here;
* a **Microsoft 365 group mailbox** -- a *different resource*. ``GET /users/{group}`` is a
  directory read an application identity is refused, the group's SMTP address is not a user
  principal, and its mail is a conversation tree rather than a message collection. A group
  mailbox addressed through the user route can therefore only ever answer ``not_found``,
  which is exactly the symptom this module removes.

How the kind is decided
-----------------------
By asking the directory, once, with the cheapest read that distinguishes them: a
``$select``-ed ``GET /groups/{id}``. A mail-enabled unified group is a group mailbox;
anything else -- including a group that exists but has no mailbox, and a ``not_found`` or
``permission_denied`` answer -- is the user route, because that is the route a user or shared
mailbox answers on and the Outlook read that follows reports its own failure truthfully.

The verdict is cached per ``(tenant, mailbox)`` for the life of the process. The cache holds
two identifiers and a string: no credential, client, adapter or message content ever enters
it, and a caller may bypass it entirely by naming the kind in the arguments.
"""

from __future__ import annotations

from typing import Any, Mapping

from .contract import Settings
from .errors import GraphError
from .execution import execute_request

#: The mailbox kinds this plugin routes between. ``user`` covers personal *and* shared
#: mailboxes: both are user objects on the same Graph route.
USER_MAILBOX = "user"
GROUP_MAILBOX = "group"
MAILBOX_KINDS = (USER_MAILBOX, GROUP_MAILBOX)

#: Directory answers that mean "this identifier is not a group I can see", and therefore
#: "use the user route". Any other failure is re-raised: a throttle or a transport failure
#: must not be silently read as a verdict about the mailbox.
_NOT_A_VISIBLE_GROUP = frozenset({"not_found", "permission_denied", "consent_required"})

#: ``groupTypes`` value that marks a Microsoft 365 (unified) group -- the only group kind
#: that has a mailbox. A security or distribution group has none.
_UNIFIED = "unified"

_GROUP_PROBE_SELECT = ["id", "mailEnabled", "groupTypes"]

#: Resolved kinds, keyed ``(tenant_id, mailbox_id)``. Values are the two strings above.
_KIND_CACHE: dict[tuple[str, str], str] = {}


def declared_kind(arguments: Mapping[str, Any] | None) -> str | None:
    """The mailbox kind the caller named, or ``None`` when it left the choice to resolution."""
    if not isinstance(arguments, Mapping):
        return None
    value = arguments.get("mailbox_type")
    if value is None:
        return None
    if value not in MAILBOX_KINDS:
        raise GraphError(
            "validation_error",
            f"mailbox_type must be one of {', '.join(MAILBOX_KINDS)}",
        )
    return str(value)


def _is_group_mailbox(group: Any) -> bool:
    """Whether a directory answer describes a mailbox-bearing Microsoft 365 group."""
    if group is None:
        return False
    if getattr(group, "mail_enabled", None) is not True:
        return False
    types = getattr(group, "group_types", None) or []
    return any(str(entry).strip().lower() == _UNIFIED for entry in types)


def known_kind(
    mailbox_id: str, *, settings: Settings, arguments: Mapping[str, Any] | None = None
) -> str | None:
    """The kind already known for this mailbox without sending a request.

    Either the caller named it, or an earlier call on this mailbox resolved it. ``None``
    means "not known yet", which is the only state that costs a directory read -- and only
    after the user route has already answered that the item is not there.
    """
    declared = declared_kind(arguments)
    if declared is not None:
        return declared
    return _KIND_CACHE.get(_cache_key(settings, mailbox_id))


def _cache_key(settings: Settings, mailbox_id: str) -> tuple[str, str]:
    return (str(settings.tenant_id), str(mailbox_id))


def resolve_mailbox_kind(
    mailbox_id: str,
    *,
    client,
    adapter,
    settings: Settings,
    arguments: Mapping[str, Any] | None = None,
) -> str:
    """The kind of mailbox ``mailbox_id`` names: ``"user"`` or ``"group"``.

    A caller-supplied ``mailbox_type`` and a cached verdict both answer without a request.
    Otherwise the directory is asked once, with the cheapest read that distinguishes the
    two, and the answer is cached for the process; a directory failure that is not a
    visible-group verdict is re-raised rather than guessed around.
    """
    known = known_kind(mailbox_id, settings=settings, arguments=arguments)
    if known is not None:
        return known
    builder = client.groups.by_group_id(mailbox_id)
    try:
        group = execute_request(
            builder,
            method="GET",
            adapter=adapter,
            configuration=_probe_configuration(builder),
        )
    except GraphError as failure:
        if failure.category not in _NOT_A_VISIBLE_GROUP:
            raise
        kind = USER_MAILBOX
    else:
        kind = GROUP_MAILBOX if _is_group_mailbox(group) else USER_MAILBOX
    _KIND_CACHE[_cache_key(settings, mailbox_id)] = kind
    return kind


def route(
    mailbox_id: str,
    *,
    client,
    adapter,
    settings: Settings,
    arguments: Mapping[str, Any] | None,
    user_route,
    group_route,
):
    """Run an Outlook read on the route the mailbox actually answers on.

    The user route is the default and stays free: a personal or shared mailbox pays nothing
    for group support -- no probe, no extra request, not even on the first call. A group
    mailbox is discovered from the only symptom it can produce, because ``/users/{group}``
    cannot answer anything else: the user route reports ``not_found``. Only then is the
    directory asked whether the identifier is a mail-enabled group, and only a ``group``
    verdict re-runs the read on the group route. The verdict is cached either way, so the
    discovery happens at most once per mailbox per process and every later call goes straight
    to the right route.

    A ``not_found`` that really is a missing item keeps its own answer: the directory says
    ``user``, nothing is retried, and the original failure is raised unchanged.
    """
    known = known_kind(mailbox_id, settings=settings, arguments=arguments)
    if known == GROUP_MAILBOX:
        return group_route()
    try:
        return user_route()
    except GraphError as failure:
        if known == USER_MAILBOX or failure.category != "not_found":
            raise
        try:
            kind = resolve_mailbox_kind(
                mailbox_id, client=client, adapter=adapter, settings=settings
            )
        except Exception:
            # The probe is an *explanation* of the failure already in hand, never a new
            # failure of its own: if the directory cannot answer, the read keeps the error
            # it actually got instead of reporting one the caller did not cause.
            raise failure from None
        if kind != GROUP_MAILBOX:
            raise
    return group_route()


def _probe_configuration(builder):
    """The typed ``$select`` of the directory probe, from the generated builder's own query."""
    from .handlers import typed_configuration

    return typed_configuration(builder, select=list(_GROUP_PROBE_SELECT))


def forget_resolved_kinds() -> None:
    """Drop every cached verdict. Used by tests; never needed on a request path."""
    _KIND_CACHE.clear()
