"""The group-mailbox route of the four Outlook reads.

A Microsoft 365 group mailbox is a conversation tree, not a message collection, so the
Outlook reads reach it through ``/groups/{id}/threads`` and the posts inside a thread rather
than ``/users/{id}/messages``. The tool surface does not change with the kind of mailbox:
``user_id`` names the mailbox and ``message_id`` names the unit a listing returns -- a
message for a user mailbox, a **conversation thread** for a group mailbox -- so the same
four operations, the same arguments and the same mailbox allowlist apply to both.

What this route can and cannot do, stated rather than papered over
------------------------------------------------------------------
* ``GET /groups/{id}/threads`` supports ``$top`` and nothing else that narrows a search: the
  service answers ``ConversationFilterOther`` to any ``$filter`` other than ``isLocked``.
  The structured search arguments are therefore applied **after** the bounded page is
  fetched, and every result says so (``filter_mode: "client"``) and reports how many threads
  were examined, so a caller can never mistake a locally filtered page for a mailbox-wide
  search. A raw OData ``filter`` is refused outright instead of being silently ignored, and
  so is ``sender``: a thread carries display names (``uniqueSenders``), not addresses, so an
  address filter could not be honoured truthfully.
* Attachments are read through ``$expand=attachments`` on the thread's posts. The per-post
  ``/attachments`` collection is refused by Exchange for an application identity, while the
  expanded navigation property carries both the metadata and ``contentBytes`` -- so a listing
  on this route still returns metadata only, and the download reuses the same verified
  size policy as the user route.
"""
from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Mapping

from .paging import paginate
from .results import normalize_result

#: The thread fields a group-mailbox listing reports, named after the user-mailbox fields
#: they correspond to so one result shape serves both routes.
THREAD_FIELDS = (
    "id",
    "subject",
    "preview",
    "from",
    "receivedDateTime",
    "hasAttachments",
)

#: The post fields a group-mailbox read reports.
POST_FIELDS = ("id", "from", "receivedDateTime", "hasAttachments", "body")

#: Search arguments the thread collection cannot express, with the reason each is refused
#: rather than quietly dropped.
UNSUPPORTED_SEARCH_ARGUMENTS: Mapping[str, str] = {
    "filter": (
        "a group mailbox thread collection accepts no OData filter (the service supports "
        "only isLocked), so a raw filter cannot be honoured"
    ),
    "sender": (
        "a group mailbox thread carries sender display names, not addresses, so an address "
        "filter cannot be honoured"
    ),
    "select": (
        "a group mailbox thread listing reports a fixed set of conversation fields, so a "
        "caller $select cannot be honoured"
    ),
}


def _text(value: Any) -> str | None:
    return value if isinstance(value, str) and value else None


def _when(value: Any) -> datetime | None:
    if not isinstance(value, datetime):
        return None
    return value if value.tzinfo is not None else value.replace(tzinfo=timezone.utc)


def _senders(thread) -> dict | None:
    """The thread's senders as the same ``{address, name}`` shape the user route reports.

    A thread records display names only, so ``name`` is the single honest field: no address
    is invented for it.
    """
    names = [
        name
        for name in (getattr(thread, "unique_senders", None) or [])
        if isinstance(name, str) and name
    ]
    return {"name": names[0]} if names else None


def project_thread(thread) -> dict:
    """One conversation thread as the metadata a mailbox listing returns."""
    projected: dict = {}
    for name, value in (
        ("id", _text(getattr(thread, "id", None))),
        ("subject", _text(getattr(thread, "topic", None))),
        ("preview", _text(getattr(thread, "preview", None))),
        ("from", _senders(thread)),
        ("hasAttachments", getattr(thread, "has_attachments", None)),
    ):
        if value is not None:
            projected[name] = value
    delivered = _when(getattr(thread, "last_delivered_date_time", None))
    if delivered is not None:
        projected["receivedDateTime"] = delivered.isoformat()
    return projected


def project_post(post, *, include_body: bool) -> dict:
    """One post as the message metadata a read returns; the body only when asked for."""
    projected: dict = {}
    identifier = _text(getattr(post, "id", None))
    if identifier is not None:
        projected["id"] = identifier
    sender = getattr(post, "from_", None) or getattr(post, "sender", None)
    email = getattr(sender, "email_address", None) if sender is not None else None
    address = _text(getattr(email, "address", None)) if email is not None else None
    if address is not None:
        structured = {"address": address}
        name = _text(getattr(email, "name", None))
        if name is not None:
            structured["name"] = name
        projected["from"] = structured
    received = _when(getattr(post, "received_date_time", None))
    if received is not None:
        projected["receivedDateTime"] = received.isoformat()
    attachments = getattr(post, "has_attachments", None)
    if attachments is not None:
        projected["hasAttachments"] = attachments
    if include_body:
        body = getattr(post, "body", None)
        content = _text(getattr(body, "content", None)) if body is not None else None
        if content is not None:
            projected["body"] = content
    return projected


def thread_matches(thread, filters: Mapping[str, Any]) -> bool:
    """Whether one thread satisfies the filters the service could not apply itself."""
    query = filters.get("query")
    if isinstance(query, str) and query.strip():
        topic = getattr(thread, "topic", None)
        if not isinstance(topic, str) or query.strip().lower() not in topic.lower():
            return False
    wanted = filters.get("has_attachments")
    if isinstance(wanted, bool) and getattr(thread, "has_attachments", None) is not wanted:
        return False
    delivered = _when(getattr(thread, "last_delivered_date_time", None))
    for name, inclusive in (("received_after", True), ("received_before", False)):
        bound = filters.get(name)
        if bound is None:
            continue
        if delivered is None:
            return False
        edge = datetime.fromisoformat(bound).astimezone(timezone.utc)
        if inclusive and delivered.astimezone(timezone.utc) < edge:
            return False
        if not inclusive and delivered.astimezone(timezone.utc) >= edge:
            return False
    return True


def thread_builder(client, mailbox_id: str):
    return client.groups.by_group_id(mailbox_id).threads


def post_builder(client, mailbox_id: str, thread_id: str):
    return (
        client.groups.by_group_id(mailbox_id)
        .threads.by_conversation_thread_id(thread_id)
        .posts
    )


def attachment_entries(posts) -> list[dict]:
    """Bounded attachment metadata of every post in a thread, never attachment content.

    Each entry names the post it belongs to, because an attachment of a group mailbox lives
    in a post rather than in the conversation itself.
    """
    entries: list[dict] = []
    for post in posts:
        post_id = _text(getattr(post, "id", None))
        for attachment in getattr(post, "attachments", None) or []:
            entries.append(
                {
                    "id": _text(getattr(attachment, "id", None)),
                    "name": _text(getattr(attachment, "name", None)),
                    "size": getattr(attachment, "size", None),
                    "content_type": _text(getattr(attachment, "content_type", None)),
                    "is_inline": getattr(attachment, "is_inline", None),
                    "attachment_type": _text(getattr(attachment, "odata_type", None)),
                    "post_id": post_id,
                }
            )
    return entries


def find_attachment(posts, attachment_id: str):
    """The attachment with this id inside a thread's posts, with the post that carries it."""
    for post in posts:
        for attachment in getattr(post, "attachments", None) or []:
            if getattr(attachment, "id", None) == attachment_id:
                return attachment, post
    return None, None


def paginate_threads(client, *, mailbox_id: str, adapter, limit: int, configuration):
    return paginate(
        thread_builder(client, mailbox_id),
        adapter=adapter,
        limit=limit,
        configuration=configuration,
    )


def search_payload(paged, filters: Mapping[str, Any]) -> dict:
    """The listing a group-mailbox search returns, with its filtering stated truthfully."""
    applied = tuple(
        name
        for name in ("query", "has_attachments", "received_after", "received_before")
        if filters.get(name) is not None
    )
    matched = [thread for thread in paged.items if thread_matches(thread, filters)]
    result = dict(paged.payload)
    result["value"] = [normalize_result(project_thread(thread)) for thread in matched]
    result["container"] = "group"
    result["unit"] = "conversation_thread"
    result["filter_mode"] = "client" if applied else "none"
    result["client_filters"] = list(applied)
    result["threads_examined"] = len(paged.items)
    return result
