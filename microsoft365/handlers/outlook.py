"""Outlook handlers: search, read, create_draft and send (WP6).

The attachment operations live in ``handlers/outlook_attachments.py`` (split out so the
attachment lane owns one file); module discovery registers them without any edit here, and
the handler registry refuses a duplicate key, so an operation cannot be registered twice.

Each handler builds the request of its operation from the arguments the administrative
argument contract declares, on the generated SDK builders, and is pinned byte-for-byte
against the verified request-information contract of the same operation
(``microsoft365/sdk_contract.py``; see ``tests/test_handlers_outlook_calendar.py``).

Action-specific rules this module owns (what the shared argument contract cannot express)
----------------------------------------------------------------------------------------
* ``outlook.search``: ``query`` and ``filter`` are two different filters, so supplying both is
  refused instead of silently dropping one; ``select`` is typed, so it reaches the request as
  ``$select`` and never as a hand-built query string. With no caller ``select`` the search
  sends :data:`DEFAULT_MESSAGE_SELECT` (candidate metadata, never the body) so a bare listing
  does not download every candidate's full content.
* ``outlook.create_draft``: a draft is **composed**, so it always needs a body and recipients;
  ``message_id`` is the alternate path of ``outlook.send`` and cannot stand in for either.
* ``outlook.send``: exactly one form -- an existing draft (``message_id``) or a composed
  message (``subject`` plus ``to_recipients``) -- because mixing them would send one silently.
  ``save_to_sent_items`` is sent only when the caller supplied it, so the service default is
  never asserted on the caller's behalf.
"""
from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Mapping

from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message import Message
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)

from ..execution import execute_request
from ..paging import paginate
from ..results import normalize_result
from . import (
    HandlerContext,
    collection_payload,
    handler,
    identifier,
    item_budget,
    open_graph_client,
    recipients,
    refuse,
    select,
    success_payload,
    text,
    typed_configuration,
)

SEARCH = "outlook.search"
READ = "outlook.read"
CREATE_DRAFT = "outlook.create_draft"
SEND = "outlook.send"

#: The compact ``$select`` a search sends when the caller named no fields of its own.
#:
#: Invoice-candidate discovery works from *metadata* -- the message id, subject, sender
#: (``from``), received time, attachment flag and the message link are what make a candidate
#: worth opening. A bare listing without a ``$select`` returns the **complete** message on
#: every page, so the body and ``bodyPreview`` of every candidate are downloaded and then
#: discarded; that is what this default prevents. ``hasAttachments`` is a *summary* flag, so a
#: message whose only attachments are inline (``Content-Id``) still reads ``false`` here -- the
#: flag is reported truthfully, not repaired -- and subject matching is containment, not
#: invoice-content detection. A caller that wants other fields (or the body, via ``read``)
#: supplies its own ``select`` and it is honored verbatim.
DEFAULT_MESSAGE_SELECT = [
    "id",
    "subject",
    "from",
    "receivedDateTime",
    "hasAttachments",
    "webLink",
]


def search_filter(arguments: Mapping, *, operation: str) -> str | None:
    """The OData ``$filter`` of a message search, or nothing when the caller sent no search.

    ``query`` is the containment filter the verified collection contract defines
    (``contains(subject,'…')``, with the quote-doubling OData requires); ``filter`` is the
    caller's own expression and is passed through untouched.
    """
    query = arguments.get("query")
    supplied = arguments.get("filter")
    if query is not None and supplied is not None:
        refuse(
            "validation_error",
            f"{operation}: query and filter are two different filters; supplying both would "
            "silently drop one",
        )
    if supplied is not None:
        if not isinstance(supplied, str) or not supplied.strip():
            refuse(
                "validation_error",
                f"{operation}: filter must be a non-blank OData filter expression",
            )
        return supplied
    conditions = []
    if query is not None:
        if not isinstance(query, str) or not query.strip():
            refuse("validation_error", f"{operation}: query must be a non-blank search string")
        conditions.append("contains(subject,'" + query.replace("'", "''") + "')")
    if "has_attachments" in arguments:
        conditions.append("hasAttachments eq " + str(arguments["has_attachments"]).lower())
    if arguments.get("sender") is not None:
        conditions.append("from/emailAddress/address eq '" + arguments["sender"].replace("'", "''") + "'")
    for name, comparison in (("received_after", "ge"), ("received_before", "lt")):
        if name in arguments:
            stamp = datetime.fromisoformat(arguments[name]).astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
            conditions.append(f"receivedDateTime {comparison} {stamp}")
    return " and ".join(conditions) or None


# ---------------------------------------------------------------- compact metadata projection
#
# A search selects candidate *metadata* off the wire, but the shared result normalization
# (``microsoft365.results``) keeps only the model fields it knows -- and ``from``,
# ``receivedDateTime`` and ``webLink`` are not among them -- so left alone the tool result
# silently loses the three fields a candidate lives on. The projection below reads each
# actual generated ``Message`` and yields plain, normalizable data for exactly the fields the
# caller's ``$select`` (or the compact default) named: the sender as a usable email/name
# structure, the date as JSON text, the link as-is, plus the id/subject/attachment flag.


def _wire_to_attr(odata_name: str) -> str:
    """The generated model attribute that carries one ``$select`` wire name."""
    return re.sub(r"(?<!^)(?=[A-Z])", "_", odata_name).lower()


def _email_name(recipient) -> dict | None:
    """A ``Recipient`` as the usable ``{address, name}`` structure the metadata promises.

    The ``from`` of a message is a known model type (``Recipient`` carrying an
    ``EmailAddress``); it is surfaced as a flat email/name structure, never the raw model, so
    a missing sender or a nameless address simply omits the field rather than a marker.
    """
    if recipient is None:
        return None
    email = getattr(recipient, "email_address", None)
    if email is None:
        return None
    address = getattr(email, "address", None)
    if not isinstance(address, str) or not address:
        return None
    structured = {"address": address}
    name = getattr(email, "name", None)
    if isinstance(name, str) and name:
        structured["name"] = name
    return structured


def _json_text(value) -> str | None:
    """A model date as bounded JSON text (ISO 8601), or nothing when absent."""
    if isinstance(value, datetime):
        return value.isoformat()
    return None


def _message_field(message: Message, odata_name: str):
    """One ``$select`` field read from an actual generated ``Message``.

    ``from`` and ``sender`` are the known Recipient model types and are surfaced as a usable
    email/name structure; ``receivedDateTime`` is surfaced as JSON text; every other field is
    the model's own attribute, handed to the shared normalization as-is (never a hand-built
    query string).
    """
    if odata_name in ("from", "sender"):
        return _email_name(message.from_ if odata_name == "from" else message.sender)
    if odata_name == "receivedDateTime":
        return _json_text(message.received_date_time)
    return getattr(message, _wire_to_attr(odata_name), None)


def _project_message(message: Message, fields) -> dict:
    """The compact candidate metadata one actual ``Message`` carries, for ``fields`` only.

    Fields the model does not carry are omitted (honest about absence), so an explicit
    ``select`` never injects a field the caller did not ask for.
    """
    projected = {}
    for odata_name in fields:
        value = _message_field(message, odata_name)
        if value is not None:
            projected[odata_name] = value
    return projected


def message_model(arguments: Mapping, *, operation: str, require_body: bool) -> Message:
    """The real generated ``Message`` a composed call sends."""
    subject = text(arguments, "subject", operation=operation)
    if subject is None:
        refuse("validation_error", f"{operation}: subject is required")
    content = text(arguments, "body", operation=operation)
    if require_body and content is None:
        refuse(
            "validation_error",
            f"{operation}: a draft needs a body; use message_id to send content that already "
            "exists",
        )
    return Message(
        subject=subject,
        body=ItemBody(content_type=BodyType.Text, content=content) if content is not None else None,
        to_recipients=recipients(arguments, "to_recipients", operation=operation, required=True),
        cc_recipients=recipients(arguments, "cc_recipients", operation=operation),
        bcc_recipients=recipients(arguments, "bcc_recipients", operation=operation),
    )


# ---------------------------------------------------------------- reads


@handler("outlook", "search")
def search(arguments: Mapping, context: HandlerContext) -> dict:
    """List one mailbox's messages, bounded by the caller's item budget."""
    user_id = identifier(arguments, "user_id", operation=SEARCH)
    filters = search_filter(arguments, operation=SEARCH)
    # A caller ``select`` is honored verbatim; with none supplied the search selects the
    # compact candidate metadata instead of pulling every candidate's complete body.
    fields = select(arguments, operation=SEARCH) or DEFAULT_MESSAGE_SELECT
    limit = item_budget(arguments, operation=SEARCH)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages
    configuration = typed_configuration(builder, top=limit, filter=filters, select=fields)
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=configuration)
    # Project the compact candidate metadata from the *actual* generated items, honoring the
    # exact select that reached the wire (a caller select never injects an unrequested field),
    # and run it through the shared normalization so every value is redacted and bounded. The
    # continuation and the observed count are the paginator's own, so the cap stays truthful.
    projected = [normalize_result(_project_message(item, fields)) for item in paged.items]
    result = dict(paged.payload)
    result["value"] = projected
    return {**collection_payload(SEARCH, paged), "result": result, "user_id": user_id}


@handler("outlook", "read")
def read(arguments: Mapping, context: HandlerContext) -> dict:
    """Read one message, with the caller's ``$select`` and nothing else."""
    user_id = identifier(arguments, "user_id", operation=READ)
    message_id = identifier(arguments, "message_id", operation=READ)
    fields = select(arguments, operation=READ)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages.by_message_id(message_id)
    configuration = typed_configuration(builder, select=fields)
    response = execute_request(builder, method="GET", configuration=configuration, adapter=adapter)
    return success_payload(READ, response)


# ---------------------------------------------------------------- writes
#
# The four writes below are implemented, exercised offline through the seam (method, URL,
# body and If-Match are pinned against the verified contract) and deliberately **not
# executable**: while the generic Hermes approval-order defect (CORE-1/CORE-2) can let a hook
# change arguments after approval was evaluated against different ones, a write must not be
# reachable from the model surface at all (R5). They are not dead code -- the registry flip of
# WP9 is the only thing between them and execution.


@handler("outlook", "create_draft")
def create_draft(arguments: Mapping, context: HandlerContext) -> dict:
    """Create a draft message, which always carries a body and recipients."""
    user_id = identifier(arguments, "user_id", operation=CREATE_DRAFT)
    if arguments.get("message_id") is not None:
        refuse(
            "validation_error",
            f"{CREATE_DRAFT}: a draft is composed from its own body and recipients, so "
            "message_id cannot stand in for either of them",
        )
    model = message_model(arguments, operation=CREATE_DRAFT, require_body=True)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages
    response = execute_request(builder, method="POST", body=model, adapter=adapter)
    return success_payload(CREATE_DRAFT, response)


@handler("outlook", "send")
def send(arguments: Mapping, context: HandlerContext) -> dict:
    """Send a composed message, or send a draft that already exists."""
    user_id = identifier(arguments, "user_id", operation=SEND)
    message_id = arguments.get("message_id")
    composed = arguments.get("subject") is not None or arguments.get("to_recipients") is not None
    if message_id is not None and composed:
        refuse(
            "validation_error",
            f"{SEND}: send either an existing draft (message_id) or a composed message "
            "(subject and to_recipients); supplying both would send one of them silently",
        )
    if message_id is None and arguments.get("subject") is None:
        refuse(
            "validation_error",
            f"{SEND}: a send needs either message_id or a composed message with subject and "
            "to_recipients",
        )

    if message_id is not None:
        draft_id = identifier(arguments, "message_id", operation=SEND)
        client, adapter = open_graph_client(context)
        builder = client.users.by_user_id(user_id).messages.by_message_id(draft_id).send
        execute_request(builder, method="POST", adapter=adapter)
        return success_payload(SEND, {"accepted": True, "form": "existing"})

    save_to_sent_items = arguments.get("save_to_sent_items")
    if save_to_sent_items is not None and type(save_to_sent_items) is not bool:
        refuse(
            "validation_error",
            f"{SEND}: save_to_sent_items must be a real boolean",
        )
    model = message_model(arguments, operation=SEND, require_body=False)

    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).send_mail
    request_body = SendMailPostRequestBody(message=model, save_to_sent_items=save_to_sent_items)
    execute_request(builder, method="POST", body=request_body, adapter=adapter)
    return success_payload(SEND, {"accepted": True, "form": "composed"})
