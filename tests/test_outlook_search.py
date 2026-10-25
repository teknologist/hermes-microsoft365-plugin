"""Invoice candidate search contracts: structured filters and bounded paging."""
from __future__ import annotations

import json
from datetime import datetime, timezone

import pytest
from msgraph import GraphServiceClient
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.message import Message
from msgraph.generated.models.message_collection_response import MessageCollectionResponse
from msgraph.generated.models.recipient import Recipient

from microsoft365 import register, registration
from tests.test_handlers_outlook_calendar import HandlerGraphAdapter
from tests.test_registration_secrets import RecordingContext

# Synthetic mailbox object ID in GUID textual form; never a real tenant identifier.
BILLINGS_ID = "d4c3b2a1-6f5e-47a6-9876-543210fedcba"

#: The compact candidate metadata the search's default ``$select`` carries, asserted here as a
#: literal (not via the handler constant) so a drift of that default fails a test.
COMPACT_MESSAGE_FIELDS = ["id", "subject", "from", "receivedDateTime", "hasAttachments", "webLink"]


def candidate_message(message_id, *, address, name=None, stamp="2026-10-01T09:30:00+00:00"):
    """A real generated ``Message`` carrying the candidate metadata: a ``from`` Recipient with
    an ``EmailAddress`` (address and, when given, a display name), a received date and a link."""
    email = EmailAddress(address=address)
    if name is not None:
        email.name = name
    return Message(
        id=message_id,
        subject="Facture #1042",
        from_=Recipient(email_address=email),
        received_date_time=datetime.fromisoformat(stamp),
        web_link=f"https://outlook.office.com/mail/inbox/{message_id}",
        has_attachments=True,
    )


@pytest.mark.parametrize("has_attachments", [True, False])
def test_registered_invoice_search_composes_filters_without_dropping_false(monkeypatch, has_attachments):
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): MessageCollectionResponse(value=[
        Message(id="message", subject="facture", has_attachments=has_attachments, received_date_time=datetime(2026, 10, 1, tzinfo=timezone.utc)),
    ])})
    monkeypatch.setattr(registration, "create_graph_client", lambda settings: GraphServiceClient(request_adapter=adapter))
    ctx = RecordingContext({"allowed_mailbox_ids": [BILLINGS_ID], "capabilities": {"outlook": {"search": True}}})
    register(ctx)
    args = {
        "action": "search", "user_id": BILLINGS_ID, "query": "facture", "has_attachments": has_attachments,
        "sender": "o'hara@example.test", "received_after": "2026-10-01T00:00:00Z", "received_before": "2026-11-01T00:00:00Z", "top": 20,
    }
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args=args) is None
    result = ctx.tools["microsoft365_outlook"]["handler"](args)
    assert result["status"] == "succeeded"
    assert result["user_id"] == BILLINGS_ID
    assert result["result"]["value"][0]["id"] == "message"
    assert result["result"]["value"][0]["hasAttachments"] is has_attachments
    query = adapter.requests[0].query_parameters
    assert query["%24top"] == 20
    assert query["%24filter"] == (
        "contains(subject,'facture') and hasAttachments eq " + str(has_attachments).lower()
        + " and from/emailAddress/address eq 'o''hara@example.test'"
        + " and receivedDateTime ge 2026-10-01T00:00:00Z and receivedDateTime lt 2026-11-01T00:00:00Z"
    )


@pytest.mark.parametrize("extra", [
    {"has_attachments": "true"}, {"sender": "not-an-address"},
    {"received_after": "2026-10-01"}, {"received_before": "2026-10-01T12:00:00"},
    {"filter": "isRead eq false", "has_attachments": True},
    {"received_after": "2026-11-01T00:00:00Z", "received_before": "2026-10-01T00:00:00Z"},
    # equal bounds are an empty interval (inclusive lower, exclusive upper)
    {"received_after": "2026-10-01T00:00:00Z", "received_before": "2026-10-01T00:00:00Z"},
    # equal instants across different zone offsets
    {"received_after": "2026-10-02T00:00:00Z", "received_before": "2026-10-01T22:00:00-02:00"},
    # the raw filter and the subject query must not be combined: the query would
    # be silently dropped
    {"filter": "isRead eq false", "query": "facture"},
])
def test_bad_structured_search_is_rejected_before_client(monkeypatch, extra):
    calls = []
    monkeypatch.setattr(registration, "create_graph_client", lambda settings: calls.append(settings))
    ctx = RecordingContext({"user_id": "main", "capabilities": {"outlook": {"search": True}}})
    register(ctx)
    args = {"action": "search", "user_id": "main", **extra}
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args=args)["action"] == "block"
    assert json.loads(ctx.tools["microsoft365_outlook"]["handler"](args))["error"] == "validation_error"
    assert calls == []


def _search_context(monkeypatch, adapter):
    """A registered search tool wired to the strict transport double."""
    monkeypatch.setattr(
        registration, "create_graph_client", lambda settings: GraphServiceClient(request_adapter=adapter)
    )
    ctx = RecordingContext({"allowed_mailbox_ids": [BILLINGS_ID], "capabilities": {"outlook": {"search": True}}})
    register(ctx)
    return ctx


def test_default_search_selects_compact_candidate_metadata_and_never_the_body(monkeypatch):
    """With no caller ``select`` the generated GET carries the compact metadata, not the body."""
    page = MessageCollectionResponse(value=[
        Message(id="message", subject="facture", has_attachments=True),
    ])
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): page})
    ctx = _search_context(monkeypatch, adapter)
    args = {"action": "search", "user_id": BILLINGS_ID, "query": "facture", "top": 20}
    result = ctx.tools["microsoft365_outlook"]["handler"](args)

    assert result["status"] == "succeeded"
    selected = adapter.requests[0].query_parameters["%24select"]
    # the compact candidate metadata, in the order the search declares it -- a literal, so a
    # drift of the default select fails here instead of being computed from the handler
    assert list(selected) == COMPACT_MESSAGE_FIELDS
    # every field the metadata promise names, and neither body field
    assert {"id", "subject", "from", "receivedDateTime", "hasAttachments", "webLink"} == set(COMPACT_MESSAGE_FIELDS)
    assert "body" not in selected and "bodyPreview" not in selected


def test_default_search_reports_the_source_mailbox_as_provenance(monkeypatch):
    """The result carries the mailbox it came from, so a candidate is never orphaned."""
    page = MessageCollectionResponse(value=[
        Message(id="message", subject="facture", has_attachments=True),
    ])
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): page})
    ctx = _search_context(monkeypatch, adapter)

    result = ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": BILLINGS_ID, "top": 20}
    )

    assert result["user_id"] == BILLINGS_ID
    assert result["result"]["value"][0]["id"] == "message"


def test_caller_select_overrides_the_compact_default(monkeypatch):
    """A caller ``select`` is honored verbatim and the compact default is not injected."""
    page = MessageCollectionResponse(value=[Message(id="message", subject="facture")])
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): page})
    ctx = _search_context(monkeypatch, adapter)
    args = {"action": "search", "user_id": BILLINGS_ID, "select": ["id", "subject", "body"], "top": 20}

    result = ctx.tools["microsoft365_outlook"]["handler"](args)

    assert result["status"] == "succeeded"
    assert list(adapter.requests[0].query_parameters["%24select"]) == ["id", "subject", "body"]


def test_registered_search_result_carries_the_compact_metadata_it_selected(monkeypatch):
    """A real Message's sender, date and link reach the tool result, not just the request.

    The ``$select`` asks for the candidate metadata, but the result must actually carry it: the
    shared normalization keeps only the fields it knows, so this pins that the sender, the
    received date and the message link survive into the payload the model sees. RED first:
    before the fix the three fields were silently dropped from the result.
    """
    message = candidate_message("AAMkAD1-msg", address="o'hara@example.test", name="O'Hara & Sons")
    page = MessageCollectionResponse(value=[message])
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): page})
    ctx = _search_context(monkeypatch, adapter)

    result = ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": BILLINGS_ID, "query": "facture", "top": 20}
    )

    assert result["status"] == "succeeded"
    # mailbox provenance: the candidate is never orphaned from its source mailbox
    assert result["user_id"] == BILLINGS_ID
    candidate = result["result"]["value"][0]
    # the known-type-model sender, surfaced as a usable email/name structure (not a raw model)
    assert candidate["from"] == {"address": "o'hara@example.test", "name": "O'Hara & Sons"}
    # the received date, as JSON text (not a model marker such as {"type": "datetime"})
    assert candidate["receivedDateTime"] == "2026-10-01T09:30:00+00:00"
    # the message link, usable as-is
    assert candidate["webLink"] == "https://outlook.office.com/mail/inbox/AAMkAD1-msg"
    # and the plain metadata a compact listing always carries
    assert candidate["id"] == "AAMkAD1-msg"
    assert candidate["subject"] == "Facture #1042"
    assert candidate["hasAttachments"] is True


def test_registered_search_caller_select_injects_no_unrequested_metadata(monkeypatch):
    """An explicit ``select`` projects only the fields asked for; the compact default is not
    smuggled into the result as unrequested payload."""
    message = candidate_message("AAMkAD1-msg", address="o'hara@example.test", name="O'Hara & Sons")
    page = MessageCollectionResponse(value=[message])
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): page})
    ctx = _search_context(monkeypatch, adapter)

    result = ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": BILLINGS_ID, "select": ["id", "subject"], "top": 20}
    )

    assert result["status"] == "succeeded"
    # the request selects exactly what was asked
    assert list(adapter.requests[0].query_parameters["%24select"]) == ["id", "subject"]
    candidate = result["result"]["value"][0]
    # ...and the result carries only those fields: no from, date, link or attachment flag
    assert set(candidate) == {"id", "subject"}
    assert "from" not in candidate and "receivedDateTime" not in candidate
    assert "webLink" not in candidate and "hasAttachments" not in candidate


def test_registered_search_second_page_preserves_dates_and_senders_with_truthful_cap(monkeypatch):
    """A second page's candidates keep their sender and date, and the continuation cap is
    truthful: the caller budget is what stops the traversal, not a page boundary."""
    first = f"https://graph.microsoft.com/v1.0/users/{BILLINGS_ID}/messages?$skiptoken=page-2"
    page_one = MessageCollectionResponse(
        value=[
            candidate_message("AAMkAD1", address="o'hara@example.test", name="O'Hara & Sons",
                              stamp="2026-10-01T09:30:00+00:00"),
            candidate_message("AAMkAD2", address="factures@acme.test", name="ACME Billing"),
        ],
        odata_count=4,
        odata_next_link=first,
    )
    page_two = MessageCollectionResponse(
        value=[
            candidate_message("AAMkAD3", address="invoices@northwind.test", name="Northwind",
                              stamp="2026-10-03T14:15:00+00:00"),
            candidate_message("AAMkAD4", address="billing@contoso.test"),
        ]
    )
    adapter = HandlerGraphAdapter(
        {
            ("GET", f"/users/{BILLINGS_ID}/messages"): page_one,
            ("GET", f"/users/{BILLINGS_ID}/messages", (("$skiptoken", "page-2"),)): page_two,
        }
    )
    ctx = _search_context(monkeypatch, adapter)

    # the caller budget (3) cuts across the page boundary: page one is fully taken, then the
    # second page is followed and one of its two candidates is taken before the budget is met.
    result = ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": BILLINGS_ID, "query": "facture", "top": 3}
    )

    assert result["status"] == "succeeded"
    assert [item["id"] for item in result["result"]["value"]] == ["AAMkAD1", "AAMkAD2", "AAMkAD3"]
    # the second-page candidate is preserved with its sender and date intact
    second_page = result["result"]["value"][2]
    assert second_page["from"] == {"address": "invoices@northwind.test", "name": "Northwind"}
    assert second_page["receivedDateTime"] == "2026-10-03T14:15:00+00:00"
    assert len(adapter.requests) == 2
    # the cap is truthful: stopped by the caller budget across two pages, truncated
    assert result["continuation"]["stop_reason"] == "caller_limit"
    assert result["continuation"]["truncated"] is True
    assert result["continuation"]["pages_fetched"] == 2
    assert result["continuation"]["items_fetched"] == 3


def test_registered_search_projects_sender_as_a_usable_email_name(monkeypatch):
    """``sender`` is the same known Recipient model as ``from``: when selected it too is
    surfaced as a flat email/name structure, and omitted when the message has no sender."""
    with_sender = candidate_message("AAMkAD1", address="o'hara@example.test", name="O'Hara & Sons")
    with_sender.sender = with_sender.from_
    page = MessageCollectionResponse(
        value=[with_sender, candidate_message("AAMkAD2", address="factures@acme.test")]
    )
    adapter = HandlerGraphAdapter({("GET", f"/users/{BILLINGS_ID}/messages"): page})
    ctx = _search_context(monkeypatch, adapter)

    result = ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": BILLINGS_ID, "query": "facture", "select": ["id", "sender"], "top": 20}
    )

    assert result["status"] == "succeeded"
    values = result["result"]["value"]
    # the known-type-model sender, surfaced as a usable email/name structure
    assert values[0]["sender"] == {"address": "o'hara@example.test", "name": "O'Hara & Sons"}
    # no sender on the second message: the field is absent, not a marker
    assert "sender" not in values[1]
