"""Group mailboxes: routing, the thread/post reads, and the host result contract.

Every request here goes through the real generated SDK builders and a strict offline
transport double, so an invented route raises instead of silently passing. Two defects are
pinned by this module:

* a registered Outlook tool used to hand the host a ``dict``, which the host rejects with
  ``tool_result_contract`` -- so a successful Graph read was reported as a failed call;
* a Microsoft 365 group mailbox was read through ``/users/{id}``, which cannot answer
  anything but ``not_found`` for one.
"""

from __future__ import annotations

import base64
import json
from urllib.parse import urlsplit

import pytest
from msgraph import GraphServiceClient
from msgraph.generated.models.conversation_thread import ConversationThread
from msgraph.generated.models.conversation_thread_collection_response import (
    ConversationThreadCollectionResponse,
)
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.models.group import Group
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message_collection_response import (
    MessageCollectionResponse,
)
from msgraph.generated.models.post import Post
from msgraph.generated.models.post_collection_response import PostCollectionResponse
from msgraph.generated.models.recipient import Recipient

from microsoft365 import mailbox, register, registration
from tests.test_handlers_outlook_calendar import HandlerGraphAdapter
from tests.test_registration_secrets import RecordingContext

#: Synthetic directory object IDs in GUID textual form; never real tenant identifiers.
USER_ID = "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d"
GROUP_ID = "d4c3b2a1-6f5e-47a6-9876-543210fedcba"
THREAD_ID = "thread"
POST_ID = "post"
ATTACHMENT_ID = "attachment"

USER_MESSAGES = ("GET", f"/users/{USER_ID}/messages")
GROUP_AS_USER_MESSAGES = ("GET", f"/users/{GROUP_ID}/messages")
GROUP_PROBE = ("GET", f"/groups/{GROUP_ID}")
GROUP_THREADS = ("GET", f"/groups/{GROUP_ID}/threads")
GROUP_POSTS = ("GET", f"/groups/{GROUP_ID}/threads/{THREAD_ID}/posts")


@pytest.fixture(autouse=True)
def forget_kinds():
    """Each test resolves the mailbox kind itself; no verdict leaks between them."""
    mailbox.forget_resolved_kinds()
    yield
    mailbox.forget_resolved_kinds()


def not_found():
    from kiota_abstractions.api_error import APIError

    return APIError("raw graph text must never surface", 404, {})


def unified_group():
    return Group(id=GROUP_ID, mail_enabled=True, group_types=["Unified"])


def thread(**overrides):
    values = {
        "id": THREAD_ID,
        "topic": "Invoice 2026-10",
        "preview": "Your invoice is attached",
        "has_attachments": True,
        "unique_senders": ["Acme Billing Ltd"],
    }
    values.update(overrides)
    return ConversationThread(**values)


def post(*, attachments=None, body="the invoice body"):
    return Post(
        id=POST_ID,
        from_=Recipient(
            email_address=EmailAddress(address="billing@example.com", name="Billing")
        ),
        has_attachments=bool(attachments),
        body=ItemBody(content=body),
        attachments=attachments,
    )


def context(monkeypatch, adapter, *, action):
    """A registered, configured operation over an offline generated SDK client."""
    calls = []
    monkeypatch.setattr(
        registration,
        "create_graph_client",
        lambda settings: calls.append(settings)
        or GraphServiceClient(request_adapter=adapter),
    )
    ctx = RecordingContext(
        {
            "user_id": USER_ID,
            "allowed_mailbox_ids": [USER_ID, GROUP_ID],
            "capabilities": {"outlook": {action: True}},
        }
    )
    register(ctx)
    return ctx, calls


def dispatch(ctx, args):
    """The registered tool's host-facing answer, decoded.

    The decode is the assertion: the host's tool contract accepts a string (or the
    multimodal envelope) and turns anything else into ``tool_result_contract``, so a result
    that is not valid JSON text never reached the model in the first place.
    """
    raw = ctx.tools["microsoft365_outlook"]["handler"](args)
    assert isinstance(raw, str), raw
    return json.loads(raw)


def paths(adapter):
    return [urlsplit(request.url).path for request in adapter.requests]


# ======================================================================================
# The host result contract
# ======================================================================================


def test_a_successful_read_reaches_the_host_as_a_string_not_a_dict(monkeypatch):
    """The defect: a dict handler payload is refused by the host as a failed tool call."""
    adapter = HandlerGraphAdapter({USER_MESSAGES: MessageCollectionResponse(value=[])})
    ctx, _ = context(monkeypatch, adapter, action="search")

    raw = ctx.tools["microsoft365_outlook"]["handler"](
        {"action": "search", "user_id": USER_ID}
    )

    assert isinstance(raw, str)
    assert json.loads(raw)["status"] == "succeeded"


def test_the_host_contract_serialization_adds_no_second_normalization():
    """Serialization is transport only: it never re-bounds, re-redacts or drops a field.

    The payload has already been through ``results.normalize_result``; a second pass here
    would be a silent, invisible truncation of an answer that was already bounded.
    """
    from microsoft365.registration import tool_result

    payload = {
        "operation": "outlook.search",
        "status": "succeeded",
        "result": {"value": [{"id": "x", "subject": "y" * 4000}]},
        "continuation": {"truncated": False},
    }

    assert json.loads(tool_result(payload)) == payload


def test_an_already_serialized_refusal_is_handed_over_untouched():
    from microsoft365.registration import tool_result

    payload = json.dumps({"error": "validation_error"})

    assert tool_result(payload) is payload


# ======================================================================================
# Routing: a user mailbox pays nothing, a group mailbox is discovered once
# ======================================================================================


def test_a_user_mailbox_search_sends_no_directory_probe(monkeypatch):
    """Personal and shared mailboxes keep the exact request sequence they always had."""
    adapter = HandlerGraphAdapter({USER_MESSAGES: MessageCollectionResponse(value=[])})
    ctx, _ = context(monkeypatch, adapter, action="search")

    assert (
        dispatch(ctx, {"action": "search", "user_id": USER_ID})["status"] == "succeeded"
    )
    assert paths(adapter) == [f"/users/{USER_ID}/messages"]


def test_a_group_mailbox_is_discovered_from_not_found_and_then_read_on_its_own_route(
    monkeypatch,
):
    """The only symptom a group mailbox can produce on the user route is not_found."""
    adapter = HandlerGraphAdapter(
        {
            GROUP_AS_USER_MESSAGES: not_found(),
            GROUP_PROBE: unified_group(),
            GROUP_THREADS: ConversationThreadCollectionResponse(value=[thread()]),
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(ctx, {"action": "search", "user_id": GROUP_ID})

    assert result["status"] == "succeeded"
    assert result["mailbox_type"] == "group"
    assert result["result"]["unit"] == "conversation_thread"
    assert result["result"]["value"] == [
        {
            "id": THREAD_ID,
            "subject": "Invoice 2026-10",
            "preview": "Your invoice is attached",
            "from": {"name": "Acme Billing Ltd"},
            "hasAttachments": True,
        }
    ]
    assert paths(adapter) == [
        f"/users/{GROUP_ID}/messages",
        f"/groups/{GROUP_ID}",
        f"/groups/{GROUP_ID}/threads",
    ]


def test_the_discovered_kind_is_remembered_so_the_probe_happens_once(monkeypatch):
    adapter = HandlerGraphAdapter(
        {
            GROUP_AS_USER_MESSAGES: not_found(),
            GROUP_PROBE: unified_group(),
            GROUP_THREADS: ConversationThreadCollectionResponse(value=[]),
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="search")

    for _ in range(3):
        assert (
            dispatch(ctx, {"action": "search", "user_id": GROUP_ID})["status"]
            == "succeeded"
        )

    assert paths(adapter) == [
        f"/users/{GROUP_ID}/messages",
        f"/groups/{GROUP_ID}",
        f"/groups/{GROUP_ID}/threads",
        f"/groups/{GROUP_ID}/threads",
        f"/groups/{GROUP_ID}/threads",
    ]


def test_an_explicit_mailbox_type_skips_the_user_route_and_the_probe(monkeypatch):
    adapter = HandlerGraphAdapter(
        {GROUP_THREADS: ConversationThreadCollectionResponse(value=[])}
    )
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(
        ctx, {"action": "search", "user_id": GROUP_ID, "mailbox_type": "group"}
    )

    assert result["status"] == "succeeded"
    assert paths(adapter) == [f"/groups/{GROUP_ID}/threads"]


def test_a_real_missing_message_keeps_its_own_not_found(monkeypatch):
    """A probe that says 'not a group' must not turn a missing item into another error."""
    adapter = HandlerGraphAdapter(
        {
            ("GET", f"/users/{USER_ID}/messages/missing"): not_found(),
            ("GET", f"/users/{USER_ID}"): not_found(),
            ("GET", f"/groups/{USER_ID}"): not_found(),
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="read")

    result = dispatch(
        ctx, {"action": "read", "user_id": USER_ID, "message_id": "missing"}
    )

    assert result["error"] == "not_found"


def test_a_probe_failure_never_replaces_the_failure_the_read_actually_got(monkeypatch):
    """A directory that cannot answer leaves the original error exactly as it was."""
    adapter = HandlerGraphAdapter(
        {("GET", f"/users/{USER_ID}/messages/missing"): not_found()}
    )
    ctx, _ = context(monkeypatch, adapter, action="read")

    result = dispatch(
        ctx, {"action": "read", "user_id": USER_ID, "message_id": "missing"}
    )

    assert result["error"] == "not_found"


def test_an_unknown_mailbox_type_is_refused_before_any_request(monkeypatch):
    adapter = HandlerGraphAdapter({})
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(
        ctx, {"action": "search", "user_id": GROUP_ID, "mailbox_type": "mailbox"}
    )

    assert result["error"] == "validation_error"
    assert adapter.requests == []


def test_a_group_mailbox_outside_the_allowlist_is_refused_before_any_request(
    monkeypatch,
):
    adapter = HandlerGraphAdapter({})
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(
        ctx,
        {
            "action": "search",
            "user_id": "ffffffff-ffff-4fff-8fff-ffffffffffff",
            "mailbox_type": "group",
        },
    )

    assert result["error"] == "validation_error"
    assert adapter.requests == []


# ======================================================================================
# The group reads themselves
# ======================================================================================


def test_group_search_states_that_it_filtered_the_page_and_not_the_mailbox(monkeypatch):
    """The thread collection has no $filter, so a narrowed search says so out loud."""
    adapter = HandlerGraphAdapter(
        {
            GROUP_THREADS: ConversationThreadCollectionResponse(
                value=[
                    thread(id="a", topic="Invoice 2026-10"),
                    thread(id="b", topic="Team lunch", has_attachments=False),
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(
        ctx,
        {
            "action": "search",
            "user_id": GROUP_ID,
            "mailbox_type": "group",
            "query": "invoice",
        },
    )["result"]

    assert result["filter_mode"] == "client"
    assert result["client_filters"] == ["query"]
    assert result["threads_examined"] == 2
    assert [item["id"] for item in result["value"]] == ["a"]
    # The query never reached the wire as a filter the service would have rejected.
    assert adapter.requests[0].query_parameters == {"%24top": 50}


@pytest.mark.parametrize("argument", ["filter", "sender", "select"])
def test_group_search_refuses_what_the_thread_collection_cannot_express(
    monkeypatch, argument
):
    """Refused, never silently dropped: a dropped filter is a wrong answer."""
    adapter = HandlerGraphAdapter({})
    ctx, _ = context(monkeypatch, adapter, action="search")
    values = {
        "filter": "isLocked eq true",
        "sender": "billing@example.com",
        "select": ["id"],
    }

    result = dispatch(
        ctx,
        {
            "action": "search",
            "user_id": GROUP_ID,
            "mailbox_type": "group",
            argument: values[argument],
        },
    )

    assert result["error"] == "validation_error"
    assert adapter.requests == []


def test_group_search_applies_the_structured_bounds_it_can_answer_locally(monkeypatch):
    from datetime import datetime, timezone

    adapter = HandlerGraphAdapter(
        {
            GROUP_THREADS: ConversationThreadCollectionResponse(
                value=[
                    thread(
                        id="old",
                        last_delivered_date_time=datetime(
                            2026, 10, 1, tzinfo=timezone.utc
                        ),
                    ),
                    thread(
                        id="kept",
                        last_delivered_date_time=datetime(
                            2026, 10, 15, tzinfo=timezone.utc
                        ),
                    ),
                    thread(
                        id="new",
                        last_delivered_date_time=datetime(
                            2026, 11, 1, tzinfo=timezone.utc
                        ),
                    ),
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(
        ctx,
        {
            "action": "search",
            "user_id": GROUP_ID,
            "mailbox_type": "group",
            "received_after": "2026-10-12T00:00:00+02:00",
            "received_before": "2026-10-19T00:00:00+02:00",
        },
    )["result"]

    assert [item["id"] for item in result["value"]] == ["kept"]
    assert result["client_filters"] == ["received_after", "received_before"]


def test_group_read_answers_with_the_posts_of_the_named_conversation(monkeypatch):
    adapter = HandlerGraphAdapter({GROUP_POSTS: PostCollectionResponse(value=[post()])})
    ctx, _ = context(monkeypatch, adapter, action="read")

    result = dispatch(
        ctx,
        {
            "action": "read",
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "mailbox_type": "group",
        },
    )

    assert result["status"] == "succeeded"
    assert result["message_id"] == THREAD_ID
    assert result["result"]["unit"] == "post"
    assert result["result"]["value"] == [
        {
            "id": POST_ID,
            "from": {"address": "billing@example.com", "name": "Billing"},
            "hasAttachments": False,
            "body": "the invoice body",
        }
    ]
    assert paths(adapter) == [f"/groups/{GROUP_ID}/threads/{THREAD_ID}/posts"]


def test_group_listing_reports_attachment_metadata_and_never_content(monkeypatch):
    adapter = HandlerGraphAdapter(
        {
            GROUP_POSTS: PostCollectionResponse(
                value=[
                    post(
                        attachments=[
                            FileAttachment(
                                id=ATTACHMENT_ID,
                                name="invoice.pdf",
                                content_type="application/pdf",
                                size=7,
                                is_inline=False,
                                content_bytes=b"private",
                            )
                        ]
                    )
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="list_attachments")

    result = dispatch(
        ctx,
        {
            "action": "list_attachments",
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "mailbox_type": "group",
        },
    )

    assert result["result"]["value"] == [
        {
            "id": ATTACHMENT_ID,
            "name": "invoice.pdf",
            "size": 7,
            "content_type": "application/pdf",
            "is_inline": False,
            "attachment_type": "#microsoft.graph.fileAttachment",
            "post_id": POST_ID,
        }
    ]
    assert "private" not in json.dumps(result)
    assert base64.b64encode(b"private").decode() not in json.dumps(result)
    # ``$expand=attachments`` is the only route an application identity may use.
    assert adapter.requests[0].query_parameters["%24expand"] == ["attachments"]


def test_group_download_returns_the_exact_bytes_of_the_named_attachment(monkeypatch):
    data = b"%PDF-1.7 invoice"
    adapter = HandlerGraphAdapter(
        {
            GROUP_POSTS: PostCollectionResponse(
                value=[
                    post(
                        attachments=[
                            FileAttachment(
                                id="other", name="x.pdf", size=1, content_bytes=b"x"
                            ),
                            FileAttachment(
                                id=ATTACHMENT_ID,
                                name="invoice.pdf",
                                content_type="application/pdf",
                                size=len(data),
                                content_bytes=data,
                            ),
                        ]
                    )
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="download_attachment")

    result = dispatch(
        ctx,
        {
            "action": "download_attachment",
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "attachment_id": ATTACHMENT_ID,
            "mailbox_type": "group",
            "save_to_disk": False,
        },
    )

    assert result["status"] == "succeeded"
    assert base64.b64decode(result["result"]["content_base64"]) == data
    assert result["result"]["size"] == len(data)
    assert result["result"]["name"] == "invoice.pdf"
    assert result["post_id"] == POST_ID


@pytest.mark.parametrize("route", ["user", "group"])
def test_a_declared_size_larger_than_the_payload_is_the_normal_case(monkeypatch, route):
    """Graph's ``size`` counts the MIME-encoded attachment; ``contentBytes`` is decoded.

    A real 28472-byte PDF attachment decodes to 28120 bytes, so demanding equality refuses
    every genuine download. The declared size is an upper bound, and it is still enforced.
    """
    data = b"%PDF-1.7 invoice"
    declared = len(data) + 352
    attachment = FileAttachment(
        id=ATTACHMENT_ID,
        name="invoice.pdf",
        content_type="application/pdf",
        size=declared,
        content_bytes=data,
    )
    if route == "group":
        adapter = HandlerGraphAdapter(
            {
                GROUP_POSTS: PostCollectionResponse(
                    value=[post(attachments=[attachment])]
                )
            }
        )
        args = {
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "mailbox_type": "group",
        }
    else:
        adapter = HandlerGraphAdapter(
            {
                (
                    "GET",
                    f"/users/{USER_ID}/messages/message/attachments/{ATTACHMENT_ID}",
                ): attachment
            }
        )
        args = {"user_id": USER_ID, "message_id": "message"}
    ctx, _ = context(monkeypatch, adapter, action="download_attachment")

    result = dispatch(
        ctx,
        {
            "action": "download_attachment",
            "attachment_id": ATTACHMENT_ID,
            "save_to_disk": False,
            **args,
        },
    )

    assert result["status"] == "succeeded", result
    assert base64.b64decode(result["result"]["content_base64"]) == data


def test_a_payload_larger_than_the_declared_size_is_still_refused(monkeypatch):
    """The bound is enforced in the direction that matters: never more bytes than declared."""
    data = b"%PDF-1.7 invoice"
    adapter = HandlerGraphAdapter(
        {
            GROUP_POSTS: PostCollectionResponse(
                value=[
                    post(
                        attachments=[
                            FileAttachment(
                                id=ATTACHMENT_ID,
                                name="invoice.pdf",
                                content_type="application/pdf",
                                size=len(data) - 1,
                                content_bytes=data,
                            )
                        ]
                    )
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="download_attachment")

    result = dispatch(
        ctx,
        {
            "action": "download_attachment",
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "attachment_id": ATTACHMENT_ID,
            "mailbox_type": "group",
            "save_to_disk": False,
        },
    )

    assert result["error"] == "precondition_failed"


def test_group_download_refuses_an_answer_that_does_not_carry_the_named_attachment(
    monkeypatch,
):
    adapter = HandlerGraphAdapter(
        {
            GROUP_POSTS: PostCollectionResponse(
                value=[
                    post(
                        attachments=[
                            FileAttachment(
                                id="other", name="x.pdf", size=1, content_bytes=b"x"
                            )
                        ]
                    )
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="download_attachment")

    result = dispatch(
        ctx,
        {
            "action": "download_attachment",
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "attachment_id": ATTACHMENT_ID,
            "mailbox_type": "group",
            "save_to_disk": False,
        },
    )

    assert result["error"] == "not_found"


def test_group_download_enforces_the_same_supported_size_policy(monkeypatch):
    from microsoft365.handlers.outlook_attachments import MAX_BASE64_ATTACHMENT_BYTES

    data = b"x" * (MAX_BASE64_ATTACHMENT_BYTES + 1)
    adapter = HandlerGraphAdapter(
        {
            GROUP_POSTS: PostCollectionResponse(
                value=[
                    post(
                        attachments=[
                            FileAttachment(
                                id=ATTACHMENT_ID,
                                name="big.pdf",
                                content_type="application/pdf",
                                size=len(data),
                                content_bytes=data,
                            )
                        ]
                    )
                ]
            )
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="download_attachment")

    result = dispatch(
        ctx,
        {
            "action": "download_attachment",
            "user_id": GROUP_ID,
            "message_id": THREAD_ID,
            "attachment_id": ATTACHMENT_ID,
            "mailbox_type": "group",
            "save_to_disk": False,
        },
    )

    assert result["error"] == "operation_not_implemented"


def test_a_mail_disabled_group_is_not_a_group_mailbox(monkeypatch):
    """A security group has no mailbox: the read keeps the user route and its failure."""
    adapter = HandlerGraphAdapter(
        {
            GROUP_AS_USER_MESSAGES: not_found(),
            GROUP_PROBE: Group(
                id=GROUP_ID, mail_enabled=False, group_types=["Unified"]
            ),
        }
    )
    ctx, _ = context(monkeypatch, adapter, action="search")

    result = dispatch(ctx, {"action": "search", "user_id": GROUP_ID})

    assert result["error"] == "not_found"
    assert paths(adapter) == [f"/users/{GROUP_ID}/messages", f"/groups/{GROUP_ID}"]
