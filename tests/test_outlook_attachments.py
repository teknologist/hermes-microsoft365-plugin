"""Attachment operations use real generated models/builders and offline transport."""
from __future__ import annotations

import base64
import hashlib
import json
from pathlib import Path
from urllib.parse import urlsplit

import pytest
from msgraph import GraphServiceClient
from msgraph.generated.models.attachment_collection_response import AttachmentCollectionResponse
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.models.item_attachment import ItemAttachment
from msgraph.generated.models.reference_attachment import ReferenceAttachment

from microsoft365 import register, registration
from tests.test_handlers_outlook_calendar import HandlerGraphAdapter
from tests.test_registration_secrets import RecordingContext

# Synthetic mailbox object IDs in GUID textual form; never real tenant identifiers.
MAIN_ID = "a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d"
BILLINGS_ID = "d4c3b2a1-6f5e-47a6-9876-543210fedcba"


def attachment_context(monkeypatch, adapter, *, action):
    """Register a real configured operation with an offline generated SDK client."""
    calls = []
    monkeypatch.setattr(registration, "create_graph_client", lambda settings: calls.append(settings) or GraphServiceClient(request_adapter=adapter))
    ctx = RecordingContext({
        "user_id": MAIN_ID, "allowed_mailbox_ids": [MAIN_ID, BILLINGS_ID],
        "capabilities": {"outlook": {action: True}},
    })
    register(ctx)
    return ctx, calls


def test_registered_attachment_listing_returns_metadata_without_content(monkeypatch):
    adapter = HandlerGraphAdapter({
        ("GET", f"/users/{BILLINGS_ID}/messages/message/attachments"): AttachmentCollectionResponse(value=[
            FileAttachment(id="attachment", name="invoice.pdf", content_type="application/pdf", size=7, is_inline=False, content_bytes=b"private"),
        ]),
    })
    ctx, calls = attachment_context(monkeypatch, adapter, action="list_attachments")
    args = {"action": "list_attachments", "user_id": BILLINGS_ID, "message_id": "message", "top": 1}
    assert ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args=args) is None
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"](args))
    assert result["status"] == "succeeded"
    assert (result["user_id"], result["message_id"]) == (BILLINGS_ID, "message")
    assert result["result"]["value"] == [{
        "id": "attachment", "name": "invoice.pdf", "size": 7,
        "content_type": "application/pdf", "is_inline": False, "attachment_type": "#microsoft.graph.fileAttachment",
    }]
    assert len(calls) == 1
    assert len(adapter.requests) == 1
    request = adapter.requests[0]
    assert request.http_method.value == "GET"
    assert urlsplit(request.url).path == f"/users/{BILLINGS_ID}/messages/message/attachments"
    assert request.query_parameters == {"%24top": 1, "%24select": ["id", "name", "size", "contentType", "isInline"]}


class AttachmentDownloadAdapter(HandlerGraphAdapter):
    """Return metadata/content only for the expected generated attachment GET."""

    def __init__(self, metadata, content):
        super().__init__()
        self.metadata = metadata
        self.content = content

    def _answer(self, request_information):
        self.requests.append(request_information)
        assert urlsplit(request_information.url).path == f"/users/{BILLINGS_ID}/messages/message/attachments/attachment"
        assert request_information.http_method.value == "GET"
        fields = request_information.query_parameters["%24select"]
        if fields == ["id", "name", "size", "contentType", "isInline"]:
            return self.metadata
        # ``contentBytes`` lives on the derived fileAttachment type: a bare name is a
        # 400 from Graph on every real download.
        assert fields == ["id", "size", "microsoft.graph.fileAttachment/contentBytes"]
        return self.content


def test_registered_download_preserves_file_attachment_bytes(monkeypatch):
    # Exactly the supported base64 bound (2048 decoded bytes): the default output is a
    # local artifact, which must not be exercised without a hermes-home isolation fixture,
    # so this pins the explicit base64 choice at its supported boundary.
    data = bytes(range(256)) * 8
    assert len(data) == MAX_BASE64_ATTACHMENT_BYTES
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", name="invoice.bin", size=len(data), content_type="application/octet-stream"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, calls = attachment_context(monkeypatch, adapter, action="download_attachment")
    # the explicit base64 output choice: the default output is a local artifact,
    # which must not be exercised without a hermes-home isolation fixture
    args = {"action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message", "attachment_id": "attachment", "save_to_disk": False}
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"](args))
    assert result["status"] == "succeeded"
    assert (result["user_id"], result["message_id"]) == (BILLINGS_ID, "message")
    assert result["result"] == {
        "name": "invoice.bin", "content_type": "application/octet-stream", "size": len(data),
        "content_base64": base64.b64encode(data).decode("ascii"),
    }
    assert len(calls) == 1
    assert len(adapter.requests) == 2


@pytest.mark.parametrize("metadata", [
    ItemAttachment(id="attachment", size=2),
    ReferenceAttachment(id="attachment", size=2),
    FileAttachment(id="attachment", size=1_048_577),
    FileAttachment(id="attachment", size=None),
])
def test_download_refuses_unsupported_or_unbounded_metadata_before_content(monkeypatch, metadata):
    adapter = AttachmentDownloadAdapter(metadata, FileAttachment(content_bytes=b"should-not-fetch"))
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message", "attachment_id": "attachment",
        "save_to_disk": False,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] in {"operation_not_implemented", "precondition_failed"}
    assert len(adapter.requests) == 1


@pytest.mark.parametrize("data,declared", [(b"", 0), (b"abc", 2)])
def test_download_empty_file_or_size_mismatch(monkeypatch, data, declared):
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=declared), FileAttachment(id="attachment", size=declared, content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message", "attachment_id": "attachment",
        "save_to_disk": False,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    if declared == 0:
        assert result["result"]["content_base64"] == ""
        assert result["result"]["size"] == 0
    else:
        assert result["error"] == "precondition_failed"


def test_download_saves_unique_profile_artifacts_without_trusting_filename(monkeypatch, tmp_path):
    import hermes_constants

    monkeypatch.setattr(hermes_constants, "get_hermes_home", lambda: tmp_path)
    data = b"%PDF-exact\x00\xff"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name="../../escape.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    args = {"action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message", "attachment_id": "attachment", "save_to_disk": True}
    records = [json.loads(ctx.tools["microsoft365_outlook"]["handler"](args)) for _ in range(2)]
    paths = []
    for result in records:
        assert isinstance(result, dict), result
        record = result["result"]
        path = Path(record["path"])
        paths.append(path)
        assert path.parent == tmp_path / "downloads" / "microsoft365"
        assert path.read_bytes() == data
        assert record["sha256"] == hashlib.sha256(data).hexdigest()
        assert record["size"] == len(data)
        assert record["name"] == "../../escape.pdf"
        assert record["content_type"] == "application/pdf"
        assert (record["user_id"], record["message_id"], record["attachment_id"]) == (BILLINGS_ID, "message", "attachment")
        assert "content_base64" not in record
    assert paths[0] != paths[1]
    assert not (tmp_path / "escape.pdf").exists()


def test_download_artifact_refuses_symlink_directory(monkeypatch, tmp_path):
    import hermes_constants

    monkeypatch.setattr(hermes_constants, "get_hermes_home", lambda: tmp_path)
    outside = tmp_path / "outside"
    outside.mkdir()
    (tmp_path / "downloads").symlink_to(outside, target_is_directory=True)
    adapter = AttachmentDownloadAdapter(FileAttachment(id="attachment", size=1), FileAttachment(id="attachment", size=1, content_bytes=b"x"))
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message", "attachment_id": "attachment", "save_to_disk": True,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "configuration_error"
    assert list(outside.iterdir()) == []


# ======================================================================================
# Slice 2A acceptance criteria beyond the partial baseline
# ======================================================================================
import os
import stat
from dataclasses import asdict

from microsoft365 import results
from microsoft365.handlers.outlook_attachments import (
    MAX_ATTACHMENT_BYTES,
    MAX_BASE64_ATTACHMENT_BYTES,
)

import microsoft365.attachments as attachments_module

ARTIFACT_DIR = "microsoft365"
ARTIFACT_SUBDIR = "downloads"


def _patched_home(monkeypatch, tmp_path):
    import hermes_constants

    monkeypatch.setattr(hermes_constants, "get_hermes_home", lambda: tmp_path)


def test_listing_follows_bounded_pagination_and_reports_truncation(monkeypatch):
    """Two pages through a validated next link; the caller's top bounds the items."""
    link = f"https://graph.microsoft.com/v1.0/users/{BILLINGS_ID}/messages/message/attachments?%24skiptoken=2"
    first = AttachmentCollectionResponse(
        value=[FileAttachment(id="a1", name="one.pdf", size=3, content_type="application/pdf")],
        odata_next_link=link,
    )
    second = AttachmentCollectionResponse(value=[
        FileAttachment(id="a2", name="two.pdf", size=3, content_type="application/pdf"),
        FileAttachment(id="a3", name="three.pdf", size=3, content_type="application/pdf"),
    ])
    adapter = HandlerGraphAdapter({
        ("GET", f"/users/{BILLINGS_ID}/messages/message/attachments"): first,
        ("GET", f"/users/{BILLINGS_ID}/messages/message/attachments", (("$skiptoken", "2"),)): second,
    })
    ctx, _ = attachment_context(monkeypatch, adapter, action="list_attachments")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "list_attachments", "user_id": BILLINGS_ID, "message_id": "message", "top": 2,
    }))
    assert result["status"] == "succeeded"
    assert [item["id"] for item in result["result"]["value"]] == ["a1", "a2"]
    assert result["continuation"] == {
        "truncated": True, "next_link_present": False, "pages_fetched": 2,
        "items_fetched": 2, "stop_reason": "caller_limit",
    }
    assert len(adapter.requests) == 2
    assert adapter.requests[1].url == link


def test_listing_mixed_attachment_types_report_type_and_exclude_content(monkeypatch):
    page = AttachmentCollectionResponse(value=[
        FileAttachment(id="f", name="inv.pdf", size=7, content_type="application/pdf",
                       is_inline=False, content_bytes=b"private-bytes"),
        ItemAttachment(id="i", name="cal.ics", size=5, content_type="text/calendar", is_inline=True),
    ])
    adapter = HandlerGraphAdapter({
        ("GET", f"/users/{BILLINGS_ID}/messages/message/attachments"): page,
    })
    ctx, _ = attachment_context(monkeypatch, adapter, action="list_attachments")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "list_attachments", "user_id": BILLINGS_ID, "message_id": "message",
    }))
    assert result["status"] == "succeeded"
    items = result["result"]["value"]
    assert [item["attachment_type"] for item in items] == [
        "#microsoft.graph.fileAttachment", "#microsoft.graph.itemAttachment",
    ]
    assert items[0]["is_inline"] is False and items[1]["is_inline"] is True
    for item in items:
        assert set(item) == {"id", "name", "size", "content_type", "is_inline", "attachment_type"}
    assert "private-bytes" not in json.dumps(result)
    # The generated SDK request selects metadata only; contentBytes is never requested.
    request = adapter.requests[0]
    assert request.query_parameters == {
        "%24top": 50, "%24select": ["id", "name", "size", "contentType", "isInline"],
    }
    assert "contentBytes" not in request.query_parameters["%24select"]


def test_download_refuses_metadata_id_mismatch_before_content(monkeypatch):
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="other-attachment", name="invoice.bin", size=1),
        FileAttachment(id="attachment", size=1, content_bytes=b"x"),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message",
        "attachment_id": "attachment", "save_to_disk": True,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "precondition_failed"
    assert len(adapter.requests) == 1


def test_download_refuses_content_id_mismatch(monkeypatch):
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", name="invoice.bin", size=3),
        FileAttachment(id="other-attachment", size=3, content_bytes=b"abc"),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message",
        "attachment_id": "attachment", "save_to_disk": True,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "precondition_failed"
    assert len(adapter.requests) == 2


@pytest.mark.parametrize("content", [None, {"id": "attachment"}])
def test_download_refuses_malformed_content_response(monkeypatch, tmp_path, content):
    _patched_home(monkeypatch, tmp_path)
    adapter = AttachmentDownloadAdapter(FileAttachment(id="attachment", size=1), content)
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message",
        "attachment_id": "attachment",
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] in {"precondition_failed", "operation_not_implemented"}
    assert len(adapter.requests) == 2
    assert not (tmp_path / ARTIFACT_SUBDIR).exists()


def test_download_refuses_artifact_oversize_from_metadata_before_content(monkeypatch):
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=MAX_ATTACHMENT_BYTES + 1),
        FileAttachment(content_bytes=b"should-not-fetch"),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    # default output choice (save_to_disk omitted) is the local artifact, capped at 10 MiB
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID, "message_id": "message",
        "attachment_id": "attachment",
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "operation_not_implemented"
    assert len(adapter.requests) == 1


def test_download_default_output_is_local_artifact(monkeypatch, tmp_path):
    _patched_home(monkeypatch, tmp_path)
    data = b"invoice-bytes"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name="invoice.csv", content_type="text/csv"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    args = {
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment",
    }
    assert "save_to_disk" not in args
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"](args))
    assert isinstance(result, dict), result
    record = result["result"]
    path = Path(record["path"])
    assert path.parent == tmp_path / ARTIFACT_SUBDIR / ARTIFACT_DIR
    assert path.read_bytes() == data
    assert path.suffix == ".csv"
    assert "content_base64" not in record


@pytest.mark.parametrize("content_type,name,expected_suffix", [
    ("application/pdf", "noext", ".pdf"),
    ("text/plain", "fake.pdf", ".txt"),
    ("image/png", "x.png", ".png"),
    ("application/xml", "x.pdf", ".xml"),
    ("application/x-unknown-format", "x.bin", ".bin"),
    (None, "x.pdf", ".bin"),
])
def test_artifact_suffix_derives_from_mime_not_from_remote_name(
    monkeypatch, tmp_path, content_type, name, expected_suffix
):
    _patched_home(monkeypatch, tmp_path)
    data = b"content"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name=name, content_type=content_type),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": True,
    }))
    assert isinstance(result, dict), result
    record = result["result"]
    path = Path(record["path"])
    assert path.suffix == expected_suffix
    assert name != path.name
    assert record["name"] == name  # the remote name stays metadata only


def test_download_empty_file_artifact(monkeypatch, tmp_path):
    _patched_home(monkeypatch, tmp_path)
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=0, name="empty.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=0, content_bytes=b""),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment",
    }))
    assert isinstance(result, dict), result
    record = result["result"]
    assert record["size"] == 0
    assert record["sha256"] == hashlib.sha256(b"").hexdigest()
    assert Path(record["path"]).read_bytes() == b""


def test_download_invalid_save_to_disk_is_blocked_before_any_client(monkeypatch):
    adapter = HandlerGraphAdapter()
    ctx, calls = attachment_context(monkeypatch, adapter, action="download_attachment")
    verdict = ctx.hooks["pre_tool_call"](tool_name="microsoft365_outlook", args={
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": "yes",
    })
    assert verdict is not None and verdict["action"] == "block"
    assert calls == []
    assert adapter.requests == []


def test_artifact_permissions_are_umask_independent(monkeypatch, tmp_path):
    _patched_home(monkeypatch, tmp_path)
    data = b"permission-probe"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name="p.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    original = os.umask(0o277)  # strips the owner write bit from every created mode
    try:
        result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
            "action": "download_attachment", "user_id": BILLINGS_ID,
            "message_id": "message", "attachment_id": "attachment",
        }))
    finally:
        os.umask(original)
    assert isinstance(result, dict), result
    path = Path(result["result"]["path"])
    assert stat.S_IMODE(path.stat().st_mode) == 0o600
    assert stat.S_IMODE(path.parent.stat().st_mode) == 0o700
    assert stat.S_IMODE((path.parent.parent).stat().st_mode) == 0o700


def test_artifact_closes_the_file_descriptor_when_fdopen_fails(monkeypatch, tmp_path):
    # Regression for the descriptor leak: when os.fdopen fails before the file object takes
    # ownership, the raw descriptor the exclusive open created must be closed (not leaked)
    # and the partial file removed -- a pre-fix save left the fd open.
    _patched_home(monkeypatch, tmp_path)
    real_open = os.open
    real_close = os.close
    file_fds: set[int] = set()
    closed_fds: set[int] = set()

    def tracking_open(*args, **kwargs):
        fd = real_open(*args, **kwargs)
        flags = args[1] if len(args) > 1 else kwargs.get("flags", 0)
        if flags & (os.O_CREAT | os.O_EXCL):
            file_fds.add(fd)
        return fd

    def tracking_close(fd, *args, **kwargs):
        closed_fds.add(fd)
        return real_close(fd, *args, **kwargs)

    monkeypatch.setattr(os, "open", tracking_open)
    monkeypatch.setattr(os, "close", tracking_close)
    monkeypatch.setattr(os, "fdopen", lambda fd, *a, **k: (_ for _ in ()).throw(OSError("fdopen failed before ownership")))

    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=1, name="f.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=1, content_bytes=b"x"),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": True,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "configuration_error"
    # The descriptor opened for the artifact was closed on the failure path (no leak).
    assert file_fds, "the exclusive file open was expected to happen"
    assert file_fds <= closed_fds
    # And the partial file was removed.
    directory = tmp_path / ARTIFACT_SUBDIR / ARTIFACT_DIR
    assert not directory.exists() or list(directory.iterdir()) == []


def test_artifact_refuses_group_writable_existing_directory(monkeypatch, tmp_path):
    # A pre-existing group/world-writable directory in the artifact path is unsafe: the
    # writer refuses it (fail closed) instead of landing an artifact on a path another
    # principal can plant entries in.
    _patched_home(monkeypatch, tmp_path)
    directory = tmp_path / ARTIFACT_SUBDIR / ARTIFACT_DIR
    directory.mkdir(parents=True)
    os.chmod(directory, 0o775)  # group-writable: unsafe
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=1, name="w.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=1, content_bytes=b"x"),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": True,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "configuration_error"
    assert list(directory.iterdir()) == []  # nothing written into the unsafe directory


def test_artifact_fails_closed_under_umask_0777(monkeypatch, tmp_path):
    # The worst-case umask strips even the owner bits off a freshly created directory, so
    # the writer cannot open it to verify or pin its mode: it fails closed (a typed error)
    # instead of writing into a directory it cannot control, and drops the partial
    # directory it created so no uncontrolled path is left behind.
    _patched_home(monkeypatch, tmp_path)
    data = b"umask-probe"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name="u.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    original = os.umask(0o777)
    try:
        result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
            "action": "download_attachment", "user_id": BILLINGS_ID,
            "message_id": "message", "attachment_id": "attachment",
        }))
    finally:
        os.umask(original)
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "configuration_error"
    # No partial artifact directory tree is left behind by the failed save.
    assert not (tmp_path / ARTIFACT_SUBDIR).exists()
    assert not (tmp_path / ARTIFACT_SUBDIR / ARTIFACT_DIR).exists()


def test_artifact_cleanup_leaves_no_partial_file_on_post_write_failure(monkeypatch, tmp_path):
    _patched_home(monkeypatch, tmp_path)

    def boom(_data):
        raise AssertionError("hash computation is inside the cleanup guard")

    monkeypatch.setattr("hashlib.sha256", boom)
    data = b"partial"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name="d.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment",
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "configuration_error"
    directory = tmp_path / ARTIFACT_SUBDIR / ARTIFACT_DIR
    assert not directory.exists() or list(directory.iterdir()) == []


def test_artifact_never_overwrites_an_existing_file(monkeypatch, tmp_path):
    _patched_home(monkeypatch, tmp_path)
    import uuid as _uuid

    monkeypatch.setattr(_uuid, "uuid4", lambda: _uuid.UUID(int=0))
    data_one, data_two = b"first-bytes", b"second-bytes"

    def run(data):
        adapter = AttachmentDownloadAdapter(
            FileAttachment(id="attachment", size=len(data), name="n.pdf", content_type="application/pdf"),
            FileAttachment(id="attachment", size=len(data), content_bytes=data),
        )
        ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
        return json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
            "action": "download_attachment", "user_id": BILLINGS_ID,
            "message_id": "message", "attachment_id": "attachment",
        }))

    first = run(data_one)
    assert isinstance(first, dict), first
    path = Path(first["result"]["path"])
    assert path.name == "00000000000000000000000000000000.pdf"
    second = run(data_two)
    second = json.loads(second) if isinstance(second, str) else second
    assert second["error"] == "configuration_error"
    assert path.read_bytes() == data_one


def test_artifact_refuses_pre_planted_symlink_file(monkeypatch, tmp_path):
    _patched_home(monkeypatch, tmp_path)
    import uuid as _uuid

    monkeypatch.setattr(_uuid, "uuid4", lambda: _uuid.UUID(int=1))
    outside = tmp_path / "outside.bin"
    outside.write_bytes(b"untouched")
    directory = tmp_path / ARTIFACT_SUBDIR / ARTIFACT_DIR
    directory.mkdir(parents=True)
    (directory / "00000000000000000000000000000001.pdf").symlink_to(outside)
    data = b"payload"
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", size=len(data), name="s.pdf", content_type="application/pdf"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment",
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "configuration_error"
    assert outside.read_bytes() == b"untouched"


def test_normalize_result_no_special_case_for_a_dict_named_content_base64():
    # A plain dict that merely *names* a content_base64 key is not a binary carrier:
    # before the fix a fast path passed every scalar through verbatim, leaking sibling
    # secrets, skipping string bounds and re-entering cycles. It must now take the exact
    # same redaction / bound / cycle path as any other mapping.
    payload = {
        "content_base64": base64.b64encode(b"invoice").decode("ascii"),
        "secret": "leak",
        "access_token": "leak-2",
        "note": "x" * 9_000,
    }
    normalized = results.normalize_result(payload)
    # The carrier field itself is short and safe: it round-trips, it is not truncated and
    # no verbatim exception let sibling scalars through unredacted or unbounded.
    assert base64.b64decode(normalized["content_base64"], validate=True) == b"invoice"
    assert "[TRUNCATED]" not in normalized["content_base64"]
    assert normalized["secret"] == "[REDACTED]"
    assert normalized["access_token"] == "[REDACTED]"
    assert normalized["note"].endswith("[TRUNCATED]")

    # A self-referential dict that also names content_base64 must be cycle-reported,
    # not recursed into (the pre-fix path re-entered without consulting the seen set).
    cyc: dict = {"content_base64": ""}
    cyc["self"] = cyc
    assert results.normalize_result(cyc)["self"] == "[CYCLE]"


def test_base64_download_at_the_supported_bound_survives_a_second_pass(monkeypatch):
    # Exactly the supported bound: 2048 decoded bytes -> 2732 base64 chars, under the
    # 4000-char normalized string budget, so a second normalization pass is exact.
    data = bytes(range(256)) * 8
    assert len(data) == MAX_BASE64_ATTACHMENT_BYTES
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", name="invoice.bin", size=len(data),
                       content_type="application/octet-stream"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": False,
    }))
    assert result["status"] == "succeeded"
    # A downstream normalization pass over the whole payload must not corrupt the bytes.
    normalized = results.normalize_result(result)
    assert normalized["result"]["content_base64"] == base64.b64encode(data).decode("ascii")
    assert "[TRUNCATED]" not in normalized["result"]["content_base64"]
    assert base64.b64decode(normalized["result"]["content_base64"], validate=True) == data


def test_base64_download_refuses_oversize_before_content_fetch(monkeypatch):
    # One byte over the supported base64 bound is refused before the bytes are fetched,
    # and the error is actionable: it recommends the local-artifact output choice.
    data = bytes(range(256)) * 8 + b"\x00"  # 2049 decoded bytes
    assert len(data) == MAX_BASE64_ATTACHMENT_BYTES + 1
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", name="invoice.bin", size=len(data),
                       content_type="application/octet-stream"),
        FileAttachment(id="attachment", size=len(data), content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": False,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "operation_not_implemented"
    assert "save_to_disk" in result["message"]
    assert len(adapter.requests) == 1  # refused from the declared metadata size, content never fetched


@pytest.mark.parametrize("declared", [None, -5, 2049])
def test_download_enforces_the_content_response_size_against_metadata_and_length(
    monkeypatch, declared
):
    # The content response's own declared size must be a nonnegative int equal to the
    # metadata size and the actual byte length; anything else is refused (it is checked,
    # not ignored).
    data = bytes(range(256)) * 8  # 2048 decoded bytes, at the supported bound
    adapter = AttachmentDownloadAdapter(
        FileAttachment(id="attachment", name="invoice.bin", size=len(data),
                       content_type="application/octet-stream"),
        FileAttachment(id="attachment", size=declared, content_bytes=data),
    )
    ctx, _ = attachment_context(monkeypatch, adapter, action="download_attachment")
    result = json.loads(ctx.tools["microsoft365_outlook"]["handler"]({
        "action": "download_attachment", "user_id": BILLINGS_ID,
        "message_id": "message", "attachment_id": "attachment", "save_to_disk": False,
    }))
    result = json.loads(result) if isinstance(result, str) else result
    assert result["error"] == "precondition_failed"
    assert len(adapter.requests) == 2  # the mismatch is detected in the content response
