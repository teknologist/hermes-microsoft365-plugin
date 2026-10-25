"""Outlook attachment handlers: list_attachments and download_attachment.

Split from ``handlers/outlook.py`` so the attachment lane owns one file. Discovery is the
whole registration mechanism (``load_handlers`` imports every module of this package), so
no registration edit is needed -- and the :func:`handler` decorator refuses a duplicate
key, so a handler cannot be registered twice, from this module or from ``outlook.py``.

Shared contract (finalized with the foundation slice; the attachment lane implements
against it and does not re-decide it)
-------------------------------------
* ``outlook.list_attachments`` is metadata only: id, name, size, content type, inline
  status and attachment type. ``contentBytes`` is never requested on the listing.
* ``outlook.download_attachment`` downloads a file attachment under an explicit,
  documented **decoded supported-size policy**: 10 MiB for a local artifact and 2048 bytes
  for explicit base64 output (whose base64 form, 2732 chars, stays under the 4000-char
  normalized string budget, so a second normalization pass cannot truncate it). The policy
  is enforced from the declared metadata size *before* the content is retrieved, from the
  content response's own declared size, and from the returned byte length *after*; it is a
  size this plugin supports and verifies -- not a claim about a network or memory bound.
* The default output choice is a **local artifact** (generated unique name, profile-scoped
  root, restrictive permissions, no overwrite, no base64 in the response). Base64 is
  available only on the *explicit* choice ``save_to_disk: False`` and only under the small
  2048-byte decoded cap, because a large base64 model response is not a useful invoice
  workflow.
* Item and reference attachments are refused with a machine-readable error before any
  content retrieval; their bytes are never claimed to have been downloaded.
* The returned type, attachment ID, byte length and limit are rechecked after the content
  retrieval as well as the declared metadata: a response that does not name the requested
  attachment is refused, never returned.
* Base64 output is integrity-bearing: the typed
  :class:`~microsoft365.results.BinaryResult` carrier is recognized by
  :func:`microsoft365.results.normalize_result` before any string rule applies, and the
  2048-byte cap keeps its base64 within the 4000-char string budget, so downstream
  normalization passes cannot silently truncate the bytes. A plain dict that merely names a
  ``content_base64`` key gets no such exception: redaction, string bounds, depth and cycle
  detection apply to it exactly as to any other mapping.
"""

from __future__ import annotations

from dataclasses import replace
from typing import Mapping

from msgraph.generated.models.file_attachment import FileAttachment

from ..attachments import save_attachment
from ..execution import execute_request
from ..group_mail import attachment_entries, find_attachment, post_builder
from ..mailbox import GROUP_MAILBOX, route
from ..paging import paginate
from ..results import binary_result, normalize_result
from . import (
    HandlerContext,
    collection_payload,
    handler,
    identifier,
    item_budget,
    open_graph_client,
    refuse,
    success_payload,
    typed_configuration,
)

LIST_ATTACHMENTS = "outlook.list_attachments"
DOWNLOAD_ATTACHMENT = "outlook.download_attachment"
#: Decoded supported-size policy for a downloaded attachment (10 MiB). Checked against the
#: declared metadata size before content retrieval and the returned length after; not a
#: network or memory bound.
MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024
#: Response cap for the explicit base64 output choice: 2048 decoded bytes, whose base64
#: form (2732 chars) stays under the 4000-char normalized string budget, so a second
#: normalization pass cannot truncate it. Above this cap the caller is told to use the
#: local-artifact output choice instead of inlining the bytes.
MAX_BASE64_ATTACHMENT_BYTES = 2048

METADATA_FIELDS = ["id", "name", "size", "contentType", "isInline"]
#: The content read's ``$select``, with ``contentBytes`` named through its **derived type**.
#:
#: ``contentBytes`` is declared on ``microsoft.graph.fileAttachment``, not on the base
#: ``microsoft.graph.attachment`` the collection is typed as, so a bare ``contentBytes``
#: is rejected by Graph with ``Could not find a property named 'contentBytes' on type
#: 'microsoft.graph.attachment'`` -- a 400 on every real download. The OData type cast is the
#: form the service accepts, and it keeps the read bounded to the three fields the download
#: actually needs instead of falling back to "select nothing and take the whole attachment".
CONTENT_FIELDS = ["id", "size", "microsoft.graph.fileAttachment/contentBytes"]
#: Why the declared ``size`` is checked as an upper bound and not as an equality.
#:
#: Graph reports an attachment's ``size`` as the length of the attachment **as it is stored
#: in the message**, which includes its MIME encoding overhead; ``contentBytes`` is the
#: decoded payload and is therefore legitimately *smaller* (a real 28472-byte PDF attachment
#: decodes to 28120 bytes). Demanding equality rejects every genuine download, so the
#: verified relationship is the true one: the decoded payload must fit inside the size the
#: service declared, and inside the supported-size policy. Nothing is loosened by this --
#: bytes beyond either bound are still refused, never truncated.
#: Posts of one group conversation a download may scan to locate the named attachment.
#: A group mailbox has no per-attachment route for an application identity, so the thread's
#: posts are the unit that carries one; the scan stays inside the paginator's own bound.
MAX_POSTS_PER_THREAD = 100


def _group_posts(client, *, adapter, mailbox_id: str, thread_id: str, limit: int):
    """The posts of one group conversation with their attachments expanded.

    ``$expand=attachments`` is the only route that works for an application identity: the
    per-post ``/attachments`` collection is refused by Exchange, while the expanded
    navigation property carries the metadata and, for a file attachment, ``contentBytes``.
    """
    builder = post_builder(client, mailbox_id, thread_id)
    configuration = typed_configuration(builder, top=limit, expand=["attachments"])
    return paginate(builder, adapter=adapter, limit=limit, configuration=configuration)


@handler("outlook", "list_attachments")
def list_attachments(arguments: Mapping, context: HandlerContext) -> dict:
    """List bounded attachment metadata, never attachment content bytes.

    On a group mailbox ``message_id`` names the conversation thread, and the listing covers
    the attachments of every post in it; each entry names the post that carries it.
    """
    operation = LIST_ATTACHMENTS
    user_id = identifier(arguments, "user_id", operation=operation)
    message_id = identifier(arguments, "message_id", operation=operation)
    limit = item_budget(arguments, operation=operation)
    client, adapter = open_graph_client(context)

    def user_route() -> dict:
        builder = (
            client.users.by_user_id(user_id)
            .messages.by_message_id(message_id)
            .attachments
        )
        configuration = typed_configuration(builder, top=limit, select=METADATA_FIELDS)
        paged = paginate(
            builder, adapter=adapter, limit=limit, configuration=configuration
        )
        metadata = [
            {
                "id": item.id,
                "name": item.name,
                "size": item.size,
                "content_type": item.content_type,
                "is_inline": item.is_inline,
                "attachment_type": item.odata_type,
            }
            for item in paged.items
        ]
        return {
            **collection_payload(
                operation,
                replace(
                    paged,
                    payload={**paged.payload, "value": normalize_result(metadata)},
                ),
            ),
            "user_id": user_id,
            "message_id": message_id,
        }

    def group_route() -> dict:
        paged = _group_posts(
            client,
            adapter=adapter,
            mailbox_id=user_id,
            thread_id=message_id,
            limit=limit,
        )
        entries = attachment_entries(paged.items)[:limit]
        return {
            **collection_payload(
                operation,
                replace(
                    paged,
                    payload={**paged.payload, "value": normalize_result(entries)},
                ),
            ),
            "user_id": user_id,
            "message_id": message_id,
            "mailbox_type": GROUP_MAILBOX,
        }

    return route(
        user_id,
        client=client,
        adapter=adapter,
        settings=context.settings,
        arguments=arguments,
        user_route=user_route,
        group_route=group_route,
    )


@handler("outlook", "download_attachment")
def download_attachment(arguments: Mapping, context: HandlerContext) -> dict:
    """Download one file attachment: a local artifact by default, base64 only on explicit choice."""
    operation = DOWNLOAD_ATTACHMENT
    user_id = identifier(arguments, "user_id", operation=operation)
    message_id = identifier(arguments, "message_id", operation=operation)
    attachment_id = identifier(arguments, "attachment_id", operation=operation)
    use_local_artifact = arguments.get("save_to_disk") is not False
    limit = MAX_ATTACHMENT_BYTES if use_local_artifact else MAX_BASE64_ATTACHMENT_BYTES
    client, adapter = open_graph_client(context)
    return route(
        user_id,
        client=client,
        adapter=adapter,
        settings=context.settings,
        arguments=arguments,
        user_route=lambda: _user_download(
            client=client,
            adapter=adapter,
            user_id=user_id,
            message_id=message_id,
            attachment_id=attachment_id,
            use_local_artifact=use_local_artifact,
            limit=limit,
        ),
        group_route=lambda: _group_download(
            arguments,
            client=client,
            adapter=adapter,
            mailbox_id=user_id,
            thread_id=message_id,
            attachment_id=attachment_id,
            use_local_artifact=use_local_artifact,
            limit=limit,
        ),
    )


def _user_download(
    *,
    client,
    adapter,
    user_id: str,
    message_id: str,
    attachment_id: str,
    use_local_artifact: bool,
    limit: int,
) -> dict:
    """The user-mailbox branch of :func:`download_attachment` (personal or shared)."""
    operation = DOWNLOAD_ATTACHMENT
    builder = (
        client.users.by_user_id(user_id)
        .messages.by_message_id(message_id)
        .attachments.by_attachment_id(attachment_id)
    )
    metadata = execute_request(
        builder,
        method="GET",
        adapter=adapter,
        configuration=typed_configuration(builder, select=METADATA_FIELDS),
    )
    if not isinstance(metadata, FileAttachment):
        refuse(
            "operation_not_implemented",
            f"{operation}: only file attachments are supported; item/reference attachments are not downloadable",
        )
    if metadata.id != attachment_id:
        refuse(
            "precondition_failed",
            f"{operation}: the returned attachment does not match the requested attachment id",
        )
    if type(metadata.size) is not int or metadata.size < 0:
        refuse(
            "precondition_failed",
            f"{operation}: attachment size metadata is missing or invalid",
        )
    if metadata.size > limit:
        if use_local_artifact:
            refuse(
                "operation_not_implemented",
                f"{operation}: attachment exceeds the supported size for a local artifact",
            )
        # Actionable: the base64 output choice is capped far below the local-artifact cap,
        # so the remedy is to retry with the local artifact instead of inlining the bytes.
        refuse(
            "operation_not_implemented",
            f"{operation}: attachment exceeds the supported size for base64 output; "
            f"retry with save_to_disk to write a local artifact",
        )
    attachment = execute_request(
        builder,
        method="GET",
        adapter=adapter,
        configuration=typed_configuration(builder, select=CONTENT_FIELDS),
    )
    if not isinstance(attachment, FileAttachment):
        refuse(
            "operation_not_implemented",
            f"{operation}: only file attachments are supported",
        )
    if attachment.id != attachment_id:
        refuse(
            "precondition_failed",
            f"{operation}: the returned attachment does not match the requested attachment id",
        )
    data = attachment.content_bytes
    # The content response's *own* declared size must be a sane nonnegative int; it is
    # checked, not ignored.
    if type(attachment.size) is not int or attachment.size < 0:
        refuse(
            "precondition_failed",
            f"{operation}: attachment content size metadata is missing or invalid",
        )
    if not isinstance(data, bytes):
        refuse(
            "precondition_failed",
            f"{operation}: attachment content is not a byte payload",
        )
    # ``refuse`` raises on every failed branch above; from here the bytes are verified.
    assert isinstance(data, bytes)
    if (
        len(data) > limit
        or attachment.size != metadata.size
        or len(data) > attachment.size
    ):
        refuse(
            "precondition_failed",
            f"{operation}: attachment content size does not match bounded metadata",
        )
    return _downloaded(
        data,
        operation=operation,
        name=metadata.name,
        content_type=metadata.content_type,
        use_local_artifact=use_local_artifact,
        identifiers={
            "user_id": user_id,
            "message_id": message_id,
            "attachment_id": attachment_id,
        },
    )


def _downloaded(
    data: bytes,
    *,
    operation: str,
    name,
    content_type,
    use_local_artifact: bool,
    identifiers: dict,
) -> dict:
    """The payload of a verified attachment download, in the caller's chosen output form.

    The two output forms of the operation live here so both mailbox routes produce exactly
    the same result shape: a private local artifact with its hash and provenance by default,
    and the typed binary carrier only on the explicit ``save_to_disk: False`` choice.
    """
    media_type = content_type or "application/octet-stream"
    if use_local_artifact:
        return success_payload(
            operation,
            {
                **save_attachment(data, content_type=content_type),
                **identifiers,
                "name": name,
                "content_type": media_type,
            },
        )
    return {
        **success_payload(
            operation, binary_result(data, content_type=media_type, name=name)
        ),
        **identifiers,
    }


def _group_download(
    arguments: Mapping,
    *,
    client,
    adapter,
    mailbox_id: str,
    thread_id: str,
    attachment_id: str,
    use_local_artifact: bool,
    limit: int,
) -> dict:
    """The group-mailbox branch of :func:`download_attachment`.

    One request answers with the thread's posts and their expanded attachments, so the
    attachment is located by the id the caller named -- an answer that does not carry it is
    refused, never substituted -- and the same verified size policy as the user route is
    enforced from the declared metadata *and* the returned byte length.
    """
    operation = DOWNLOAD_ATTACHMENT
    paged = _group_posts(
        client,
        adapter=adapter,
        mailbox_id=mailbox_id,
        thread_id=thread_id,
        limit=MAX_POSTS_PER_THREAD,
    )
    attachment, post = find_attachment(paged.items, attachment_id)
    if attachment is None:
        refuse(
            "not_found",
            f"{operation}: the conversation carries no attachment with the requested id",
        )
    if not isinstance(attachment, FileAttachment):
        refuse(
            "operation_not_implemented",
            f"{operation}: only file attachments are supported; item/reference attachments are not downloadable",
        )
    if type(attachment.size) is not int or attachment.size < 0:
        refuse(
            "precondition_failed",
            f"{operation}: attachment size metadata is missing or invalid",
        )
    if attachment.size > limit:
        if use_local_artifact:
            refuse(
                "operation_not_implemented",
                f"{operation}: attachment exceeds the supported size for a local artifact",
            )
        refuse(
            "operation_not_implemented",
            f"{operation}: attachment exceeds the supported size for base64 output; "
            f"retry with save_to_disk to write a local artifact",
        )
    data = attachment.content_bytes
    if not isinstance(data, bytes):
        refuse(
            "precondition_failed",
            f"{operation}: attachment content is not a byte payload",
        )
    assert isinstance(data, bytes)
    if len(data) > limit or len(data) > attachment.size:
        refuse(
            "precondition_failed",
            f"{operation}: attachment content size does not match bounded metadata",
        )
    return {
        **_downloaded(
            data,
            operation=operation,
            name=attachment.name,
            content_type=attachment.content_type,
            use_local_artifact=use_local_artifact,
            identifiers={
                "user_id": mailbox_id,
                "message_id": thread_id,
                "attachment_id": attachment_id,
            },
        ),
        "post_id": getattr(post, "id", None),
        "mailbox_type": GROUP_MAILBOX,
    }
