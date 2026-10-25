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
CONTENT_FIELDS = ["id", "size", "contentBytes"]


@handler("outlook", "list_attachments")
def list_attachments(arguments: Mapping, context: HandlerContext) -> dict:
    """List bounded attachment metadata, never attachment content bytes."""
    operation = LIST_ATTACHMENTS
    user_id = identifier(arguments, "user_id", operation=operation)
    message_id = identifier(arguments, "message_id", operation=operation)
    limit = item_budget(arguments, operation=operation)
    client, adapter = open_graph_client(context)
    builder = client.users.by_user_id(user_id).messages.by_message_id(message_id).attachments
    configuration = typed_configuration(builder, top=limit, select=METADATA_FIELDS)
    paged = paginate(builder, adapter=adapter, limit=limit, configuration=configuration)
    metadata = [{
        "id": item.id, "name": item.name, "size": item.size,
        "content_type": item.content_type, "is_inline": item.is_inline,
        "attachment_type": item.odata_type,
    } for item in paged.items]
    return {
        **collection_payload(operation, replace(paged, payload={**paged.payload, "value": normalize_result(metadata)})),
        "user_id": user_id, "message_id": message_id,
    }


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
    builder = client.users.by_user_id(user_id).messages.by_message_id(message_id).attachments.by_attachment_id(attachment_id)
    metadata = execute_request(
        builder, method="GET", adapter=adapter,
        configuration=typed_configuration(builder, select=METADATA_FIELDS),
    )
    if not isinstance(metadata, FileAttachment):
        refuse("operation_not_implemented", f"{operation}: only file attachments are supported; item/reference attachments are not downloadable")
    if metadata.id != attachment_id:
        refuse("precondition_failed", f"{operation}: the returned attachment does not match the requested attachment id")
    if type(metadata.size) is not int or metadata.size < 0:
        refuse("precondition_failed", f"{operation}: attachment size metadata is missing or invalid")
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
        builder, method="GET", adapter=adapter,
        configuration=typed_configuration(builder, select=CONTENT_FIELDS),
    )
    if not isinstance(attachment, FileAttachment):
        refuse("operation_not_implemented", f"{operation}: only file attachments are supported")
    if attachment.id != attachment_id:
        refuse("precondition_failed", f"{operation}: the returned attachment does not match the requested attachment id")
    data = attachment.content_bytes
    # The content response's *own* declared size must be a sane nonnegative int; it is
    # checked, not ignored.
    if type(attachment.size) is not int or attachment.size < 0:
        refuse("precondition_failed", f"{operation}: attachment content size metadata is missing or invalid")
    if not isinstance(data, bytes):
        refuse("precondition_failed", f"{operation}: attachment content is not a byte payload")
    # ``refuse`` raises on every failed branch above; from here the bytes are verified.
    assert isinstance(data, bytes)
    if len(data) > limit or attachment.size != metadata.size or attachment.size != len(data):
        refuse("precondition_failed", f"{operation}: attachment content size does not match bounded metadata")
    if use_local_artifact:
        return success_payload(operation, {
            **save_attachment(data, content_type=metadata.content_type),
            "user_id": user_id, "message_id": message_id, "attachment_id": attachment_id,
            "name": metadata.name, "content_type": metadata.content_type or "application/octet-stream",
        })
    return {
        **success_payload(operation, binary_result(
            data, content_type=metadata.content_type or "application/octet-stream", name=metadata.name,
        )),
        "user_id": user_id, "message_id": message_id, "attachment_id": attachment_id,
    }
