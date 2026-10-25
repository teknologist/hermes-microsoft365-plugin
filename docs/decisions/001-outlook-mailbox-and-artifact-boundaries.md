# ADR-001: Outlook mailbox allowlist and local attachment artifacts

## Status

Accepted boundary decisions. Transfer caps and buffered-versus-streaming implementation
remain provisional operational choices, not promises of a hard memory/network limit.

## Context

Invoice discovery must cover the main mailbox and the Billings shared mailbox without
removing resource isolation from other Microsoft 365 services. Large model-visible
base64 responses are unsuitable as the default way to inspect invoice files. The plugin
already has generated Graph builders, validation, execution and pagination boundaries.

## Decision

- Extend the same Outlook tool with opt-in `search`, `read`, `list_attachments` and
  `download_attachment` reads. Keep all external write operations disabled.
- A nonempty `allowed_mailbox_ids` replaces the configured `user_id` boundary only for
  Outlook. Each entry is a verbatim GUID-shaped object ID; reject malformed strings,
  whitespace, wildcards, UPNs and case-variant duplicates. Exact call matching preserves
  configured spelling. Do not implicitly include `user_id`.
- Check the target before credentials and again against final dispatch arguments.
  Omitted/empty allowlists retain legacy behavior, including its unconfigured-user case.
  Calendar, To Do and other services retain their existing boundary.
- Do not expose configured IDs in preflight/errors. Preserve omitted allowlist settings
  during dashboard saves, including saves from older clients.
- Separate candidate discovery from classification and financial actions. Subject query
  is containment, not attachment search. All downloaded content is untrusted data.
- Return local artifacts by default under the active profile's fixed download root.
  Generate unique names, reject symlink traversal, use restrictive permissions and
  exclusive creation, and clean up failed writes. Original filename and MIME are
  metadata, not trusted paths or proof of format.
- Return hashes and source mailbox/message/attachment IDs. Identical content from
  distinct messages retains distinct provenance; do not silently deduplicate or delete.
- Reuse the existing generated SDK and execution seam. No second Graph client,
  database, scheduler, OCR service or financial submission path is introduced.

## Alternatives rejected

- Disabling the resource boundary or accepting arbitrary tenant mailboxes expands
  authority beyond the two-mailbox use case.
- Applying the Outlook allowlist to every service silently changes established access
  decisions outside this workflow.
- Remote-filename output paths permit traversal/collision problems; arbitrary local
  paths unnecessarily expand write authority.
- Default base64 responses risk context bloat or corrupt truncation. Artifact paths
  keep tool results compact and bytes available for explicit inspection.

## Consequences and verification

Tenant-side application access restrictions and consent are still required. Local
allowlisting does not confer authorization. Operators must configure real object IDs;
legacy `user_id` compatibility is deliberately preserved rather than normalized.

Current ordinary file support is limited to 10 MiB decoded for local artifacts, with a
smaller explicit base64 mode. Metadata prechecks and byte-length rechecks do not bound
SDK buffering. Streaming/hard transport limits require separate design and evidence.
Item/reference attachments are unsupported. Search/list pagination is bounded and a
cap hit is not evidence of a complete mailbox scan.

Tests use synthetic data and generated SDK request contracts. They cannot prove live
tenant permissions, remote filter behavior or real invoice classification. Installation,
publication and live mailbox smoke tests are separate authorization gates.
