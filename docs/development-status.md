# Development handoff — 2026-10-25

Dated evidence. Recheck Git and installed state before continuing. Base implementation commit: `fd6d842` (`feat/invoice-mail-workflow`). Real tenant/mailbox identifiers and credentials are intentionally omitted.

## Status of the two priorities below: both resolved, live

The two blockers recorded in the previous handoff were reproduced, diagnosed and fixed, and
the four Outlook reads now run end to end against the real tenant through the registered
host tool path. Three defects were found, two of which the offline suite could not see:

1. **Host result contract.** A registered Outlook tool returned the handler's `dict`. The
   host's registry accepts a string (or the multimodal envelope) and turns anything else
   into `tool_result_contract`, so *every* successful Graph read was reported as a failed
   call. Serialization now happens at the registration boundary (`registration.tool_result`),
   where `microsoft365_preflight` already did it; handlers keep their contract-pinned `dict`
   and nothing is re-bounded or re-redacted.
2. **"Billings" is not a shared mailbox.** `GET /directoryObjects/{id}` identifies it as a
   mail-enabled **unified group**. `/users/{group}` is a directory read an application
   identity is refused, and the group's SMTP address is not a user principal, so the user
   route could only ever answer `not_found` — which is exactly what it did. Group mail lives
   at `/groups/{id}/threads` and the posts inside a thread, under `Group.Read.All` (already
   granted; the app's roles are `Group.Read.All` and `Mail.Read`). No delegated flow is
   needed and none was implemented. Attachments come back only through `$expand=attachments`
   on the posts: the per-post `/attachments` collection answers `ErrorAccessDenied` 403 for
   an application identity.
3. **Two latent defects in the existing user-mailbox download**, found only by running it
   live. `$select=contentBytes` is a 400 from Graph (`contentBytes` is on
   `microsoft.graph.fileAttachment`, not on the base `attachment` type) — now selected
   through the OData type cast. And the declared `size` counts the MIME-encoded attachment
   while `contentBytes` is decoded, so the equality check refused every genuine download —
   now verified as an upper bound, with over-size still refused.

### Live evidence (2026-10-25, registered host tool path, read-only)

| Check | Result |
| --- | --- |
| `search` user mailbox | succeeded, 3 messages, host type `str` |
| `read` user message | succeeded |
| `search` group mailbox | succeeded, 5 conversation threads, `mailbox_type: group` |
| `read` group conversation | succeeded, 1 post, sender address projected |
| `list_attachments` group | succeeded, 2 PDFs, metadata only, no `contentBytes` in the result |
| `download_attachment` group | succeeded, 28120-byte PDF, `%PDF-` magic, SHA-256 matches, mode 0600 |
| `list_attachments` user | succeeded |
| `download_attachment` user | succeeded, 74526-byte artifact, hash verified on disk |
| non-allowlisted mailbox | blocked at the hook, before any credential work |

Offline suite: **1223 passed**, coverage **92.22%** (`make check`). No mail was mutated; no
message content, tenant identifier or credential is recorded here or in the tests.

### Remaining gaps, stated rather than glossed

- **A true shared mailbox was not live-verified**: none is configured in
  `allowed_mailbox_ids`, and the application's roles do not permit enumerating the directory
  to find one. It takes the same `/users/{id}` route under the same `Mail.Read` role as the
  personal mailbox that *was* verified, and is covered by the offline suite — but that is a
  structural argument, not tenant evidence. Add the mailbox's object ID to the allowlist to
  confirm it.
- **Deployment has not happened.** The installed copy under the active profile is still the
  old one; this repository's working tree is what was verified. Installation and restart
  remain separate authorization gates — follow the
  [deployment runbook](runbooks/hermes-deployment.md).

## Delivered implementation

- Outlook-only mailbox allowlist, legacy compatibility, dashboard round-trip and final-argument validation before credentials.
- Structured search filters, subject containment, compact projections and bounded pagination.
- Metadata-only attachment listing; private generated local artifacts with hash and provenance.
- Supported decoded size: 10 MiB for artifacts, 2048 bytes for explicit base64. These do not bound SDK JSON buffering. Item/reference attachment downloads are refused.
- Mailbox-kind routing: personal, shared and Microsoft 365 group mailboxes behind one unchanged tool surface (`microsoft365/mailbox.py`, `microsoft365/group_mail.py`).
- Host-facing result serialization at the registration boundary (`registration.tool_result`).
- Full suite: **1223 passed**, coverage **92.22%**.
- One effective user plugin was observed. A shadowed source-tree copy was left untouched. Changing this checkout does not update the installed copy.

## Resolved: host result contract (was Priority 1)

The real main-mailbox search returned:

```text
Tool handler returned unsupported result type: dict
error_type: tool_result_contract
```

Root cause: `tools/registry.py::_normalize_handler_result` in the host accepts a string or
the multimodal envelope and converts anything else into that error. The registered
`service_handler` returned the handler payload directly, while `microsoft365_preflight`
already serialized its own. The fix is one boundary function, `registration.tool_result`,
applied where the host is handed the result; no normalization or redaction is bypassed and
an already-serialized refusal is passed through untouched.

Regression: `tests/test_group_mailbox.py` asserts the host-facing answer is a `str` that
decodes to the payload, that serialization adds no second pass, and that a string result is
handed over unchanged. The existing registered-tool tests were asserting the rejected
`dict` shape and now decode, which is the RED→GREEN record.

## Resolved: "Billings" is a group mailbox, not a shared one (was Priority 2)

`GET /directoryObjects/{id}` answers `#microsoft.graph.group`, mail-enabled, `groupTypes:
["Unified"]`. That settles the three hypotheses the previous handoff could not separate: the
ID was correct, the mailbox is provisioned differently, and application access was never the
problem. `/users/{group}` returns `Authorization_RequestDenied` for the directory read and
`not_found` for mail, because a group mailbox is simply not addressable there.

The implemented route is `/groups/{id}/threads` and the posts inside a thread, under the
`Group.Read.All` role the application already holds. **No delegated flow is needed**, none
was implemented, and `client.py` still rejects every mode but `application`. The allowlist,
the read-only gates and the application mode are unchanged.

Verified service constraints, each reflected in the code rather than assumed:

- the thread collection rejects every `$filter` but `isLocked` (`ConversationFilterOther`),
  so narrowing is applied to the fetched page and reported as such;
- the per-post `/attachments` collection answers `ErrorAccessDenied` 403 for an application
  identity, while `$expand=attachments` on the posts returns metadata *and* `contentBytes`;
- an attachment's declared `size` includes MIME encoding overhead and exceeds the decoded
  `contentBytes` length.

## Pending end-to-end acceptance

The requested invoice scan was October 12–18, 2026 in Paris time:
`[2026-10-12T00:00:00+02:00, 2026-10-19T00:00:00+02:00)`.
The read path now works, so this scan can be run: search both allowlisted mailboxes,
inspect PDF attachment metadata and report coverage/truncation. On the group mailbox the
date bounds are applied to the fetched page, so a truncated traversal must be partitioned
and `threads_examined` / `filter_mode` reported honestly. Candidates are not confirmed
invoices. Download only as authorized; never mutate mail or submit financial records.

Local historical plans/reports under `.hermes/` are ignored and may not exist in a fresh clone. This handoff, the source/tests and ADR are the portable starting point. Follow the [deployment runbook](runbooks/hermes-deployment.md) for installation, restart and rollback. Publication and live verification are separate gates.
