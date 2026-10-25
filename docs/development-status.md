# Development handoff — 2026-10-25

Dated evidence, not a claim that live integration works. Recheck Git and installed state before continuing. Base implementation commit: `fd6d842` (`feat/invoice-mail-workflow`). Real tenant/mailbox identifiers and credentials are intentionally omitted.

## Delivered implementation

- Outlook-only mailbox allowlist, legacy compatibility, dashboard round-trip and final-argument validation before credentials.
- Structured search filters, subject containment, compact projections and bounded pagination.
- Metadata-only attachment listing; private generated local artifacts with hash and provenance.
- Supported decoded size: 10 MiB for artifacts, 2048 bytes for explicit base64. These do not bound SDK JSON buffering. Item/reference attachment downloads are refused.
- Full suite after Azure Identity alignment: **1194 passed**. Earlier coverage run: **91.90%**. Installed validation/doctor and actual configured preflight passed. These are not proof of live mailbox access.
- One effective user plugin was observed. A shadowed source-tree copy was left untouched. Changing this checkout does not update the installed copy.

## Priority 1: host result contract

The real main-mailbox search returned:

```text
Tool handler returned unsupported result type: dict
error_type: tool_result_contract
```

Root cause is unresolved. This is not evidence that no matching messages exist. Inspect `registration.py`, `results.py` and the installed host's tool-dispatch contract. Reproduce a synthetic successful Graph response through the actual host invocation path; offline tests missed this boundary. Locate the mismatch before choosing serialization changes; do not add a generic normalization/redaction bypass.

Acceptance: demonstrate RED/GREEN on the host-facing regression, preserve items/pagination/artifact provenance and sanitized errors, pass the full suite, then obtain usable results from an explicitly authorized live read. Do not patch Hermes core from this repository.

## Priority 2: shared mailbox identity and access

After local allowlist validation succeeded, both the dated query and a minimal unfiltered shared-mailbox search returned:

```text
not_found: the requested microsoft 365 item does not exist or is not visible
```

That does not establish whether the ID is incorrect, the mailbox is provisioned differently, or application access is restricted. It does not establish that delegated access will work. The user's hypothesis is that Billings may be accessible through delegated access; investigation stopped before verification or an architectural decision.

Next steps:

1. Confirm mailbox identity/type and tenant-side application access with an authorized administrator. Keep real IDs out of committed evidence.
2. Preserve safely available status/code/request-correlation details without logging tokens or message contents. Distinguish local validation, Graph errors and host-result errors.
3. Verify current Microsoft documentation for shared-mail delegated scopes and the signed-in user's mailbox permissions. Application roles and delegated scopes are different models.
4. If delegated OAuth is needed, propose an ADR covering interactive consent, account selection, scoped secure token caching/refresh and tenant/account isolation. Obtain approval before implementing a new flow or initiating consent.

`client.py` uses `ClientSecretCredential` and rejects modes other than application. Changing `authentication_mode` to `delegated` is not sufficient: contract entries mark that mode unimplemented. Preserve the allowlist, read-only gates and existing application mode in future work.

Acceptance: identify the resource and chosen auth model with evidence; demonstrate an authorized read and rejection of an out-of-scope target. Never put production messages/invoices in fixtures.

## Pending end-to-end acceptance

The requested invoice scan was October 12–18, 2026 in Paris time:
`[2026-10-12T00:00:00+02:00, 2026-10-19T00:00:00+02:00)`.
No usable email/PDF results were returned. Once access and result handling work, search both explicitly allowlisted mailboxes, inspect PDF attachment metadata and report coverage/truncation. Candidates are not confirmed invoices. Download only as authorized; never mutate mail or submit financial records.

Local historical plans/reports under `.hermes/` are ignored and may not exist in a fresh clone. This handoff, the source/tests and ADR are the portable starting point. Follow the [deployment runbook](runbooks/hermes-deployment.md) for installation, restart and rollback. Publication and live verification are separate gates.
