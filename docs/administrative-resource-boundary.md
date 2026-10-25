# Administrative resource boundary

The default product decision is **single-user binding** when `settings.user_id` is configured. For every operation whose registry metadata requires `user_id`, a call carrying a different user is rejected before argument dispatch, secret lookup, credential creation, or Graph client construction. The rejection is deliberately value-free: tenant, client, and user identifiers are never returned by preflight or boundary errors.

## Decision: explicit Outlook mailbox allowlist

`settings.allowed_mailbox_ids` allows a main mailbox and a shared mailbox without
removing the single-user boundary from Calendar, To Do, or Teams. A nonempty list
replaces the `user_id` boundary **only for Outlook**; `user_id` is not implicitly
added. Matching is literal and exact, with no wildcard semantics or normalization.
Absent/empty lists retain previous behavior, including no configured single-user
pin when `user_id` is empty. Each entry must have the GUID textual shape (8-4-4-4-12 hexadecimal characters),
without whitespace or normalization. Wildcards, UPNs, malformed entries and
case-variant duplicates are rejected; error messages omit configured values. The common
validator enforces the policy before approval, registered dispatch, and direct
`service_tool_handler` dispatch with settings.

Outlook-only preflight accepts a nonempty allowlist instead of a configured
`user_id`; selecting another user-scoped service still requires `user_id`.
Dashboard reads remain redacted, and saves preserve an omitted allowlist so older
clients cannot silently erase the policy. An explicit empty list clears it.

This plugin does not infer tenant-wide access from application permissions. Planner operations are the explicit exception to the *user identifier requirement* because their endpoint metadata does not require `user_id`; that does not grant tenant-wide authorization, consent, or execution.

Preflight derives `tenant_id`, `client_id`, and `user_id` requirements from selected operation metadata. Application mode checks the scoped application secret; delegated mode does not read it. Delegated authentication remains not implemented, and all write operations remain non-executable.
