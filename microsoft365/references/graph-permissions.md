# Microsoft Graph permission matrix

This is a conservative pre-release matrix for local administrative reporting. It does not assert tenant consent or remote success. Runtime metadata in `microsoft365/contract.py` is authoritative and tested against the complete operation list.

| Service / operation | Application role(s) | Application status | Delegated status |
|---|---|---|---|
| Outlook search/read/list_attachments/download_attachment | `Mail.Read` | supported | not implemented |
| Outlook draft | `Mail.ReadWrite` | supported | not implemented |
| Outlook send | `Mail.Send` | supported | not implemented |
| SharePoint search/read | `Sites.Read.All` | supported | not implemented |
| SharePoint download/upload | `Files.Read.All` / `Files.ReadWrite.All` | supported | not implemented |
| OneDrive search/read/download/upload | `Files.Read.All` / `Files.ReadWrite.All` | supported | not implemented |
| Calendar search/create/update | `Calendars.Read` / `Calendars.ReadWrite` | supported | not implemented |
| Teams list teams/channels | `Team.ReadBasic.All` / `Channel.ReadBasic.All` | supported | not implemented |
| Teams message search | none claimed | unsupported auth mode | requires delegated implementation |
| Teams send message | none claimed | unsupported auth mode | requires delegated implementation |
| To Do (5 operations) | endpoint-specific, see "Microsoft To Do endpoint claims" | reads supported, writes `not_verified` | not implemented |
| Planner (6 operations) | endpoint-specific, see "Planner endpoint claims" | contract verified offline, not executable | not implemented |

## Outlook attachment permission claims

`outlook.list_attachments` uses `GET /users/{user_id}/messages/{message_id}/attachments`.
`outlook.download_attachment` uses `GET /users/{user_id}/messages/{message_id}/attachments/{attachment_id}`
for metadata and selected file content. Both require application `Mail.Read` as the
least-privileged permission documented by Microsoft for message attachments. No
`Mail.ReadWrite`, `Mail.Send` or delegated shared-mail scope is needed by this
application-mode read-only workflow. Tenant consent and mailbox authorization remain
separate prerequisites; the local allowlist cannot grant either.

Official endpoint references consulted for this extension:

- https://learn.microsoft.com/en-us/graph/api/message-list-attachments
- https://learn.microsoft.com/en-us/graph/api/attachment-get

Claims remain `documented_not_verified`: documentation and offline SDK tests are not
live tenant evidence. Listing requests metadata only; downloading supports ordinary
file attachments, not item/reference attachments. Search is invoice-candidate discovery,
not automatic invoice classification, OCR, reconciliation or financial submission.

## Planner endpoint claims

Planner is reported one endpoint at a time. The blanket `Planner list/read` and `Planner create/update`
rows this matrix used to carry are gone: `Tasks.Read.All` and `Tasks.ReadWrite.All` are not a single
claim, each endpoint carries its own. `microsoft365/contract.py` declares one `PlannerEndpoint` per row
below and derives each operation's permission tuple, container, required identifiers and required
headers from those endpoints, so a claim cannot drift from the endpoint it belongs to, and
`tests/test_planner_contracts.py` fails if this table and the registry disagree.

| Operation | Endpoint | Application role | Permission claim | Implementation status | Recorded endpoint reference |
|---|---|---|---|---|---|
| `planner.list_plans` | `GET /planner/plans` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/planner-list-plans |
| `planner.list_plans` | `GET /groups/{group_id}/planner/plans` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/planner-list-plans |
| `planner.list_buckets` | `GET /planner/plans/{plan_id}/buckets` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/planner-list-buckets |
| `planner.list_tasks` | `GET /planner/plans/{plan_id}/tasks` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks |
| `planner.read` | `GET /planner/tasks/{planner_task_id}` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/plannertask-get |
| `planner.create_tasks` | `POST /planner/tasks` | `Tasks.ReadWrite.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/planner-post-tasks |
| `planner.update_tasks` | `PATCH /planner/tasks/{planner_task_id}` | `Tasks.ReadWrite.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/plannertask-update |

### Container semantics (decision recorded by WP2)

A Planner request is only valid against one container, and the identifiers that identify that container
must be non-blank before anything is sent:

- `planner.list_plans` accepts an **optional** `group_id`. Omitted (or `None`), it reads the tenant-wide
  route `GET /planner/plans`; supplied, it reads the group-owned route
  `GET /groups/{group_id}/planner/plans`. A blank `group_id` is refused -- it never falls back silently
  to the tenant-wide route.
- `planner.list_buckets` and `planner.list_tasks` index by **plan**: `plan_id` is a required path
  identifier, and the path has no calling form without it.
- `planner.read` and `planner.update_tasks` index by **task**: `planner_task_id` is a required path
  identifier.
- `planner.create_tasks` carries both container identifiers in the **request body** (`planId`,
  `bucketId`) together with the task title: creation is bucket-scoped, and the path `POST /planner/tasks`
  carries no identifier at all.
- `planner.update_tasks` additionally requires an `If-Match` ETag header. The other endpoints here
  (`planner.list_plans`, `planner.list_buckets`, `planner.list_tasks`, `planner.read`,
  `planner.create_tasks`) declare no required header and send no `If-Match`.

`validate_planner_identifiers()` in `microsoft365/contract.py` refuses a blank or absent required
identifier -- and a blank optional one -- naming the endpoint the identifier belongs to, before any
client, approval or token work happens.

### Claim legend

- `documented_not_verified` -- the application role recorded for that endpoint. No endpoint page was
  re-read in this offline milestone and no tenant has been contacted (rule R10; remote verification is
  WP16), so the claim is **not** verified. Roles on rows without a recorded reference are inherited from
  this matrix's earlier revision and carry no page citation yet; every claim must be recomputed from its
  endpoint's permission table before the catalog release.
- `verified` -- reserved. Only for a claim recomputed against the endpoint reference **and** confirmed on
  a test tenant. No Planner claim may use it yet.

### Implementation status legend

- `implemented` -- every endpoint the operation uses is pinned by a strict offline request-contract test and a registered handler. Reads may be executable when explicitly included in the registry set; writes remain withheld until CORE-1/CORE-2.

## Microsoft To Do endpoint claims

To Do is reported one endpoint at a time. The blanket `To Do list/search/read` and
`To Do create/update` rows this matrix used to carry are gone: the read role and the write role
are not one claim, and neither is a To Do-wide fact. `microsoft365/contract.py` declares one
`TodoEndpoint` per row below and derives each operation's permission tuple, container, required
identifiers, required headers, contract cases and claim statuses from those endpoints, so a claim
cannot drift from the endpoint it belongs to. `tests/test_todo_contracts.py` fails if this table
and the registry disagree.

| Operation | Endpoint | Application role | Permission claim | Implementation status | Recorded endpoint reference |
|---|---|---|---|---|---|
| `todo.list_task_lists` | `GET /users/{user_id}/todo/lists` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/todo-list-lists |
| `todo.search` | `GET /users/{user_id}/todo/lists/{todo_list_id}/tasks` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks |
| `todo.read` | `GET /users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}` | `Tasks.Read.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/todotask-get |
| `todo.create_tasks` | `POST /users/{user_id}/todo/lists/{todo_list_id}/tasks` | `Tasks.ReadWrite.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks |
| `todo.update_tasks` | `PATCH /users/{user_id}/todo/lists/{todo_list_id}/tasks/{todo_task_id}` | `Tasks.ReadWrite.All` | `documented_not_verified` | `implemented` | https://learn.microsoft.com/en-us/graph/api/todotask-update |

The `sdk_contract` case of each row, in the same order, is `todo_lists`, `todo_tasks`, `todo_read`,
`todo_create_tasks` and `todo_update_tasks`. Only the task listing carries a recorded reference; the
other four say `not recorded` rather than naming a page that was not re-read in this offline
milestone. To Do addresses its container exclusively by path, so `user_id` -- and `todo_list_id`,
and `todo_task_id` for the item endpoints -- must be non-blank before a request is built.

### Application-mode write support is `not_verified`, and not promoted

`todo.create_tasks` and `todo.update_tasks` report the application-mode support status
`not_verified` in the administrative registry (`operation_status("application", "todo", …)`) and
stay **non-executable**. The status is derived from the endpoint table above by
`todo_write_support()`, never asserted, and the recorded reason is:

> the application permission of this write is unverified: no endpoint reference was re-read offline and no test tenant was contacted (R10; remote verification is WP16)

`ApplicationWriteSupport` refuses to construct a `supported` status without recorded evidence, and
`TodoEndpoint` refuses a `verified` permission claim without recorded evidence, so this cannot be
promoted by editing one string. A promoted claim is also not an execution: `supported` reports
authentication-mode support only, and `executable` remains whatever the registry says -- today
`False` for every To Do operation, because no handler exists.

### Task field policy

A To Do write may carry exactly the fields the endpoint contract records -- `title`, `body`,
`due_date_time` and `status` -- and nothing else. The generated `TodoTask` model declares more
properties than this milestone can re-read against the endpoint reference offline, so the writable
set stays deliberately minimal and fail-closed: any other field is refused with the case name
instead of being sent and silently dropped by Graph. `subject` is refused in particular because no
To Do model has that property -- To Do uses `title` -- and the generated model's own deserializers
are the authority for which `$select` names a task read may ask for. `status` is mapped to the
generated `TaskStatus` member before the body is built, because Kiota's writer emits nothing at all
for an enum field that is not a generated member.

### Filtering by status

The search path sends the pinned `contains(title,'…')` filter and no other. Acceptance of a
`$filter` on `status` is not recorded offline, so no status filter is sent and a status match stays
client-side, limited to the tasks a caller has already retrieved (`todo_task_status_matches()`). A
value the generated `TaskStatus` enum does not declare is refused rather than being dropped from
the request.

### To Do claim legend

- `documented_not_verified` on every To Do row: the application role recorded for that endpoint.
  No endpoint page was re-read in this offline milestone and no tenant has been contacted (R10;
  remote verification is WP16), so no To Do claim is verified.
- `not_verified` for the application-mode write status of `todo.create_tasks` and
  `todo.update_tasks`: the endpoints' own claims are unverified, so their support cannot be
  promoted, and the operations cannot be executed.

Official endpoint references:

- https://learn.microsoft.com/en-us/graph/api/search-query
- https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams
- https://learn.microsoft.com/en-us/graph/api/channel-list
- https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks
- https://learn.microsoft.com/en-us/graph/api/planner-post-tasks
- https://learn.microsoft.com/en-us/graph/api/plannertask-update
- https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks

## WP-D static evidence refresh

The endpoint rows above and the index below were recomputed against official Microsoft Graph endpoint pages. Static documentation is not tenant or token evidence: every permission claim remains `documented_not_verified`, remote verification remains `not_tested`, and no Graph call was made.

| Operation | Application role(s) | Delegated scope(s) | Application mode | Delegated mode | Official endpoint documentation |
|---|---|---|---|---|---|
| `teams.list_teams` | `Team.ReadBasic.All` | `Team.ReadBasic.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/user-list-joinedteams |
| `teams.list_channels` | `Channel.ReadBasic.All` | `Channel.ReadBasic.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/channel-list |
| `teams.search_messages` | `none claimed` | `Chat.Read, ChannelMessage.Read.All` | `unsupported_auth_mode` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/search-query |
| `teams.send_messages` | `none claimed` | `ChannelMessage.Send` | `unsupported_auth_mode` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/channel-post-messages |
| `WP-D planner.list_plans` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/planner-list-plans |
| `WP-D planner.list_buckets` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/planner-list-buckets |
| `WP-D planner.list_tasks` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/plannerplan-list-tasks |
| `WP-D planner.read` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/plannertask-get |
| `WP-D planner.create_tasks` | `Tasks.ReadWrite.All` | `Tasks.ReadWrite` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/planner-post-tasks |
| `WP-D planner.update_tasks` | `Tasks.ReadWrite.All` | `Tasks.ReadWrite` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/plannertask-update |
| `WP-D todo.list_task_lists` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/todo-list-lists |
| `WP-D todo.search` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks |
| `WP-D todo.read` | `Tasks.Read.All` | `Tasks.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/todotask-get |
| `WP-D todo.create_tasks` | `Tasks.ReadWrite.All` | `Tasks.ReadWrite` | `not_verified` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/todotasklist-post-tasks |
| `WP-D todo.update_tasks` | `Tasks.ReadWrite.All` | `Tasks.ReadWrite` | `not_verified` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/todotask-update |
| `outlook.search` | `Mail.Read` | `Mail.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/user-list-messages |
| `outlook.read` | `Mail.Read` | `Mail.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/message-get |
| `outlook.create_draft` | `Mail.ReadWrite` | `Mail.ReadWrite` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/user-post-messages |
| `outlook.send` | `Mail.Send` | `Mail.Send` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/user-sendmail; https://learn.microsoft.com/en-us/graph/api/message-send |
| `calendar.search` | `Calendars.Read` | `Calendars.Read` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/user-list-events |
| `calendar.create_events` | `Calendars.ReadWrite` | `Calendars.ReadWrite` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/user-post-events |
| `calendar.update_events` | `Calendars.ReadWrite` | `Calendars.ReadWrite` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/event-update |
| `sharepoint.search` | `Sites.Read.All` | `Sites.Read.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/site-list; https://learn.microsoft.com/en-us/graph/api/driveitem-search; https://learn.microsoft.com/en-us/graph/api/site-get |
| `sharepoint.read` | `Sites.Read.All` | `Sites.Read.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-get |
| `sharepoint.download_files` | `Files.Read.All` | `Files.Read.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-get-content |
| `sharepoint.upload_files` | `Files.ReadWrite.All` | `Files.ReadWrite.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-put-content |
| `onedrive.search` | `Files.Read.All` | `Files.Read.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-search |
| `onedrive.read` | `Files.Read.All` | `Files.Read.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-get |
| `onedrive.download_files` | `Files.Read.All` | `Files.Read.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-get-content |
| `onedrive.upload_files` | `Files.ReadWrite.All` | `Files.ReadWrite.All` | `supported` | `not_implemented` | https://learn.microsoft.com/en-us/graph/api/driveitem-put-content |

### Static limitations

- Official endpoint pages establish documented least-privileged roles/scopes and routes, but cannot establish tenant admin consent, token contents, conditional access, service availability, or successful authorization.
- Application roles and delegated scopes are different permission systems. The matrix does not derive one from the other.
- Teams message search and send remain claim-free and unsupported in application mode. Delegated mode is recorded but not implemented.
- To Do application writes remain `not_verified` and non-executable. No unsupported or unverified operation is promoted by this refresh.
