# Development and architecture

## Setup

Use Python 3.13 (packaging permits >=3.11,<3.14), `uv`, and a compatible Hermes source checkout. The plugin uses Hermes APIs but does not vendor or automatically install Hermes core.

From the repository root:

```sh
uv venv --python 3.13 .venv
uv sync --frozen --extra test
export HERMES_SOURCE="/absolute/path/to/hermes-agent"
make check
```

Reuse a suitable existing `.venv`; do not rebuild it while parallel workers run tests. `HERMES_SOURCE` must contain `agent/` and `hermes_cli/`. A typical source installation is `$HOME/.hermes/hermes-agent`, but inspect the machine. If Hermes is already importable in the test environment, omit the variable. CI pins a specific host checkout in `.github/workflows/ci.yml`.

`tests/__init__.py` prevents Hermes' test package from shadowing local fixtures. The Makefile puts this checkout first on `PYTHONPATH`. In worktrees sharing an editable environment, verify `microsoft365.__file__` points into the owning checkout.

## Commands

| Command | Purpose |
| --- | --- |
| `make test` | Full offline suite |
| `make test TEST_ARGS="tests/test_mailbox_allowlist.py -q"` | Focused regression |
| `make check` | Lock consistency, byte compilation, full suite with 90% coverage gate, diff whitespace |
| `make validate` | Hermes validation/doctor of this checkout; does not install |
| `make build` | Wheel in ignored `dist/` |

Override `PYTHON`, `UV` or `HERMES` for explicit executable paths. Test temporary files live under ignored `.hermes/test-tmp`. Never put synthetic attachment files in the real profile download root. There is no configured mandatory formatter/type checker; do not invent a lint gate or report one as passed.

## Runtime map

| Area | Responsibility |
| --- | --- |
| `contract.py` | Operation inventory, auth/execution statuses, settings and resource boundaries |
| `validation.py` | Schemas and cross-field checks before client construction |
| `registration.py` | Conditional tools, hook, final dispatch and host-facing result boundary |
| `client.py` | Execution-time scoped secret lookup and application client credentials |
| `handlers/` | Discovered handler registry; Outlook search/read and attachment handlers have separate modules |
| `sdk_contract.py` | Independent expectations using real generated SDK builders |
| `execution.py`, `errors.py` | Async execution and sanitized error taxonomy |
| `paging.py`, `results.py` | Bounded validated continuations and result normalization |
| `attachments.py` | Profile-scoped private/exclusive artifact writes and cleanup |
| `files.py` | Drive transfer/path safety |
| `dashboard/plugin_api.py`, `desktop/plugin.js` | Configuration UI/backend |
| `plugin.yaml`, `preflight.py` | Host manifest and local readiness |

Paths above are relative to `microsoft365/`. Reuse these seams instead of adding another Graph client or authorization boundary. Do not let arbitrary result dictionaries bypass redaction or size limits.

## Tests to consult

- Boundary/config: `test_mailbox_allowlist.py`, `test_validation.py`, `test_dashboard_api.py`, `test_contract.py`, `test_registration_schemas.py`.
- Search: `test_outlook_search.py`, `test_handlers_outlook_calendar.py`, `test_paging.py`, `test_sdk_contract.py`.
- Attachments: `test_outlook_attachments.py`, `test_handler_ownership_split.py`, `test_paging.py`. There is no `tests/test_results.py`.
- Host: `test_probe_compatibility.py`, `test_registration_secrets.py`, `test_execution.py`, `test_manifest.py`. Offline host probes do not replace an end-to-end invocation through the installed host.

Use synthetic generated SDK models and strict transport doubles. Assert independent expected requests/output, byte integrity and zero credential/client calls for rejected inputs. Preserve false booleans, timezone-aware bounds and truthful truncation. Inventory updates must add real operations, not weaken invariants.

## Compatibility

Read exact pins from `pyproject.toml`, `plugin.yaml` and `uv.lock`; align CI assertions when changing pins. The next host dependency rebuild may reject pins even if an existing process imports cached packages successfully. Repair only a proven conflict, sync and rerun tests.

A Python 3.14 Hermes bootstrap test is not proof that this package's Python <3.14 packaging permits installation there. Record interpreter, imported plugin path and SDK versions. Separate source tests, host admission, configured dispatch, installed state and tenant verification.
