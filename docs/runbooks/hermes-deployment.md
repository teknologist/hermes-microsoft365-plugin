# Deploy, update and restart in Hermes

This runbook is for an **explicitly authorized** deployment. Documentation changes or a passing test suite do not authorize replacing the installed plugin, editing live settings, restarting services, publishing commits or making tenant calls.

## 1. Establish the target and recovery point

Run from the development repository:

```sh
git status --short
git branch --show-current
git remote -v
hermes --version
hermes plugins show microsoft365
hermes plugins install --help
hermes plugins enable --help
hermes gateway status
```

Resolve the intended active profile before setting paths. The default home is `$HOME/.hermes`; a selected profile may use a different home. Use the same profile for copy, CLI, settings, secrets and restart. Inspect effective plugin discovery: user and source-tree copies may coexist, but only the intended copy should register. Never delete a bundled copy just because it is shadowed.

Set explicit paths in an operator shell (substitute the resolved home):

```sh
REPO="$PWD"
PROFILE_HOME="${HERMES_HOME:-$HOME/.hermes}"
TARGET="$PROFILE_HOME/plugins/microsoft365"
BACKUP="$PROFILE_HOME/plugin-backups/microsoft365-$(date +%Y%m%d-%H%M%S)"
test -f "$REPO/microsoft365/plugin.yaml"
test ! -L "$TARGET"
mkdir -p "$(dirname "$BACKUP")"
if test -d "$TARGET"; then cp -a "$TARGET" "$BACKUP"; fi
```

Backups belong outside plugin discovery roots. Preserve live config/secrets; do not print them or include them in Git. Record the backup and source commit. Stop if the target has unknown local modifications or does not represent the intended profile.

## 2. Verify source and dependency compatibility

```sh
make check
make validate
```

Inspect the current host's dependency requirements, not merely its version banner. This deployment previously encountered an Azure Identity conflict: the current host required 1.26.0 while the plugin pinned 1.25.3. The plugin was automatically disabled by dependency resolution. The current code aligns on 1.26.0; do not assume that value fits every future host.

If a resolver conflict occurs, change only the justified pin, align `plugin.yaml`, `pyproject.toml`, `uv.lock` and CI expectations, run `uv lock`, `uv sync --extra test`, and repeat tests. Do not use `--no-deps` to bypass a broken active environment or edit Hermes core to force compatibility.

## 3. Choose the deployment route

### Local development copy (current workflow)

For the manually managed user plugin directory, copy the **plugin subdirectory**, not the entire repository:

```sh
mkdir -p "$TARGET"
rsync -a --exclude '__pycache__/' --exclude '*.pyc' \
  "$REPO/microsoft365/" "$TARGET/"
hermes plugins enable microsoft365 --no-allow-tool-override
```

This is a local directory deployment, not a catalog installation. No `--delete` is used: inspect retired source files explicitly and remove only confirmed plugin-owned obsolete files after backup, because stale modules can double-register handlers. Do not overwrite unknown user files. Confirm the copied files match source bytes (e.g. SHA-256 comparison excluding caches); a zero exit code from the copy is not sufficient.

`enable` resolves required dependencies. A successful enable does not reload an already-running process or prove it uses those new dependencies. Preserve existing operation flags; change only authorized settings via `hermes config set plugins.entries.microsoft365.settings.<key> <value>`. Never widen mailbox scope or put the client secret in YAML.

### Tracked Git installation / updates

Use this route only when publication and changing installation provenance are authorized. Check the CLI's current help and saved source first. A bare catalog name can select the upstream repository rather than the user's fork.

For a new tracked installation, use the exact approved fork and published immutable commit:

```sh
hermes plugins install https://github.com/teknologist/hermes-microsoft365-plugin.git \
  --ref FULL_40_CHARACTER_PUBLISHED_COMMIT --yes-deps --enable
```

For an existing tracked installation whose saved source is verified:

```sh
hermes plugins update microsoft365
```

Updates may follow a different remote commit; record and verify what was installed. Do not force a source change, remove a manual install, or silently replace local modifications. `hermes plugins adopt` applies to a genuinely self-cloned plugin directory with correct provenance, not this copied subdirectory. Commit authorization does not imply push authorization. If signing fails, ask before making a one-off unsigned commit; retain global signing settings.

## 4. Verify the installed copy

```sh
hermes plugins validate "$TARGET"
hermes plugins doctor "$TARGET" --ci
hermes plugins show microsoft365
```

Confirm source hashes, enabled state, effective manifest/import path and exactly one effective registration. Unconfigured validation probes may show only preflight and warn about conditional service tools; verify actual configured registration independently.

Invoke `microsoft365_preflight` through the actual host tool dispatcher. Require `locally_ready: true`, SDK availability, no configuration errors and the intended selected operations. Never display credential values. A direct Python helper can lack the host's secret scope and report a false missing-secret diagnosis.

## 5. Restart the process that actually serves tools

**A gateway-hosted agent must not restart its own gateway process.** The runtime blocks this because termination propagates to the command before it can finish. Hand the following commands to a separate operator shell outside the gateway; do not evade the guard using background jobs, detached processes or another agent.

For an installed messaging-gateway service, in that external shell:

```sh
hermes gateway status
hermes gateway restart
hermes gateway status
```

Check service status and startup logs; verify the new process actually loaded the intended plugin. If no managed service exists, do not install one as a workaround. Use the existing supervisor/foreground launch method with explicit authorization.

For **Hermes Desktop**, fully quit and reopen the app when dependency changes require it. Restarting the messaging gateway is not proof that Desktop's backend reloaded. For a standalone dashboard or CLI, restart that specific process through its existing launch method. A new chat/schema refresh alone is not proof of dependency replacement.

After the correct restart, refresh tool schemas and run the actual preflight again. If permitted, make one bounded read with the explicit target mailbox ID. Compare saved scope with actual validation; stale configuration is a hypothesis until restart/retry confirms it. Do not bypass a boundary rejection.

## 6. Live smoke test and rollback

Only with separate authorization: a tiny date-window search, metadata listing for a known message, then one small attachment download. Verify local byte count/hash/provenance and unchanged mailbox state. Record sanitized errors and truthful pagination; never claim success from an offline fixture.

For rollback, stop serving the changed copy through the authorized external process-control path, preserve the failed installation for diagnosis, and restore the saved plugin directory to `TARGET`. Re-resolve dependencies with `hermes plugins enable microsoft365 --no-allow-tool-override`, validate the restored copy and restart the correct host externally. Restoring old source may recreate a dependency conflict; stop rather than forcing an incompatible environment. Restore configuration only if this deployment changed it and only from an approved backup.

## Handoff checklist

- Source commit/diff and tests recorded; installation path and backup recorded.
- Installed bytes/effective discovery verified; dependency resolution succeeded; enabled state read back.
- Correct host restarted by an authorized operator; post-restart preflight checked.
- Live success or exact remaining blocker recorded separately; no secrets or production content in Git.

CLI forms above were inspected on the development host. Recheck `--help` on another version. This runbook has not restarted or redeployed any service merely by being written.
