# Contributing

This plugin is pre-release. Changes must preserve every operation in `microsoft365.contract.OPERATIONS`; removing or renaming an operation requires explicit product authorization.

## Local checks

```bash
python -m pytest -q
python -m compileall -q microsoft365
python -m hermes_cli.plugin_validate microsoft365
hermes plugins doctor microsoft365 --ci
detect-secrets scan > secret-scan.json
python -c 'import json; data=json.load(open("secret-scan.json")); assert not data["results"], data["results"]'
```

Use test-driven development for runtime behavior: add a focused failing test, confirm the expected failure, implement the smallest fix, and run the full suite. SDK contract tests must use real generated `msgraph-sdk==1.64.0` builders/models and Kiota request information. Do not replace generated builder chains with permissive `__getattr__` fakes.

Do not use real credentials, tenant data, or Graph network calls in unit tests. Never commit `.env`, tokens, secrets, tenant responses, or generated evidence that claims remote verification.

Use focused conventional commits (`test:`, `feat:`, `fix:`, `docs:`, `chore:`).

## Review and merge

External pull requests require a green `ci-gate` check and maintainer (Code Owner)
approval before merge; see `docs/review-policy.md` for the acceptance criteria and
approval flow.
