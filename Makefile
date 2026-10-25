PYTHON ?= .venv/bin/python
UV ?= uv
HERMES ?= hermes
TEST_ARGS ?= tests/ -q
export PYTHONPATH := $(CURDIR)$(if $(HERMES_SOURCE),:$(HERMES_SOURCE))$(if $(PYTHONPATH),:$(PYTHONPATH))

.PHONY: test check validate build

test:
	@mkdir -p .hermes/test-tmp
	TMPDIR="$(CURDIR)/.hermes/test-tmp" $(PYTHON) -m pytest $(TEST_ARGS)

check:
	$(UV) lock --check
	$(PYTHON) -m compileall -q microsoft365
	$(MAKE) test TEST_ARGS="tests/ -q --cov=microsoft365 --cov-report=term --cov-fail-under=90"
	git diff --check

validate:
	$(HERMES) plugins validate "$(CURDIR)/microsoft365"
	$(HERMES) plugins doctor "$(CURDIR)/microsoft365" --ci

build:
	$(UV) build --wheel --out-dir dist
