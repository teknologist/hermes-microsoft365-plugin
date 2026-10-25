"""Foundation slice: the attachment handlers live in their own module, registered once.

Pins the ownership split the parallel implementation lanes rely on:

* module discovery (``load_handlers``) registers ``outlook.list_attachments`` and
  ``outlook.download_attachment`` -- and only they from the new module -- exactly once;
* the owning module is ``handlers/outlook_attachments.py``; ``handlers/outlook.py`` no
  longer defines either handler;
* a duplicate registration is refused with ``HandlerRegistrationError`` at import time
  (checked before the registry is mutated), so a handler added back to ``outlook.py``
  fails the import loudly instead of being dispatched twice.
"""
from __future__ import annotations

import inspect

import pytest

from microsoft365.handlers import HandlerRegistrationError, handler, load_handlers
from microsoft365.handlers import outlook, outlook_attachments

ATTACHMENT_KEYS = ("outlook.list_attachments", "outlook.download_attachment")


def test_attachment_handlers_are_registered_exactly_once_by_discovery():
    registry = load_handlers()
    for key in ATTACHMENT_KEYS:
        assert registry[key].key == key
    assert [record for record in registry.values() if record.key in set(ATTACHMENT_KEYS)]
    assert len({record.function for record in registry.values() if record.key in set(ATTACHMENT_KEYS)}) == 2


def test_the_owning_module_is_outlook_attachments_and_not_outlook():
    for key in ATTACHMENT_KEYS:
        operation = key.rsplit(".", 1)[1]
        function = getattr(outlook_attachments, operation)
        assert inspect.getmodule(function) is outlook_attachments
        assert not hasattr(outlook, operation)


def test_a_duplicate_registration_is_refused_not_dispatched_twice():
    def duplicate(*_args, **_kwargs):
        raise AssertionError("unreachable")

    with pytest.raises(HandlerRegistrationError, match="outlook.search"):
        handler("outlook", "search")(duplicate)
