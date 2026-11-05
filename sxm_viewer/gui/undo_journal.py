"""One chronological undo order for the main window (Qt-free).

The app keeps several independent undo stacks (the main preview's display
history, quick-crop entries, per-file adjustments, collection edits). Each
push also records its *source* here, so Ctrl+Z can undo the most recent
action wherever it happened instead of draining the stacks in a fixed
priority order.

Tokens and stack entries are popped together from the newest end, so they
stay matched even when a stack trims its oldest entries. A token whose
handler finds nothing to undo (its entry was removed some other way) is
skipped.
"""
from __future__ import annotations

import weakref
from typing import Callable, Dict, Iterable


class UndoJournal:
    """Tokens are ``(source, target)``. ``target`` names *which* object holds
    the entry when one source has several holders (e.g. the main preview and
    each pop-up canvas); it is held weakly, so a closed pop-up's tokens are
    skipped instead of keeping it alive."""

    def __init__(self, limit: int = 500):
        self._tokens: list = []
        self._limit = int(limit)

    def record(self, source: str, target=None) -> None:
        ref = None
        if target is not None:
            try:
                ref = weakref.ref(target)
            except TypeError:
                ref = lambda t=target: t  # noqa: E731 - not weak-referenceable
        self._tokens.append((str(source), ref))
        if len(self._tokens) > self._limit:
            del self._tokens[: len(self._tokens) - self._limit]

    def clear(self) -> None:
        self._tokens.clear()

    def __len__(self) -> int:
        return len(self._tokens)

    @property
    def sources(self) -> list:
        return [source for source, _ref in self._tokens]

    def undo(self, handlers: Dict[str, Callable[..., bool]]) -> bool:
        """Undo the newest recorded action. Returns True if one was undone.
        Handlers take the token's target (None when recorded without one)."""
        while self._tokens:
            source, ref = self._tokens.pop()
            handler = handlers.get(source)
            if handler is None:
                continue
            target = None
            if ref is not None:
                target = ref()
                if target is None:
                    continue  # its holder is gone
            try:
                if handler(target):
                    return True
            except Exception:
                continue
        return False


class JournaledList(list):
    """A list-backed undo stack whose ``append`` also records ``source``."""

    def __init__(self, journal: UndoJournal, source: str, items: Iterable = ()):
        super().__init__(items)
        self._journal = journal
        self._source = source

    def append(self, item) -> None:  # noqa: D401 - list API
        super().append(item)
        self._journal.record(self._source)
