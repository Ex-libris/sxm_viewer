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

from typing import Callable, Dict, Iterable


class UndoJournal:
    def __init__(self, limit: int = 500):
        self._tokens: list = []
        self._limit = int(limit)

    def record(self, source: str) -> None:
        self._tokens.append(str(source))
        if len(self._tokens) > self._limit:
            del self._tokens[: len(self._tokens) - self._limit]

    def clear(self) -> None:
        self._tokens.clear()

    def __len__(self) -> int:
        return len(self._tokens)

    def undo(self, handlers: Dict[str, Callable[[], bool]]) -> bool:
        """Undo the newest recorded action. Returns True if one was undone."""
        while self._tokens:
            handler = handlers.get(self._tokens.pop())
            if handler is None:
                continue
            try:
                if handler():
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
