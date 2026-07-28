"""Activity log panel: batched append of status messages to a text box.

First extraction in the effort to reduce ``SXMGridViewer``'s god-class
character (see ``docs/refactor/DUPLICATION_INVENTORY.md`` item R1). The
attribute-cohesion analysis (``scripts/analysis/attribute_cohesion.py``)
scored this group at **75% isolation** - the highest in the class - meaning
almost every method touching this state touched nothing else, so it lifts
out with essentially no ripple.

Owns what used to be three ``SXMGridViewer`` attributes
(``activity_log_box``, ``_activity_log_pending``,
``_activity_log_flush_timer``) and three methods
(``_append_activity_log``, ``_flush_activity_log_pending``,
``_on_clear_activity_log``).

**Why batching**: log lines arrive from a signal that can fire hundreds of
times during a folder load, and appending to a ``QPlainTextEdit`` per line
is slow enough to visibly stall the UI. Messages accumulate and flush on a
short single-shot timer instead.
"""
from __future__ import annotations

from datetime import datetime

from .._shared import QtCore, QtWidgets
from .debounce import ACCUMULATE, Debouncer

FLUSH_INTERVAL_MS = 60
MAX_BLOCKS = 500


class ActivityLog(QtCore.QObject):
    """Batched writer for the activity-log text box.

    Construct with the widget it drives; ``append`` is signal-compatible
    with ``log_emitter.message_logged``.
    """

    def __init__(self, box: QtWidgets.QPlainTextEdit, parent=None):
        super().__init__(parent)
        self._box = box
        # Ordered payload, so entries are kept here rather than in the
        # Debouncer's set (which is unordered by design); the debouncer
        # only owns the timing.
        self._pending: list[str] = []
        self._debounce = Debouncer(self.flush, interval_ms=FLUSH_INTERVAL_MS,
                                   mode=ACCUMULATE, parent=self)
        try:
            self._box.document().setMaximumBlockCount(MAX_BLOCKS)
        except Exception:
            pass

    @property
    def widget(self):
        return self._box

    def append(self, message: str):
        """Queue a timestamped message; flushes on the next timer tick."""
        if self._box is None:
            return
        entry = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
        self._pending.append(entry)
        # Payload lives in self._pending (order matters); pass a sentinel
        # so the debouncer knows something is outstanding.
        self._debounce.schedule(("pending",))

    def flush(self):
        """Write everything queued, in one append."""
        self._debounce.take()
        if self._box is None:
            self._pending = []
            return
        pending, self._pending = self._pending, []
        if not pending:
            return
        # RuntimeError = Qt "wrapped C++ object deleted" during teardown;
        # that is the only failure mode worth surviving here, so catch it
        # narrowly rather than swallowing every exception.
        try:
            self._box.appendPlainText("\n".join(pending))
            self._scroll_to_end()
        except RuntimeError:
            return
        # Keep the log visibly moving during long synchronous work (folder
        # loads) without letting the user click mid-operation.
        QtWidgets.QApplication.processEvents(
            QtCore.QEventLoop.ExcludeUserInputEvents, 5)

    def clear(self):
        self._pending = []
        self._debounce.cancel()
        if self._box is not None:
            try:
                self._box.clear()
            except RuntimeError:
                pass

    def _scroll_to_end(self):
        bar = self._box.verticalScrollBar()
        bar.setValue(bar.maximum())


__all__ = ["ActivityLog"]
