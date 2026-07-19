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
        self._pending: list[str] = []
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.flush)
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
        if not self._timer.isActive():
            self._timer.start(FLUSH_INTERVAL_MS)

    def flush(self):
        """Write everything queued, in one append."""
        if self._box is None:
            self._pending = []
            return
        pending, self._pending = self._pending, []
        if not pending:
            return
        try:
            self._box.appendPlainText("\n".join(pending))
            self._scroll_to_end()
        except Exception:
            # Fall back to line-at-a-time so one bad entry cannot lose the
            # rest of the batch.
            for entry in pending:
                try:
                    self._box.appendPlainText(entry)
                except Exception:
                    pass
        try:
            # Keep the log visibly moving during long synchronous work
            # (folder loads) without letting the user click mid-operation.
            QtWidgets.QApplication.processEvents(
                QtCore.QEventLoop.ExcludeUserInputEvents, 5)
        except Exception:
            pass

    def clear(self):
        self._pending = []
        if self._box is not None:
            try:
                self._box.clear()
            except Exception:
                pass

    def _scroll_to_end(self):
        try:
            bar = self._box.verticalScrollBar()
            bar.setValue(bar.maximum())
        except Exception:
            pass


__all__ = ["ActivityLog"]
