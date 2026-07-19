"""Debounced callbacks: coalesce bursts of requests into one deferred call.

Extracted because the attribute-cohesion analysis found the same
hand-rolled shape - ``_X_pending`` + ``_X_timer`` + ``_flush_X()`` -
implemented independently **four times** on ``SXMGridViewer`` (activity
log, preview request, thumbnail render state, compact histogram). Each
copy carried its own subtly different guard code, and each contributed
2-3 more attributes to an already 800-attribute class.

Two payload policies cover every existing use:

* **latest** - keep only the most recent payload (a preview request:
  older requests are obsolete the moment a newer one arrives);
* **accumulate** - union successive payloads into a set (thumbnail paths
  needing a re-render: every one still has to happen).

Both are single-shot: the timer is armed on the first request and left
alone by later ones, so a continuous stream still flushes every
``interval_ms`` rather than being starved.
"""
from __future__ import annotations

from .._shared import QtCore

LATEST = "latest"
ACCUMULATE = "accumulate"


class Debouncer(QtCore.QObject):
    """Call ``callback`` once, ``interval_ms`` after the first request.

    ``callback`` is invoked with no arguments; read the coalesced payload
    with :meth:`take` (which also clears it). Keeping the two separate
    means a callback that bails out early can leave the payload for the
    next flush simply by not calling ``take``.
    """

    def __init__(self, callback, interval_ms=0, mode=LATEST, parent=None):
        super().__init__(parent)
        self._callback = callback
        self._interval_ms = int(interval_ms)
        self._mode = mode
        self._payload = set() if mode == ACCUMULATE else None
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._fire)

    # --- scheduling -----------------------------------------------------

    def schedule(self, payload=None):
        """Record ``payload`` and arm the timer if not already armed.

        In ACCUMULATE mode ``payload`` is treated as an iterable whose
        items are unioned into the pending set; in LATEST mode it replaces
        whatever was pending.

        Falls back to firing synchronously if the timer cannot be started
        (e.g. called from a non-GUI thread during teardown) - the original
        hand-rolled copies all had this same guard.
        """
        if self._mode == ACCUMULATE:
            if payload:
                self._payload.update(payload)
        else:
            self._payload = payload
        try:
            if not self._timer.isActive():
                self._timer.start(self._interval_ms)
        except RuntimeError:
            self._fire()

    def rearm(self):
        """Re-arm the timer without touching the pending payload.

        For the "a new request arrived while we were busy" case: the
        callback returned early leaving the payload in place, and it must
        be retried rather than dropped.
        """
        if self._payload in (None, set()):
            return
        try:
            if not self._timer.isActive():
                self._timer.start(self._interval_ms)
        except RuntimeError:
            self._fire()

    def cancel(self):
        """Drop the pending payload and disarm."""
        self._stop_timer()
        self._payload = set() if self._mode == ACCUMULATE else None

    def flush(self):
        """Run the callback now, if anything is pending."""
        self._stop_timer()
        self._fire()

    def _stop_timer(self):
        # RuntimeError is the Qt "wrapped C++ object has been deleted"
        # case during teardown - the only thing stop() realistically
        # raises, so catch it narrowly rather than swallowing everything.
        try:
            self._timer.stop()
        except RuntimeError:
            pass

    @property
    def is_active(self):
        try:
            return self._timer.isActive()
        except RuntimeError:
            return False

    # --- payload --------------------------------------------------------

    def take(self):
        """Return and clear the coalesced payload."""
        if self._mode == ACCUMULATE:
            payload, self._payload = self._payload, set()
            return payload
        payload, self._payload = self._payload, None
        return payload

    def peek(self):
        return self._payload

    def _fire(self):
        try:
            self._callback()
        except Exception:
            # A debounced UI refresh must never take the app down; the
            # hand-rolled versions all swallowed here too.
            pass


__all__ = ["Debouncer", "LATEST", "ACCUMULATE"]
