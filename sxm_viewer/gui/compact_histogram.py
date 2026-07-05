"""Compact, always-visible histogram control for the main toolbar.

A small sibling of the "Histogram & Range" dialog (gui/controllers/histogram.py):
same drag-a-limit-line / Auto (1-99%) / Reset interaction, but rendered inline in
the toolbar instead of a modal popup. This widget owns no viewer state - it only
renders whatever it's told via set_data()/set_range() and emits signals for the
main window to act on; all clim math (percentile/finite-range/apply) stays in
main_window.py so this widget and the full dialog always agree.
"""
from __future__ import annotations

import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from .._shared import QtCore, QtWidgets

_HIST_BINS = 48
_DRAG_TOL_FRACTION = 0.04  # wider than the full dialog's 1%: fewer pixels here


class CompactHistogramWidget(QtWidgets.QWidget):
    """Inline histogram with draggable min/max lines, plus Auto and Reset buttons."""

    # (lo, hi, final) - final=True on drag-release / Auto / Reset (commit + undo);
    # False on intermediate drag motion (live preview only).
    climChanged = QtCore.pyqtSignal(float, float, bool)
    autoRequested = QtCore.pyqtSignal()
    resetRequested = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._vmin = None
        self._vmax = None
        self._lo = None
        self._hi = None
        self._dragging = None  # None | "lo" | "hi"
        self._lines = (None, None)

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        self._fig = Figure(figsize=(1.2, 0.3), dpi=100)
        self._fig.subplots_adjust(left=0.01, right=0.99, top=0.97, bottom=0.03)
        self._ax = self._fig.add_subplot(111)
        self._style_axes()
        self._canvas = FigureCanvas(self._fig)
        self._canvas.setMinimumSize(70, 30)
        self._canvas.setMaximumSize(150, 30)
        self._canvas.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        self._canvas.setToolTip(
            "Contrast range for the current view. Drag either line to adjust; "
            "use Auto for a robust 1-99% stretch, or Reset for the true min/max."
        )
        layout.addWidget(self._canvas)

        self._auto_btn = QtWidgets.QToolButton(self)
        self._auto_btn.setText("Auto")
        self._auto_btn.setToolTip("Set range to the 1st-99th percentile of the current view")
        self._auto_btn.setSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        self._auto_btn.clicked.connect(self.autoRequested.emit)
        layout.addWidget(self._auto_btn)

        self._reset_btn = QtWidgets.QToolButton(self)
        self._reset_btn.setText("Reset")
        self._reset_btn.setToolTip("Set range to the true finite min/max of the current view")
        self._reset_btn.setSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Fixed)
        self._reset_btn.clicked.connect(self.resetRequested.emit)
        layout.addWidget(self._reset_btn)

        self._set_controls_enabled(False)

        self._canvas.mpl_connect("button_press_event", self._on_press)
        self._canvas.mpl_connect("motion_notify_event", self._on_motion)
        self._canvas.mpl_connect("button_release_event", self._on_release)

        self.clear()

    def _style_axes(self):
        self._ax.set_xticks([])
        self._ax.set_yticks([])
        for spine in self._ax.spines.values():
            spine.set_visible(False)

    def _set_controls_enabled(self, enabled: bool):
        self._auto_btn.setEnabled(enabled)
        self._reset_btn.setEnabled(enabled)
        self._canvas.setEnabled(enabled)

    # ---------- public API ----------

    def clear(self):
        """Show an empty/placeholder state (no view loaded)."""
        self._vmin = self._vmax = self._lo = self._hi = None
        self._dragging = None
        self._lines = (None, None)
        self._ax.clear()
        self._style_axes()
        self._canvas.draw_idle()
        self._set_controls_enabled(False)

    def set_data(self, vmin, vmax, lo, hi, finite):
        """Repopulate the histogram and marker positions for a new view."""
        if vmin is None or vmax is None or finite is None or len(finite) == 0:
            self.clear()
            return
        self._vmin = float(vmin)
        self._vmax = float(vmax)
        self._lo = float(lo)
        self._hi = float(hi)
        self._ax.clear()
        self._style_axes()
        try:
            hist, edges = np.histogram(finite, bins=_HIST_BINS)
            self._ax.fill_between(edges[:-1], hist, step="post", color="#4a90e2", linewidth=0)
        except Exception:
            pass
        self._ax.set_xlim(self._vmin, self._vmax)
        l0 = self._ax.axvline(self._lo, color="#d81b60", linestyle="-", linewidth=1.2)
        l1 = self._ax.axvline(self._hi, color="#d81b60", linestyle="-", linewidth=1.2)
        self._lines = (l0, l1)
        self._canvas.draw_idle()
        self._set_controls_enabled(True)

    def set_range(self, lo, hi):
        """Move the marker lines only, without recomputing the histogram bars."""
        if self._vmin is None:
            return
        self._lo, self._hi = float(lo), float(hi)
        self._update_lines()

    # ---------- drag interaction (ported from controllers/histogram.py) ----------

    def _update_lines(self):
        l0, l1 = self._lines
        if l0 is None or l1 is None:
            return
        l0.set_xdata([self._lo, self._lo])
        l1.set_xdata([self._hi, self._hi])
        self._canvas.draw_idle()

    def _on_press(self, event):
        if event.button != 1 or event.xdata is None or self._lo is None:
            return
        span = max(abs(self._hi - self._lo), 1e-12)
        tol = span * _DRAG_TOL_FRACTION
        if abs(event.xdata - self._lo) < tol:
            self._dragging = "lo"
        elif abs(event.xdata - self._hi) < tol:
            self._dragging = "hi"
        else:
            self._dragging = None

    def _on_motion(self, event):
        if self._dragging is None or event.xdata is None:
            return
        x = float(event.xdata)
        if self._dragging == "lo":
            self._lo = min(x, self._hi)
        else:
            self._hi = max(x, self._lo)
        self._update_lines()
        self.climChanged.emit(self._lo, self._hi, False)

    def _on_release(self, event):
        if self._dragging is None:
            return
        self._dragging = None
        self.climChanged.emit(self._lo, self._hi, True)
