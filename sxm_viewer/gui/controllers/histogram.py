"""Histogram & range dialog helpers for preview canvases."""
from __future__ import annotations

from pathlib import Path

import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from ..._shared import QtWidgets, QtCore


def open_histogram_dialog(owner, canvas):
    """Show histogram dialog for a given preview canvas."""
    if canvas is None or not getattr(canvas, "views", None):
        QtWidgets.QMessageBox.information(owner, "Histogram", "No image loaded in preview.")
        return

    views = list(canvas.views)
    dialog_parent = None
    try:
        dialog_parent = canvas.window() if canvas is not None else None
    except Exception:
        dialog_parent = None
    if dialog_parent is None:
        dialog_parent = owner
    dlg = QtWidgets.QDialog(dialog_parent)
    dlg.setWindowTitle("Histogram & Range")
    dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
    layout = QtWidgets.QVBoxLayout(dlg)

    selector = None
    if len(views) > 1:
        selector = QtWidgets.QComboBox()
        for v in views:
            title = v.get("title") or Path(str(v.get("path", ""))).name
            selector.addItem(title)
        layout.addWidget(selector)

    fig = Figure(figsize=(4.5, 3))
    canvas_hist = FigureCanvas(fig)
    ax = fig.add_subplot(111)
    layout.addWidget(canvas_hist)

    spins_layout = QtWidgets.QHBoxLayout()
    spins_layout.addWidget(QtWidgets.QLabel("Min:"))
    spin_lo = QtWidgets.QDoubleSpinBox()
    spin_lo.setDecimals(6)
    spin_lo.setMinimum(-1e12)
    spin_lo.setMaximum(1e12)
    spins_layout.addWidget(spin_lo, 1)
    spins_layout.addWidget(QtWidgets.QLabel("Max:"))
    spin_hi = QtWidgets.QDoubleSpinBox()
    spin_hi.setDecimals(6)
    spin_hi.setMinimum(-1e12)
    spin_hi.setMaximum(1e12)
    spins_layout.addWidget(spin_hi, 1)
    layout.addLayout(spins_layout)

    btn_row = QtWidgets.QHBoxLayout()
    auto_btn = QtWidgets.QPushButton("Auto (1-99%)")
    reset_btn = QtWidgets.QPushButton("Reset")
    live_cb = QtWidgets.QCheckBox("Live preview")
    live_cb.setChecked(True)
    apply_btn = QtWidgets.QPushButton("Apply")
    btn_row.addWidget(auto_btn)
    btn_row.addWidget(reset_btn)
    btn_row.addWidget(live_cb)
    btn_row.addStretch(1)
    btn_row.addWidget(apply_btn)
    layout.addLayout(btn_row)

    state = {"view_idx": 0, "lines": (None, None), "finite": None, "dragging": None, "undo_pushed": False}

    def load_view(idx: int):
        idx = max(0, min(idx, len(views) - 1))
        state["view_idx"] = idx
        view = views[idx]
        vmin, vmax, finite = owner._view_finite_values(view)
        state["finite"] = finite
        ax.clear()
        if finite is None:
            ax.set_title("No finite data")
            canvas_hist.draw_idle()
            return
        hist, edges = np.histogram(finite, bins=256)
        ax.plot(edges[:-1], hist, color="#4a90e2")
        ax.set_xlabel("Value")
        ax.set_ylabel("Count")
        lo, hi = view.get("clim", (vmin, vmax))
        spin_lo.blockSignals(True)
        spin_hi.blockSignals(True)
        spin_lo.setMinimum(vmin)
        spin_hi.setMinimum(vmin)
        spin_lo.setMaximum(vmax)
        spin_hi.setMaximum(vmax)
        spin_lo.setValue(lo)
        spin_hi.setValue(hi)
        spin_lo.blockSignals(False)
        spin_hi.blockSignals(False)
        l0 = ax.axvline(lo, color="#d81b60", linestyle="--")
        l1 = ax.axvline(hi, color="#d81b60", linestyle="--")
        state["lines"] = (l0, l1)
        ax.set_title(view.get("title") or Path(str(view.get("path", ""))).name)
        canvas_hist.draw_idle()

    def update_lines():
        l0, l1 = state.get("lines", (None, None))
        if l0 is None or l1 is None:
            return
        lo = spin_lo.value()
        hi = spin_hi.value()
        l0.set_xdata([lo, lo])
        l1.set_xdata([hi, hi])
        canvas_hist.draw_idle()

    def _set_spin_values(lo, hi, block=False):
        if block:
            spin_lo.blockSignals(True)
            spin_hi.blockSignals(True)
        spin_lo.setValue(lo)
        spin_hi.setValue(hi)
        if block:
            spin_lo.blockSignals(False)
            spin_hi.blockSignals(False)

    def apply_current(close=False):
        idx = state.get("view_idx", 0)
        view = views[idx]
        lo, hi = spin_lo.value(), spin_hi.value()
        if lo > hi:
            lo, hi = hi, lo
        if not state.get("undo_pushed", False):
            try:
                canvas.push_undo_state("histogram_range")
                state["undo_pushed"] = True
            except Exception:
                pass
        owner._apply_clim_to_view(canvas, view, lo, hi)
        if close:
            dlg.accept()

    def on_auto():
        finite = state.get("finite")
        if finite is None:
            return
        try:
            lo, hi = np.percentile(finite, [1, 99])
        except Exception:
            return
        _set_spin_values(float(lo), float(hi), block=True)
        update_lines()
        maybe_live_apply()

    def on_reset():
        view = views[state.get("view_idx", 0)]
        vmin, vmax, _ = owner._view_finite_values(view)
        if vmin is None:
            return
        _set_spin_values(vmin, vmax, block=True)
        update_lines()
        maybe_live_apply()

    def maybe_live_apply():
        if live_cb.isChecked():
            apply_current(close=False)

    def _on_press(event):
        if event.button != 1 or event.xdata is None:
            return
        l0, l1 = state.get("lines", (None, None))
        if l0 is None or l1 is None:
            return
        lo_val = spin_lo.value()
        hi_val = spin_hi.value()
        span = max(abs(hi_val - lo_val), 1e-12)
        tol = span * 0.01
        if abs(event.xdata - lo_val) < tol:
            state["dragging"] = "lo"
        elif abs(event.xdata - hi_val) < tol:
            state["dragging"] = "hi"
        else:
            state["dragging"] = None

    def _on_motion(event):
        if state.get("dragging") is None or event.xdata is None:
            return
        x = float(event.xdata)
        lo_val = spin_lo.value()
        hi_val = spin_hi.value()
        if state["dragging"] == "lo":
            lo_val = min(x, hi_val)
        elif state["dragging"] == "hi":
            hi_val = max(x, lo_val)
        _set_spin_values(lo_val, hi_val, block=True)
        update_lines()
        maybe_live_apply()

    def _on_release(event):
        if state.get("dragging") is None:
            return
        state["dragging"] = None
        maybe_live_apply()

    spin_lo.valueChanged.connect(update_lines)
    spin_hi.valueChanged.connect(update_lines)
    spin_lo.valueChanged.connect(maybe_live_apply)
    spin_hi.valueChanged.connect(maybe_live_apply)
    apply_btn.clicked.connect(lambda: apply_current(close=True))
    auto_btn.clicked.connect(on_auto)
    reset_btn.clicked.connect(on_reset)
    if selector:
        selector.currentIndexChanged.connect(load_view)

    try:
        cid_press = canvas_hist.mpl_connect("button_press_event", _on_press)
        cid_motion = canvas_hist.mpl_connect("motion_notify_event", _on_motion)
        cid_release = canvas_hist.mpl_connect("button_release_event", _on_release)
        fig.canvas.setProperty("hist_cids", (cid_press, cid_motion, cid_release))
    except Exception:
        pass

    profile_was_enabled = bool(getattr(canvas, "profile_enabled", False))
    profile_user_flag = bool(getattr(canvas, "_profile_user_enabled", profile_was_enabled))
    exported_profile_state = None
    if profile_was_enabled:
        try:
            exported_profile_state = canvas.export_profile_state()
        except Exception:
            exported_profile_state = None
        try:
            canvas.enable_profile(False)
        except Exception:
            pass
        try:
            canvas._profile_user_enabled = False
        except Exception:
            pass

    load_view(0)
    dlg.exec_()

    if profile_was_enabled:
        try:
            if exported_profile_state:
                canvas.import_profile_state(exported_profile_state, emit=False)
        except Exception:
            pass
        try:
            canvas.enable_profile(True)
        except Exception:
            pass
        try:
            canvas._profile_user_enabled = profile_user_flag
        except Exception:
            pass
    else:
        try:
            canvas.enable_profile(False)
        except Exception:
            pass
