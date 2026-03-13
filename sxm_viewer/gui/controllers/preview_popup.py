"""Helpers for building preview pop-out dialogs."""
from __future__ import annotations

from ..._shared import QtWidgets, QtCore
from ..canvases.detail_preview_canvas import MultiPreviewCanvas
from ...processing.filters import FILTER_DEFINITIONS, _gaussian_available
from .profile import PopupProfileController


def spawn_preview_popup(owner, views, title=None):
    """Create a preview popup dialog reusing the existing owner logic."""
    if not views:
        return None

    dlg = QtWidgets.QDialog(owner)
    dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
    dlg.setWindowFlags(
        dlg.windowFlags()
        | QtCore.Qt.WindowMinimizeButtonHint
        | QtCore.Qt.WindowMaximizeButtonHint
        | QtCore.Qt.WindowSystemMenuHint
    )
    dlg.setMinimumSize(0, 0)
    dlg.setWindowTitle(title or "Preview")
    layout = QtWidgets.QVBoxLayout(dlg)
    layout.setContentsMargins(6, 6, 6, 6)
    controls_grid = QtWidgets.QGridLayout()
    controls_grid.setContentsMargins(4, 4, 4, 4)
    controls_grid.setHorizontalSpacing(6)
    controls_grid.setVerticalSpacing(4)
    layout.setSizeConstraint(QtWidgets.QLayout.SetNoConstraint)

    def _tool_button(label, tooltip):
        btn = QtWidgets.QToolButton()
        btn.setText(label)
        btn.setCheckable(True)
        btn.setAutoRaise(False)
        btn.setToolTip(tooltip)
        return btn

    measure_btn = _tool_button("Profile", "Toggle profile measurement (drag line)")
    angle_btn = _tool_button("Angle", "Toggle angle tool (Ctrl+Shift+click to add vertex)")
    scale_btn = _tool_button("Scale", "Toggle scale bar")
    clear_btn = QtWidgets.QToolButton()
    clear_btn.setText("Clear overlays")
    clear_btn.setToolTip("Clear profiles, angles and related overlays")
    filter_btn = QtWidgets.QToolButton()
    filter_btn.setText("Filter")
    filter_btn.setToolTip("Apply scan-level filters (flatten, plane, etc.)")
    filter_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    hist_btn = QtWidgets.QToolButton()
    hist_btn.setText("Histogram")
    hist_btn.setToolTip("Show histogram and adjust display range")
    hist_btn.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
    layout_cb = QtWidgets.QComboBox()
    layout_cb.addItems(["Grid", "Stacked"])
    relative_cb = QtWidgets.QCheckBox("Relative axes")
    frame_fill_cb = QtWidgets.QCheckBox("Frame fill")
    frame_fill_cb.setToolTip("Hide title/axes ticks/colorbar and fill the frame to the window")
    controls_grid.addWidget(measure_btn, 0, 0)
    controls_grid.addWidget(angle_btn, 0, 1)
    controls_grid.addWidget(scale_btn, 0, 2)
    controls_grid.addWidget(filter_btn, 0, 3)
    controls_grid.addWidget(hist_btn, 0, 4)
    controls_grid.addWidget(clear_btn, 1, 0)
    controls_grid.addWidget(QtWidgets.QLabel("Layout"), 1, 1)
    controls_grid.addWidget(layout_cb, 1, 2)
    controls_grid.addWidget(relative_cb, 1, 3, 1, 2)
    controls_grid.addWidget(frame_fill_cb, 1, 5)
    controls_grid.setColumnStretch(6, 1)

    # Use a default that we immediately adapt to the
    # aspect ratio of the underlying image so the popup
    # is created snugly around the content.
    canvas = MultiPreviewCanvas(dlg, figsize=(4, 3))
    try:
        canvas.set_compact_size_hints(True)
        canvas.setMinimumSize(0, 0)
        canvas.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
    except Exception:
        pass
    try:
        canvas.set_show_title(getattr(owner, "show_preview_title", True))
    except Exception:
        pass
    try:
        canvas.set_show_molecules(getattr(owner, "show_molecules", True))
    except Exception:
        pass
    try:
        canvas.set_show_acquisition_overlay(getattr(owner, "show_acquisition_overlay", False))
    except Exception:
        pass
    try:
        canvas.set_profile_label_mode(getattr(owner, "profile_label_mode", "length"))
    except Exception:
        pass
    try:
        # Keep data undistorted in popups; combined with square canvas sizing
        # this prevents rectangular stretching while still using available area.
        canvas.set_fit_to_canvas(False)
    except Exception:
        pass
    canvas.set_view_layout(getattr(owner.preview_canvas, "_view_layout", "grid"))
    _square_resize_busy = {"active": False}
    _popup_resize_threshold_px = 2
    _last_square_target = {"w": -1, "h": -1}
    _mode_state = {"frame_fill": False}
    resize_sync_timer = QtCore.QTimer(dlg)
    resize_sync_timer.setSingleShot(True)
    resize_sync_timer.setInterval(45)
    resize_settle_timer = QtCore.QTimer(dlg)
    resize_settle_timer.setSingleShot(True)
    resize_settle_timer.setInterval(180)

    def _enforce_square_dialog():
        if _square_resize_busy["active"]:
            return
        try:
            if dlg.isMaximized() or dlg.isFullScreen():
                return
        except Exception:
            return
        try:
            _square_resize_busy["active"] = True
            layout.activate()
            margins = layout.contentsMargins()
            controls_w = max(controls_grid.sizeHint().width(), 1)
            controls_h = max(controls_grid.sizeHint().height(), 1)
            canvas_min = canvas.minimumSizeHint()
            min_side = max(controls_w, canvas_min.width(), canvas_min.height(), 1)
            avail_w = max(1, dlg.width() - margins.left() - margins.right())
            avail_h = max(1, dlg.height() - margins.top() - margins.bottom() - controls_h)
            side = min(avail_w, avail_h)
            side = max(side, min_side)
            target_w = side + margins.left() + margins.right()
            target_h = side + controls_h + margins.top() + margins.bottom()
            dlg_min = dlg.minimumSizeHint()
            target_w = max(target_w, dlg_min.width())
            target_h = max(target_h, dlg_min.height())
            if (
                target_w == _last_square_target["w"]
                and target_h == _last_square_target["h"]
                and abs(target_w - dlg.width()) <= 1
                and abs(target_h - dlg.height()) <= 1
            ):
                return
            if abs(target_w - dlg.width()) > 1 or abs(target_h - dlg.height()) > 1:
                _last_square_target["w"] = int(target_w)
                _last_square_target["h"] = int(target_h)
                dlg.resize(target_w, target_h)
        except Exception:
            pass
        finally:
            _square_resize_busy["active"] = False

    def _resize_to_canvas(force=False):
        try:
            layout.activate()
            if force:
                dlg.adjustSize()
                # Allow shrinking after the initial tight fit.
                dlg.setMinimumSize(0, 0)
                _enforce_square_dialog()
        except Exception:
            pass

    def _enforce_square_when_idle():
        try:
            if QtWidgets.QApplication.mouseButtons() != QtCore.Qt.NoButton:
                resize_settle_timer.start()
                return
        except Exception:
            pass
        _enforce_square_dialog()

    def _schedule_resize(force=False):
        if force:
            QtCore.QTimer.singleShot(0, lambda: _resize_to_canvas(force=True))
            return
        try:
            resize_sync_timer.start()
            resize_settle_timer.start()
        except Exception:
            QtCore.QTimer.singleShot(0, lambda: _resize_to_canvas(force=False))

    resize_sync_timer.timeout.connect(lambda: _resize_to_canvas(force=False))
    resize_settle_timer.timeout.connect(_enforce_square_when_idle)

    # Try to adapt the canvas figure size to the image aspect
    # so that `adjustSize()` produces a tight dialog around the
    # displayed frame (axes + colorbar).
    try:
        base = 5.0
        v0 = views[0]
        arr0 = v0.get("arr")
        if arr0 is not None:
            import numpy as _np
            a = _np.asarray(arr0)
            if a.ndim >= 2 and a.shape[0] > 0:
                h, w = a.shape[0], a.shape[1]
                aspect = float(w) / float(h) if h else 1.0
                if aspect >= 1.0:
                    fig_w = base * aspect
                    fig_h = base
                else:
                    fig_w = base
                    fig_h = base / aspect
                try:
                    canvas.fig.set_size_inches(fig_w, fig_h, forward=True)
                except Exception:
                    canvas.fig.set_size_inches(fig_w, fig_h)
    except Exception:
        pass

    canvas.set_views([owner._copy_view_for_popup(v) for v in views])
    canvas.set_views_callback(lambda _=None: _schedule_resize(force=False))
    canvas.enable_scale_bar(owner.scale_bar_cb.isChecked())
    canvas._detail_dark = bool(getattr(owner, "detail_dark_view", False))
    canvas._detail_grid = bool(getattr(owner, "detail_grid_view", False))
    canvas.set_crop_callback(
        lambda v: spawn_preview_popup(
            owner,
            [owner._copy_view_for_popup(v)],
            title=owner._friendly_view_title(v, default="Cropped view"),
        )
    )
    canvas.set_double_click_callback(
        lambda v=None: spawn_preview_popup(
            owner,
            [owner._copy_view_for_popup(v)] if v else [],
            title=owner._friendly_view_title(v, default="Preview copy") if v else "Preview copy",
        )
    )
    canvas.set_filter_menu_callback(lambda menu, view, c=canvas: owner._populate_canvas_filter_menu(menu, c, view))
    canvas.set_stp_export_callback(owner._export_view_as_stp)
    canvas.set_window_arrange_callback(owner.on_arrange_popouts)
    seq = views[0].get("crop_sequence") if views else None
    if hasattr(owner, "quick_crop_controller"):
        owner.quick_crop_controller.register_popup(seq, dlg)
    canvas.enable_fixed_crop_quick_mode(owner.quick_crop_mode)
    canvas.show_fixed_crop_template(owner.show_crop_template_overlay)
    canvas.show_fixed_crop_history(owner.show_crop_history_overlay)
    try:
        canvas.set_molecule_palette(owner.molecule_palette, notify=False)
        canvas.set_molecule_palette_callback(owner._on_molecule_palette_changed)
        owner._popup_canvases.append(canvas)
    except Exception:
        pass

    measure_initial = bool(
        getattr(owner.preview_canvas, "_profile_user_enabled", getattr(owner.preview_canvas, "profile_enabled", False))
    )
    angle_initial = getattr(owner.preview_canvas, "angle_enabled", False)
    profile_controller = PopupProfileController(owner, canvas, title or "Profile")
    profile_controller.set_initial_state(measure_initial)
    canvas.enable_angle(angle_initial)
    measure_btn.setChecked(measure_initial)
    angle_btn.setChecked(angle_initial)
    scale_btn.setChecked(owner.scale_bar_cb.isChecked())
    hist_menu = QtWidgets.QMenu(hist_btn)
    hist_menu.addAction("Adjust…", lambda: owner._open_histogram_dialog(canvas))
    hist_menu.addAction("Auto (1–99%)", lambda: owner._auto_contrast(canvas))
    hist_menu.addAction("Reset range", lambda: owner._reset_contrast(canvas))
    hist_btn.setMenu(hist_menu)
    hist_btn.clicked.connect(lambda _: owner._open_histogram_dialog(canvas))

    filter_menu = QtWidgets.QMenu(filter_btn)
    for key, info in FILTER_DEFINITIONS.items():
        act = QtWidgets.QAction(info["label"], filter_menu)
        if info.get("needs_gaussian") and not _gaussian_available():
            act.setEnabled(False)
            act.setToolTip("Requires scipy or OpenCV.")
        act.triggered.connect(lambda _, k=key: owner._apply_filter_to_canvas(canvas, filter_key=k))
        filter_menu.addAction(act)
    filter_menu.addSeparator()
    filter_menu.addAction("Custom pipeline...", lambda: owner._open_custom_filter_for_canvas(canvas))
    filter_menu.addAction("Clear filter", lambda: owner._apply_filter_to_canvas(canvas, pipeline=[]))
    filter_btn.setMenu(filter_menu)
    filter_btn.clicked.connect(lambda _: owner._open_custom_filter_for_canvas(canvas))

    layout_cb.setCurrentText("Stacked" if canvas._view_layout == "stacked" else "Grid")
    rel_initial = any(bool(v.get("relative_axes")) for v in views if isinstance(v, dict))
    relative_cb.setChecked(rel_initial)
    canvas.set_relative_axes_override(rel_initial)
    frame_fill_cb.setChecked(False)

    def _toggle_angle(checked):
        try:
            canvas.enable_angle(bool(checked))
        except Exception:
            pass

    def _toggle_scale(checked):
        canvas.enable_scale_bar(bool(checked))
        try:
            canvas._redraw()
        except Exception:
            pass

    def _toggle_layout(text):
        canvas.set_view_layout("stacked" if text.lower().startswith("stacked") else "grid")
        _schedule_resize(force=False)

    def _toggle_frame_fill(checked):
        _mode_state["frame_fill"] = bool(checked)
        try:
            canvas.set_frame_fill_mode(_mode_state["frame_fill"])
        except Exception:
            pass
        _last_square_target["w"] = -1
        _last_square_target["h"] = -1
        _schedule_resize(force=False)

    measure_btn.toggled.connect(profile_controller.toggle_measure)
    angle_btn.toggled.connect(_toggle_angle)
    scale_btn.toggled.connect(_toggle_scale)
    layout_cb.currentTextChanged.connect(_toggle_layout)
    frame_fill_cb.toggled.connect(_toggle_frame_fill)
    def _toggle_relative(checked):
        canvas.set_relative_axes_override(bool(checked))
        _schedule_resize(force=False)

    relative_cb.toggled.connect(_toggle_relative)

    def _clear_overlays():
        try:
            canvas.clear_angle_measurement()
            canvas._clear_profile_artists()
            canvas._clear_saved_profile_artists(notify=False)
            canvas.profile_pts = None
            canvas._emit_profile_state()
            canvas.draw_idle()
        except Exception:
            pass

    clear_btn.clicked.connect(_clear_overlays)
    controls_widget = QtWidgets.QWidget(dlg)
    controls_widget.setLayout(controls_grid)
    layout.addWidget(controls_widget, 0)
    layout.addWidget(canvas, 1)
    canvas.setFocus()

    class _PopupKeyFilter(QtCore.QObject):
        def __init__(self, cvs):
            super().__init__(cvs)
            self.canvas = cvs

        def eventFilter(self, obj, event):
            if event.type() == QtCore.QEvent.Resize:
                try:
                    new_size = event.size()
                    old_size = event.oldSize()
                    dw = abs(int(new_size.width()) - int(old_size.width()))
                    dh = abs(int(new_size.height()) - int(old_size.height()))
                    if max(dw, dh) >= _popup_resize_threshold_px:
                        _schedule_resize(force=False)
                except Exception:
                    _schedule_resize(force=False)
                return False
            if event.type() == QtCore.QEvent.Wheel:
                try:
                    if event.modifiers() & QtCore.Qt.ControlModifier:
                        _schedule_resize(force=False)
                except Exception:
                    pass
                return False
            if event.type() == QtCore.QEvent.KeyPress:
                if (event.modifiers() & QtCore.Qt.ControlModifier) and event.key() == QtCore.Qt.Key_D:
                    try:
                        spawn_preview_popup(
                            owner,
                            [owner._copy_view_for_popup(v) for v in self.canvas.views],
                            title="Preview copy",
                        )
                    except Exception:
                        pass
                    event.accept()
                    return True
            return False

    key_filter = _PopupKeyFilter(canvas)
    dlg.installEventFilter(key_filter)
    canvas.installEventFilter(key_filter)
    _schedule_resize(force=True)
    dlg.show()
    owner._popup_refs.append(dlg)
    if hasattr(owner, "quick_crop_controller"):
        owner.quick_crop_controller.update_popup_actions()

    def _on_popup_closed(_=None):
        if dlg in owner._popup_refs:
            owner._popup_refs.remove(dlg)
        if hasattr(owner, "quick_crop_controller"):
            owner.quick_crop_controller.update_popup_actions()
        profile_controller.dispose()

    def _remove_popup_canvas(_=None):
        if canvas in getattr(owner, "_popup_canvases", []):
            owner._popup_canvases.remove(canvas)

    dlg.finished.connect(_on_popup_closed)
    dlg.finished.connect(_remove_popup_canvas)
    return dlg
