"""Helpers for building preview pop-out dialogs."""
from __future__ import annotations

from ..._shared import QtWidgets, QtCore
from ..canvases.detail_preview_canvas import MultiPreviewCanvas
from ..dialogs.profile_dialog import ProfileDialog
from ...processing.filters import FILTER_DEFINITIONS, _gaussian_available


def spawn_preview_popup(owner, views, title=None):
    """Create a preview popup dialog reusing the existing owner logic."""
    if not views:
        return None

    dlg = QtWidgets.QDialog(owner)
    dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
    dlg.setWindowFlags(dlg.windowFlags() | QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowMaximizeButtonHint)
    dlg.setWindowTitle(title or "Preview")
    layout = QtWidgets.QVBoxLayout(dlg)
    controls_bar = QtWidgets.QHBoxLayout()
    controls_bar.setContentsMargins(6, 6, 6, 6)
    controls_bar.setSpacing(10)

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
    controls_bar.addWidget(measure_btn)
    controls_bar.addWidget(angle_btn)
    controls_bar.addWidget(scale_btn)
    controls_bar.addWidget(filter_btn)
    controls_bar.addWidget(hist_btn)
    controls_bar.addWidget(clear_btn)
    controls_bar.addSpacing(6)
    controls_bar.addWidget(QtWidgets.QLabel("Layout"))
    controls_bar.addWidget(layout_cb)
    controls_bar.addStretch(1)

    canvas = MultiPreviewCanvas(dlg, figsize=(6, 5))
    try:
        canvas.set_show_title(getattr(owner, "show_preview_title", True))
    except Exception:
        pass
    try:
        canvas.set_show_molecules(getattr(owner, "show_molecules", True))
    except Exception:
        pass
    canvas.set_view_layout(getattr(owner.preview_canvas, "_view_layout", "grid"))
    canvas.set_views([owner._copy_view_for_popup(v) for v in views])
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
    seq = views[0].get("crop_sequence") if views else None
    owner._register_crop_popup(seq, dlg)
    canvas.enable_fixed_crop_quick_mode(owner.quick_crop_mode)
    canvas.show_fixed_crop_template(owner.show_crop_template_overlay)
    canvas.show_fixed_crop_history(owner.show_crop_history_overlay)
    try:
        canvas.set_molecule_palette(owner.molecule_palette, notify=False)
        canvas.set_molecule_palette_callback(owner._on_molecule_palette_changed)
        owner._popup_canvases.append(canvas)
    except Exception:
        pass

    measure_initial = bool(getattr(owner.preview_canvas, "_profile_user_enabled", getattr(owner.preview_canvas, "profile_enabled", False)))
    angle_initial = getattr(owner.preview_canvas, "angle_enabled", False)
    canvas.enable_profile(measure_initial)
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
    popup_state = {"profile_dialog": None}

    def _ensure_profile_dialog(active, saved):
        dlg_prof = popup_state.get("profile_dialog")
        if dlg_prof is None:
            unit = None
            y_label = None
            if canvas.views:
                try:
                    unit = canvas.views[0].get("unit")
                    y_label = canvas.views[0].get("colorbar_label") or canvas.views[0].get("unit")
                except Exception:
                    pass
            dlg_prof = ProfileDialog(active, saved, parent=dlg, unit=unit, y_label=y_label)
            dlg_prof.setWindowTitle((title or "Profile") + " (popup)")
            try:
                dlg_prof.set_context_source(canvas, dark=owner.detail_dark_view, grid=owner.detail_grid_view)
            except Exception:
                pass
            dlg_prof.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
            dlg_prof.finished.connect(lambda _=None: popup_state.__setitem__("profile_dialog", None))
            popup_state["profile_dialog"] = dlg_prof
        dlg_prof.update_profiles(active, saved)
        dlg_prof.show()
        try:
            dlg_prof.raise_()
            dlg_prof.activateWindow()
        except Exception:
            pass

    def _compute_profiles_from_canvas():
        active = None
        try:
            active = canvas._build_profile_data(
                canvas.profile_pts,
                color=getattr(canvas, "_active_profile_color", "#fbc02d"),
                view=canvas.views[0] if canvas.views else None,
            )
        except Exception:
            active = None
        saved = []
        try:
            for entry in getattr(canvas, "_saved_profiles", []):
                data = entry.get("data")
                if data is None:
                    data = canvas._build_profile_data(entry.get("pts"), color=entry.get("color"), view=canvas.views[0] if canvas.views else None)
                if data:
                    saved.append(data)
        except Exception:
            saved = []
        return active, saved

    def _dispatch_profile_dialog(active=None, saved=None):
        if not getattr(canvas, "_profile_user_enabled", False):
            return
        if active is None and saved is None:
            active, saved = _compute_profiles_from_canvas()
        if not active and not saved:
            return
        _ensure_profile_dialog(active, saved)

    def _refresh_profile_dialog_from_canvas():
        if not getattr(canvas, "_profile_user_enabled", False):
            return
        active, saved = _compute_profiles_from_canvas()
        if not active and not saved:
            return
        _dispatch_profile_dialog(active, saved)

    try:
        canvas.set_profile_callback(_dispatch_profile_dialog)
    except Exception:
        canvas.profile_callback = _dispatch_profile_dialog
    try:
        canvas.set_profile_state_callback(lambda _state=None: _refresh_profile_dialog_from_canvas())
    except Exception:
        canvas._profile_state_callback = lambda _state=None: _refresh_profile_dialog_from_canvas()

    def _toggle_measure(checked):
        def _force_profile_dialog():
            try:
                active = canvas._build_profile_data(
                    canvas.profile_pts,
                    color=getattr(canvas, "_active_profile_color", "#fbc02d"),
                    view=canvas.views[0] if canvas.views else None,
                )
            except Exception:
                active = None
            saved = []
            try:
                for entry in getattr(canvas, "_saved_profiles", []):
                    data = entry.get("data")
                    if data is None:
                        data = canvas._build_profile_data(entry.get("pts"), color=entry.get("color"), view=canvas.views[0] if canvas.views else None)
                    if data:
                        saved.append(data)
            except Exception:
                saved = []
            if getattr(canvas, "_profile_user_enabled", False) and (active or saved):
                _ensure_profile_dialog(active, saved)

        try:
            canvas.enable_profile(bool(checked))
            canvas._profile_user_enabled = bool(checked)
            if not checked and popup_state.get("profile_dialog"):
                popup_state["profile_dialog"].close()
            if checked:
                try:
                    a, s = _compute_profiles_from_canvas()
                    _dispatch_profile_dialog(a, s)
                except Exception:
                    pass
                try:
                    canvas._emit_profile()
                except Exception:
                    _force_profile_dialog()
                else:
                    _force_profile_dialog()
            _refresh_profile_dialog_from_canvas()
        except Exception:
            pass

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

    measure_btn.toggled.connect(_toggle_measure)
    angle_btn.toggled.connect(_toggle_angle)
    scale_btn.toggled.connect(_toggle_scale)
    layout_cb.currentTextChanged.connect(_toggle_layout)

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
    layout.addLayout(controls_bar)
    layout.addWidget(canvas, 1)
    canvas.setFocus()

    class _PopupKeyFilter(QtCore.QObject):
        def __init__(self, cvs):
            super().__init__(cvs)
            self.canvas = cvs

        def eventFilter(self, obj, event):
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
    dlg.resize(760, 620)
    dlg.show()
    owner._popup_refs.append(dlg)
    if hasattr(owner, "_update_crop_popup_actions"):
        owner._update_crop_popup_actions()

    def _on_popup_closed(_=None):
        if dlg in owner._popup_refs:
            owner._popup_refs.remove(dlg)
        if hasattr(owner, "_update_crop_popup_actions"):
            owner._update_crop_popup_actions()

    def _remove_popup_canvas(_=None):
        if canvas in getattr(owner, "_popup_canvases", []):
            owner._popup_canvases.remove(canvas)

    dlg.finished.connect(_on_popup_closed)
    dlg.finished.connect(_remove_popup_canvas)
    return dlg
