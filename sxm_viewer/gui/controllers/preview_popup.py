"""Helpers for building preview pop-out dialogs."""
from __future__ import annotations

from ..._shared import QtWidgets, QtCore
from ..canvases.detail_preview_canvas import MultiPreviewCanvas
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
    layout.setSpacing(0)
    layout.setSizeConstraint(QtWidgets.QLayout.SetNoConstraint)

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
        # Keep data undistorted in popups.
        canvas.set_fit_to_canvas(False)
    except Exception:
        pass

    source_canvas = getattr(owner, "preview_canvas", None)
    canvas.set_view_layout(getattr(source_canvas, "_view_layout", "grid"))
    try:
        canvas.set_show_profile_overlays(getattr(source_canvas, "_show_profile_overlays", True))
        canvas.set_show_angle_overlays(getattr(source_canvas, "_show_angle_overlays", True))
        canvas.set_show_shortcut_hint(getattr(source_canvas, "_show_shortcut_hint", True))
    except Exception:
        pass

    _square_resize_busy = {"active": False}
    _popup_resize_threshold_px = 2
    _last_square_target = {"w": -1, "h": -1}
    resize_sync_timer = QtCore.QTimer(dlg)
    resize_sync_timer.setSingleShot(True)
    resize_sync_timer.setInterval(40)
    resize_settle_timer = QtCore.QTimer(dlg)
    resize_settle_timer.setSingleShot(True)
    resize_settle_timer.setInterval(200)

    def _minimum_square_side():
        try:
            hint = canvas.sizeHint()
            min_hint = canvas.minimumSizeHint()
            side = max(
                int(hint.width()),
                int(hint.height()),
                int(min_hint.width()),
                int(min_hint.height()),
                152,
            )
            font_scale = float(getattr(canvas, "_view_font_scale", 1.0))
            if font_scale > 1.0:
                side += int(44.0 * (font_scale - 1.0))
            return side
        except Exception:
            return 180

    def _enforce_square_dialog():
        if _square_resize_busy["active"]:
            return
        try:
            if dlg.isMaximized() or dlg.isFullScreen():
                return
        except Exception:
            return
        try:
            if QtWidgets.QApplication.mouseButtons() != QtCore.Qt.NoButton:
                return
        except Exception:
            pass
        try:
            _square_resize_busy["active"] = True
            layout.activate()
            margins = layout.contentsMargins()
            min_side = _minimum_square_side()
            avail_w = max(1, dlg.width() - margins.left() - margins.right())
            avail_h = max(1, dlg.height() - margins.top() - margins.bottom())
            side = max(min(avail_w, avail_h), min_side)
            target_w = side + margins.left() + margins.right()
            target_h = side + margins.top() + margins.bottom()
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
                dlg.resize(int(target_w), int(target_h))
        except Exception:
            pass
        finally:
            _square_resize_busy["active"] = False

    def _resize_to_canvas(force=False):
        try:
            layout.activate()
            if force:
                dlg.adjustSize()
                dlg.setMinimumSize(0, 0)
                _last_square_target["w"] = -1
                _last_square_target["h"] = -1
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
    try:
        canvas.set_plot_font_family_callback(lambda fam: owner.set_plot_font_family(fam))
        canvas.set_plot_font_family(getattr(owner, "_plot_font_family", "sans-serif"))
    except Exception:
        pass
    def _on_popup_canvas_state_changed(_=None):
        _schedule_resize(force=False)
        try:
            if hasattr(owner, "_on_canvas_display_options_changed"):
                owner._on_canvas_display_options_changed(canvas)
        except Exception:
            pass
    canvas.set_views_callback(_on_popup_canvas_state_changed)
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
    canvas.set_histogram_dialog_callback(lambda c: owner._open_histogram_dialog(c))
    canvas.set_stp_export_callback(owner._export_view_as_stp)
    canvas.set_window_arrange_callback(owner.on_arrange_popouts)
    canvas.set_copy_feedback_handler(lambda view=None, info=None, host=dlg: owner._on_view_copied(view, info, target=host))

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

    rel_override = getattr(source_canvas, "_relative_axes_override", None)
    if rel_override is None:
        rel_override = any(bool(v.get("relative_axes")) for v in views if isinstance(v, dict))
    canvas.set_relative_axes_override(rel_override)

    frame_fill_initial = bool(getattr(source_canvas, "_frame_fill_mode", False))
    if frame_fill_initial:
        try:
            canvas.set_frame_fill_mode(True)
        except Exception:
            pass

    measure_initial = bool(
        getattr(source_canvas, "_profile_user_enabled", getattr(source_canvas, "profile_enabled", False))
    )
    angle_initial = bool(getattr(source_canvas, "angle_enabled", False))
    profile_controller = PopupProfileController(owner, canvas, title or "Profile")
    profile_controller.set_initial_state(measure_initial)
    canvas.set_angle_tool_enabled(angle_initial)

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
                        _schedule_resize(force=True)
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
