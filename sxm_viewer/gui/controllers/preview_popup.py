"""Helpers for building preview pop-out dialogs."""
from __future__ import annotations

import copy

import numpy as np

from ..._shared import QtWidgets, QtCore
from ..canvases.detail_preview_canvas import MultiPreviewCanvas
from .profile import PopupProfileController


def _popup_image_size_text(view):
    if not isinstance(view, dict) or not view:
        return ""
    override = view.get("_popup_image_size_override")
    width = None
    height = None
    unit = str(view.get("axis_unit") or "").strip()
    if isinstance(override, dict):
        try:
            width = float(override.get("width"))
            height = float(override.get("height"))
            unit = str(override.get("unit") or unit).strip()
        except Exception:
            width = None
            height = None
    extent = view.get("extent_raw")
    if extent is None:
        extent = view.get("extent")
    if (width is None or height is None) and extent is not None:
        try:
            x0, x1, y1, y0 = extent
            width = abs(float(x1) - float(x0))
            height = abs(float(y0) - float(y1))
        except Exception:
            width = None
            height = None
    if width is None or height is None:
        try:
            arr = np.asarray(view.get("arr"))
            if arr.ndim >= 2:
                height = float(arr.shape[0])
                width = float(arr.shape[1])
                unit = "px"
        except Exception:
            return ""
    if width is None or height is None:
        return ""
    if not unit:
        unit = "px" if view.get("extent") is None and view.get("extent_raw") is None else "nm"

    def _fmt(value):
        value = float(value)
        if unit == "px":
            return str(int(round(value)))
        if abs(value - round(value)) < 1e-6:
            return str(int(round(value)))
        if abs(value) >= 100:
            return f"{value:.0f}"
        if abs(value) >= 10:
            return f"{value:.1f}".rstrip("0").rstrip(".")
        return f"{value:.3g}"

    return f"{_fmt(width)} x {_fmt(height)} {unit}".strip()


def _popup_window_title(owner, dlg, view=None, *, title=None, default="Preview"):
    if view is None:
        try:
            views = getattr(getattr(dlg, "_preview_canvas", None), "views", None)
            view = (views or [None])[0]
        except Exception:
            view = None
    size_text = _popup_image_size_text(view)
    base = title
    if base is None:
        base = getattr(dlg, "_popup_title_base", None)
    if not base:
        try:
            base = owner._friendly_view_title(view, default=default)
        except Exception:
            base = default
    base = str(base or default).strip() or default
    if size_text:
        prefix = f"{size_text} | "
        if base.startswith(prefix):
            base = base[len(prefix):].strip() or default
    try:
        dlg._popup_title_base = base
    except Exception:
        pass
    return f"{size_text} | {base}" if size_text else base


def _resolve_popup_channel_source(owner, views):
    if not views or len(views) != 1:
        return None
    view = views[0] or {}
    meta = view.get("meta") or {}
    file_path = view.get("path") or meta.get("path") or meta.get("file_path")
    channel_idx = view.get("channel_idx")
    if channel_idx is None:
        channel_idx = meta.get("channel_index")
    if not file_path or channel_idx is None:
        return None
    try:
        channel_idx = int(channel_idx)
    except Exception:
        return None
    header, fds = owner.headers.get(str(file_path), (None, None))
    if header is None or not fds or len(fds) <= 1 or channel_idx < 0 or channel_idx >= len(fds):
        return None
    return {
        "file_path": str(file_path),
        "channel_idx": channel_idx,
        "header": header,
        "fds": fds,
    }


def _shift_view_relative_zero(owner, view, enabled: bool):
    def _normalized_relative_clim(arr, clim):
        if clim is None:
            auto_fn = getattr(owner, "_auto_preview_clim", None)
            if callable(auto_fn):
                try:
                    return auto_fn(arr, relative_zero=True)
                except Exception:
                    return None
            return None
        try:
            _lo, _hi = clim
            hi_val = float(_hi)
        except Exception:
            return None
        finite = np.asarray(arr, dtype=float)
        finite = finite[np.isfinite(finite)]
        if finite.size:
            hi_val = max(hi_val, float(np.nanmax(finite)))
        hi_val = max(hi_val, 0.0)
        if hi_val <= 0.0:
            return None
        return (0.0, hi_val)

    new_view = owner._copy_view_for_popup(view)
    arr = new_view.get("arr")
    if arr is None:
        new_view["display_relative_zero"] = bool(enabled)
        if not enabled:
            new_view["zero_offset"] = None
            new_view.pop("relative_zero_source_clim", None)
        return new_view
    arr_np = np.asarray(arr, dtype=float)
    is_relative = bool(new_view.get("display_relative_zero", False))
    try:
        zero_offset = float(new_view.get("zero_offset")) if new_view.get("zero_offset") is not None else None
    except Exception:
        zero_offset = None
    if enabled and not is_relative:
        finite = arr_np[np.isfinite(arr_np)]
        zero_offset = float(np.nanmin(finite)) if finite.size else 0.0
        orig_clim = new_view.get("clim")
        if orig_clim is not None:
            try:
                lo0, hi0 = orig_clim
                new_view["relative_zero_source_clim"] = (float(lo0), float(hi0))
            except Exception:
                new_view.pop("relative_zero_source_clim", None)
        new_view["arr"] = arr_np - zero_offset
        clim = new_view.get("clim")
        if clim is not None:
            try:
                lo, hi = clim
                rel_clim = _normalized_relative_clim(
                    new_view["arr"],
                    (float(lo) - zero_offset, float(hi) - zero_offset),
                )
                if rel_clim is not None:
                    new_view["clim"] = rel_clim
                else:
                    new_view.pop("clim", None)
            except Exception:
                new_view.pop("clim", None)
        else:
            rel_clim = _normalized_relative_clim(new_view["arr"], None)
            if rel_clim is not None:
                new_view["clim"] = rel_clim
    elif not enabled and is_relative:
        zero_offset = float(zero_offset or 0.0)
        new_view["arr"] = arr_np + zero_offset
        source_clim = new_view.pop("relative_zero_source_clim", None)
        if source_clim is not None:
            try:
                lo, hi = source_clim
                new_view["clim"] = (float(lo), float(hi))
            except Exception:
                new_view.pop("clim", None)
        else:
            clim = new_view.get("clim")
            if clim is not None:
                try:
                    lo, hi = clim
                    new_view["clim"] = (float(lo) + zero_offset, float(hi) + zero_offset)
                except Exception:
                    new_view.pop("clim", None)
        zero_offset = None
    else:
        if enabled and zero_offset is None:
            finite = arr_np[np.isfinite(arr_np)]
            zero_offset = float(np.nanmin(finite)) if finite.size else 0.0
        if enabled:
            if new_view.get("relative_zero_source_clim") is None and new_view.get("clim") is not None:
                try:
                    lo, hi = new_view.get("clim")
                    new_view["relative_zero_source_clim"] = (float(lo) + float(zero_offset or 0.0), float(hi) + float(zero_offset or 0.0))
                except Exception:
                    new_view.pop("relative_zero_source_clim", None)
            rel_clim = _normalized_relative_clim(new_view.get("arr"), new_view.get("clim"))
            if rel_clim is not None:
                new_view["clim"] = rel_clim
            else:
                new_view.pop("clim", None)
        else:
            new_view.pop("relative_zero_source_clim", None)
    new_view["display_relative_zero"] = bool(enabled)
    new_view["zero_offset"] = zero_offset if enabled else None
    return new_view


def _apply_popup_display_state(owner, views, *, relative_zero: bool):
    return [_shift_view_relative_zero(owner, view, relative_zero) for view in (views or [])]


def spawn_preview_popup(owner, views, title=None, *, show_immediately=True, restore_mode=False, source_canvas=None):
    """Create a preview popup dialog reusing the existing owner logic."""
    if not views:
        return None

    dlg = QtWidgets.QDialog(owner)
    try:
        # Detach the popup from the main window as an owned native dialog so
        # the viewer can be brought in front by simply clicking it.
        dlg.setParent(None, dlg.windowFlags())
        dlg.setWindowFlag(QtCore.Qt.Window, True)
        dlg.setWindowIcon(owner.windowIcon())
    except Exception:
        pass
    dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
    dlg.setWindowFlags(
        dlg.windowFlags()
        | QtCore.Qt.WindowMinimizeButtonHint
        | QtCore.Qt.WindowMaximizeButtonHint
        | QtCore.Qt.WindowSystemMenuHint
    )
    dlg.setMinimumSize(0, 0)
    dlg.setWindowTitle(_popup_window_title(owner, dlg, (views or [None])[0], title=title, default="Preview"))
    dlg._preview_resize_paused = not bool(show_immediately)

    layout = QtWidgets.QVBoxLayout(dlg)
    layout.setContentsMargins(6, 6, 6, 6)
    layout.setSpacing(0)
    layout.setSizeConstraint(QtWidgets.QLayout.SetNoConstraint)
    popup_source = _resolve_popup_channel_source(owner, views)
    popup_display_state = {
        "relative_zero": bool((views[0] or {}).get("display_relative_zero", getattr(owner, "display_units_relative", False)))
        if views else bool(getattr(owner, "display_units_relative", False))
    }

    # Use a default that we immediately adapt to the
    # aspect ratio of the underlying image so the popup
    # is created snugly around the content.
    canvas = MultiPreviewCanvas(dlg, figsize=(4, 3))
    try:
        canvas._undo_suspend_depth += 1
    except Exception:
        pass
    try:
        canvas.set_render_suspended(True)
    except Exception:
        pass
    try:
        canvas.set_compact_size_hints(True)
        canvas.setMinimumSize(0, 0)
        canvas.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        canvas._show_image_size_overlay = False
    except Exception:
        pass

    source_canvas = source_canvas or getattr(owner, "preview_canvas", None)
    rel_override = getattr(source_canvas, "_relative_axes_override", None)
    if rel_override is None:
        rel_override = any(bool(v.get("relative_axes")) for v in views if isinstance(v, dict))
    frame_fill_initial = bool(getattr(source_canvas, "_frame_fill_mode", False))
    measure_initial = bool(
        getattr(source_canvas, "_profile_user_enabled", getattr(source_canvas, "profile_enabled", False))
    )
    angle_initial = bool(getattr(source_canvas, "angle_enabled", False))
    profile_state_initial = None
    inherit_profile_state = not any((view or {}).get("crop_sequence") is not None for view in (views or []))
    if inherit_profile_state and source_canvas is not None and hasattr(source_canvas, "export_profile_state"):
        try:
            profile_state_initial = copy.deepcopy(source_canvas.export_profile_state())
        except Exception:
            profile_state_initial = None

    if restore_mode:
        try:
            canvas._show_title = bool(getattr(owner, "show_preview_title", True))
            canvas.show_molecules = bool(getattr(owner, "show_molecules", True))
            canvas._show_acquisition_overlay = bool(getattr(owner, "show_acquisition_overlay", False))
            canvas._profile_label_mode = str(getattr(owner, "profile_label_mode", "length") or "length")
            canvas._fit_to_canvas = False
            canvas._view_layout = str(getattr(source_canvas, "_view_layout", "grid") or "grid")
            canvas._show_profile_overlays = bool(getattr(source_canvas, "_show_profile_overlays", True))
            canvas._show_angle_overlays = bool(getattr(source_canvas, "_show_angle_overlays", True))
            canvas._show_shortcut_hint = bool(getattr(source_canvas, "_show_shortcut_hint", True))
            canvas._show_molecule_gizmo = bool(getattr(source_canvas, "_show_molecule_gizmo", getattr(owner, "show_molecule_gizmo", False)))
            canvas._detail_dark = bool(getattr(owner, "detail_dark_view", False))
            canvas._detail_grid = bool(getattr(owner, "detail_grid_view", False))
            font_family = str(getattr(owner, "_plot_font_family", "sans-serif") or "sans-serif")
            canvas._font_family = font_family
            settings = dict(getattr(canvas, "_scale_bar_settings", {}) or {})
            settings["font_family"] = font_family
            canvas._scale_bar_settings = settings
            canvas.scale_bar_enabled = bool(owner.scale_bar_cb.isChecked())
            if canvas.scale_bar_enabled:
                canvas._connect_scale_bar_events()
            canvas._relative_axes_override = rel_override
            canvas.molecule_palette = (getattr(owner, "molecule_palette", "avogadro") or "avogadro").lower()
            if frame_fill_initial:
                canvas._frame_fill_prev_state = {
                    "show_ticks": bool(getattr(canvas, "_show_ticks", True)),
                    "show_colorbar": bool(getattr(canvas, "_show_colorbar", True)),
                    "show_title": bool(getattr(canvas, "_show_title", True)),
                    "fit_to_canvas": bool(getattr(canvas, "_fit_to_canvas", False)),
                }
                canvas._show_ticks = False
                canvas._show_colorbar = False
                canvas._show_title = False
                canvas._fit_to_canvas = False
                canvas._frame_fill_mode = True
        except Exception:
            pass
    else:
        try:
            canvas.set_show_title(getattr(owner, "show_preview_title", True))
        except Exception:
            pass
        try:
            canvas.set_show_molecules(getattr(owner, "show_molecules", True))
        except Exception:
            pass
        try:
            canvas.set_show_molecule_gizmo(bool(getattr(source_canvas, "_show_molecule_gizmo", getattr(owner, "show_molecule_gizmo", False))))
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
        canvas.set_view_layout(getattr(source_canvas, "_view_layout", "grid"))
        try:
            canvas.set_show_profile_overlays(getattr(source_canvas, "_show_profile_overlays", True))
            canvas.set_show_angle_overlays(getattr(source_canvas, "_show_angle_overlays", True))
            canvas.set_show_shortcut_hint(getattr(source_canvas, "_show_shortcut_hint", True))
        except Exception:
            pass
        try:
            canvas._detail_dark = bool(getattr(owner, "detail_dark_view", False))
            canvas._detail_grid = bool(getattr(owner, "detail_grid_view", False))
            canvas._relative_axes_override = rel_override
            canvas.molecule_palette = (getattr(owner, "molecule_palette", "avogadro") or "avogadro").lower()
            canvas.scale_bar_enabled = bool(owner.scale_bar_cb.isChecked())
            if canvas.scale_bar_enabled:
                canvas._connect_scale_bar_events()
            if frame_fill_initial:
                canvas._frame_fill_prev_state = {
                    "show_ticks": bool(getattr(canvas, "_show_ticks", True)),
                    "show_colorbar": bool(getattr(canvas, "_show_colorbar", True)),
                    "show_title": bool(getattr(canvas, "_show_title", True)),
                    "fit_to_canvas": bool(getattr(canvas, "_fit_to_canvas", False)),
                }
                canvas._show_ticks = False
                canvas._show_colorbar = False
                canvas._show_title = False
                canvas._fit_to_canvas = False
                canvas._frame_fill_mode = True
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

    def _enforce_square_dialog(*, respect_min_side: bool = True):
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
            side = min(avail_w, avail_h)
            if respect_min_side:
                side = max(side, min_side)
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
                try:
                    if dlg.isMaximized() or dlg.isFullScreen():
                        canvas.updateGeometry()
                        canvas.draw_idle()
                        return
                except Exception:
                    pass
                dlg.adjustSize()
                dlg.setMinimumSize(0, 0)
                _last_square_target["w"] = -1
                _last_square_target["h"] = -1
                _enforce_square_dialog(respect_min_side=True)
        except Exception:
            pass

    def _enforce_square_when_idle():
        try:
            if QtWidgets.QApplication.mouseButtons() != QtCore.Qt.NoButton:
                resize_settle_timer.start()
                return
        except Exception:
            pass
        # After a user drag-resize, keep the popup square without forcing it
        # back up to the latest font-derived content hint.
        _enforce_square_dialog(respect_min_side=False)

    def _schedule_resize(force=False):
        if getattr(dlg, "_preview_resize_paused", False):
            return
        if force:
            QtCore.QTimer.singleShot(0, lambda: _resize_to_canvas(force=True))
            return
        try:
            resize_sync_timer.start()
            resize_settle_timer.start()
        except Exception:
            QtCore.QTimer.singleShot(0, lambda: _resize_to_canvas(force=False))

    def _resume_popup_resize(*, force=False):
        dlg._preview_resize_paused = False
        if force:
            _schedule_resize(force=True)

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

    canvas.set_views(_apply_popup_display_state(owner, views, relative_zero=popup_display_state["relative_zero"]))
    if profile_state_initial is not None and hasattr(canvas, "import_profile_state"):
        try:
            canvas.import_profile_state(profile_state_initial, emit=False)
            if profile_state_initial.get("saved") and hasattr(canvas, "set_show_profile_overlays"):
                canvas.set_show_profile_overlays(True)
            canvas.draw_idle()
        except Exception:
            pass
    try:
        canvas.set_plot_font_family_callback(lambda fam: owner.set_plot_font_family(fam))
        if not restore_mode:
            canvas.set_plot_font_family(getattr(owner, "_plot_font_family", "sans-serif"))
    except Exception:
        pass
    def _on_popup_canvas_state_changed(_=None):
        if not getattr(canvas, "_popup_style_resize_lock", False):
            _schedule_resize(force=False)
        try:
            dlg.setWindowTitle(_popup_window_title(owner, dlg))
        except Exception:
            pass
        try:
            if hasattr(owner, "_on_canvas_display_options_changed"):
                owner._on_canvas_display_options_changed(canvas)
        except Exception:
            pass
    canvas.set_views_callback(_on_popup_canvas_state_changed)
    canvas.set_crop_callback(lambda v, c=canvas: owner._on_preview_crop(v, c))
    canvas.set_virtual_copy_callback(lambda v: owner._create_virtual_copy_from_popup_view(v))
    canvas.set_double_click_callback(
        lambda v=None: spawn_preview_popup(
            owner,
            [owner._copy_view_for_popup(v)] if v else [],
            title=owner._friendly_view_title(v, default="Preview copy") if v else "Preview copy",
            source_canvas=canvas,
        )
    )
    canvas.set_filter_menu_callback(lambda menu, view, c=canvas: owner._populate_canvas_filter_menu(menu, c, view))
    canvas.set_histogram_dialog_callback(lambda c: owner._open_histogram_dialog(c))
    canvas.set_histogram_auto_callback(lambda c: owner._auto_contrast(c))
    canvas.set_histogram_reset_callback(lambda c: owner._reset_contrast(c))
    canvas.set_compare_menu_callback(
        lambda action, view, c=canvas: owner.on_compare_menu_action(action, view, c),
        state_cb=owner.compare_menu_state,
    )
    if hasattr(canvas, "set_collection_menu_callback"):
        canvas.set_collection_menu_callback(
            lambda action, view, c=canvas: owner.collection_controller.handle_canvas_menu_action(action, view, c),
            help_cb=owner.on_collection_help,
        )
    canvas.set_stp_export_callback(owner._export_view_as_stp)
    canvas.set_window_arrange_callback(owner.on_arrange_popouts)
    canvas.set_window_minimize_callback(owner.on_minimize_popouts)
    canvas.set_window_restore_callback(owner.on_restore_popouts)
    canvas.set_window_close_callback(owner.on_close_popouts)
    canvas.set_copy_feedback_handler(lambda view=None, info=None, host=dlg: owner._on_view_copied(view, info, target=host))
    canvas.set_display_relative_zero_menu_callback(
        lambda enabled: _set_popup_relative_zero(enabled),
        state_cb=lambda: popup_display_state["relative_zero"],
        tooltip="Display values relative to the current zero/reference",
    )
    canvas.set_apply_popup_style_callback(
        lambda: owner._apply_popup_style_to_all(canvas),
        tooltip="Copy font size, typography and display layout from this popup to the other open pop-outs",
    )

    seq = views[0].get("crop_sequence") if views else None
    if hasattr(owner, "quick_crop_controller"):
        owner.quick_crop_controller.register_popup(seq, dlg)
    canvas.enable_fixed_crop_quick_mode(owner.quick_crop_mode)
    canvas.show_fixed_crop_template(bool(owner.show_crop_template_overlay))
    canvas.show_fixed_crop_history(owner.show_crop_history_overlay)
    try:
        canvas.set_molecule_palette_callback(owner._on_molecule_palette_changed)
        if hasattr(canvas, "set_recent_molecule_callback"):
            canvas._recent_molecule_paths = list(getattr(owner, "recent_molecules", []) or [])
            canvas.set_recent_molecule_callback(owner._on_recent_molecules_updated)
        owner._popup_canvases.append(canvas)
    except Exception:
        pass
    profile_controller = PopupProfileController(owner, canvas, title or "Profile")
    canvas.export_profile_dialog_state = profile_controller.export_dialog_state
    canvas.restore_profile_dialog_state = profile_controller.restore_dialog_state
    if not restore_mode:
        profile_controller.set_initial_state(measure_initial)
        canvas.set_angle_tool_enabled(angle_initial)

    try:
        canvas._undo_suspend_depth = max(0, getattr(canvas, "_undo_suspend_depth", 0) - 1)
    except Exception:
        pass

    def _set_popup_relative_zero(enabled):
        popup_display_state["relative_zero"] = bool(enabled)
        canvas._popup_relative_zero_enabled = bool(enabled)
        if not getattr(canvas, "views", None):
            return
        canvas.set_views(
            _apply_popup_display_state(owner, canvas.views, relative_zero=popup_display_state["relative_zero"]),
            preserve_profiles=True,
        )

    canvas._popup_relative_zero_enabled = bool(popup_display_state["relative_zero"])
    canvas._popup_relative_zero_setter = _set_popup_relative_zero

    popup_header = None
    if popup_source is not None:
        popup_header = QtWidgets.QWidget(dlg)
        popup_header_layout = QtWidgets.QHBoxLayout(popup_header)
        popup_header_layout.setContentsMargins(0, 0, 0, 6)
        popup_header_layout.setSpacing(6)
        popup_header_layout.addWidget(QtWidgets.QLabel("Channel", popup_header))
        channel_prev_btn = QtWidgets.QToolButton(popup_header)
        channel_prev_btn.setArrowType(QtCore.Qt.LeftArrow)
        channel_prev_btn.setAutoRaise(True)
        channel_prev_btn.setToolTip("Previous channel in this popup")
        popup_header_layout.addWidget(channel_prev_btn)
        channel_combo = QtWidgets.QComboBox(popup_header)
        channel_combo.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        for idx, fd in enumerate(popup_source["fds"]):
            label = fd.get("Caption", fd.get("FileName", f"chan{idx}"))
            channel_combo.addItem(f"{idx}: {label}")
            channel_combo.setItemData(idx, label, QtCore.Qt.ToolTipRole)
        channel_combo.setCurrentIndex(int(popup_source["channel_idx"]))
        popup_header_layout.addWidget(channel_combo, 1)
        channel_next_btn = QtWidgets.QToolButton(popup_header)
        channel_next_btn.setArrowType(QtCore.Qt.RightArrow)
        channel_next_btn.setAutoRaise(True)
        channel_next_btn.setToolTip("Next channel in this popup")
        popup_header_layout.addWidget(channel_next_btn)
        layout.addWidget(popup_header)

        channel_sync = {"active": False}

        def _sync_popup_channel_buttons():
            idx = channel_combo.currentIndex()
            count = channel_combo.count()
            channel_prev_btn.setEnabled(idx > 0)
            channel_next_btn.setEnabled(0 <= idx < count - 1)

        def _apply_popup_channel(idx):
            if channel_sync["active"]:
                return
            try:
                idx = int(idx)
            except Exception:
                return
            if idx < 0 or idx >= channel_combo.count():
                return
            current_view = canvas.views[0] if getattr(canvas, "views", None) else {}
            cmap_override = current_view.get("cmap")
            try:
                zoom_states = canvas.export_zoom_states()
            except Exception:
                zoom_states = None
            try:
                bundle = owner._build_single_channel_view(
                    popup_source["file_path"],
                    idx,
                    cmap_override=cmap_override,
                    use_local_cmap=True,
                )
            except Exception:
                bundle = None
            if not bundle:
                return
            next_view = bundle["view"]
            seq = current_view.get("crop_sequence")
            if seq is not None and source_canvas is not None and hasattr(source_canvas, "get_fixed_crop_history_entry"):
                try:
                    crop_entry = source_canvas.get_fixed_crop_history_entry(seq)
                except Exception:
                    crop_entry = None
                if crop_entry is not None:
                    try:
                        rebuilt_view = canvas.build_view_from_fixed_crop_entry(next_view, crop_entry)
                    except Exception:
                        rebuilt_view = None
                    if rebuilt_view:
                        next_view = rebuilt_view
            channel_sync["active"] = True
            try:
                popup_source["channel_idx"] = idx
                canvas._popup_channel_source = dict(popup_source)
                preserve_profiles = True
                canvas.set_views(
                    _apply_popup_display_state(
                        owner,
                        [next_view],
                        relative_zero=popup_display_state["relative_zero"],
                    ),
                    preserve_profiles=preserve_profiles,
                )
                if zoom_states:
                    try:
                        canvas.apply_zoom_states(zoom_states)
                    except Exception:
                        pass
                dlg.setWindowTitle(
                    _popup_window_title(
                        owner,
                        dlg,
                        next_view,
                        title=owner._friendly_view_title(next_view, default="Preview"),
                        default="Preview",
                    )
                )
                channel_combo.blockSignals(True)
                channel_combo.setCurrentIndex(idx)
                channel_combo.blockSignals(False)
                _sync_popup_channel_buttons()
                # Preserve the user's manually resized popup size when
                # switching channels; only grow if the refreshed content
                # now requires more room.
                _schedule_resize(force=False)
            finally:
                channel_sync["active"] = False

        channel_combo.currentIndexChanged.connect(_apply_popup_channel)
        channel_prev_btn.clicked.connect(lambda: channel_combo.setCurrentIndex(max(0, channel_combo.currentIndex() - 1)))
        channel_next_btn.clicked.connect(lambda: channel_combo.setCurrentIndex(min(channel_combo.count() - 1, channel_combo.currentIndex() + 1)))
        _sync_popup_channel_buttons()
        canvas._popup_channel_source = dict(popup_source)
        dlg._preview_channel_combo = channel_combo

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
                        if dlg.isMaximized() or dlg.isFullScreen():
                            return False
                        _schedule_resize(force=True)
                except Exception:
                    pass
                return False
            if event.type() in (QtCore.QEvent.WindowActivate, QtCore.QEvent.MouseButtonPress, QtCore.QEvent.FocusIn):
                try:
                    if hasattr(owner, "_set_active_preview_popup"):
                        owner._set_active_preview_popup(dlg, self.canvas)
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
                            source_canvas=self.canvas,
                        )
                    except Exception:
                        pass
                    event.accept()
                    return True
            return False

    key_filter = _PopupKeyFilter(canvas)
    dlg.installEventFilter(key_filter)
    canvas.installEventFilter(key_filter)

    dlg._preview_popup_schedule_resize = _schedule_resize
    dlg._resume_preview_resize = _resume_popup_resize
    dlg._preview_canvas = canvas
    dlg._popup_title_base = str(title or "").strip() or owner._friendly_view_title((views or [None])[0], default="Preview")
    if show_immediately:
        dlg.show()
        if not restore_mode:
            try:
                canvas.set_render_suspended(False)
            except Exception:
                pass
        _resume_popup_resize(force=True)
        if hasattr(owner, "_set_active_preview_popup"):
            try:
                owner._set_active_preview_popup(dlg, canvas)
            except Exception:
                pass
    elif not restore_mode:
        try:
            canvas.set_render_suspended(False)
        except Exception:
            pass
    owner._popup_refs.append(dlg)
    if hasattr(owner, "quick_crop_controller"):
        owner.quick_crop_controller.update_popup_actions()

    def _on_popup_closed(_=None):
        remember_cb = getattr(owner, "_remember_closed_preview_popup", None)
        if callable(remember_cb):
            try:
                remember_cb(dlg, canvas)
            except Exception:
                pass
        if dlg in owner._popup_refs:
            owner._popup_refs.remove(dlg)
        if hasattr(owner, "quick_crop_controller"):
            owner.quick_crop_controller.update_popup_actions()
        profile_controller.dispose()
        if hasattr(owner, "_clear_active_preview_popup"):
            try:
                owner._clear_active_preview_popup(dlg)
            except Exception:
                pass

    def _remove_popup_canvas(_=None):
        if canvas in getattr(owner, "_popup_canvases", []):
            owner._popup_canvases.remove(canvas)

    dlg.finished.connect(_on_popup_closed)
    dlg.finished.connect(_remove_popup_canvas)
    return dlg
