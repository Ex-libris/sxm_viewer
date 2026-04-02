"""Measurement helpers for SXMGridViewer."""
from __future__ import annotations

from ..._shared import (
    QtCore,
    QtGui,
    QtWidgets,
    QIcon,
    QPixmap,
    QImage,
    QPainter,
    QPen,
    QBrush,
    FigureCanvas,
    Figure,
    Line2D,
    colormaps,
    np,
    Path,
    defaultdict,
    OrderedDict,
    datetime,
    hashlib,
    itertools,
    io,
    json,
    math,
    os,
    sys,
    threading,
    _scipy_ndimage,
    log_status,
    matplotlib,
)
from ..detail_panels import ProfileDialog

def _on_start_profile(viewer, force_enable=False):
    # toggle interactive line profile mode
    views = getattr(viewer.preview_canvas, 'views', [])
    if not views:
        if force_enable:
            viewer._pending_profile_enable = True
        elif not getattr(viewer, '_pending_profile_enable', False):
            QtWidgets.QMessageBox.information(viewer, "Measure profile", "No image to measure. Load a channel first.")
        return
    viewer._pending_profile_enable = False
    active = getattr(viewer.preview_canvas, 'profile_enabled', False)
    if force_enable and active:
        return
    if not active:
        # enter profile mode
        viewer._disable_angle_mode()
        viewer.preview_canvas.set_profile_callback(viewer._on_profile_updated)
        if hasattr(viewer.preview_canvas, 'set_profile_highlight_callback'):
            viewer.preview_canvas.set_profile_highlight_callback(viewer._on_canvas_overlay_highlight)
        viewer.preview_canvas.enable_profile(True)
        try:
            viewer.preview_canvas.setFocus(QtCore.Qt.OtherFocusReason)
        except Exception:
            pass
        try: viewer.measure_profile_btn.setText('Exit profile')
        except Exception: pass
        viewer.meta_box.setPlainText("Profile mode: drag the yellow endpoints on the main image. Close to exit.")
        viewer._profile_dialog = None
    elif not force_enable:
        viewer._disable_profile_mode()


def _on_start_angle(viewer, force_enable=False):
    canvas = getattr(viewer, 'preview_canvas', None)
    if canvas is None:
        return
    views = getattr(canvas, 'views', [])
    if not views:
        if force_enable:
            viewer._pending_angle_enable = True
        elif not getattr(viewer, '_pending_angle_enable', False):
            QtWidgets.QMessageBox.information(viewer, "Measure angle", "No image to measure. Load a channel first.")
        return
    viewer._pending_angle_enable = False
    active = bool(getattr(canvas, 'angle_enabled', False))
    if force_enable and active:
        return
    if not active:
        viewer._disable_profile_mode()
        if hasattr(canvas, 'set_angle_callback'):
            canvas.set_angle_callback(viewer._on_angle_updated)
        canvas.enable_angle(True)
        try:
            canvas.setFocus(QtCore.Qt.OtherFocusReason)
        except Exception:
            pass
        try:
            viewer.measure_angle_btn.setText('Exit angle')
        except Exception:
            pass
        if hasattr(viewer, 'angle_value_label'):
            viewer.angle_value_label.setText("Angle: --")
    else:
        viewer._disable_angle_mode()


def _disable_profile_mode(viewer):
    canvas = getattr(viewer, 'preview_canvas', None)
    if canvas is None:
        return
    viewer._pending_profile_enable = False
    try:
        canvas.enable_profile(False)
    except Exception:
        pass
    try:
        if hasattr(canvas, 'set_profile_highlight_callback'):
            canvas.set_profile_highlight_callback(None)
    except Exception:
        pass
    try:
        if hasattr(canvas, 'clear_saved_profiles'):
            canvas.clear_saved_profiles()
    except Exception:
        pass
    try:
        viewer.measure_profile_btn.setText('Measure profile')
    except Exception:
        pass
    try:
        if hasattr(viewer, '_profile_dialog') and viewer._profile_dialog is not None:
            viewer._profile_dialog.close()
            viewer._profile_dialog = None
    except Exception:
        pass
    viewer._disable_angle_mode()


def _disable_angle_mode(viewer, reset_button=True):
    canvas = getattr(viewer, 'preview_canvas', None)
    if canvas is not None and hasattr(canvas, 'enable_angle'):
        try:
            canvas.enable_angle(False)
        except Exception:
            pass
    if reset_button:
        try:
            viewer.measure_angle_btn.setText('Measure angle')
        except Exception:
            pass
    if hasattr(viewer, 'angle_value_label'):
        viewer.angle_value_label.setText("Angle: --")


def _on_exit_profile_mode(viewer):
    viewer._disable_profile_mode()


def _on_profile_dialog_closed(viewer):
    dlg = getattr(viewer, '_profile_dialog', None)
    remember_cb = getattr(viewer, "_remember_closed_main_profile_dialog", None)
    if callable(remember_cb):
        try:
            remember_cb(dlg)
        except Exception:
            pass
    canvas = getattr(viewer, 'preview_canvas', None)
    if canvas is not None:
        try:
            if hasattr(canvas, 'deactivate_profile_tool'):
                canvas.deactivate_profile_tool(clear_active=True, clear_saved=True)
            else:
                canvas.enable_profile(False)
        except Exception:
            pass
        try:
            if hasattr(canvas, 'set_profile_highlight_callback'):
                canvas.set_profile_highlight_callback(None)
        except Exception:
            pass
    viewer._pending_profile_enable = False
    try:
        viewer.measure_profile_btn.setText('Measure profile')
    except Exception:
        pass
    viewer._profile_dialog = None


def _on_clear_profile_measurement(viewer):
    canvas = getattr(viewer, 'preview_canvas', None)
    if canvas is None:
        return
    was_enabled = getattr(canvas, 'profile_enabled', False)
    try:
        canvas.profile_pts = None
    except Exception:
        pass
    try:
        if hasattr(canvas, 'clear_saved_profiles'):
            canvas.clear_saved_profiles()
    except Exception:
        pass
    try:
        if hasattr(viewer, '_profile_dialog') and viewer._profile_dialog is not None:
            viewer._profile_dialog.close()
            viewer._profile_dialog = None
    except Exception:
        pass
    if was_enabled:
        try:
            canvas.enable_profile(False)
            canvas.enable_profile(True)
            viewer.measure_profile_btn.setText('Exit profile')
        except Exception:
            pass
    else:
        try:
            viewer.measure_profile_btn.setText('Measure profile')
        except Exception:
            pass
    if hasattr(canvas, 'clear_angle_measurement'):
        try:
            canvas.clear_angle_measurement()
        except Exception:
            pass


def _on_profile_updated(viewer, active_profile, saved_profiles):
    # create or update a persistent profile dialog
    try:
        if not active_profile and not saved_profiles:
            viewer._last_profile_payload = None
            if hasattr(viewer, '_profile_dialog') and viewer._profile_dialog is not None:
                try:
                    viewer._profile_dialog.close()
                except Exception:
                    pass
                viewer._profile_dialog = None
            return
        viewer._last_profile_payload = (active_profile, list(saved_profiles or []))
        y_label = None
        try:
            if viewer.last_preview:
                file_key, channel_idx = viewer.last_preview
                header, fds = viewer.headers.get(str(file_key), (None, None))
                if fds and 0 <= int(channel_idx) < len(fds):
                    fd = fds[int(channel_idx)]
                    y_label = fd.get('Caption', fd.get('FileName', f"chan{channel_idx}"))
        except Exception:
            y_label = None
        ref_unit = None
        if active_profile:
            ref_unit = active_profile.get('unit')
        elif saved_profiles:
            ref_unit = saved_profiles[0].get('unit')
        activate_cb = None
        canvas = getattr(viewer, 'preview_canvas', None)
        if canvas is not None and hasattr(canvas, 'activate_saved_profile'):
            def _activate(idx):
                try:
                    if canvas.activate_saved_profile(idx):
                        return True
                except Exception:
                    pass
                return False
            activate_cb = _activate
        highlight_cb = None
        if canvas is not None and hasattr(canvas, 'highlight_saved_profile'):
            def _highlight(idx):
                try:
                    canvas.highlight_saved_profile(idx)
                    return True
                except Exception:
                    return False
            highlight_cb = _highlight
        marker_select_cb = None
        if canvas is not None and hasattr(canvas, 'set_profile_marker_key'):
            def _marker_select(idx):
                try:
                    canvas.set_profile_marker_key(idx)
                except Exception:
                    pass
            marker_select_cb = _marker_select
        def _set_preserve(enabled):
            try:
                viewer.preserve_profiles_on_channel_change = bool(enabled)
                viewer.config['preserve_profiles_on_channel_change'] = bool(enabled)
                save_config(viewer.config)
            except Exception:
                pass
        delete_cb = None
        if canvas is not None and hasattr(canvas, 'remove_saved_profile'):
            def _delete(idx):
                try:
                    return canvas.remove_saved_profile(idx)
                except Exception:
                    return False
            delete_cb = _delete
        add_overlay_cb = None
        if canvas is not None and hasattr(canvas, 'snapshot_active_profile'):
            def _add_overlay():
                try:
                    canvas.snapshot_active_profile()
                except Exception:
                    pass
            add_overlay_cb = _add_overlay
        label_scale_cb = None
        if canvas is not None and hasattr(canvas, 'set_profile_label_scale'):
            label_scale_cb = canvas.set_profile_label_scale
        marker_update_cb = None
        if canvas is not None and hasattr(canvas, 'set_profile_marker_positions'):
            def _marker_update(positions, domain):
                try:
                    canvas.set_profile_marker_positions(positions, domain=domain, emit=False)
                except Exception:
                    pass
            marker_update_cb = _marker_update
        style_update_cb = None
        if canvas is not None and hasattr(canvas, 'set_profile_style'):
            def _style_update(profile_key, **changes):
                try:
                    return canvas.set_profile_style(profile_key, **changes)
                except Exception:
                    return False
            style_update_cb = _style_update
        palette_cb = None
        if canvas is not None and hasattr(canvas, 'apply_profile_palette'):
            def _palette_update(name):
                try:
                    return canvas.apply_profile_palette(name)
                except Exception:
                    return False
            palette_cb = _palette_update
        if not hasattr(viewer, '_profile_dialog') or viewer._profile_dialog is None:
            dark_pref = bool(getattr(viewer, 'dark_mode', False))
            viewer._profile_dialog = ProfileDialog(active_profile, saved_profiles, parent=viewer, unit=ref_unit, y_label=y_label,
                                                  activate_overlay_callback=activate_cb,
                                                  highlight_overlay_callback=highlight_cb,
                                                  label_scale_callback=label_scale_cb,
                                                  delete_overlay_callback=delete_cb,
                                                  marker_update_callback=marker_update_cb,
                                                  marker_select_callback=marker_select_cb,
                                                  add_overlay_callback=add_overlay_cb,
                                                  style_update_callback=style_update_cb,
                                                  palette_callback=palette_cb,
                                                  dark_mode=dark_pref)
            if hasattr(viewer._profile_dialog, "detach_as_workspace_window"):
                viewer._profile_dialog.detach_as_workspace_window()
            if hasattr(viewer._profile_dialog, 'set_preserve_profiles_callback'):
                viewer._profile_dialog.set_preserve_profiles_callback(
                    _set_preserve, enabled=getattr(viewer, 'preserve_profiles_on_channel_change', True)
                )
            try:
                viewer._profile_dialog.move(viewer._next_popup_pos(offset=30))
            except Exception:
                pass
            viewer._profile_dialog.finished.connect(lambda _=None: _on_profile_dialog_closed(viewer))
            viewer._profile_dialog.show()
        else:
            if hasattr(viewer._profile_dialog, 'set_label_scale_callback'):
                viewer._profile_dialog.set_label_scale_callback(label_scale_cb)
            if hasattr(viewer._profile_dialog, 'set_marker_update_callback'):
                viewer._profile_dialog.set_marker_update_callback(marker_update_cb)
            if hasattr(viewer._profile_dialog, 'set_marker_select_callback'):
                viewer._profile_dialog.set_marker_select_callback(marker_select_cb)
            if hasattr(viewer._profile_dialog, 'set_add_overlay_callback'):
                viewer._profile_dialog.set_add_overlay_callback(add_overlay_cb)
            if hasattr(viewer._profile_dialog, 'set_style_update_callback'):
                viewer._profile_dialog.set_style_update_callback(style_update_cb)
            if hasattr(viewer._profile_dialog, 'set_palette_callback'):
                viewer._profile_dialog.set_palette_callback(palette_cb)
            if hasattr(viewer._profile_dialog, 'set_preserve_profiles_callback'):
                viewer._profile_dialog.set_preserve_profiles_callback(
                    _set_preserve, enabled=getattr(viewer, 'preserve_profiles_on_channel_change', True)
                )
            viewer._profile_dialog.update_profiles(active_profile, saved_profiles,
                                                  activate_overlay_callback=activate_cb,
                                                  highlight_overlay_callback=highlight_cb)
        if canvas is not None and hasattr(viewer._profile_dialog, 'set_context_source'):
            try:
                viewer._profile_dialog.set_context_source(canvas, dark=viewer.detail_dark_view, grid=viewer.detail_grid_view)
            except Exception:
                pass
        if canvas is not None and hasattr(canvas, 'set_profile_marker_callback'):
            def _marker_from_canvas(positions, domain):
                dlg = getattr(viewer, '_profile_dialog', None)
                if dlg is None:
                    return
                try:
                    dlg.set_marker_positions(positions, domain=domain)
                except Exception:
                    pass
            try:
                canvas.set_profile_marker_callback(_marker_from_canvas)
            except Exception:
                pass
    except Exception as exc:
        try:
            log_status(f"Profile dialog error: {exc}")
        except Exception:
            pass


def _on_angle_updated(viewer, info):
    if not hasattr(viewer, 'angle_value_label'):
        return
    if not info:
        viewer.angle_value_label.setText("Angle: --")
        return
    frame_index = info.get('frame_index')
    total_frames = info.get('total_frames')
    angle_text = "Angle"
    if frame_index is not None and total_frames:
        angle_text += f" ({frame_index + 1}/{total_frames})"
    angle_text += f": {info.get('angle_deg', 0.0):.2f}"
    unit = info.get('unit')
    if unit:
        angle_text += f" | L1={info.get('len_a', 0.0):.3f} {unit} L2={info.get('len_b', 0.0):.3f} {unit}"
    viewer.angle_value_label.setText(angle_text)


def _on_show_profile_window(viewer):
    try:
        log_status("Show profile window requested.")
    except Exception:
        pass
    dlg = getattr(viewer, '_profile_dialog', None)
    if dlg is not None:
        canvas = getattr(viewer, 'preview_canvas', None)
        if canvas is not None and hasattr(canvas, 'export_profile_datasets'):
            try:
                active, saved = canvas.export_profile_datasets()
                if active or saved:
                    viewer._on_profile_updated(active, saved)
            except Exception:
                pass
        try:
            log_status("Profile dialog already exists; showing.")
        except Exception:
            pass
        dlg.show()
        try:
            dlg.raise_()
            dlg.activateWindow()
        except Exception:
            pass
        return
    canvas = getattr(viewer, 'preview_canvas', None)
    if canvas is not None and hasattr(canvas, 'export_profile_datasets'):
        try:
            active, saved = canvas.export_profile_datasets()
            if active or saved:
                try:
                    log_status("Rebuilding profile dialog from current canvas datasets.")
                except Exception:
                    pass
                viewer._on_profile_updated(active, saved)
                return
        except Exception:
            pass
    payload = getattr(viewer, '_last_profile_payload', None)
    if payload and (payload[0] or payload[1]):
        try:
            log_status("Rebuilding profile dialog from last payload.")
        except Exception:
            pass
        viewer._on_profile_updated(payload[0], payload[1])
    else:
        try:
            log_status("No profile data available for dialog.")
        except Exception:
            pass
        # Open an empty dialog so users can still access the preview/context view.
        y_label = None
        ref_unit = None
        try:
            if viewer.last_preview:
                file_key, channel_idx = viewer.last_preview
                header, fds = viewer.headers.get(str(file_key), (None, None))
                if fds and 0 <= int(channel_idx) < len(fds):
                    fd = fds[int(channel_idx)]
                    y_label = fd.get('Caption', fd.get('FileName', f"chan{channel_idx}"))
                    ref_unit = fd.get('PhysUnit') or None
        except Exception:
            y_label = None
            ref_unit = None
        try:
            viewer._profile_dialog = ProfileDialog(None, [], parent=viewer, unit=ref_unit, y_label=y_label)
            if hasattr(viewer._profile_dialog, "detach_as_workspace_window"):
                viewer._profile_dialog.detach_as_workspace_window()
            viewer._profile_dialog.move(viewer._next_popup_pos(offset=30))
            viewer._profile_dialog.show()
            if canvas is not None and hasattr(viewer._profile_dialog, 'set_context_source'):
                viewer._profile_dialog.set_context_source(
                    canvas,
                    dark=getattr(viewer, 'detail_dark_view', False),
                    grid=getattr(viewer, 'detail_grid_view', False),
                )
        except Exception:
            QtWidgets.QMessageBox.information(
                viewer, "Profile measurement",
                "No profile data available. Start measuring first."
            )


def _on_canvas_overlay_highlight(viewer, idx):
    dlg = getattr(viewer, '_profile_dialog', None)
    if dlg is not None:
        try:
            dlg.select_overlay(idx)
        except Exception:
            pass


def export_profile_dialog_state(viewer):
    """Capture open state and geometry for the main Profile measurement window."""
    dlg = getattr(viewer, "_profile_dialog", None)
    if dlg is None:
        return None
    try:
        if not dlg.isVisible():
            return None
    except Exception:
        pass
    state = {"open": True}
    try:
        geo = dlg.geometry()
        state["geometry"] = [int(geo.x()), int(geo.y()), int(geo.width()), int(geo.height())]
    except Exception:
        pass
    try:
        state["window_state"] = int(dlg.windowState())
    except Exception:
        pass
    return state


def restore_profile_dialog_state(viewer, state):
    """Restore the main Profile measurement window from session state."""
    if not isinstance(state, dict) or not bool(state.get("open")):
        return None
    _on_show_profile_window(viewer)
    dlg = getattr(viewer, "_profile_dialog", None)
    if dlg is None:
        return None
    geom = state.get("geometry")
    if geom and len(geom) == 4:
        try:
            x, y, w, h = [int(v) for v in geom]
            dlg.setGeometry(x, y, w, h)
        except Exception:
            pass
    window_state = state.get("window_state")
    if window_state is not None:
        try:
            dlg.setWindowState(QtCore.Qt.WindowStates(int(window_state)))
        except Exception:
            pass
    try:
        dlg.show()
    except Exception:
        pass
    return dlg
__all__ = [
    "_on_start_profile",
    "_on_start_angle",
    "_disable_profile_mode",
    "_disable_angle_mode",
    "_on_exit_profile_mode",
    "_on_clear_profile_measurement",
    "_on_profile_updated",
    "_on_angle_updated",
    "_on_show_profile_window",
    "_on_canvas_overlay_highlight",
    "export_profile_dialog_state",
    "restore_profile_dialog_state",
]




