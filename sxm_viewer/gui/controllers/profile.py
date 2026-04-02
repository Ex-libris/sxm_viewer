"""Profile dialog coordination helpers for preview pop-outs."""
from __future__ import annotations

from typing import Optional, Tuple, List

from ..._shared import QtCore, QtWidgets
from ..dialogs.profile_dialog import ProfileDialog


class PopupProfileController:
    """Manages profile dialog state for a popup canvas."""

    def __init__(self, owner, canvas, title: Optional[str] = None):
        self.owner = owner
        self.canvas = canvas
        self.title = title or "Profile"
        self._dialog: Optional[ProfileDialog] = None
        self._install_callbacks()

    # ------------------------------------------------------------------
    def _install_callbacks(self):
        def _profile_cb(active, saved):
            self.dispatch_dialog(active, saved)

        def _state_cb(_state=None):
            self.refresh_from_canvas()

        try:
            self.canvas.set_profile_callback(_profile_cb)
        except Exception:
            self.canvas.profile_callback = _profile_cb
        try:
            self.canvas.set_profile_state_callback(_state_cb)
        except Exception:
            self.canvas._profile_state_callback = _state_cb

    # ------------------------------------------------------------------
    def dispose(self):
        if self._dialog:
            try:
                self._deregister_dialog(self._dialog)
                self._dialog.close()
            except Exception:
                pass
            self._dialog = None

    # ------------------------------------------------------------------
    def set_initial_state(self, enabled: bool):
        try:
            self.canvas.enable_profile(bool(enabled))
        except Exception:
            pass
        try:
            self.canvas._profile_user_enabled = bool(enabled)
        except Exception:
            pass

    def _profile_updates_enabled(self) -> bool:
        """Keep the dialog live while a profile exists in move-only mode after Ctrl-draw."""
        canvas = self.canvas
        return bool(
            getattr(canvas, "_profile_user_enabled", False)
            or getattr(canvas, "profile_enabled", False)
            or getattr(canvas, "_profile_move_only", False)
        )

    # ------------------------------------------------------------------
    def toggle_measure(self, checked: bool):
        def _force_dialog():
            active, saved = self._compute_profiles_from_canvas()
            if active or saved:
                self._ensure_profile_dialog(active, saved)

        try:
            self.canvas.enable_profile(bool(checked))
            self.canvas._profile_user_enabled = bool(checked)
        except Exception:
            pass
        if not checked:
            self.dispose()
            return
        self.refresh_from_canvas()
        try:
            self.canvas._emit_profile()
        except Exception:
            pass
        _force_dialog()

    # ------------------------------------------------------------------
    def dispatch_dialog(self, active=None, saved=None):
        if not self._profile_updates_enabled():
            return
        if active is None and saved is None:
            active, saved = self._compute_profiles_from_canvas()
        if not active and not saved:
            return
        self._ensure_profile_dialog(active, saved)

    # ------------------------------------------------------------------
    def refresh_from_canvas(self):
        if not self._profile_updates_enabled():
            return
        active, saved = self._compute_profiles_from_canvas()
        if not active and not saved:
            return
        self._ensure_profile_dialog(active, saved)

    # ------------------------------------------------------------------
    def _compute_profiles_from_canvas(self) -> Tuple[Optional[dict], List[dict]]:
        canvas = self.canvas
        if not getattr(canvas, "views", None):
            return None, []
        try:
            active = canvas._build_profile_data(
                canvas.profile_pts,
                color=getattr(canvas, "_active_profile_color", "#fbc02d"),
                lw=getattr(canvas, "_active_profile_lw", 2.0),
                line_style=getattr(canvas, "_active_profile_line_style", "-"),
                marker_style=getattr(canvas, "_active_profile_marker_style", "o"),
                marker_size=getattr(canvas, "_active_profile_marker_size", 7.0),
                view=canvas.views[0] if canvas.views else None,
                live_profile_ref=canvas._profile_live_ref(None) if hasattr(canvas, "_profile_live_ref") else None,
            )
        except Exception:
            active = None
        saved: List[dict] = []
        try:
            for entry in getattr(canvas, "_saved_profiles", []):
                data = entry.get("data")
                if data is None:
                    data = canvas._build_profile_data(
                        entry.get("pts"),
                        color=entry.get("color"),
                        lw=entry.get("lw"),
                        line_style=entry.get("line_style"),
                        marker_style=entry.get("marker_style"),
                        marker_size=entry.get("marker_size"),
                        view=canvas.views[0] if canvas.views else None,
                        live_profile_ref=canvas._profile_live_ref(entry=entry) if hasattr(canvas, "_profile_live_ref") else None,
                    )
                if data:
                    saved.append(data)
        except Exception:
            saved = []
        return active, saved

    # ------------------------------------------------------------------
    def _build_dialog_callbacks(self):
        """Mirror the main profile-dialog callbacks for popup canvases."""
        canvas = self.canvas
        activate_cb = None
        if canvas is not None and hasattr(canvas, "activate_saved_profile"):
            def _activate(idx):
                try:
                    if canvas.activate_saved_profile(idx):
                        return True
                except Exception:
                    pass
                return False

            activate_cb = _activate

        highlight_cb = None
        if canvas is not None and hasattr(canvas, "highlight_saved_profile"):
            def _highlight(idx):
                try:
                    canvas.highlight_saved_profile(idx)
                    return True
                except Exception:
                    return False

            highlight_cb = _highlight

        marker_select_cb = None
        if canvas is not None and hasattr(canvas, "set_profile_marker_key"):
            def _marker_select(idx):
                try:
                    canvas.set_profile_marker_key(idx)
                except Exception:
                    pass

            marker_select_cb = _marker_select

        delete_cb = None
        if canvas is not None and hasattr(canvas, "remove_saved_profile"):
            def _delete(idx):
                try:
                    return canvas.remove_saved_profile(idx)
                except Exception:
                    return False

            delete_cb = _delete

        add_overlay_cb = None
        if canvas is not None and hasattr(canvas, "snapshot_active_profile"):
            def _add_overlay():
                try:
                    canvas.snapshot_active_profile()
                except Exception:
                    pass

            add_overlay_cb = _add_overlay

        label_scale_cb = None
        if canvas is not None and hasattr(canvas, "set_profile_label_scale"):
            label_scale_cb = canvas.set_profile_label_scale

        marker_update_cb = None
        if canvas is not None and hasattr(canvas, "set_profile_marker_positions"):
            def _marker_update(positions, domain):
                try:
                    canvas.set_profile_marker_positions(positions, domain=domain, emit=False)
                except Exception:
                    pass

            marker_update_cb = _marker_update

        style_update_cb = None
        if canvas is not None and hasattr(canvas, "set_profile_style"):
            def _style_update(profile_key, **changes):
                try:
                    return canvas.set_profile_style(profile_key, **changes)
                except Exception:
                    return False

            style_update_cb = _style_update

        palette_cb = None
        if canvas is not None and hasattr(canvas, "apply_profile_palette"):
            def _palette_update(name):
                try:
                    return canvas.apply_profile_palette(name)
                except Exception:
                    return False

            palette_cb = _palette_update

        return (
            activate_cb,
            highlight_cb,
            delete_cb,
            add_overlay_cb,
            label_scale_cb,
            marker_update_cb,
            marker_select_cb,
            style_update_cb,
            palette_cb,
        )

    # ------------------------------------------------------------------
    def _ensure_profile_dialog(self, active, saved):
        (
            activate_cb,
            highlight_cb,
            delete_cb,
            add_overlay_cb,
            label_scale_cb,
            marker_update_cb,
            marker_select_cb,
            style_update_cb,
            palette_cb,
        ) = self._build_dialog_callbacks()
        if saved and hasattr(self.canvas, "set_show_profile_overlays"):
            try:
                if not bool(getattr(self.canvas, "_show_profile_overlays", True)):
                    self.canvas.set_show_profile_overlays(True)
            except Exception:
                pass
        dlg = self._dialog
        created = dlg is None
        if dlg is None:
            unit = None
            y_label = None
            try:
                view = self.canvas.views[0]
                unit = view.get("unit")
                y_label = view.get("colorbar_label") or view.get("unit")
            except Exception:
                pass
            dlg = ProfileDialog(
                active,
                saved,
                parent=self.owner,
                unit=unit,
                y_label=y_label,
                activate_overlay_callback=activate_cb,
                highlight_overlay_callback=highlight_cb,
                label_scale_callback=label_scale_cb,
                delete_overlay_callback=delete_cb,
                marker_update_callback=marker_update_cb,
                marker_select_callback=marker_select_cb,
                add_overlay_callback=add_overlay_cb,
                style_update_callback=style_update_cb,
                palette_callback=palette_cb,
            )
            dlg.setWindowTitle(f"{self.title} (popup)")
            if hasattr(dlg, "detach_as_workspace_window"):
                dlg.detach_as_workspace_window()
            try:
                dlg.set_context_source(
                    self.canvas,
                    dark=getattr(self.owner, "detail_dark_view", False),
                    grid=getattr(self.owner, "detail_grid_view", False),
                )
            except Exception:
                pass
            self._register_dialog(dlg)
            dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
            dlg.finished.connect(lambda _=None: self._clear_dialog())
            self._dialog = dlg
        else:
            if hasattr(dlg, "set_label_scale_callback"):
                dlg.set_label_scale_callback(label_scale_cb)
            if hasattr(dlg, "set_marker_update_callback"):
                dlg.set_marker_update_callback(marker_update_cb)
            if hasattr(dlg, "set_marker_select_callback"):
                dlg.set_marker_select_callback(marker_select_cb)
            if hasattr(dlg, "set_add_overlay_callback"):
                dlg.set_add_overlay_callback(add_overlay_cb)
            if hasattr(dlg, "set_delete_overlay_callback"):
                dlg.set_delete_overlay_callback(delete_cb)
            else:
                dlg._delete_overlay_cb = delete_cb
            if hasattr(dlg, "set_style_update_callback"):
                dlg.set_style_update_callback(style_update_cb)
            if hasattr(dlg, "set_palette_callback"):
                dlg.set_palette_callback(palette_cb)
            dlg.update_profiles(
                active,
                saved,
                activate_overlay_callback=activate_cb,
                highlight_overlay_callback=highlight_cb,
            )
        if created:
            dlg.show()
            self._dock_dialog_near_canvas(dlg)
            try:
                dlg.raise_()
                dlg.activateWindow()
            except Exception:
                pass

    def _clear_dialog(self):
        dlg = self._dialog
        remember_cb = getattr(self.owner, "_remember_closed_popup_profile_dialog", None)
        if callable(remember_cb):
            try:
                remember_cb(self, dlg)
            except Exception:
                pass
        try:
            if self.canvas is not None and hasattr(self.canvas, "deactivate_profile_tool"):
                self.canvas.deactivate_profile_tool(clear_active=True, clear_saved=True)
        except Exception:
            pass
        self._deregister_dialog(dlg)
        self._dialog = None

    def export_dialog_state(self):
        """Capture open/geometry state for popup profile dialogs in sessions."""
        dlg = self._dialog
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

    def restore_dialog_state(self, state):
        """Reopen a popup profile dialog after canvas/profile state is restored."""
        if not isinstance(state, dict) or not bool(state.get("open")):
            return None
        self.refresh_from_canvas()
        dlg = self._dialog
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

    def _register_dialog(self, dlg):
        if dlg is None:
            return
        dialogs = getattr(self.owner, "_profile_dialogs", None)
        if dialogs is not None and dlg not in dialogs:
            dialogs.append(dlg)
        refs = getattr(self.owner, "_popup_refs", None)
        if refs is not None and dlg not in refs:
            refs.append(dlg)
        controller = getattr(self.owner, "quick_crop_controller", None)
        if controller:
            controller.update_popup_actions()

    def _deregister_dialog(self, dlg):
        if dlg is None:
            return
        dialogs = getattr(self.owner, "_profile_dialogs", None)
        if dialogs and dlg in dialogs:
            dialogs.remove(dlg)
        refs = getattr(self.owner, "_popup_refs", None)
        if refs and dlg in refs:
            refs.remove(dlg)
        controller = getattr(self.owner, "quick_crop_controller", None)
        if controller:
            controller.update_popup_actions()

    def _dock_dialog_near_canvas(self, dlg):
        if dlg is None:
            return
        try:
            source_window = self.canvas.window()
        except Exception:
            source_window = None
        if source_window is None or source_window is self.owner:
            return
        if not source_window.isVisible():
            return
        try:
            src_geo = source_window.frameGeometry()
        except Exception:
            return
        target = QtCore.QPoint(int(src_geo.right() + 16), int(src_geo.top()))
        width = dlg.frameGeometry().width()
        height = dlg.frameGeometry().height()
        screen = None
        try:
            screen = QtWidgets.QApplication.screenAt(src_geo.center())
        except Exception:
            screen = None
        if screen is None:
            screen = QtWidgets.QApplication.primaryScreen()
        bounds = screen.availableGeometry() if screen else None
        if bounds:
            max_x = bounds.right() - width
            max_y = bounds.bottom() - height
            new_x = min(max(bounds.left(), target.x()), max_x)
            new_y = min(max(bounds.top(), target.y()), max_y)
            dlg.move(new_x, new_y)
        else:
            dlg.move(target)

"""__all__ is intentionally omitted; controller used via class import."""
