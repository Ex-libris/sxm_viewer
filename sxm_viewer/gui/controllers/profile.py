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
        if not getattr(self.canvas, "_profile_user_enabled", False):
            return
        if active is None and saved is None:
            active, saved = self._compute_profiles_from_canvas()
        if not active and not saved:
            return
        self._ensure_profile_dialog(active, saved)

    # ------------------------------------------------------------------
    def refresh_from_canvas(self):
        if not getattr(self.canvas, "_profile_user_enabled", False):
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
                view=canvas.views[0] if canvas.views else None,
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
                        view=canvas.views[0] if canvas.views else None,
                    )
                if data:
                    saved.append(data)
        except Exception:
            saved = []
        return active, saved

    # ------------------------------------------------------------------
    def _ensure_profile_dialog(self, active, saved):
        dlg = self._dialog
        if dlg is None:
            unit = None
            y_label = None
            try:
                view = self.canvas.views[0]
                unit = view.get("unit")
                y_label = view.get("colorbar_label") or view.get("unit")
            except Exception:
                pass
            dlg = ProfileDialog(active, saved, parent=self.owner, unit=unit, y_label=y_label)
            dlg.setWindowTitle(f"{self.title} (popup)")
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
            dlg.update_profiles(active, saved)
        dlg.show()
        self._dock_dialog_near_canvas(dlg)
        try:
            dlg.raise_()
            dlg.activateWindow()
        except Exception:
            pass

    def _clear_dialog(self):
        self._deregister_dialog(self._dialog)
        self._dialog = None

    def _register_dialog(self, dlg):
        if dlg is None:
            return
        refs = getattr(self.owner, "_popup_refs", None)
        if refs is not None and dlg not in refs:
            refs.append(dlg)
        controller = getattr(self.owner, "quick_crop_controller", None)
        if controller:
            controller.update_popup_actions()

    def _deregister_dialog(self, dlg):
        if dlg is None:
            return
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
