"""Spectroscopy selection/comparison helpers."""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtCore, QtWidgets
from ..spectroscopy import popups as spectro_popups


class SpectroCompareController:
    """Encapsulates spectroscopy selection + popup orchestration."""

    def __init__(self, viewer):
        self.viewer = viewer

    # ------------------------------------------------------------------
    def handle_preview_click(self, spec, event=None):
        modifiers = self._event_modifiers(event)
        file_key = str(spec.get('image_key') or spec.get('path') or '') if spec else ''
        return self._handle_activation(spec, file_key, is_matrix_hint=False, modifiers=modifiers)

    def handle_marker_click(self, spec, file_key, is_matrix_hint, modifiers):
        return self._handle_activation(spec, file_key, is_matrix_hint=is_matrix_hint, modifiers=modifiers)

    # ------------------------------------------------------------------
    def spec_identity_key(self, spec):
        if not spec:
            return None
        base = spec.get('path')
        try:
            base = str(Path(base))
        except Exception:
            base = str(base)
        idx = spec.get('matrix_index')
        if idx is not None:
            return f"{base}#idx{idx}"
        x = spec.get('x')
        y = spec.get('y')
        if x is not None or y is not None:
            try:
                x_val = float(x) if x is not None else ''
                y_val = float(y) if y is not None else ''
                return f"{base}#pos{round(x_val,6)}_{round(y_val,6)}"
            except Exception:
                return f"{base}#pos{x}_{y}"
        order_idx = spec.get('order_idx')
        if order_idx is not None:
            return f"{base}#order{order_idx}"
        return base

    # ------------------------------------------------------------------
    def toggle_multi_spec_selection(self, spec):
        viewer = self.viewer
        if not spec:
            return
        key = self.spec_identity_key(spec) or str(Path(spec.get('path')))
        if key in viewer._multi_spec_selection_keys:
            viewer._multi_spec_selection = [s for s in viewer._multi_spec_selection if self.spec_identity_key(s) != key]
            viewer._multi_spec_selection_keys.remove(key)
        else:
            viewer._multi_spec_selection.append(spec)
            viewer._multi_spec_selection_keys.add(key)
        self.update_spec_selection_label()
        if not viewer._multi_spec_selection:
            viewer._multi_single_popup_anchor = None
        if len(viewer._multi_spec_selection) >= 2:
            self.open_multi_popup()

    def clear_multi_spec_selection(self):
        viewer = self.viewer
        viewer._multi_spec_selection = []
        viewer._multi_spec_selection_keys = set()
        viewer._multi_single_popup_anchor = None
        viewer._last_clicked_spec = None
        for dlg in list(getattr(viewer, "_multi_spectro_popups", [])):
            try:
                dlg.close()
            except Exception:
                pass
        viewer._multi_spectro_popups = []
        self.update_spec_selection_label()

    def update_spec_selection_label(self):
        viewer = self.viewer
        count = len(getattr(viewer, "_multi_spec_selection", []))
        if hasattr(viewer, 'spec_selection_label'):
            viewer.spec_selection_label.setText(f"Spectra selected: {count}")

    # ------------------------------------------------------------------
    def prime_multi_selection_anchor(self, current_spec):
        viewer = self.viewer
        if viewer._multi_spec_selection or not getattr(viewer, "_last_clicked_spec", None):
            return
        candidate = viewer._last_clicked_spec
        if candidate is None:
            return
        if current_spec and self.spec_identity_key(candidate) == self.spec_identity_key(current_spec):
            return
        key = self.spec_identity_key(candidate)
        if not key or key in viewer._multi_spec_selection_keys:
            viewer._last_clicked_spec = None
            return
        viewer._multi_spec_selection.append(candidate)
        viewer._multi_spec_selection_keys.add(key)
        self.update_spec_selection_label()
        viewer._last_clicked_spec = None

    # ------------------------------------------------------------------
    def append_spec_to_single_popup(self, spec):
        viewer = self.viewer
        if not spec:
            return
        key = self.spec_identity_key(spec)
        if not key:
            return
        dlg = self._active_single_popup()
        if dlg is None:
            dlg = self.ensure_single_popup(spec)
            if dlg:
                viewer._multi_single_popup_anchor = key
            return
        if key == viewer._multi_single_popup_anchor:
            return
        if hasattr(dlg, "add_external_spectrum"):
            try:
                dlg.add_external_spectrum(spec)
            except Exception:
                pass

    def ensure_single_popup(self, spec):
        viewer = self.viewer
        if not spec:
            return None
        key = self.spec_identity_key(spec)
        if key and getattr(viewer, "_spectro_popups", None):
            for dlg in list(viewer._spectro_popups):
                dlg_spec = getattr(dlg, "spec", None)
                if dlg_spec and self.spec_identity_key(dlg_spec) == key:
                    try:
                        dlg.raise_()
                        dlg.activateWindow()
                    except Exception:
                        pass
                    return dlg
        return self.open_single_popup(spec)

    def open_single_popup(self, spec):
        viewer = self.viewer
        if not viewer._spectros_loaded:
            viewer.ensure_spectros_loaded(refresh=False)
        return spectro_popups._open_spectroscopy_popup(viewer, spec)

    def open_multi_popup(self):
        viewer = self.viewer
        if not viewer._spectros_loaded:
            viewer.ensure_spectros_loaded(refresh=False)
        return spectro_popups._open_multi_spectroscopy_popup(viewer)

    # ------------------------------------------------------------------
    def _handle_activation(self, spec, file_key, is_matrix_hint, modifiers):
        viewer = self.viewer
        if not spec or not viewer.show_spectra:
            return False
        if modifiers & QtCore.Qt.ShiftModifier:
            self.prime_multi_selection_anchor(spec)
            key = self.spec_identity_key(spec) if spec else None
            already_selected = bool(key and key in getattr(viewer, "_multi_spec_selection_keys", set()))
            self.toggle_multi_spec_selection(spec)
            added = bool(spec and key and not already_selected and key in viewer._multi_spec_selection_keys)
            if added:
                self.append_spec_to_single_popup(spec)
                try:
                    viewer._highlight_spectrum_entry(spec)
                except Exception:
                    pass
            return True
        self.clear_multi_spec_selection()
        viewer._last_clicked_spec = spec
        force_matrix = bool(modifiers & QtCore.Qt.ControlModifier)
        is_matrix = (
            is_matrix_hint
            or viewer._is_matrix_spec(spec)
            or (force_matrix and spec.get('matrix_index') is not None)
        )
        if is_matrix and file_key:
            viewer._open_matrix_explorer_for_file(file_key)
        else:
            self.open_single_popup(spec)
        try:
            viewer._highlight_spectrum_entry(spec)
        except Exception:
            pass
        return True

    # ------------------------------------------------------------------
    def _active_single_popup(self):
        viewer = self.viewer
        anchor = getattr(viewer, "_multi_single_popup_anchor", None)
        if not anchor:
            return None
        dlg = self._single_popup_for_key(anchor)
        if dlg is None:
            viewer._multi_single_popup_anchor = None
        return dlg

    def _single_popup_for_key(self, key):
        viewer = self.viewer
        if not key or not getattr(viewer, "_spectro_popups", None):
            return None
        for dlg in list(viewer._spectro_popups):
            dlg_spec = getattr(dlg, "spec", None)
            if dlg_spec and self.spec_identity_key(dlg_spec) == key:
                if getattr(dlg, "isVisible", None):
                    try:
                        if dlg.isVisible():
                            return dlg
                    except Exception:
                        continue
                else:
                    return dlg
        return None

    # ------------------------------------------------------------------
    @staticmethod
    def _event_modifiers(event):
        mods = QtCore.Qt.NoModifier
        if event is None:
            return mods
        try:
            if hasattr(event, "modifiers"):
                return event.modifiers()
            gui_evt = getattr(event, "guiEvent", None)
            if gui_evt is not None and hasattr(gui_evt, "modifiers"):
                return gui_evt.modifiers()
        except Exception:
            pass
        return mods
