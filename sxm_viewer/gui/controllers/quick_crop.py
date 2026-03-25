"""Quick crop control helpers."""
from __future__ import annotations

import math
from functools import partial
from pathlib import Path

from ..._shared import QtWidgets, QtGui, QtCore
from ...config import save_config


class QuickCropController:
    """Encapsulates quick-crop UI/state coordination for the main window."""

    def __init__(self, viewer):
        self.viewer = viewer
        self.history_entries = []
        self.popup_stack = []
        self.selected_sequences = []
        self.active_sequence = None

    # ------------------------------------------------------------------
    def set_mode(self, enabled: bool, save: bool = True):
        viewer = self.viewer
        enabled = bool(enabled)
        viewer.quick_crop_mode = enabled
        if save:
            viewer.config['quick_crop_mode'] = viewer.quick_crop_mode
            save_config(viewer.config)
        act = getattr(viewer, 'fixed_crop_quick_act', None)
        if act is not None:
            try:
                act.blockSignals(True)
                act.setChecked(enabled)
                act.blockSignals(False)
            except Exception:
                pass
        btn = getattr(viewer, 'quick_crop_btn', None)
        if btn is not None:
            try:
                btn.blockSignals(True)
                btn.setChecked(enabled)
                btn.setText("Quick crop: On" if enabled else "Quick crop: Off")
                btn.blockSignals(False)
            except Exception:
                pass
        detail_widget = getattr(viewer, 'quick_crop_detail_widget', None)
        if detail_widget is not None:
            try:
                detail_widget.setVisible(enabled)
            except Exception:
                pass
        canvases = [getattr(viewer, 'preview_canvas', None)] + list(getattr(viewer, '_popup_canvases', []))
        for canv in canvases:
            if canv is None:
                continue
            try:
                canv.enable_fixed_crop_quick_mode(enabled)
            except Exception:
                continue
        if enabled:
            self.apply_template_from_controls()
        self.update_hint()

    # ------------------------------------------------------------------
    def update_hint(self):
        viewer = self.viewer
        label = getattr(viewer, 'quick_crop_hint_lbl', None)
        if label is None:
            return
        overlay_state = "visible" if getattr(viewer, 'show_crop_history_overlay', False) else "hidden"
        if viewer.quick_crop_mode:
            text = (
                "Shift+drag to size, click to spawn crops. "
                "Right-click -> Quick tools -> Edit crop frame (or Ctrl+E) enables drag/resize/rotate, Enter applies. "
                "Ctrl+Z undo, Ctrl+Shift+W close latest pop-out. "
                f"History overlay is {overlay_state}. "
                "Ctrl+Shift+R reapplies the current real-size template, Ctrl+Shift+H toggles history overlay, Ctrl+Shift+T toggles template overlay."
            )
        else:
            text = "Quick crop mode is off. Press Ctrl+Shift+C to toggle."
        label.setText(text)

    # ------------------------------------------------------------------
    def sync_template_controls(self):
        viewer = self.viewer
        width_spin = getattr(viewer, 'quick_crop_real_width_spin', None)
        if width_spin is None:
            return
        try:
            real_width, real_height, real_unit = viewer.preview_canvas.get_fixed_crop_template_real_size()
        except Exception:
            real_width = real_height = 0.0
            real_unit = "nm"
        if width_spin is not None:
            width_spin.blockSignals(True)
            if real_width > 0:
                width_spin.setValue(real_width)
            width_spin.blockSignals(False)
        height_spin = getattr(viewer, 'quick_crop_real_height_spin', None)
        if height_spin is not None:
            height_spin.blockSignals(True)
            if real_height > 0:
                height_spin.setValue(real_height)
            height_spin.blockSignals(False)
        if real_width > 0 and real_height > 0:
            viewer._quick_crop_aspect = real_width / max(0.001, real_height)
        unit_lbl = getattr(viewer, 'quick_crop_real_unit_lbl', None)
        if unit_lbl is not None:
            unit_lbl.setText(real_unit or "nm")
        info_lbl = getattr(viewer, 'quick_crop_real_px_info_lbl', None)
        if info_lbl is not None:
            try:
                px_dims = viewer.preview_canvas.get_fixed_crop_template_size()
            except Exception:
                px_dims = ()
            if len(px_dims) == 2 and px_dims[0] and px_dims[1]:
                info_lbl.setText(
                    f"W: {real_width:.3f} {real_unit} ({px_dims[0]} px)  |  H: {real_height:.3f} {real_unit} ({px_dims[1]} px)"
                )
            else:
                info_lbl.setText("")

    # ------------------------------------------------------------------
    def on_real_spin_changed(self, sender=None):
        viewer = self.viewer
        width_spin = getattr(viewer, 'quick_crop_real_width_spin', None)
        height_spin = getattr(viewer, 'quick_crop_real_height_spin', None)
        if width_spin is None or height_spin is None:
            return
        lock_cb = getattr(viewer, 'quick_crop_lock_aspect_cb', None)
        if lock_cb is not None and lock_cb.isChecked():
            aspect = viewer._quick_crop_aspect or 1.0
            if sender is width_spin:
                new_w = width_spin.value()
                new_h = max(0.01, new_w / max(0.001, aspect))
                height_spin.blockSignals(True)
                height_spin.setValue(new_h)
                height_spin.blockSignals(False)
            elif sender is height_spin:
                new_h = height_spin.value()
                new_w = max(0.01, new_h * aspect)
                width_spin.blockSignals(True)
                width_spin.setValue(new_w)
                width_spin.blockSignals(False)
        square_cb = getattr(viewer, 'quick_crop_square_cb', None)
        if square_cb is not None and square_cb.isChecked():
            val = width_spin.value()
            height_spin.blockSignals(True)
            height_spin.setValue(val)
            height_spin.blockSignals(False)
        self.apply_template_from_controls()

    # ------------------------------------------------------------------
    def apply_template_from_controls(self):
        viewer = self.viewer
        width_spin = getattr(viewer, 'quick_crop_real_width_spin', None)
        height_spin = getattr(viewer, 'quick_crop_real_height_spin', None)
        if width_spin is None or height_spin is None:
            return
        square = False
        square_cb = getattr(viewer, 'quick_crop_square_cb', None)
        if square_cb is not None:
            square = bool(square_cb.isChecked())
        real_w = width_spin.value()
        real_h = height_spin.value()
        success = False
        try:
            success = viewer.preview_canvas.set_fixed_crop_template_real_size(real_w, real_h, square=square)
        except Exception:
            success = False
        if not success:
            try:
                h, w = viewer.preview_canvas.get_main_view_shape()
            except Exception:
                h = w = 0
            if w > 0 and h > 0:
                aspect = real_h / real_w if real_w not in (0, None) else 1.0
                base_w = max(2, min(int(max(2, w * 0.25)), w))
                base_h = base_w if square else max(2, min(int(round(base_w * aspect)), h))
                try:
                    success = viewer.preview_canvas.set_fixed_crop_template_size(base_w, base_h, square=square)
                except Exception:
                    success = False
        if success:
            viewer._quick_crop_last_real_size = [real_w, real_h]
            if real_w > 0 and real_h > 0:
                viewer._quick_crop_aspect = real_w / max(0.001, real_h)
            self.sync_template_controls()

    # ------------------------------------------------------------------
    def register_popup(self, seq, dlg):
        if dlg is None:
            return
        entry = {"seq": seq, "dialog": dlg}
        self.popup_stack.append(entry)
        dlg.finished.connect(lambda _=None, d=dlg: self.unregister_popup(d))
        close_shortcut = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+W"), dlg)
        close_shortcut.setContext(QtCore.Qt.WidgetShortcut)
        close_shortcut.activated.connect(dlg.close)
        dlg._quick_crop_close_shortcut = close_shortcut
        self.update_active_sequence_from_stack()
        self.update_popup_actions()

    def unregister_popup(self, dlg):
        if dlg is None:
            return
        changed = False
        seq_removed = None
        new_stack = []
        for entry in self.popup_stack:
            if entry.get("dialog") is dlg:
                changed = True
                seq_removed = entry.get("seq")
                continue
            new_stack.append(entry)
        if changed:
            self.popup_stack = new_stack
            if seq_removed is not None:
                self.finalize_removed_crop(seq_removed)
            else:
                self.refresh_history_panel()
        self.update_popup_actions()

    # ------------------------------------------------------------------
    def close_popup(self, seq=None):
        target = None
        if seq is None:
            if self.popup_stack:
                target = self.popup_stack[-1]
        else:
            for entry in reversed(self.popup_stack):
                if entry.get("seq") == seq:
                    target = entry
                    break
        if target is None:
            return False
        dlg = target.get("dialog")
        if dlg:
            try:
                dlg.close()
            except Exception:
                pass
        self.unregister_popup(dlg)
        return True

    def close_latest_popup(self):
        self.close_popup(None)

    def close_all_popups(self):
        stack = list(self.popup_stack)
        for entry in stack:
            dlg = entry.get("dialog")
            if dlg:
                try:
                    dlg.close()
                except Exception:
                    pass
        self.popup_stack.clear()
        self.set_active_sequence(None, keep_selected=False)
        self.refresh_history_panel()
        self.update_popup_actions()

    # ------------------------------------------------------------------
    def set_active_sequence(self, seq, keep_selected=True):
        self.active_sequence = seq
        if not keep_selected:
            self.selected_sequences = []
        viewer = self.viewer
        canvases = [getattr(viewer, 'preview_canvas', None)] + list(getattr(viewer, '_popup_canvases', []))
        for canv in canvases:
            if canv is None:
                continue
            try:
                canv.set_fixed_crop_history_highlight(seq)
            except Exception:
                continue

    def update_active_sequence_from_stack(self):
        seq = None
        if self.popup_stack:
            seq = self.popup_stack[-1].get("seq")
        self.set_active_sequence(seq)

    # ------------------------------------------------------------------
    def focus_history_entry(self, seq, multi=False):
        if seq is None:
            return
        self.update_selected_sequences(seq, multi=multi)
        self.set_active_sequence(seq, keep_selected=True)
        for entry in self.popup_stack:
            if entry.get("seq") == seq:
                dlg = entry.get("dialog")
                if dlg:
                    try:
                        state = dlg.windowState()
                        if state & QtCore.Qt.WindowMinimized:
                            dlg.setWindowState(state & ~QtCore.Qt.WindowMinimized)
                        dlg.raise_()
                        dlg.activateWindow()
                    except Exception:
                        pass
                break
        self.refresh_history_panel()

    def focus_history_entry_with_shift(self, seq, _checked=False):
        modifiers = QtWidgets.QApplication.keyboardModifiers()
        multi = bool(modifiers & QtCore.Qt.ShiftModifier)
        self.focus_history_entry(seq, multi=multi)

    def update_selected_sequences(self, seq, multi=False):
        if seq is None:
            self.selected_sequences = []
            return
        selected = list(self.selected_sequences)
        if not multi:
            selected = [seq]
        else:
            if seq in selected:
                selected.remove(seq)
            else:
                selected.append(seq)
        self.selected_sequences = selected

    # ------------------------------------------------------------------
    def finalize_removed_crop(self, seq):
        if seq is None:
            return
        if seq in self.selected_sequences:
            self.selected_sequences.remove(seq)
        entry = self.viewer.preview_canvas.remove_fixed_crop_history_entry(seq)
        if entry:
            self.update_active_sequence_from_stack()
        else:
            self.refresh_history_panel()

    def undo_last_crop(self):
        entry = self.viewer.preview_canvas.undo_fixed_crop_entry()
        if entry:
            seq = entry.get("sequence")
            if seq is not None:
                self.close_popup(seq)

    def clear_history(self):
        self.viewer.preview_canvas.clear_fixed_crop_history()
        self.close_all_popups()

    # ------------------------------------------------------------------
    def on_history_updated(self, entries):
        self.history_entries = list(entries or [])
        self.refresh_history_panel()
        try:
            width, height = self.viewer.preview_canvas.get_fixed_crop_template_size()
        except Exception:
            width = height = 0
        if width and height:
            self.sync_template_controls()
        self.update_active_sequence_from_stack()

    # ------------------------------------------------------------------
    def refresh_history_panel(self):
        viewer = self.viewer
        layout = getattr(viewer, 'crop_history_layout', None)
        if layout is None:
            return
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()
        entries = self.history_entries or []
        selected_seqs = self.cleanup_selected_sequences()
        label = getattr(viewer, 'crop_history_label', None)
        if label is not None:
            label_text = "Crop history"
            if entries:
                label_text = f"{label_text} ({len(entries)})"
            label.setText(label_text)
        display = entries[-12:]
        active_seq = selected_seqs[-1] if selected_seqs else self.active_sequence
        for entry in reversed(display):
            seq = entry.get("sequence")
            frame = QtWidgets.QFrame()
            frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
            frame.setFrameShadow(QtWidgets.QFrame.Raised)
            entry_layout = QtWidgets.QHBoxLayout(frame)
            entry_layout.setContentsMargins(6, 4, 6, 4)
            entry_layout.setSpacing(8)
            color = entry.get("color") or self._crop_color_for_seq(seq)
            real_size = entry.get("real_size", (0.0, 0.0))
            unit = entry.get("unit") or "nm"
            if any(real_size):
                size_text = f"{real_size[0]:.2f} × {real_size[1]:.2f} {unit}"
            else:
                pixel_bounds = entry.get("pixel_bounds")
                if pixel_bounds:
                    px_width = int(abs(pixel_bounds[1] - pixel_bounds[0]) + 1)
                    px_height = int(abs(pixel_bounds[3] - pixel_bounds[2]) + 1)
                    size_text = f"{px_width} × {px_height} px"
                else:
                    size_text = "Unknown size"
            label_widget = QtWidgets.QLabel(f"#{seq} - {size_text}")
            label_widget.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
            entry_layout.addWidget(label_widget)
            select_btn = QtWidgets.QToolButton()
            select_btn.setText("Select")
            select_btn.setCheckable(True)
            select_btn.setChecked(seq in selected_seqs)
            select_btn.setToolTip("Shift+click to multiselect, click to highlight")
            select_btn.clicked.connect(partial(self.focus_history_entry_with_shift, seq))
            entry_layout.addWidget(select_btn)
            virtual_btn = QtWidgets.QToolButton()
            virtual_btn.setText("Virtual copy")
            virtual_btn.setToolTip("Add this crop snapshot as a virtual thumbnail")
            virtual_btn.setEnabled(bool(entry.get("view_snapshot")))
            virtual_btn.clicked.connect(partial(viewer._create_virtual_copy_from_history, seq))
            entry_layout.addWidget(virtual_btn)
            statuses = []
            if seq == active_seq:
                statuses.append("Active")
                frame.setStyleSheet(f"QFrame {{ border: 1px solid {color}; background-color: rgba(255, 102, 255, 0.08); }}")
            elif seq in selected_seqs:
                statuses.append("Selected")
                frame.setStyleSheet(f"QFrame {{ border: 1px solid {color}; background-color: rgba(143, 223, 255, 0.08); }}")
            else:
                frame.setStyleSheet(f"QFrame {{ border: 1px solid {color}; }}")
            if self._has_popup_for_seq(seq):
                statuses.append("Open")
            if statuses:
                icons = []
                if "Active" in statuses:
                    icons.append("✔")
                if "Open" in statuses:
                    icons.append("●")
                status_lbl = QtWidgets.QLabel(" ".join(icons) or " ")
                status_lbl.setToolTip(", ".join(statuses))
                status_lbl.setStyleSheet(f"color: {color}; font-weight: bold;")
                entry_layout.addWidget(status_lbl)
            entry_layout.addStretch(1)
            close_btn = QtWidgets.QToolButton()
            close_btn.setText("Close view")
            close_btn.setToolTip("Close the pop-out that corresponds to this crop")
            close_btn.clicked.connect(lambda _, s=seq: self.close_popup(s))
            entry_layout.addWidget(close_btn)
            remove_btn = QtWidgets.QToolButton()
            remove_btn.setText("Remove overlay")
            remove_btn.setToolTip("Drop this crop and its overlay from the preview")
            remove_btn.clicked.connect(lambda _, s=seq: self.remove_history_entry(s))
            entry_layout.addWidget(remove_btn)
            layout.addWidget(frame)
        layout.addStretch(1)
        viewer.crop_history_panel.setVisible(bool(entries) or getattr(viewer, 'show_crop_history_overlay', False))
        viewer.quick_crop_undo_btn.setEnabled(bool(entries))
        viewer.quick_crop_clear_btn.setEnabled(bool(entries))
        viewer.quick_crop_close_btn.setEnabled(bool(self.popup_stack))
        export_btn = getattr(viewer, 'quick_crop_export_btn', None)
        if export_btn is not None:
            export_btn.setEnabled(bool(selected_seqs))
        tile_btn = getattr(viewer, 'quick_crop_tile_btn', None)
        if tile_btn is not None:
            self.update_popup_actions()

    # ------------------------------------------------------------------
    def cleanup_selected_sequences(self):
        entries = self.history_entries or []
        valid = {entry.get("sequence") for entry in entries if entry.get("sequence") is not None}
        selected = [seq for seq in self.selected_sequences if seq in valid]
        if selected != self.selected_sequences:
            self.selected_sequences = selected
        return selected

    def _has_popup_for_seq(self, seq):
        if seq is None:
            return False
        for entry in self.popup_stack:
            if entry.get("seq") == seq:
                dlg = entry.get("dialog")
                if dlg and dlg.isVisible():
                    return True
        return False

    def update_popup_actions(self):
        viewer = self.viewer
        tile_btn = getattr(viewer, 'quick_crop_tile_btn', None)
        if tile_btn is None:
            return
        alive = []
        cleaned = []
        for dlg in list(getattr(viewer, '_popup_refs', [])):
            if dlg is None:
                continue
            try:
                if dlg.isVisible():
                    alive.append(dlg)
                cleaned.append(dlg)
            except RuntimeError:
                continue
        viewer._popup_refs = cleaned
        tile_btn.setEnabled(bool(alive))

    def _crop_color_for_seq(self, seq: int):
        palette = [
            "#ff66ff", "#66c2ff", "#ffa600", "#00c896",
            "#c084ff", "#ff6b6b", "#4dd0e1", "#ffd166",
        ]
        try:
            return palette[seq % len(palette)]
        except Exception:
            return palette[0]

    def arrange_popups(self):
        viewer = self.viewer
        popups = [dlg for dlg in getattr(viewer, '_popup_refs', []) if dlg and dlg.isVisible()]
        if not popups:
            return
        screen = QtWidgets.QApplication.primaryScreen()
        geom = screen.availableGeometry() if screen else QtWidgets.QApplication.desktop().availableGeometry()
        cols = int(math.ceil(math.sqrt(len(popups))))
        rows = int(math.ceil(len(popups) / cols))
        tile_w = max(280, geom.width() // cols)
        tile_h = max(240, geom.height() // rows)
        for idx, dlg in enumerate(popups):
            row, col = divmod(idx, cols)
            x = geom.left() + col * tile_w
            y = geom.top() + row * tile_h
            dlg.setGeometry(x + 10, y + 10, tile_w - 20, tile_h - 20)
            dlg.raise_()
            dlg.activateWindow()

    # ------------------------------------------------------------------
    def export_selected_crops(self):
        selected = list(self.selected_sequences)
        if not selected:
            QtWidgets.QMessageBox.information(self.viewer, "Export crops", "Select one or more crops from the history first.")
            return
        directory = QtWidgets.QFileDialog.getExistingDirectory(self.viewer, "Export crops")
        if not directory:
            return
        fmt, ok = QtWidgets.QInputDialog.getItem(
            self.viewer,
            "Export format",
            "Format:",
            ["PNG", "SVG"],
            0,
            False,
        )
        if not ok or not fmt:
            return
        fmt = fmt.lower()
        dpi = 300 if fmt == "png" else 150
        saved = []
        failed = []
        for seq in selected:
            entry = self.get_history_entry(seq)
            if entry is None:
                failed.append(seq)
                continue
            fig = self.viewer.preview_canvas.render_crop_entry_figure(entry)
            if fig is None:
                failed.append(seq)
                continue
            title = self.viewer._sanitize_filename_component(entry.get("title") or f"crop_{seq}") or f"crop_{seq}"
            path = Path(directory) / f"{title}_{seq}.{fmt}"
            try:
                fig.savefig(str(path), format=fmt, dpi=dpi, bbox_inches="tight", pad_inches=0.02)
                saved.append(str(path))
            except Exception:
                failed.append(seq)
            finally:
                try:
                    import matplotlib.pyplot as _plt
                    _plt.close(fig)
                except Exception:
                    pass
        summary = []
        if saved:
            summary.append(f"Saved {len(saved)} file{'s' if len(saved) != 1 else ''}.")
        if failed:
            summary.append("Could not render crops: " + ", ".join(f"#{s}" for s in failed) + ".")
        if not summary:
            summary.append("No crops exported.")
        QtWidgets.QMessageBox.information(self.viewer, "Export crops", "\n".join(summary))

    # ------------------------------------------------------------------
    def remove_history_entry(self, seq):
        if seq is None:
            return
        if not self.close_popup(seq):
            self.finalize_removed_crop(seq)

    def get_history_entry(self, seq):
        if seq is None:
            return None
        for entry in self.history_entries:
            if entry.get("sequence") == seq:
                return entry
        return None
