"""Enhanced canvas window with modern UI/UX and polished aesthetics."""
from __future__ import annotations
from pathlib import Path
from typing import TYPE_CHECKING
import math

from ..._shared import QtCore, QtGui, QtWidgets, np
from ..constants import (
    CANVAS_ALIGN_GAP,
    CANVAS_ALIGN_MARGIN,
    CANVAS_DROP_OFFSET,
    CANVAS_SPLITTER_SIZES,
    CANVAS_WINDOW_SIZE,
)
from ...data.io import parse_header
from ...processing.detection import _find_topography_channel
from .canvas_items import CanvasImageItem
from . import canvas_window_actions
from . import canvas_window_ui
from .canvas_state import (
    capture_state,
    delete_selected,
    push_undo_state,
    redo,
    restore_state,
    undo,
)
from .canvas_view import CanvasGraphicsView

if TYPE_CHECKING:
    from typing import Optional


def _safe_float(text):
    try:
        return float(text)
    except Exception:
        return None


class ExperimentalCanvasWindow(QtWidgets.QDialog):
    def __init__(self, viewer, parent=None):
        super().__init__(parent)
        self.viewer = viewer
        self.setWindowTitle("Enhanced Scientific Canvas")
        try:
            screen = QtWidgets.QApplication.screenAt(QtGui.QCursor.pos()) or QtWidgets.QApplication.primaryScreen()
            avail = screen.availableGeometry() if screen is not None else None
        except Exception:
            avail = None
        target_w, target_h = CANVAS_WINDOW_SIZE
        if avail is not None:
            target_w = min(target_w, max(860, int(avail.width() * 0.72)))
            target_h = min(target_h, max(620, int(avail.height() * 0.78)))
        self.setMinimumSize(840, 580)
        self.resize(target_w, target_h)
        self.setSizeGripEnabled(True)
        self._drop_offset = QtCore.QPointF(*CANVAS_DROP_OFFSET)
        self._selected_item: Optional[CanvasImageItem] = None
        self._sync_colorbars = False
        self._kind_cmap = {
            "topo": "afmhot",
            "current": "Blues_r",
            "df": "gray",
        }
        self._sync_by_channel = True
        self._show_overlay_info = True
        self._show_overlay_file = False
        self._last_aligned_width: float | None = None
        self._grid_locked = False  # prevents automatic resizing
        self._global_show_title = False
        self._global_show_colorbar = True
        self._global_show_colorbar_ticks = True
        self._global_text_scale = 1.0
        self._global_text_color: QtGui.QColor | None = None
        self._global_show_scale_bar = False
        self._global_scale_bar_length_nm: float | None = None
        self._metadata_bar_default = False
        self._metadata_unit_default = True
        self._colorbar_mode = "bottom"
        self._display_preset_name = "Custom"
        self._suppress_preset_sync = False
        self._undo_stack = []
        self._undo_index = -1
        self._file_scale_bars = {}
        self._restoring = False
        self._global_show_molecules = bool(getattr(self.viewer, "show_molecules", True))
        self._molecule_palette = str(getattr(self.viewer, "molecule_palette", "pymol") or "pymol").lower()
        self._recent_molecule_paths = list(getattr(self.viewer, "recent_molecules", []) or [])

        self.scene = QtWidgets.QGraphicsScene(self)
        self.view = CanvasGraphicsView(self)
        self.view.setScene(self.scene)
        self._dark = bool(getattr(self.viewer, "dark_mode", False))
        self._apply_styles(self._dark)
        self.view.set_background_color(self._workspace_color(self._dark))

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Build UI
        toolbar_widget = self._build_toolbar()
        main_layout.addWidget(toolbar_widget)

        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self.left_controls = canvas_window_ui.build_left_controls(self)
        self.splitter.addWidget(self.left_controls)
        self.splitter.addWidget(self.view)
        self.splitter.addWidget(self._build_inspector())
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 4)
        self.splitter.setStretchFactor(2, 1)
        self.splitter.setSizes(list(CANVAS_SPLITTER_SIZES))
        self.splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.splitter, 1)

        self.status_label = QtWidgets.QLabel("Ready")
        canvas_window_ui.apply_status_style(self)
        main_layout.addWidget(self.status_label)

        # Re-apply after the full widget tree exists so late-created controls pick up the theme.
        self._apply_styles(self._dark)
        self.view.set_background_color(self._workspace_color(self._dark))

        self.scene.selectionChanged.connect(self._on_selection_changed)
        self._set_display_preset_combo("Custom")
        self._push_undo_state()

    def showEvent(self, event):
        super().showEvent(event)
        try:
            self._set_inspector_focus(bool(self._selected_item))
        except Exception:
            pass

    def _set_inspector_focus(self, has_selection: bool):
        splitter = getattr(self, "splitter", None)
        if splitter is None:
            return
        try:
            total = max(600, splitter.size().width())
        except Exception:
            total = CANVAS_WINDOW_SIZE[0]
        left_width = 170
        inspector_width = 260 if has_selection else 220
        canvas_width = max(360, total - left_width - inspector_width)
        splitter.setSizes([left_width, canvas_width, inspector_width])

    def _workspace_color(self, dark: bool) -> QtGui.QColor:
        return QtGui.QColor("#1f2328" if dark else "#f3efe8")

    def _create_icon_button(self, text: str, icon_text: str = "", tooltip: str = "") -> QtWidgets.QPushButton:
        """Create a button with optional icon."""
        display_text = f"{icon_text} {text}" if icon_text else text
        btn = QtWidgets.QPushButton(display_text)
        if tooltip:
            btn.setToolTip(tooltip)
        return btn

    def _create_toolbar_section(self, title: str, widgets: list) -> QtWidgets.QWidget:
        return canvas_window_ui.create_toolbar_section(title, widgets)

    def _build_toolbar(self):
        return canvas_window_ui.build_toolbar(self)

    def _create_toolbar_group(self, title):
        return canvas_window_ui.create_toolbar_group(title)

    def _create_separator(self):
        return canvas_window_ui.create_separator()

    def set_dark_mode(self, dark: bool):
        """Update the canvas theme to match the main viewer mode."""
        self._dark = bool(dark)
        self._apply_styles(self._dark)
        try:
            self.view.set_background_color(self._workspace_color(self._dark))
        except Exception:
            pass

    def _apply_styles(self, dark: bool | None = None):
        """Apply scientific GUI styling - high contrast, clear organization."""
        if dark is None:
            dark = bool(getattr(self.viewer, "dark_mode", False))
        canvas_window_ui.apply_styles(self, dark=bool(dark))
        canvas_window_ui.apply_status_style(self)
        try:
            # Force a re-polish so existing widgets pick up the new stylesheet.
            self.style().unpolish(self)
            self.style().polish(self)
            self.update()
        except Exception:
            pass

    def _build_inspector(self):
        return canvas_window_ui.build_inspector(self)

    def _hline(self):
        line = QtWidgets.QFrame()
        line.setFrameShape(QtWidgets.QFrame.HLine)
        line.setFrameShadow(QtWidgets.QFrame.Sunken)
        return line

    def _set_inspector_enabled(self, enabled: bool):
        for widget in (
            self.colorbar_edit,
            self.cmap_combo,
            self.vmin_edit,
            self.vmax_edit,
            self.auto_range_btn,
            self.copy_range_btn,
            self.duplicate_btn,
            self.remove_btn,
        ):
            widget.setEnabled(enabled)

    def _selected_canvas_items(self):
        try:
            return [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        except Exception:
            return []

    def _apply_to_canvas_items(self, fn, *, selected_only: bool = False):
        items = self._selected_canvas_items() if selected_only else [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        for item in items:
            try:
                fn(item)
            except Exception:
                continue
        return items

    def _set_display_preset_combo(self, name: str):
        combo = getattr(self, "display_preset_combo", None)
        if combo is None:
            return
        label = str(name or "Custom")
        idx = combo.findText(label)
        if idx < 0:
            idx = combo.findText("Custom")
        if idx < 0:
            return
        try:
            self._suppress_preset_sync = True
            combo.setCurrentIndex(idx)
        finally:
            self._suppress_preset_sync = False
        self._display_preset_name = combo.currentText() or label

    def _mark_display_preset_custom(self):
        if getattr(self, "_suppress_preset_sync", False):
            return
        self._set_display_preset_combo("Custom")

    def _apply_display_preset(self, name: str, *, push_undo: bool = True):
        preset = str(name or "").strip().lower()
        if preset not in ("clean", "analysis", "publication"):
            return
        try:
            self._suppress_preset_sync = True
            if preset == "clean":
                self._on_global_show_title_toggled(False)
                self._on_overlay_info_toggled(False)
                self._on_overlay_file_toggled(False)
                self._on_metadata_bar_toggled(False)
                self._on_metadata_unit_toggled(False)
                self._on_scale_bar_toggled(True)
                self._apply_global_show_colorbar(False)
                self._apply_global_show_colorbar_ticks(False)
                self._on_colorbar_position_changed("Hidden")
            elif preset == "analysis":
                self._on_global_show_title_toggled(True)
                self._on_overlay_info_toggled(False)
                self._on_overlay_file_toggled(False)
                self._on_metadata_bar_toggled(False)
                self._on_metadata_unit_toggled(False)
                self._on_scale_bar_toggled(True)
                self._apply_global_show_colorbar(True)
                self._apply_global_show_colorbar_ticks(True)
                self._on_colorbar_position_changed("Right")
            elif preset == "publication":
                self._on_global_show_title_toggled(False)
                self._on_overlay_info_toggled(False)
                self._on_overlay_file_toggled(False)
                self._on_metadata_bar_toggled(False)
                self._on_metadata_unit_toggled(False)
                self._on_scale_bar_toggled(True)
                self._apply_global_show_colorbar(False)
                self._apply_global_show_colorbar_ticks(False)
                self._on_colorbar_position_changed("Hidden")
        finally:
            self._suppress_preset_sync = False
        self._set_display_preset_combo(preset.title())
        self.status_label.setText(f"Display preset: {preset.title()}")
        if push_undo:
            self._push_undo_state()

    def _on_apply_display_preset_clicked(self):
        combo = getattr(self, "display_preset_combo", None)
        if combo is None:
            return
        name = combo.currentText()
        if name == "Custom":
            return
        self._apply_display_preset(name, push_undo=True)

    def _on_item_overlay_chip_toggled(self, item: CanvasImageItem, key: str):
        self._selected_item = item
        self._on_selection_changed()
        label_map = {
            "title": "Title",
            "scale": "Scale bar",
            "cbar": "Colorbar",
            "meta": "Metadata",
            "unit": "Unit badge",
            "file": "Filename badge",
        }
        self.status_label.setText(f"Toggled {label_map.get(key, key)} for selected tile")
        self._mark_display_preset_custom()
        self._push_undo_state()

    def _on_selection_changed(self):
        selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        item = selected[0] if selected else None
        self._selected_item = item
        if item is None:
            self._set_inspector_focus(False)
            if hasattr(self, "selection_hint"):
                self.selection_hint.setText("Select a tile to edit its labels, scale bar, colormap and export settings.")
            self.file_label.setText("-")
            self.channel_label.setText("-")
            self.colorbar_edit.setText("")
            self.text_scale_slider.setValue(int(round(self._global_text_scale * 100)))
            try:
                self.font_color_auto_check.blockSignals(True)
                self.font_color_auto_check.setChecked(self._global_text_color is None)
            finally:
                self.font_color_auto_check.blockSignals(False)
            try:
                self.scale_bar_check.blockSignals(True)
                self.scale_bar_check.setChecked(self._global_show_scale_bar)
            finally:
                self.scale_bar_check.blockSignals(False)
            left_combo = getattr(self, "left_scale_bar_combo", None)
            if left_combo is not None:
                try:
                    left_combo.blockSignals(True)
                    left_combo.setCurrentText("Auto" if self._global_scale_bar_length_nm is None else f"{self._global_scale_bar_length_nm:g} nm")
                finally:
                    left_combo.blockSignals(False)
            if hasattr(self, "canvas_molecules_check"):
                self.canvas_molecules_check.blockSignals(True)
                self.canvas_molecules_check.setChecked(self._global_show_molecules)
                self.canvas_molecules_check.blockSignals(False)
            self.vmin_edit.setText("")
            self.vmax_edit.setText("")
            self.stats_label.setText("-")
            self._set_inspector_enabled(False)
            return
        self._set_inspector_enabled(True)
        self._set_inspector_focus(True)
        if hasattr(self, "selection_hint"):
            self.selection_hint.setText("Use the on-tile chips for quick toggles, or refine the selected tile from the tabs below.")
        self.file_label.setText(Path(item.file_path).name)
        self.channel_label.setText(str(item.channel_index))
        self.colorbar_edit.setText(item.colorbar_label)
        try:
            self.text_scale_slider.blockSignals(True)
            self.text_scale_slider.setValue(int(round(self._global_text_scale * 100)))
        finally:
            self.text_scale_slider.blockSignals(False)
        try:
            self.font_color_auto_check.blockSignals(True)
            self.font_color_auto_check.setChecked(self._global_text_color is None)
        finally:
            self.font_color_auto_check.blockSignals(False)
        try:
            self.scale_bar_check.blockSignals(True)
            self.scale_bar_check.setChecked(self._global_show_scale_bar)
        finally:
            self.scale_bar_check.blockSignals(False)
        if hasattr(self, "canvas_molecules_check"):
            try:
                self.canvas_molecules_check.blockSignals(True)
                self.canvas_molecules_check.setChecked(self._global_show_molecules)
            finally:
                self.canvas_molecules_check.blockSignals(False)
        try:
            self.scale_bar_combo.blockSignals(True)
            if self._global_scale_bar_length_nm is None:
                self.scale_bar_combo.setCurrentText("Auto")
            else:
                label = f"{self._global_scale_bar_length_nm:g} nm"
                idx = self.scale_bar_combo.findText(label)
                self.scale_bar_combo.setCurrentIndex(idx if idx >= 0 else 0)
        finally:
            self.scale_bar_combo.blockSignals(False)
        left_combo = getattr(self, "left_scale_bar_combo", None)
        if left_combo is not None:
            try:
                left_combo.blockSignals(True)
                if self._global_scale_bar_length_nm is None:
                    left_combo.setCurrentText("Auto")
                else:
                    label = f"{self._global_scale_bar_length_nm:g} nm"
                    idx = left_combo.findText(label)
                    left_combo.setCurrentIndex(idx if idx >= 0 else 0)
            finally:
                left_combo.blockSignals(False)
        self.cmap_combo.setCurrentText(item.cmap)
        self.vmin_edit.setText("" if item.vmin is None else str(item.vmin))
        self.vmax_edit.setText("" if item.vmax is None else str(item.vmax))
        arr = item.data_array
        try:
            stats_text = (
                f"Shape: {arr.shape[0]} x {arr.shape[1]}\n"
                f"Min: {np.nanmin(arr):.3e}\n"
                f"Max: {np.nanmax(arr):.3e}\n"
                f"Mean: {np.nanmean(arr):.3e}\n"
                f"Std: {np.nanstd(arr):.3e}"
            )
        except Exception:
            stats_text = "Stats: N/A"
        self.stats_label.setText(stats_text)
        n_selected = len([i for i in self.scene.items() if isinstance(i, CanvasImageItem) and i.isSelected()])
        self.status_label.setText(f"{n_selected} selected | {len(self.scene.items())} total items")

    def _on_colorbar_changed(self):
        if self._selected_item is None:
            return
        self._selected_item.set_colorbar_label(self.colorbar_edit.text().strip())
        self._push_undo_state()

    def _on_text_scale_changed(self, value: int):
        scale = max(0.01, min(2.4, value / 100.0))
        self._global_text_scale = scale
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item._fixed_text_scale_value = scale
                item._use_fixed_text_scale = True
                # Clear any alignment-locked text scale so the slider takes effect.
                item.set_locked_text_scale(None)
                item._update_rendered_pixmap()
        self._push_undo_state()

    def _on_font_color_auto_toggled(self, checked: bool):
        self._global_text_color = None if checked else self._global_text_color
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_text_color_override(self._global_text_color)

    def _on_font_color_pick(self):
        color = QtWidgets.QColorDialog.getColor(self._global_text_color or QtGui.QColor("#ffffff"), self, "Select font color")
        if not color.isValid():
            return
        self._global_text_color = color
        self.font_color_auto_check.setChecked(False)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_text_color_override(self._global_text_color)

    def _on_scale_bar_toggled(self, checked: bool):
        self._global_show_scale_bar = bool(checked)
        for attr_name in ("scale_bar_check", "toolbar_scale_bar_check"):
            widget = getattr(self, attr_name, None)
            if widget is not None:
                try:
                    widget.blockSignals(True)
                    widget.setChecked(self._global_show_scale_bar)
                finally:
                    widget.blockSignals(False)
        self._apply_to_canvas_items(lambda item: item.set_show_scale_bar(self._global_show_scale_bar))
        self._mark_display_preset_custom()

    def _on_scale_bar_size_changed(self, text: str):
        for combo_name in ("scale_bar_combo", "left_scale_bar_combo"):
            combo = getattr(self, combo_name, None)
            if combo is not None and combo.currentText() != text:
                try:
                    combo.blockSignals(True)
                    combo.setCurrentText(text)
                finally:
                    combo.blockSignals(False)
        if text.lower().startswith("auto"):
            self._global_scale_bar_length_nm = None
        else:
            try:
                self._global_scale_bar_length_nm = float(text.split()[0])
            except Exception:
                self._global_scale_bar_length_nm = None
        for item in self.scene.items():
            if not isinstance(item, CanvasImageItem):
                continue
            length = self._convert_scale_bar_length(item._axis_unit, self._global_scale_bar_length_nm)
            item.set_scale_bar_length(length)

    def _selected_target_items(self):
        selected = self._selected_canvas_items()
        if not selected and self._selected_item is not None:
            selected = [self._selected_item]
        return selected

    def _persist_recent_molecule_path(self, path: str):
        try:
            norm = str(Path(path).resolve())
        except Exception:
            norm = str(path)
        recent = [norm]
        for old in list(self._recent_molecule_paths):
            if old != norm and old not in recent:
                recent.append(old)
        self._recent_molecule_paths = recent[:8]
        try:
            self.viewer.recent_molecules = list(self._recent_molecule_paths)
        except Exception:
            pass
        try:
            if hasattr(self.viewer, "_on_recent_molecules_updated"):
                self.viewer._on_recent_molecules_updated(self._recent_molecule_paths)
        except Exception:
            pass

    def _persist_item_molecules(self, item: CanvasImageItem):
        try:
            store = getattr(self.viewer, "molecule_overlays", None)
            if isinstance(store, dict):
                store[str(item.file_path)] = item.export_molecule_state()
        except Exception:
            pass

    def _on_canvas_show_molecules_toggled(self, checked: bool):
        self._global_show_molecules = bool(checked)
        widget = getattr(self, "canvas_molecules_check", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setChecked(self._global_show_molecules)
            finally:
                widget.blockSignals(False)
        self._apply_to_canvas_items(lambda item: item.set_show_molecules(self._global_show_molecules))
        self.status_label.setText("Canvas molecules shown" if self._global_show_molecules else "Canvas molecules hidden")

    def _choose_canvas_molecule_path(self):
        recent = [p for p in self._recent_molecule_paths if p]
        chosen_path = None
        if recent:
            menu = QtWidgets.QMenu(self)
            actions = {}
            for path in recent[:8]:
                act = menu.addAction(str(path))
                actions[act] = path
            browse_act = menu.addAction("Browse...")
            anchor = QtGui.QCursor.pos()
            btn = getattr(self, "canvas_molecule_load_btn", None)
            if btn is not None:
                anchor = btn.mapToGlobal(btn.rect().bottomLeft())
            chosen = menu.exec_(anchor)
            if chosen in actions:
                chosen_path = actions[chosen]
            elif chosen != browse_act:
                return None
        if chosen_path:
            return chosen_path
        start_dir = ""
        if recent:
            try:
                start_dir = str(Path(recent[0]).parent)
            except Exception:
                start_dir = ""
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load Molecule",
            start_dir,
            "Molecule Files (*.xyz *.pdb *.mol);;All Files (*)",
        )
        return path or None

    def _on_canvas_load_molecule(self):
        items = self._selected_target_items()
        if not items:
            self.status_label.setText("Select one or more canvas tiles to load a molecule")
            return
        path = self._choose_canvas_molecule_path()
        if not path:
            return
        loaded = 0
        for item in items:
            try:
                item.set_molecule_palette(self._molecule_palette)
                if item.add_molecule_from_path(path):
                    item.set_show_molecules(True)
                    self._persist_item_molecules(item)
                    loaded += 1
            except Exception:
                continue
        if loaded:
            self._persist_recent_molecule_path(path)
            self._on_canvas_show_molecules_toggled(True)
            self._push_undo_state()
            self.status_label.setText(f"Loaded molecule onto {loaded} canvas tile(s)")

    def _on_canvas_clear_molecules(self):
        items = self._selected_target_items()
        if not items:
            self.status_label.setText("Select one or more canvas tiles to clear molecules")
            return
        cleared = 0
        for item in items:
            try:
                if item.export_molecule_state():
                    item.clear_molecules()
                    self._persist_item_molecules(item)
                    cleared += 1
            except Exception:
                continue
        if cleared:
            self._push_undo_state()
            self.status_label.setText(f"Cleared molecules from {cleared} canvas tile(s)")

    def _convert_scale_bar_length(self, unit: str, length_nm: float | None) -> float | None:
        if length_nm is None:
            return None
        unit_norm = (unit or "").strip().lower()
        if unit_norm in ("a", "å", "angstrom", "angstroms"):
            return length_nm * 10.0
        return length_nm

    def _on_cmap_changed(self, name: str):
        if self._selected_item is None or not name:
            return
        self._selected_item.set_cmap(name)
        kind = self._selected_item.kind or self._infer_kind_for_item(self._selected_item)
        if kind:
            self._kind_cmap[kind] = name
            if self._sync_by_channel:
                for item in self.scene.items():
                    if isinstance(item, CanvasImageItem):
                        item_kind = item.kind or self._infer_kind_for_item(item)
                        if item_kind == kind:
                            item.set_cmap(name)
        if self._sync_colorbars:
            self._sync_all_colorbars()
        self._push_undo_state()

    def _on_apply_cmap_to_selected(self, name: str):
        if not name:
            return
        selected = self._selected_canvas_items()
        if not selected and self._selected_item is not None:
            selected = [self._selected_item]
        if not selected:
            return
        for item in selected:
            item.set_cmap(name)
            kind = item.kind or self._infer_kind_for_item(item)
            if kind:
                self._kind_cmap[kind] = name
        if self._sync_colorbars:
            self._sync_all_colorbars()
        self._push_undo_state()

    def _on_copy_cmap(self):
        if self._selected_item is None:
            return
        cmap = self._selected_item.cmap
        for item in self._selected_canvas_items():
            item.set_cmap(cmap)
        self._push_undo_state()

    def _on_range_changed(self):
        if self._selected_item is None:
            return
        vmin = _safe_float(self.vmin_edit.text())
        vmax = _safe_float(self.vmax_edit.text())
        if vmin is None or vmax is None:
            return
        self._selected_item.set_range(vmin, vmax)
        if self._sync_colorbars:
            self._sync_all_colorbars()
        self._push_undo_state()

    def _on_auto_range(self):
        if self._selected_item is None:
            return
        self._selected_item.set_range(None, None)
        self.vmin_edit.setText("")
        self.vmax_edit.setText("")
        if self._sync_colorbars:
            self._sync_all_colorbars()
        self._push_undo_state()

    def _on_auto_range_selected(self):
        selected = self._selected_canvas_items()
        if not selected and self._selected_item is not None:
            selected = [self._selected_item]
        if not selected:
            return
        for item in selected:
            item.set_range(None, None)
        if self._selected_item in selected:
            self.vmin_edit.setText("")
            self.vmax_edit.setText("")
        if self._sync_colorbars:
            self._sync_all_colorbars()
        self._push_undo_state()

    def _on_copy_range(self):
        if self._selected_item is None:
            return
        vmin = self._selected_item.vmin
        vmax = self._selected_item.vmax
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem) and item.isSelected():
                item.set_range(vmin, vmax)
        self._push_undo_state()

    def _on_metadata_bar_toggled(self, checked: bool):
        self._metadata_bar_default = bool(checked)
        widget = getattr(self, "toolbar_metadata_bar_check", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setChecked(self._metadata_bar_default)
            finally:
                widget.blockSignals(False)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
        self._mark_display_preset_custom()

    def _on_metadata_unit_toggled(self, checked: bool):
        self._metadata_unit_default = bool(checked)
        widget = getattr(self, "toolbar_unit_badge_check", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setChecked(self._metadata_unit_default)
            finally:
                widget.blockSignals(False)
        self._apply_to_canvas_items(lambda item: item.set_metadata_unit_visible(self._metadata_unit_default))
        self._mark_display_preset_custom()

    def _on_global_show_title_toggled(self, checked: bool):
        self._global_show_title = bool(checked)
        widget = getattr(self, "toolbar_show_title_check", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setChecked(self._global_show_title)
            finally:
                widget.blockSignals(False)
        self._apply_to_canvas_items(lambda item: item.set_show_title(self._global_show_title))
        self._mark_display_preset_custom()

    def _on_duplicate_item(self):
        if self._selected_item is None:
            return
        state = self._selected_item.to_state()
        item = self._add_view_from_header(Path(state["file_path"]), int(state["channel_index"]), cmap_override=state.get("cmap"))
        if item:
            item.apply_state(state)
            item.setPos(item.pos() + QtCore.QPointF(30, 30))
        self._push_undo_state()

    def _on_sync_colorbars_toggled(self, checked: bool):
        self._sync_colorbars = checked
        if checked:
            self._sync_all_colorbars()

    def _on_sync_by_channel_toggled(self, checked: bool):
        self._sync_by_channel = bool(checked)
        if checked:
            self._sync_colors_by_channel()

    def _on_overlay_info_toggled(self, checked: bool):
        self._show_overlay_info = bool(checked)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_overlay(self._show_overlay_info, self._show_overlay_file)
                item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
                item.set_metadata_unit_visible(self._metadata_unit_default)
        self._mark_display_preset_custom()

    def _on_overlay_file_toggled(self, checked: bool):
        self._show_overlay_file = bool(checked)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_overlay(self._show_overlay_info, self._show_overlay_file)
                item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
                item.set_metadata_unit_visible(self._metadata_unit_default)
                item.set_metadata_file_visible(self._show_overlay_file)
        self._mark_display_preset_custom()

    def _metadata_bar_visible_default(self) -> bool:
        return bool(self._metadata_bar_default)

    def _on_colorbar_position_changed(self, text: str):
        mode = text.lower()
        if mode == "hidden":
            mode = "none"
        mode = mode if mode in ("bottom", "top", "left", "right", "inset", "none") else "bottom"
        self._colorbar_mode = mode
        widget = getattr(self, "colorbar_mode_combo", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setCurrentText(mode.capitalize() if mode != "none" else "Hidden")
            finally:
                widget.blockSignals(False)
        self._apply_colorbar_mode_to_all(mode)
        self._mark_display_preset_custom()

    def _apply_colorbar_mode_to_all(self, mode: str):
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_colorbar_mode(mode)
        self.status_label.setText(f"Colorbar mode: {mode.capitalize()}")

    def _on_global_show_colorbar_toggled(self, checked: bool):
        self._apply_global_show_colorbar(checked)

    def _on_global_show_colorbar_ticks_toggled(self, checked: bool):
        self._apply_global_show_colorbar_ticks(checked)

    def _apply_global_show_colorbar(self, show: bool):
        self._global_show_colorbar = bool(show)
        widget = getattr(self, "show_colorbar_check", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setChecked(self._global_show_colorbar)
            finally:
                widget.blockSignals(False)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_colorbar(self._global_show_colorbar)
        self._mark_display_preset_custom()

    def _apply_global_show_colorbar_ticks(self, show: bool):
        self._global_show_colorbar_ticks = bool(show)
        widget = getattr(self, "colorbar_ticks_check", None)
        if widget is not None:
            try:
                widget.blockSignals(True)
                widget.setChecked(self._global_show_colorbar_ticks)
            finally:
                widget.blockSignals(False)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_colorbar_ticks(self._global_show_colorbar_ticks)
        self._mark_display_preset_custom()

    def _on_canvas_color_clicked(self):
        color = QtWidgets.QColorDialog.getColor(self.view.backgroundBrush().color(), self, "Canvas color")
        if color.isValid():
            self.view.set_background_color(color)
            for item in self.scene.items():
                if isinstance(item, CanvasImageItem):
                    item.set_frame_color(color)

    def _sync_all_colorbars(self):
        if not self._sync_colorbars or self._selected_item is None:
            return
        vmin = self._selected_item.vmin
        vmax = self._selected_item.vmax
        cmap = self._selected_item.cmap
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_range(vmin, vmax)
                item.set_cmap(cmap)

    def _sync_colors_by_channel(self):
        for item in self.scene.items():
            if not isinstance(item, CanvasImageItem):
                continue
            kind = item.kind or self._infer_kind_for_item(item)
            if kind is None:
                continue
            cmap = self._kind_cmap.get(kind)
            if cmap:
                item.set_cmap(cmap)

    def _display_channel_label(self, kind: str | None, unit_display: str | None) -> str:
        if kind == "df":
            base = "Δf"
            unit = unit_display or "Hz"
        elif kind == "current":
            base = "I_tunnel"
            unit = unit_display or "A"
        elif kind == "topo":
            base = "Topography"
            unit = unit_display or ""
        else:
            base = ""
            unit = unit_display or ""
        if unit:
            return f"{base} ({unit})" if base else f"{unit}"
        return base

    def handle_drop(self, payloads: list[dict], paths: list[str]):
        try:
            groups = []
            for payload in payloads:
                items = payload.get("items")
                file_path = payload.get("file_path")
                cmap = payload.get("cmap")
                channel_idx = payload.get("channel_index")
                # Handle multi-item payloads from thumbnail drags
                if items and isinstance(items, (list, tuple, set)):
                    for path_str in items:
                        try:
                            path_obj = Path(path_str)
                            if channel_idx is not None:
                                try:
                                    idx = int(channel_idx)
                                except Exception:
                                    idx = None
                                if idx is not None:
                                    self._add_view_from_header(path_obj, idx, cmap_override=cmap, place=True)
                                    continue
                            group = self._add_kind_views_for_header(path_obj, cmap_override=cmap)
                            if group:
                                groups.append(group)
                        except Exception as exc:
                            QtWidgets.QMessageBox.warning(self, "Canvas drop", f"Unable to load {path_str}: {exc}")
                    continue
                if not file_path:
                    continue
                try:
                    group = self._add_kind_views_for_header(Path(file_path), cmap_override=cmap)
                    if group:
                        groups.append(group)
                        continue
                    if channel_idx is not None:
                        try:
                            idx = int(channel_idx)
                        except Exception:
                            idx = None
                        if idx is not None:
                            self._add_view_from_header(Path(file_path), idx, cmap_override=cmap, place=True)
                except Exception as exc:
                    QtWidgets.QMessageBox.warning(self, "Canvas drop", f"Unable to load view: {exc}")
            for path in paths:
                try:
                    file_groups = self._add_views_from_file(Path(path))
                    if file_groups:
                        groups.extend(file_groups)
                except Exception as exc:
                    QtWidgets.QMessageBox.warning(self, "Canvas drop", f"Unable to load {path}: {exc}")
            if groups:
                self._arrange_by_kind(groups)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Canvas drop", f"Unexpected error: {exc}")

    def _add_views_from_file(self, path: Path):
        if not path.exists():
            return
        suffix = path.suffix.lower()
        if suffix == ".txt":
            try:
                header, fds = parse_header(path)
            except Exception:
                return
            return [self._add_kind_views_for_header(path, header=header, fds=fds)]
        if suffix == ".int":
            resolved = self._resolve_header_for_int(path)
            if resolved is None:
                txt_path, _ = QtWidgets.QFileDialog.getOpenFileName(
                    self,
                    "Select header for dropped .int",
                    str(path.parent),
                    "SXM headers (*.txt)",
                )
                if txt_path:
                    resolved = self._resolve_header_for_int(path, header_path=Path(txt_path))
            if resolved is None:
                QtWidgets.QMessageBox.warning(self, "Canvas drop", f"No .txt header references {path.name}")
                return
            header_path, header, fds, idx = resolved
            return [self._add_kind_views_for_header(header_path, header=header, fds=fds)]

    def _resolve_header_for_int(self, int_path: Path, header_path: Path | None = None):
        candidates = []
        if header_path is not None:
            candidates.append(Path(header_path))
        else:
            direct = int_path.with_suffix(".txt")
            if direct.exists():
                candidates.append(direct)
            candidates.extend(int_path.parent.glob("*.txt"))
        seen = set()
        for cand in candidates:
            if cand in seen:
                continue
            seen.add(cand)
            try:
                header, fds = parse_header(cand)
            except Exception:
                continue
            for idx, fd in enumerate(fds):
                fname = fd.get("FileName", "")
                if Path(str(fname)).name.lower() == int_path.name.lower():
                    return cand, header, fds, idx
        return None

    def _add_view_from_header(
        self,
        header_path: Path,
        channel_idx: int,
        cmap_override: str | None = None,
        *,
        place: bool = True,
        kind: str | None = None,
    ):
        header_path = Path(header_path)
        header, fds = None, None
        file_key = str(header_path)
        if file_key in getattr(self.viewer, "headers", {}):
            header, fds = self.viewer.headers.get(file_key, (None, None))
        if header is None or fds is None:
            try:
                header, fds = parse_header(header_path)
            except Exception:
                return None
        try:
            channel_idx = int(channel_idx)
        except Exception:
            channel_idx = 0
        if not fds:
            return None
        if channel_idx < 0 or channel_idx >= len(fds):
            channel_idx = max(0, min(len(fds) - 1, channel_idx))
        fd = fds[channel_idx]
        base_extent = self.viewer._header_extent(header)
        unit_norm, arr_base = self.viewer._get_filtered_channel_array(file_key, channel_idx, header, fd)
        arr_adj, adj_extent = self.viewer._apply_adjustments_for_channel(file_key, channel_idx, arr_base, base_extent)
        disp_extent = self.viewer._display_extent(adj_extent, header)
        unit_display, arr_display, _ = self.viewer._scale_unit_for_display(unit_norm, arr_adj)
        caption = fd.get("Caption", fd.get("FileName", f"chan{channel_idx}"))
        title = caption
        colorbar_label = self._display_channel_label(kind, unit_display) or caption
        if cmap_override is None and kind in self._kind_cmap:
            cmap = self._kind_cmap.get(kind)
        else:
            cmap = cmap_override
        if not cmap:
            cmap = self.viewer.preview_cmap_combo.currentText() or self.viewer.preview_cmap
        try:
            # Match the preview orientation (imshow origin lower); canvas sometimes appeared vertically mirrored.
            arr_display = np.flipud(arr_display)
        except Exception:
            pass
        axis_unit = header.get('XPhysUnit') or header.get('YPhysUnit') or header.get('ScanUnit') or ''
        if not axis_unit:
            axis_unit = 'px' if disp_extent is None else 'nm'
        date = str(header.get('Date', '') or '').strip()
        time_txt = str(header.get('Time', '') or '').strip()
        datetime_txt = " ".join([t for t in (date, time_txt) if t]).strip()
        overlay_label = self._display_channel_label(kind, unit_display) or caption
        overlay_txt = overlay_label
        if datetime_txt:
            overlay_txt = f"{overlay_label} | {datetime_txt}"
        file_overlay = header_path.name
        # Choose an initial canvas width based on available viewport space to avoid oversized tiles.
        try:
            window_width = float(self.width())
            canvas_width_area = window_width * 0.65  # account for inspector panel
        except Exception:
            canvas_width_area = 900.0

        existing_items = [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        if not existing_items:
            # First drop: make the primary item large for legibility
            target_cols = 2.0
            total_gap_space = 80.0 + (24.0 * (target_cols - 1))
            default_width = (canvas_width_area - total_gap_space) / target_cols
            default_width = max(340.0, min(520.0, default_width))
        else:
            # Subsequent items: moderate size grid
            num_columns = 3.0
            total_gap_space = 80.0 + (24.0 * (num_columns - 1))  # margins + gaps
            default_width = (canvas_width_area - total_gap_space) / num_columns
            default_width = max(240.0, min(320.0, default_width))

        item = CanvasImageItem(
            arr_display,
            cmap=cmap,
            title=title,
            colorbar_label=colorbar_label,
            file_path=str(header_path),
            channel_index=channel_idx,
            unit=unit_display,
            canvas_width=default_width,
        )
        self.scene.addItem(item)
        item.set_kind(kind)
        item.set_scale_info(disp_extent, axis_unit)
        item.set_overlay_text(overlay_txt, file_overlay)
        item.set_show_title(self._global_show_title)
        item.set_show_overlay(self._show_overlay_info, self._show_overlay_file)
        item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
        item.set_metadata_unit_visible(self._metadata_unit_default)
        item.set_metadata_file_visible(self._show_overlay_file)
        item.set_show_colorbar(self._global_show_colorbar)
        item.set_show_colorbar_ticks(self._global_show_colorbar_ticks)
        item.set_colorbar_mode(self._colorbar_mode)
        item._fixed_text_scale_value = self._global_text_scale
        item._use_fixed_text_scale = True
        item.set_text_color_override(self._global_text_color)
        item.set_molecule_palette(self._molecule_palette)
        item.set_show_scale_bar(self._global_show_scale_bar)
        item.set_scale_bar_length(self._convert_scale_bar_length(axis_unit, self._global_scale_bar_length_nm))
        item.set_parent_window(self)
        try:
            molecule_state = list((getattr(self.viewer, "molecule_overlays", {}) or {}).get(file_key, []) or [])
        except Exception:
            molecule_state = []
        if molecule_state:
            item.set_molecule_state(molecule_state)
        item.set_show_molecules(self._global_show_molecules and bool(molecule_state))
        if file_key not in self._file_scale_bars:
            self._file_scale_bars[file_key] = item._scale_bar_spec()[0] if item._scale_bar_spec() else None
        item.set_scale_bar_length(self._file_scale_bars.get(file_key))
        item.set_frame_color(self.view.backgroundBrush().color())
        if place:
            self._place_item(item)
        try:
            self.scene.clearSelection()
            item.setSelected(True)
            self._selected_item = item
        except Exception:
            pass
        self.status_label.setText(f"Added {caption}")
        self._push_undo_state()
        return item

    def _add_kind_views_for_header(
        self,
        header_path: Path,
        *,
        header: dict | None = None,
        fds: list | None = None,
        cmap_override: str | None = None,
    ):
        header_path = Path(header_path)
        if header is None or fds is None:
            try:
                header, fds = parse_header(header_path)
            except Exception:
                return None
        if not fds:
            return None
        indices = self._find_kind_channel_indices(fds)
        group = {}
        for kind, idx in indices.items():
            item = self._add_view_from_header(
                header_path,
                idx,
                cmap_override=cmap_override,
                place=False,
                kind=kind,
            )
            if item is not None:
                group[kind] = item
        if group:
            shared_width = None
            for item in group.values():
                try:
                    width = float(item.get_canvas_width())
                except Exception:
                    width = None
                if width is not None:
                    shared_width = width if shared_width is None else max(shared_width, width)
            if shared_width is not None:
                for item in group.values():
                    try:
                        item.set_canvas_width(shared_width)
                    except Exception:
                        continue
        return group if group else None

    def _find_kind_channel_indices(self, fds: list) -> dict:
        indices = {}
        topo_idx = _find_topography_channel(fds)
        if topo_idx is not None:
            indices["topo"] = topo_idx
        current_idx = self._find_channel_by_tokens(
            fds,
            tokens=("it_to_pc", "it to pc", "it-to-pc", "current"),
            avoid=("setpoint", "feedback"),
        )
        if current_idx is not None:
            indices["current"] = current_idx
        df_idx = self._find_channel_by_tokens(
            fds,
            tokens=("df", "d f", "frequency shift", "freq shift"),
            avoid=("dft",),
        )
        if df_idx is not None:
            indices["df"] = df_idx
        return indices

    def _find_channel_by_tokens(self, fds: list, tokens: tuple, avoid: tuple = ()) -> int | None:
        def normalize(text: str) -> str:
            cleaned = []
            for ch in text.lower():
                cleaned.append(ch if ch.isalnum() else " ")
            return " ".join("".join(cleaned).split())

        for idx, fd in enumerate(fds):
            fname = normalize(fd.get("FileName", "") or "")
            if fname:
                if any(bad in fname for bad in avoid):
                    continue
                for tok in tokens:
                    if tok in fname:
                        return idx
            raw = f"{fd.get('Caption','')} {fd.get('FileName','')} {fd.get('PhysUnit','')}"
            norm = normalize(raw)
            if any(bad in norm for bad in avoid):
                continue
            for tok in tokens:
                if tok in norm:
                    return idx
        return None

    def _arrange_by_kind(self, groups: list[dict]):
        return canvas_window_actions.arrange_by_kind(self, groups)

    def _reflow_items_in_grid(self, items, target_width=None):
        tiles = [i for i in items if isinstance(i, CanvasImageItem)]
        if not tiles:
            return
        if target_width is None:
            widths = [float(i.boundingRect().width()) for i in tiles if i.boundingRect().width() > 0]
            if widths:
                target_width = float(np.median(widths))
            else:
                target_width = 280.0
        cols = max(1, int(round(math.sqrt(len(tiles)))))
        rows = int(math.ceil(len(tiles) / cols))
        ordered = sorted(tiles, key=lambda it: (round(it.pos().y(), 2), round(it.pos().x(), 2)))
        margin = CANVAS_ALIGN_MARGIN
        gap_x = CANVAS_ALIGN_GAP
        gap_y = CANVAS_ALIGN_GAP
        index = 0
        y = margin
        for _ in range(rows):
            row_items = ordered[index : index + cols]
            index += cols
            if not row_items:
                continue
            row_height = max(float(item.boundingRect().height()) for item in row_items)
            x = margin
            for item in row_items:
                item.setPos(x, y)
                width = float(item.boundingRect().width()) or target_width
                x += width + gap_x
            y += row_height + gap_y

    def _on_align_selected(self):
        selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        if not selected:
            return
        if len(selected) == 1:
            # Treat single selection as "match all to this size"
            self._apply_size_from_reference(selected[0], items=None, reflow=True)
            return
        min_x = min(item.pos().x() for item in selected)
        for item in selected:
            item.setPos(min_x, item.pos().y())
        self._push_undo_state()

    def _reset_locked_alignment(self):
        """Completely reset alignment state for all items."""
        self._last_aligned_width = None
        self._grid_locked = False  # unlock grid
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_locked_text_scale(None)
        self.status_label.setText("Alignment reset - items can be freely resized")
        self._push_undo_state()

    def _break_alignment_for_item(self, item: CanvasImageItem):
        """Break alignment lock for a specific item that was manually resized."""
        if self._grid_locked:
            # Keep global lock but allow this item to change text scale
            item.set_locked_text_scale(None)
        else:
            item.set_locked_text_scale(None)

    def _on_align_by_channels(self):
        items = [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        if not items:
            return
        selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        ref_item = selected[0] if selected else items[0]
        target_width = ref_item.get_canvas_width()
        self._last_aligned_width = target_width
        target_scale = ref_item._effective_text_scale()
        for item in items:
            item.set_canvas_width(target_width)
            item.set_locked_text_scale(target_scale)
        self._grid_locked = True
        groups = {}
        for item in items:
            kind = item.kind or self._infer_kind_for_item(item)
            if kind is None:
                continue
            groups.setdefault(item.file_path, {})[kind] = item
        if not groups:
            self._reflow_items_in_grid(items, target_width)
            self.status_label.setText(
                f"\U0001f512 Grid locked at {target_width:.0f}px width - click Reset alignment to unlock"
            )
            self._push_undo_state()
            return
        # Determine ordering: keep existing x-position ordering if available, else by file name
        columns = []
        for file_path, group in groups.items():
            min_x = min((item.pos().x() for item in group.values()), default=0.0)
            columns.append((min_x, str(file_path), group))
        columns.sort(key=lambda entry: (round(entry[0], 3), entry[1]))
        # Build kind order using a priority list, then alphabetical fallback
        all_kinds = set()
        for _, _, group in columns:
            all_kinds.update(group.keys())
        priority = ["topo", "topography", "z", "current", "it", "df", "freq", "phase"]
        def _kind_key(k):
            k_low = str(k).lower()
            if k_low in priority:
                return (0, priority.index(k_low))
            return (1, k_low)
        kinds = sorted(all_kinds, key=_kind_key)
        margin = CANVAS_ALIGN_MARGIN
        gap_x = CANVAS_ALIGN_GAP
        gap_y = CANVAS_ALIGN_GAP
        col_widths = []
        for _, _, group in columns:
            width = max((item.boundingRect().width() for item in group.values()), default=target_width or 200.0)
            col_widths.append(max(width, 200.0))
        row_heights = []
        for kind in kinds:
            height = 0.0
            for _, _, group in columns:
                item = group.get(kind)
                if item is not None:
                    height = max(height, item.boundingRect().height())
            row_heights.append(max(height, 0.0))
        for col_idx, (_, _, group) in enumerate(columns):
            x = margin + sum(col_widths[:col_idx]) + gap_x * col_idx
            for row_idx, kind in enumerate(kinds):
                item = group.get(kind)
                if item is None:
                    continue
                y = margin + sum(row_heights[:row_idx]) + gap_y * row_idx
                item.setPos(x, y)
        self.status_label.setText(
            f"\U0001f512 Grid locked at {target_width:.0f}px width - click Reset alignment to unlock"
        )
        self._push_undo_state()

    def _apply_size_from_reference(self, ref_item: CanvasImageItem, items=None, reflow: bool = True):
        if items is None:
            items = [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        target_width = ref_item.get_canvas_width()
        target_scale = ref_item._effective_text_scale()
        for item in items:
            item.set_canvas_width(target_width)
            item.set_locked_text_scale(target_scale)
        if reflow:
            self._reflow_items_in_grid(items, target_width)
        self._last_aligned_width = target_width
        self._grid_locked = True
        self.status_label.setText(
            f"\U0001f512 Matched size to reference ({target_width:.0f}px) - click Reset alignment to unlock"
        )
        self._push_undo_state()

    def _propagate_resize(self, source: CanvasImageItem, new_width: float, text_scale: float | None = None):
        selected = []
        try:
            selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        except Exception:
            selected = []
        if len(selected) <= 1:
            return
        for item in selected:
            if item is source:
                continue
            item.set_canvas_width(new_width)
            if text_scale is not None:
                item.set_locked_text_scale(text_scale)

    def _finalize_resize_group(self, source: CanvasImageItem):
        selected = []
        try:
            selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        except Exception:
            selected = []
        if len(selected) <= 1:
            return
        target_width = source.get_canvas_width()
        self._reflow_items_in_grid(selected, target_width)
        self._push_undo_state()

    def _on_polish_layout(self):
        items = [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        if not items:
            return
        target_width = self._last_aligned_width
        if target_width is None:
            widths = [float(i.boundingRect().width()) for i in items if i.boundingRect().width() > 0]
            target_width = float(np.median(widths)) if widths else 300.0
        for item in items:
            try:
                item.set_canvas_width(target_width)
                item.set_show_colorbar(self._global_show_colorbar)
                item.set_show_colorbar_ticks(self._global_show_colorbar_ticks)
                item.set_show_overlay(self._show_overlay_info, self._show_overlay_file)
                item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
                item.set_metadata_unit_visible(self._metadata_unit_default)
                item.set_show_title(self._global_show_title)
                item._fixed_text_scale_value = self._global_text_scale
                item._use_fixed_text_scale = True
                item.set_locked_text_scale(None)
                item._update_rendered_pixmap()
            except Exception:
                continue
        self._last_aligned_width = target_width
        self._grid_locked = True
        self._reflow_items_in_grid(items, target_width)
        self.status_label.setText("Polished layout: normalized tile sizes and annotations")
        self._push_undo_state()
        return
        kinds = ["topo", "current", "df"]
        columns = []
        for file_path, group in groups.items():
            min_x = min((item.pos().x() for item in group.values()), default=0.0)
            columns.append((min_x, file_path, group))
        columns.sort(key=lambda entry: entry[0])
        margin = CANVAS_ALIGN_MARGIN
        gap_x = CANVAS_ALIGN_GAP
        gap_y = CANVAS_ALIGN_GAP
        col_widths = []
        for _, _, group in columns:
            width = max((item.boundingRect().width() for item in group.values()), default=200.0)
            col_widths.append(max(width, 200.0))
        row_heights = []
        for kind in kinds:
            height = 0.0
            for _, _, group in columns:
                item = group.get(kind)
                if item is not None:
                    height = max(height, item.boundingRect().height())
            row_heights.append(max(height, 0.0))
        for col_idx, (_, _, group) in enumerate(columns):
            x = margin + sum(col_widths[:col_idx]) + gap_x * col_idx
            for row_idx, kind in enumerate(kinds):
                item = group.get(kind)
                if item is None:
                    continue
                y = margin + sum(row_heights[:row_idx]) + gap_y * row_idx
                item.setPos(x, y)
        self.status_label.setText(
            f"🔒 Grid locked at {target_width:.0f}px width - click Reset alignment to unlock"
        )
        self._push_undo_state()

    def _infer_kind_for_item(self, item: CanvasImageItem) -> str | None:
        file_key = str(item.file_path)
        header, fds = self.viewer.headers.get(file_key, (None, None))
        if header is None or fds is None:
            try:
                header, fds = parse_header(Path(file_key))
            except Exception:
                return None
        if not fds:
            return None
        indices = self._find_kind_channel_indices(fds)
        for kind, idx in indices.items():
            if idx == item.channel_index:
                item.set_kind(kind)
                return kind
        return None

    def _apply_layout(self, layout_type: str):
        return canvas_window_actions.apply_layout(self, layout_type)

    def _on_export_image(self):
        return canvas_window_actions.on_export_image(self)

    def _on_save_canvas(self):
        return canvas_window_actions.on_save_canvas(self)

    def _on_load_canvas(self):
        return canvas_window_actions.on_load_canvas(self)

    def _delete_selected(self):
        delete_selected(self)

    def _on_remove_item(self):
        self._delete_selected()

    def _handle_canvas_key(self, event: QtGui.QKeyEvent) -> bool:
        if event is None:
            return False
        mods = event.modifiers()
        key = event.key()
        if key in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
            self._delete_selected()
            event.accept()
            return True
        if mods & QtCore.Qt.ControlModifier and key == QtCore.Qt.Key_Z:
            self._undo()
            event.accept()
            return True
        if mods & QtCore.Qt.ControlModifier and key == QtCore.Qt.Key_Y:
            self._redo()
            event.accept()
            return True
        return False

    def keyPressEvent(self, event: QtGui.QKeyEvent):
        if self._handle_canvas_key(event):
            return
        super().keyPressEvent(event)

    def _capture_state(self):
        return capture_state(self)

    def _restore_state(self, state):
        restore_state(self, state)

    def _push_undo_state(self):
        push_undo_state(self)

    def _undo(self):
        undo(self)

    def _redo(self):
        redo(self)

    def _place_item(self, item: CanvasImageItem):
        return canvas_window_actions.place_item(self, item)






