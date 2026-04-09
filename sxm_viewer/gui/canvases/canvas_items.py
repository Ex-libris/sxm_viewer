"""Canvas items and context menus for the scientific canvas."""
from __future__ import annotations

import io

from ..._shared import QtCore, QtGui, QtWidgets, np, matplotlib
from .canvas_rendering import render_tile_mpl, render_tile_figure_mpl, _text_color_for_frame
from .molecular_overlay import Molecule, MoleculePropertiesDialog, available_atom_palettes, get_atom_color


def _append_canvas_menu_actions(menu: QtWidgets.QMenu, parent, view):
    actions = {}
    if parent is None or view is None:
        return actions

    actions["align_selected"] = menu.addAction("Align selected")
    actions["align_by_channel"] = menu.addAction("Align by channel")
    actions["reset_alignment"] = menu.addAction("Reset alignment")
    menu.addSeparator()

    range_menu = menu.addMenu("Range")
    actions["auto_range"] = range_menu.addAction("Auto range for selected")
    actions["copy_range"] = range_menu.addAction("Copy range to selected")
    range_menu.addSeparator()
    actions["sync_ranges"] = range_menu.addAction("Sync ranges")
    actions["sync_ranges"].setCheckable(True)
    actions["sync_ranges"].setChecked(bool(getattr(parent, "_sync_colorbars", False)))

    cmap_menu = menu.addMenu("Colormap")
    actions["copy_cmap"] = cmap_menu.addAction("Copy colormap to selected")
    cmap_menu.addSeparator()
    actions["sync_colors_by_channel"] = cmap_menu.addAction("Sync colors by channel")
    actions["sync_colors_by_channel"].setCheckable(True)
    actions["sync_colors_by_channel"].setChecked(bool(getattr(parent, "_sync_by_channel", False)))
    cmap_actions = {}
    for cmap_name in ("viridis", "plasma", "magma", "inferno", "cividis", "afmhot", "gray", "Blues_r", "RdBu_r"):
        cmap_actions[cmap_menu.addAction(cmap_name)] = cmap_name
    actions["cmap_actions"] = cmap_actions

    menu.addSeparator()
    overlay_menu = menu.addMenu("Overlay")
    actions["overlay_info"] = overlay_menu.addAction("Channel/date")
    actions["overlay_info"].setCheckable(True)
    actions["overlay_info"].setChecked(bool(getattr(parent, "_show_overlay_info", False)))
    actions["overlay_file"] = overlay_menu.addAction("Filename")
    actions["overlay_file"].setCheckable(True)
    actions["overlay_file"].setChecked(bool(getattr(parent, "_show_overlay_file", False)))

    view_menu = menu.addMenu("View")
    actions["show_grid"] = view_menu.addAction("Show grid")
    actions["show_grid"].setCheckable(True)
    actions["show_grid"].setChecked(bool(getattr(view, "_show_grid", False)))
    actions["snap_grid"] = view_menu.addAction("Snap to grid")
    actions["snap_grid"].setCheckable(True)
    actions["snap_grid"].setChecked(bool(getattr(view, "_snap_to_grid", False)))
    actions["canvas_color"] = view_menu.addAction("Canvas color...")

    display_menu = menu.addMenu("Display")
    actions["show_metadata_bar"] = display_menu.addAction("Show metadata bar")
    actions["show_metadata_bar"].setCheckable(True)
    actions["show_metadata_bar"].setChecked(bool(getattr(parent, "_metadata_bar_default", True)))
    actions["show_unit_badge"] = display_menu.addAction("Show unit badge")
    actions["show_unit_badge"].setCheckable(True)
    actions["show_unit_badge"].setChecked(bool(getattr(parent, "_metadata_unit_default", True)))
    actions["show_title"] = display_menu.addAction("Show title")
    actions["show_title"].setCheckable(True)
    actions["show_title"].setChecked(bool(getattr(parent, "_global_show_title", False)))
    display_menu.addSeparator()
    actions["show_colorbar"] = display_menu.addAction("Show colorbar")
    actions["show_colorbar"].setCheckable(True)
    actions["show_colorbar"].setChecked(bool(getattr(parent, "_global_show_colorbar", True)))
    actions["show_colorbar_ticks"] = display_menu.addAction("Show colorbar ticks")
    actions["show_colorbar_ticks"].setCheckable(True)
    actions["show_colorbar_ticks"].setChecked(bool(getattr(parent, "_global_show_colorbar_ticks", True)))
    actions["show_scale_bar"] = display_menu.addAction("Show scale bar")
    actions["show_scale_bar"].setCheckable(True)
    actions["show_scale_bar"].setChecked(bool(getattr(parent, "_global_show_scale_bar", False)))
    cbar_menu = display_menu.addMenu("Colorbar position")
    cbar_position_actions = {}
    current_mode = str(getattr(parent, "_colorbar_mode", "bottom") or "bottom")
    for mode, label in (
        ("bottom", "Bottom"),
        ("top", "Top"),
        ("left", "Left"),
        ("right", "Right"),
        ("inset", "Inset"),
        ("none", "Hidden"),
    ):
        act = cbar_menu.addAction(label)
        act.setCheckable(True)
        act.setChecked(current_mode == mode)
        cbar_position_actions[act] = mode
    actions["cbar_position_actions"] = cbar_position_actions

    molecules_menu = menu.addMenu("Molecules")
    actions["show_molecules"] = molecules_menu.addAction("Show molecules")
    actions["show_molecules"].setCheckable(True)
    actions["show_molecules"].setChecked(bool(getattr(parent, "_global_show_molecules", False)))
    molecules_menu.addSeparator()
    actions["load_molecule"] = molecules_menu.addAction("Load onto selected...")
    actions["clear_molecules"] = molecules_menu.addAction("Clear from selected")

    layout_menu = menu.addMenu("Layout")
    actions["layout_2x2"] = layout_menu.addAction("2x2")
    actions["layout_1x3"] = layout_menu.addAction("1x3")
    actions["layout_3x1"] = layout_menu.addAction("3x1")
    return actions

class AlignmentGuide(QtWidgets.QGraphicsLineItem):
    """Visual guide shown when items are aligned."""
    def __init__(self, x1, y1, x2, y2):
        super().__init__(x1, y1, x2, y2)
        pen = QtGui.QPen(QtGui.QColor(100, 150, 255, 180), 1, QtCore.Qt.DashLine)
        self.setPen(pen)
        self.setZValue(1000)


class RubberBandSelection(QtWidgets.QGraphicsRectItem):
    """Visual rubber band for drag selection."""
    def __init__(self):
        super().__init__()
        self.setPen(QtGui.QPen(QtGui.QColor(100, 150, 255), 1, QtCore.Qt.DashLine))
        self.setBrush(QtGui.QBrush(QtGui.QColor(100, 150, 255, 30)))
        self.setZValue(999)


class CanvasImageItem(QtWidgets.QGraphicsObject):
    resize_handle_size = 12

    def __init__(
        self,
        arr: np.ndarray,
        *,
        cmap: str,
        title: str,
        colorbar_label: str,
        file_path: str,
        channel_index: int,
        unit: str | None = None,
        vmin: float | None = None,
        vmax: float | None = None,
        canvas_width: float = 280.0,
    ):
        super().__init__()
        self.setFlags(
            QtWidgets.QGraphicsItem.ItemIsSelectable
            | QtWidgets.QGraphicsItem.ItemIsMovable
            | QtWidgets.QGraphicsItem.ItemSendsGeometryChanges
        )
        self.setAcceptHoverEvents(True)
        self.setAcceptedMouseButtons(QtCore.Qt.LeftButton | QtCore.Qt.RightButton | QtCore.Qt.MiddleButton)
        self._arr = np.asarray(arr)
        self._cmap = cmap
        self._title = title
        self._colorbar_label = colorbar_label or ""
        self._unit = unit or ""
        self._file_path = str(file_path)
        self._channel_index = int(channel_index)
        self._vmin = vmin
        self._vmax = vmax
        self._title_height = 18
        self._colorbar_height = 10
        self._colorbar_pad_y = 4
        self._colorbar_padding_x = 6
        self._show_title = False
        self._show_colorbar = True
        self._show_colorbar_ticks = True
        self._canvas_width = float(canvas_width)
        self._full_dpi = 320
        self._fast_dpi = 180
        self._fast_render = False
        self._colorbar_width = 16
        self._colorbar_mode = "bottom"
        self._use_fixed_text_scale = True
        self._fixed_text_scale_value = 1.0
        self._show_scale_bar = False
        self._text_color_override: QtGui.QColor | None = None
        self._rendered_pixmap: QtGui.QPixmap | None = None
        self._render_timer = QtCore.QTimer()
        self._render_timer.setSingleShot(True)
        self._render_timer.setInterval(60)
        self._render_timer.timeout.connect(self._render_now)
        self._render_pending = False
        self._resizing = False
        self._resize_origin = None
        self._resize_size = None
        self._resize_start_canvas_width: float | None = None
        self._kind = None
        self._keep_aspect = True
        self._locked_text_scale: float | None = None
        self._image_aspect = self._compute_image_aspect()
        self._extent = None
        self._axis_unit = ""
        self._show_overlay_main = False
        self._show_overlay_file = False
        self._overlay_main_text = ""
        self._overlay_file_text = ""
        self._frame_color = None
        self._base_image_width = float(max(1.0, self._canvas_width))
        self._parent_window = None
        self._scale_bar_length = None
        self._metadata_height = 24
        self._metadata_padding = 8
        self._metadata_bar_visible = True
        self._metadata_unit_visible = True
        self._metadata_file_visible = False
        self._metadata_left_text = ""
        self._metadata_right_text = ""
        self._quick_chip_rects: dict[str, QtCore.QRectF] = {}
        self._molecule_state: list[dict] = []
        self._show_molecules = False
        self._molecule_palette = "pymol"
        self._show_hydrogens = True
        self._molecule_drag_idx: int | None = None
        self._molecule_drag_mode: str | None = None
        self._molecule_drag_start_data: tuple[float, float] | None = None
        self._molecule_drag_start_scene: QtCore.QPointF | None = None
        self._molecule_drag_start_offset = None
        self._molecule_drag_start_angles = None
        self._molecule_history: list[list[dict]] = []
        self._molecule_props_dialog = None
        self._refresh_metadata_text()
        self._render_pending = True
        self._render_now()

    def boundingRect(self) -> QtCore.QRectF:
        return QtCore.QRectF(self._rect)

    def _resize_handle_rect(self) -> QtCore.QRectF:
        size = self.resize_handle_size
        return QtCore.QRectF(
            self._rect.right() - size - 2,
            self._rect.bottom() - size - 2,
            size,
            size,
        )

    def _tile_image_size(self) -> tuple[float, float]:
        width = max(20.0, self._canvas_width)
        height = max(20.0, self._canvas_height())
        return width, height

    def _compute_image_aspect(self) -> float:
        try:
            height, width = self._arr.shape
            return max(1e-6, float(width) / max(1.0, float(height)))
        except Exception:
            return 1.0

    def _tile_total_width(self) -> float:
        width, _ = self._tile_image_size()
        if self._show_colorbar and self._colorbar_mode in ("left", "right"):
            width += self._colorbar_thickness() + self._colorbar_padding_x
        return width

    def _tile_total_height(self) -> float:
        _, height = self._tile_image_size()
        extra = 0.0
        if self._show_colorbar and self._colorbar_mode in ("bottom", "top"):
            extra += self._colorbar_thickness() + self._colorbar_pad_y
        extra += self._metadata_bar_height()
        return height + extra

    def _metadata_bar_height(self) -> float:
        if not self._metadata_bar_visible:
            return 0.0
        if not (self._metadata_left_text or self._metadata_right_text):
            return 0.0
        return self._metadata_height + (self._metadata_padding * 2)

    def _canvas_height(self) -> float:
        aspect = max(1e-6, self._image_aspect)
        return self._canvas_width / aspect

    def _effective_text_scale(self) -> float:
        if self._locked_text_scale is not None:
            return self._locked_text_scale
        if self._use_fixed_text_scale:
            return self._fixed_text_scale_value
        return self._text_scale_for_width(self._canvas_width)

    def _colorbar_thickness(self) -> float:
        scale = self._effective_text_scale()
        return max(8.0, 12.0 * scale)

    def _scale_bar_spec(self):
        if not self._extent or not self._axis_unit or self._axis_unit == "px":
            return None
        try:
            x0, x1, y1, y0 = self._extent
            width = abs(float(x1) - float(x0))
        except Exception:
            return None
        if width <= 0:
            return None
        if self._scale_bar_length:
            return self._scale_bar_length, width
        targets = [0.2 * width, 0.1 * width, 0.3 * width]
        candidates = self._scale_bar_candidates()
        best = None
        best_err = None
        for t in targets:
            for cand in candidates:
                err = abs(cand - t)
                if best is None or err < best_err:
                    best = cand
                    best_err = err
        if best is None:
            return None
        return best, width

    def _scale_bar_candidates(self) -> list[float]:
        base_nm = [0.5, 1, 2, 3, 5, 10, 20, 50, 100, 200, 500]
        unit = (self._axis_unit or "").strip().lower()
        if unit in ("a", "å", "angstrom", "angstroms"):
            return [val * 10.0 for val in base_nm]
        return base_nm

    def _scale_bar_width(self) -> float | None:
        if not self._extent or not self._axis_unit or self._axis_unit == "px":
            return None
        try:
            x0, x1, y1, y0 = self._extent
            return abs(float(x1) - float(x0))
        except Exception:
            return None

    def _update_rendered_pixmap(self):
        self._render_pending = True
        if not self._render_timer.isActive():
            self._render_timer.start()

    def _render_now(self):
        if not self._render_pending:
            return
        self._render_pending = False
        if self._arr is None:
            return
        width = max(2, int(round(self._tile_total_width())))
        height = max(2, int(round(self._tile_total_height())))
        metadata_height = self._metadata_bar_height() if self._metadata_bar_visible and (self._metadata_left_text or self._metadata_right_text) else 0.0
        text_scale = self._effective_text_scale()
        frame_color = self._frame_color.name() if isinstance(self._frame_color, QtGui.QColor) else "#070707"
        if self._text_color_override is not None and self._text_color_override.isValid():
            text_color = self._text_color_override.name()
        else:
            text_color = _text_color_for_frame(frame_color)
        show_overlay_main = self._show_overlay_main and not self._metadata_bar_visible
        show_overlay_file = self._show_overlay_file and not self._metadata_bar_visible
        scale_spec = self._scale_bar_spec()
        scale_length = scale_spec[0] if scale_spec else None
        scale_width = scale_spec[1] if scale_spec else None
        pixmap = render_tile_mpl(
            self._arr,
            cmap=self._cmap,
            vmin=self._vmin,
            vmax=self._vmax,
            title=self._title,
            colorbar_label=self._colorbar_label,
            width_px=width,
            height_px=height,
            dpi=self._fast_dpi if self._fast_render else self._full_dpi,
            show_colorbar=self._show_colorbar,
            show_colorbar_ticks=self._show_colorbar_ticks,
            show_title=self._show_title,
            show_metadata=self._metadata_bar_visible and bool(self._metadata_left_text or self._metadata_right_text),
            metadata_left=self._metadata_left_text,
            metadata_right=self._metadata_right_text,
            show_overlay_main=show_overlay_main,
            overlay_main=self._overlay_main_text,
            show_overlay_file=show_overlay_file,
            overlay_file=self._overlay_file_text,
            cbar_position=self._colorbar_mode,
            metadata_height=metadata_height,
            frame_color=frame_color,
            text_scale=text_scale,
            text_color=text_color,
            show_scale_bar=self._show_scale_bar,
            scale_bar_length=scale_length,
            scale_bar_unit=self._axis_unit,
            scale_bar_width=scale_width,
            extent=self._extent,
            show_molecules=self._show_molecules,
            molecules=self._molecule_state,
            molecule_palette=self._molecule_palette,
            show_hydrogens=self._show_hydrogens,
        )
        self.prepareGeometryChange()
        self._rendered_pixmap = pixmap
        self._rect = QtCore.QRectF(0, 0, pixmap.width(), pixmap.height())
        self.update()

    def set_canvas_width(self, width: float):
        width = max(50.0, float(width))
        if abs(width - self._canvas_width) < 1e-6:
            return
        self._canvas_width = width
        self._update_rendered_pixmap()

    def get_canvas_width(self) -> float:
        return self._canvas_width

    def reset_to_data_size(self):
        self.set_canvas_width(max(120.0, self._base_image_width))

    def paint(self, painter: QtGui.QPainter, option, widget=None):
        if self._rendered_pixmap is None:
            return
        painter.setRenderHint(QtGui.QPainter.SmoothPixmapTransform)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        bg_color = self._frame_color or QtGui.QColor("#070707")
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(QtGui.QBrush(bg_color))
        painter.drawRect(self._rect)
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.drawPixmap(QtCore.QPointF(0, 0), self._rendered_pixmap)
        if self.isSelected():
            pen = QtGui.QPen(QtGui.QColor("#4a90e2"), 2)
            pen.setCosmetic(True)
            painter.setPen(pen)
            painter.setBrush(QtCore.Qt.NoBrush)
            painter.drawRect(self._rect)
            self._draw_quick_toggle_chips(painter)
        else:
            self._quick_chip_rects = {}

    def _chip_specs(self):
        return [
            ("title", "T", self._show_title, "Toggle title"),
            ("scale", "S", self._show_scale_bar, "Toggle scale bar"),
            ("cbar", "C", self._show_colorbar, "Toggle colorbar"),
            ("meta", "M", self._metadata_bar_visible, "Toggle metadata bar"),
            ("unit", "U", self._metadata_unit_visible, "Toggle unit badge"),
            ("file", "F", self._metadata_file_visible, "Toggle filename badge"),
        ]

    def _draw_quick_toggle_chips(self, painter: QtGui.QPainter):
        chips = self._chip_specs()
        if not chips:
            self._quick_chip_rects = {}
            return
        text_color = self._text_color_override if self._text_color_override is not None else None
        if text_color is None or not text_color.isValid():
            text_color = QtGui.QColor(_text_color_for_frame(self._frame_color.name() if isinstance(self._frame_color, QtGui.QColor) else "#070707"))
        inactive_text = QtGui.QColor("#c9d1d9")
        active_fill = QtGui.QColor("#1f6feb")
        inactive_fill = QtGui.QColor(15, 18, 22, 210)
        border_col = QtGui.QColor(255, 255, 255, 35)
        chip_h = 20.0
        chip_w = 24.0
        gap = 4.0
        x = 8.0
        y = 8.0
        self._quick_chip_rects = {}
        font = painter.font()
        font.setPointSizeF(max(7.5, font.pointSizeF() if font.pointSizeF() > 0 else 8.0))
        font.setBold(True)
        painter.setFont(font)
        for key, label, enabled, _tip in chips:
            rect = QtCore.QRectF(x, y, chip_w, chip_h)
            path = QtGui.QPainterPath()
            path.addRoundedRect(rect, 6.0, 6.0)
            painter.setPen(QtGui.QPen(border_col, 1))
            painter.setBrush(active_fill if enabled else inactive_fill)
            painter.drawPath(path)
            painter.setPen(text_color if enabled else inactive_text)
            painter.drawText(rect, QtCore.Qt.AlignCenter, label)
            self._quick_chip_rects[key] = rect
            x += chip_w + gap

    def _chip_at_pos(self, pos: QtCore.QPointF) -> str | None:
        for key, rect in self._quick_chip_rects.items():
            if rect.contains(pos):
                return key
        return None

    def _toggle_quick_chip(self, key: str):
        if key == "title":
            self.set_show_title(not self._show_title)
        elif key == "scale":
            self.set_show_scale_bar(not self._show_scale_bar)
        elif key == "cbar":
            self.set_show_colorbar(not self._show_colorbar)
        elif key == "meta":
            self.set_metadata_bar_visible(not self._metadata_bar_visible)
        elif key == "unit":
            self.set_metadata_unit_visible(not self._metadata_unit_visible)
        elif key == "file":
            self.set_metadata_file_visible(not self._metadata_file_visible)
        else:
            return
        if self._parent_window is not None:
            self._parent_window._on_item_overlay_chip_toggled(self, key)

    def hoverEnterEvent(self, event):
        if self._resize_handle_rect().contains(event.pos()):
            self.setCursor(QtCore.Qt.SizeFDiagCursor)
        else:
            self.setCursor(QtCore.Qt.OpenHandCursor)

    def hoverLeaveEvent(self, event):
        self.setCursor(QtCore.Qt.ArrowCursor)

    def hoverMoveEvent(self, event):
        chip_key = self._chip_at_pos(event.pos())
        if chip_key:
            tip = next((tip for key, _label, _enabled, tip in self._chip_specs() if key == chip_key), "")
            self.setToolTip(tip)
            self.setCursor(QtCore.Qt.PointingHandCursor)
        elif self._resize_handle_rect().contains(event.pos()):
            self.setCursor(QtCore.Qt.SizeFDiagCursor)
        else:
            self.setToolTip("")
            self.setCursor(QtCore.Qt.OpenHandCursor)
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.RightButton:
            if not self.isSelected():
                self.setSelected(True)
            event.accept()
            return
        chip_key = self._chip_at_pos(event.pos())
        if event.button() == QtCore.Qt.LeftButton and chip_key:
            if not self.isSelected():
                self.setSelected(True)
            self._toggle_quick_chip(chip_key)
            event.accept()
            return
        if event.button() == QtCore.Qt.LeftButton and self._resize_handle_rect().contains(event.pos()):
            self._resizing = True
            self._fast_render = True
            self._resize_origin = event.pos()
            self._resize_size = QtCore.QSizeF(self._rect.width(), self._rect.height())
            self._resize_start_canvas_width = self._canvas_width
            event.accept()
            return
        mol_idx = self._molecule_hit(event.pos())
        if mol_idx is not None and event.button() in (QtCore.Qt.LeftButton, QtCore.Qt.MiddleButton):
            molecules = self._molecule_objects()
            if 0 <= mol_idx < len(molecules):
                data_pos = self._local_to_data(event.pos())
                if data_pos is not None:
                    self._push_molecule_snapshot()
                    self._molecule_drag_idx = mol_idx
                    self._molecule_drag_start_data = data_pos
                    self._molecule_drag_start_scene = event.scenePos()
                    self._molecule_drag_start_offset = molecules[mol_idx].offset.copy()
                    self._molecule_drag_start_angles = molecules[mol_idx].angles.copy()
                    if event.button() == QtCore.Qt.MiddleButton or (event.modifiers() & QtCore.Qt.ControlModifier and event.modifiers() & QtCore.Qt.ShiftModifier):
                        self._molecule_drag_mode = "rotate_3d"
                    elif event.modifiers() & QtCore.Qt.ShiftModifier:
                        self._molecule_drag_mode = "rotate_z"
                    else:
                        self._molecule_drag_mode = "translate"
                    event.accept()
                    return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # Alt+drag to duplicate
        if event.modifiers() & QtCore.Qt.AltModifier and not hasattr(self, '_alt_duplicated'):
            if self._parent_window:
                self._parent_window._on_duplicate_item()
                self._alt_duplicated = True
                event.accept()
                return

        if self._molecule_drag_idx is not None:
            molecules = self._molecule_objects()
            if 0 <= self._molecule_drag_idx < len(molecules):
                mol = molecules[self._molecule_drag_idx]
                if self._molecule_drag_mode == "translate":
                    data_pos = self._local_to_data(event.pos())
                    if data_pos is not None and self._molecule_drag_start_data is not None and self._molecule_drag_start_offset is not None:
                        dx = data_pos[0] - self._molecule_drag_start_data[0]
                        dy = data_pos[1] - self._molecule_drag_start_data[1]
                        mol.offset = self._molecule_drag_start_offset + np.array([dx, dy, 0.0], dtype=float)
                elif self._molecule_drag_mode == "rotate_z":
                    data_pos = self._local_to_data(event.pos())
                    if data_pos is not None and self._molecule_drag_start_data is not None and self._molecule_drag_start_offset is not None and self._molecule_drag_start_angles is not None:
                        center = self._molecule_drag_start_offset
                        v_start = np.array([self._molecule_drag_start_data[0] - center[0], self._molecule_drag_start_data[1] - center[1]], dtype=float)
                        v_now = np.array([data_pos[0] - center[0], data_pos[1] - center[1]], dtype=float)
                        if np.linalg.norm(v_start) > 1e-6 and np.linalg.norm(v_now) > 1e-6:
                            a0 = np.arctan2(v_start[1], v_start[0])
                            a1 = np.arctan2(v_now[1], v_now[0])
                            mol.angles = self._molecule_drag_start_angles.copy()
                            mol.angles[2] += float(np.degrees(a1 - a0))
                elif self._molecule_drag_mode == "rotate_3d":
                    if self._molecule_drag_start_scene is not None and self._molecule_drag_start_angles is not None:
                        dx = event.scenePos().x() - self._molecule_drag_start_scene.x()
                        dy = event.scenePos().y() - self._molecule_drag_start_scene.y()
                        mol.angles = self._molecule_drag_start_angles.copy()
                        mol.angles[0] += dy * 0.45
                        mol.angles[1] += dx * 0.45
                self._store_molecule_objects(molecules)
            event.accept()
            return
        
        if self._resizing and self._resize_origin is not None:
            delta = event.pos() - self._resize_origin
            if self._keep_aspect:
                delta_amount = delta.x() if abs(delta.x()) >= abs(delta.y()) else delta.y()
            else:
                delta_amount = delta.x()
            start_width = self._resize_start_canvas_width or self._canvas_width
            new_width = max(60.0, start_width + delta_amount)
            self.set_canvas_width(new_width)
            if self._parent_window is not None:
                self._parent_window._propagate_resize(self, new_width, text_scale=self._effective_text_scale())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if hasattr(self, '_alt_duplicated'):
            delattr(self, '_alt_duplicated')

        if self._molecule_drag_idx is not None:
            self._molecule_drag_idx = None
            self._molecule_drag_mode = None
            self._molecule_drag_start_data = None
            self._molecule_drag_start_scene = None
            self._molecule_drag_start_offset = None
            self._molecule_drag_start_angles = None
            event.accept()
            if self._parent_window is not None:
                self._parent_window._persist_item_molecules(self)
                self._parent_window._push_undo_state()
            return
        
        if self._resizing:
            self._resizing = False
            self._fast_render = False
            self._resize_origin = None
            self._resize_size = None
            self._resize_start_canvas_width = None
            event.accept()
            # Only break alignment lock if user actually changed size
            if self._parent_window is not None:
                self._parent_window._break_alignment_for_item(self)
                self._parent_window._finalize_resize_group(self)
                self._parent_window._push_undo_state()
            return
        super().mouseReleaseEvent(event)
        if self._parent_window is not None:
            self._parent_window._push_undo_state()

    def contextMenuEvent(self, event):
        mol_idx = self._molecule_hit(event.pos())
        if mol_idx is not None:
            self._show_molecule_menu(event.screenPos(), mol_idx)
            event.accept()
            return
        menu = QtWidgets.QMenu()

        duplicate_action = menu.addAction("Duplicate")
        menu.addSeparator()
        copy_svg_action = menu.addAction("Copy as SVG (vector)")
        copy_svg_selected = menu.addAction("Copy selected as SVG (vector)")
        save_svg_action = menu.addAction("Save as SVG...")
        save_pdf_action = menu.addAction("Save as PDF...")
        menu.addSeparator()
        bring_forward = menu.addAction("Bring Forward")
        send_backward = menu.addAction("Send Backward")
        menu.addSeparator()
        lock_aspect = menu.addAction("Lock Aspect Ratio")
        lock_aspect.setCheckable(True)
        lock_aspect.setChecked(self._keep_aspect)
        menu.addSeparator()
        reset_size = menu.addAction("Reset to Original Size")
        menu.addSeparator()
        delete_action = menu.addAction("Delete")

        parent = self._parent_window
        canvas_actions = {}
        if parent is not None:
            menu.addSeparator()
            canvas_menu = menu.addMenu("Canvas")
            canvas_actions = _append_canvas_menu_actions(canvas_menu, parent, getattr(parent, "view", None))

        selected_items = []
        try:
            if self.scene() is not None:
                selected_items = [i for i in self.scene().selectedItems() if isinstance(i, CanvasImageItem)]
        except Exception:
            selected_items = []
        copy_svg_selected.setEnabled(bool(selected_items))
        action = menu.exec_(event.screenPos())

        if action is not None:
            if action == duplicate_action:
                if self._parent_window:
                    self._parent_window._on_duplicate_item()
            elif action == delete_action:
                if self._parent_window:
                    self._parent_window._on_remove_item()
            elif action == copy_svg_action:
                self._copy_svg_to_clipboard()
            elif action == copy_svg_selected:
                self._copy_selected_svg()
            elif action == save_svg_action:
                self._save_vector_to_file("svg")
            elif action == save_pdf_action:
                self._save_vector_to_file("pdf")
            elif action == bring_forward:
                self.setZValue(self.zValue() + 1)
            elif action == send_backward:
                self.setZValue(self.zValue() - 1)
            elif action == lock_aspect:
                self._keep_aspect = lock_aspect.isChecked()
            elif action == reset_size:
                self.reset_to_data_size()
            else:
                self._handle_canvas_menu_action(action, canvas_actions)

        event.accept()

    def _handle_canvas_menu_action(self, action, canvas_actions):
        if not canvas_actions:
            return
        parent = self._parent_window
        view = getattr(parent, "view", None)
        if parent is None or view is None:
            return
        cmap_actions = canvas_actions.get("cmap_actions") or {}
        cbar_position_actions = canvas_actions.get("cbar_position_actions") or {}
        if action == canvas_actions.get("align_selected"):
            parent._on_align_selected()
        elif action == canvas_actions.get("align_by_channel"):
            parent._on_align_by_channels()
        elif action == canvas_actions.get("reset_alignment"):
            parent._reset_locked_alignment()
        elif action == canvas_actions.get("auto_range"):
            parent._on_auto_range_selected()
        elif action == canvas_actions.get("copy_range"):
            parent._on_copy_range()
        elif action == canvas_actions.get("sync_ranges"):
            checked = canvas_actions["sync_ranges"].isChecked()
            if hasattr(parent, "sync_cbar_check"):
                parent.sync_cbar_check.setChecked(checked)
            else:
                parent._on_sync_colorbars_toggled(checked)
        elif action == canvas_actions.get("copy_cmap"):
            parent._on_copy_cmap()
        elif action == canvas_actions.get("sync_colors_by_channel"):
            checked = canvas_actions["sync_colors_by_channel"].isChecked()
            if hasattr(parent, "sync_by_channel_check"):
                parent.sync_by_channel_check.setChecked(checked)
            else:
                parent._on_sync_by_channel_toggled(checked)
        elif action in cmap_actions:
            parent._on_apply_cmap_to_selected(cmap_actions[action])
        elif action == canvas_actions.get("overlay_info"):
            checked = canvas_actions["overlay_info"].isChecked()
            if hasattr(parent, "overlay_info_check"):
                parent.overlay_info_check.setChecked(checked)
            else:
                parent._on_overlay_info_toggled(checked)
        elif action == canvas_actions.get("overlay_file"):
            checked = canvas_actions["overlay_file"].isChecked()
            if hasattr(parent, "overlay_file_check"):
                parent.overlay_file_check.setChecked(checked)
            else:
                parent._on_overlay_file_toggled(checked)
        elif action == canvas_actions.get("show_grid"):
            checked = canvas_actions["show_grid"].isChecked()
            if hasattr(parent, "show_grid_check"):
                parent.show_grid_check.setChecked(checked)
            else:
                view.set_show_grid(checked)
        elif action == canvas_actions.get("snap_grid"):
            checked = canvas_actions["snap_grid"].isChecked()
            if hasattr(parent, "snap_grid_check"):
                parent.snap_grid_check.setChecked(checked)
            else:
                view.set_snap_to_grid(checked)
        elif action == canvas_actions.get("canvas_color"):
            parent._on_canvas_color_clicked()
        elif action == canvas_actions.get("show_metadata_bar"):
            parent._on_metadata_bar_toggled(canvas_actions["show_metadata_bar"].isChecked())
        elif action == canvas_actions.get("show_unit_badge"):
            parent._on_metadata_unit_toggled(canvas_actions["show_unit_badge"].isChecked())
        elif action == canvas_actions.get("show_title"):
            parent._on_global_show_title_toggled(canvas_actions["show_title"].isChecked())
        elif action == canvas_actions.get("show_colorbar"):
            parent._on_global_show_colorbar_toggled(canvas_actions["show_colorbar"].isChecked())
        elif action == canvas_actions.get("show_colorbar_ticks"):
            parent._on_global_show_colorbar_ticks_toggled(canvas_actions["show_colorbar_ticks"].isChecked())
        elif action == canvas_actions.get("show_scale_bar"):
            parent._on_scale_bar_toggled(canvas_actions["show_scale_bar"].isChecked())
        elif action == canvas_actions.get("show_molecules"):
            parent._on_canvas_show_molecules_toggled(canvas_actions["show_molecules"].isChecked())
        elif action == canvas_actions.get("load_molecule"):
            parent._on_canvas_load_molecule()
        elif action == canvas_actions.get("clear_molecules"):
            parent._on_canvas_clear_molecules()
        elif action in cbar_position_actions:
            parent._on_colorbar_position_changed(cbar_position_actions[action])
        elif action == canvas_actions.get("layout_2x2"):
            parent._apply_layout("2x2")
        elif action == canvas_actions.get("layout_1x3"):
            parent._apply_layout("1x3")
        elif action == canvas_actions.get("layout_3x1"):
            parent._apply_layout("3x1")

    def set_title(self, title: str):
        self._title = title or ""
        self._update_rendered_pixmap()

    def set_colorbar_label(self, label: str):
        self._colorbar_label = label or ""
        self._update_rendered_pixmap()

    def set_cmap(self, cmap: str):
        self._cmap = cmap or self._cmap
        self._update_rendered_pixmap()

    def set_range(self, vmin: float | None, vmax: float | None):
        self._vmin = vmin
        self._vmax = vmax
        self._update_rendered_pixmap()

    def set_show_colorbar(self, show: bool):
        self._show_colorbar = show
        self._update_rendered_pixmap()

    def set_show_colorbar_ticks(self, show: bool):
        self._show_colorbar_ticks = bool(show)
        self._update_rendered_pixmap()

    def set_colorbar_mode(self, mode: str):
        normalized = mode.lower()
        mode_map = {
            "bottom": "bottom",
            "top": "top",
            "left": "left",
            "right": "right",
            "inset": "inset",
            "none": "none",
            "hidden": "none",
        }
        normalized = mode_map.get(normalized, "bottom")
        if normalized == self._colorbar_mode:
            return
        self._colorbar_mode = normalized
        self._update_rendered_pixmap()

    def set_show_title(self, show: bool):
        self._show_title = bool(show)
        self._update_rendered_pixmap()

    def set_scale_info(self, extent, axis_unit: str | None):
        self._extent = extent
        self._axis_unit = axis_unit or ""
        self._refresh_metadata_text()
        self._update_rendered_pixmap()

    def set_overlay_text(self, main_text: str, file_text: str | None = None):
        self._overlay_main_text = main_text or ""
        if file_text is not None:
            self._overlay_file_text = file_text or ""
        self._refresh_metadata_text()
        self._update_rendered_pixmap()

    def set_show_overlay(self, show_main: bool, show_file: bool | None = None):
        self._show_overlay_main = bool(show_main)
        if show_file is not None:
            self._show_overlay_file = bool(show_file)
        self._refresh_metadata_text()
        self._update_rendered_pixmap()

    def set_metadata_bar_visible(self, visible: bool):
        self._metadata_bar_visible = bool(visible)
        self._update_rendered_pixmap()

    def set_metadata_unit_visible(self, visible: bool):
        self._metadata_unit_visible = bool(visible)
        self._refresh_metadata_text()
        self._update_rendered_pixmap()

    def set_metadata_file_visible(self, visible: bool):
        self._metadata_file_visible = bool(visible)
        self._refresh_metadata_text()
        self._update_rendered_pixmap()

    def _refresh_metadata_text(self):
        right_parts = []
        if self._metadata_unit_visible and self._axis_unit:
            right_parts.append(self._axis_unit)
        if self._metadata_file_visible and self._overlay_file_text:
            right_parts.append(self._overlay_file_text)
        self._metadata_left_text = ""
        self._metadata_right_text = " | ".join(right_parts)

    def _text_scale_factor(self, img_rect: QtCore.QRectF) -> float:
        if self._locked_text_scale is not None:
            return self._locked_text_scale
        width = max(40.0, img_rect.width())
        ratio = width / max(self._base_image_width, 1.0)
        return max(0.6, min(2.4, ratio * 1.1))

    def _text_scale_for_width(self, width: float) -> float:
        ratio = width / max(self._base_image_width, 1.0)
        return max(0.6, min(2.4, ratio * 1.1))

    def set_locked_text_scale(self, scale: float | None):
        self._locked_text_scale = scale
        self.update()

    def to_state(self) -> dict:
        rect = self._rect
        return {
            "file_path": self._file_path,
            "channel_index": self._channel_index,
            "cmap": self._cmap,
            "title": self._title,
            "colorbar_label": self._colorbar_label,
            "vmin": self._vmin,
            "vmax": self._vmax,
            "pos": [self.pos().x(), self.pos().y()],
            "size": [rect.width(), rect.height()],
            "canvas_width": self._canvas_width,
            "show_colorbar": self._show_colorbar,
            "show_colorbar_ticks": self._show_colorbar_ticks,
            "show_title": self._show_title,
            "show_metadata_bar": self._metadata_bar_visible,
            "show_metadata_unit": self._metadata_unit_visible,
            "show_metadata_file": self._metadata_file_visible,
            "show_scale_bar": self._show_scale_bar,
            "colorbar_mode": self._colorbar_mode,
            "kind": self._kind,
            "text_scale": self._fixed_text_scale_value if self._use_fixed_text_scale else None,
            "show_molecules": self._show_molecules,
            "molecule_palette": self._molecule_palette,
            "molecule_state": self.export_molecule_state(),
            "show_hydrogens": self._show_hydrogens,
        }

    def apply_state(self, state: dict):
        self.set_title(state.get("title") or self._title)
        self.set_colorbar_label(state.get("colorbar_label") or self._colorbar_label)
        self.set_cmap(state.get("cmap") or self._cmap)
        vmin = state.get("vmin")
        vmax = state.get("vmax")
        self.set_range(vmin, vmax)
        self.set_show_colorbar(state.get("show_colorbar", True))
        self.set_show_colorbar_ticks(state.get("show_colorbar_ticks", True))
        self.set_show_title(state.get("show_title", self._show_title))
        self.set_metadata_bar_visible(state.get("show_metadata_bar", self._metadata_bar_visible))
        self.set_metadata_unit_visible(state.get("show_metadata_unit", self._metadata_unit_visible))
        self.set_metadata_file_visible(state.get("show_metadata_file", self._metadata_file_visible))
        self.set_show_scale_bar(state.get("show_scale_bar", self._show_scale_bar))
        self.set_show_molecules(state.get("show_molecules", self._show_molecules))
        self.set_molecule_palette(state.get("molecule_palette", self._molecule_palette))
        self.set_molecule_state(state.get("molecule_state") or [])
        self._show_hydrogens = bool(state.get("show_hydrogens", self._show_hydrogens))
        self.set_colorbar_mode(state.get("colorbar_mode", self._colorbar_mode))
        self._kind = state.get("kind", self._kind)
        canvas_width = state.get("canvas_width")
        if canvas_width is not None:
            self.set_canvas_width(float(canvas_width))
        else:
            size = state.get("size") or []
            if len(size) == 2:
                self.set_canvas_width(max(80.0, float(size[0])))
        pos = state.get("pos") or []
        if len(pos) == 2:
            self.setPos(float(pos[0]), float(pos[1]))
        ts = state.get("text_scale")
        if ts is not None:
            self._fixed_text_scale_value = max(0.01, min(2.4, float(ts)))
            self._use_fixed_text_scale = True

    @property
    def file_path(self) -> str:
        return self._file_path

    @property
    def channel_index(self) -> int:
        return self._channel_index

    @property
    def cmap(self) -> str:
        return self._cmap

    @property
    def title(self) -> str:
        return self._title

    @property
    def colorbar_label(self) -> str:
        return self._colorbar_label

    @property
    def vmin(self) -> float | None:
        return self._vmin

    @property
    def vmax(self) -> float | None:
        return self._vmax

    def image_size(self) -> tuple[float, float]:
        return self._canvas_width, self._canvas_height()

    def itemChange(self, change, value):
        if change == QtWidgets.QGraphicsItem.ItemSelectedChange:
            return value
        elif change == QtWidgets.QGraphicsItem.ItemPositionChange:
            if self._parent_window and getattr(self._parent_window, "_grid_locked", False):
                pass
            return value
        return super().itemChange(change, value)

    @property
    def kind(self) -> str | None:
        return self._kind

    def set_kind(self, kind: str | None):
        self._kind = kind

    def set_frame_color(self, color: QtGui.QColor | None):
        self._frame_color = color
        self._update_rendered_pixmap()

    def set_parent_window(self, window):
        self._parent_window = window

    def set_scale_bar_length(self, length: float | None):
        self._scale_bar_length = length
        self._update_rendered_pixmap()

    def set_show_scale_bar(self, show: bool):
        self._show_scale_bar = bool(show)
        self._update_rendered_pixmap()

    def export_molecule_state(self) -> list[dict]:
        return [dict(entry) for entry in (self._molecule_state or [])]

    def set_molecule_state(self, state):
        payload = []
        for entry in state or []:
            try:
                if isinstance(entry, Molecule):
                    payload.append(entry.to_dict())
                elif isinstance(entry, dict):
                    payload.append(dict(entry))
            except Exception:
                continue
        self._molecule_state = payload
        self._update_rendered_pixmap()

    def add_molecule_from_path(self, path) -> bool:
        try:
            mol = Molecule(path)
        except Exception:
            return False
        try:
            if self._extent and len(self._extent) == 4:
                x0, x1, y1, y0 = [float(v) for v in self._extent]
                mol.offset = np.array([(x0 + x1) / 2.0, (y0 + y1) / 2.0, 0.0], dtype=float)
        except Exception:
            pass
        payload = self.export_molecule_state()
        payload.append(mol.to_dict())
        self._molecule_state = payload
        self._update_rendered_pixmap()
        return True

    def clear_molecules(self):
        self._molecule_state = []
        self._update_rendered_pixmap()

    def set_show_molecules(self, show: bool):
        self._show_molecules = bool(show)
        self._update_rendered_pixmap()

    def set_molecule_palette(self, palette: str):
        self._molecule_palette = str(palette or "pymol").lower()
        self._update_rendered_pixmap()

    def _push_molecule_snapshot(self):
        snap = self.export_molecule_state()
        self._molecule_history.append(snap)
        if len(self._molecule_history) > 20:
            self._molecule_history = self._molecule_history[-20:]

    def undo_last_molecule_change(self) -> bool:
        if not self._molecule_history:
            return False
        self.set_molecule_state(self._molecule_history.pop())
        return True

    def _image_rect(self) -> QtCore.QRectF:
        width, height = self._tile_image_size()
        x = 0.0
        y = 0.0
        if self._show_colorbar and self._colorbar_mode == "left":
            x += self._colorbar_thickness() + self._colorbar_padding_x
        if self._show_colorbar and self._colorbar_mode == "top":
            y += self._colorbar_thickness() + self._colorbar_pad_y
        return QtCore.QRectF(x, y, width, height)

    def _normalized_extent(self):
        if not self._extent or len(self._extent) != 4:
            return None
        try:
            x0, x1, y1, y0 = [float(v) for v in self._extent]
        except Exception:
            return None
        return float(x0), float(x1), min(float(y0), float(y1)), max(float(y0), float(y1))

    def _local_to_data(self, pos: QtCore.QPointF):
        rect = self._image_rect()
        if not rect.contains(pos):
            return None
        extent = self._normalized_extent()
        if extent is None:
            return None
        x0, x1, ymin, ymax = extent
        if rect.width() <= 0 or rect.height() <= 0:
            return None
        fx = (pos.x() - rect.left()) / rect.width()
        fy = (pos.y() - rect.top()) / rect.height()
        x = x0 + fx * (x1 - x0)
        y = ymax - fy * (ymax - ymin)
        return float(x), float(y)

    def _data_to_local(self, x: float, y: float) -> QtCore.QPointF | None:
        rect = self._image_rect()
        extent = self._normalized_extent()
        if extent is None or rect.width() <= 0 or rect.height() <= 0:
            return None
        x0, x1, ymin, ymax = extent
        if abs(x1 - x0) < 1e-12 or abs(ymax - ymin) < 1e-12:
            return None
        fx = (float(x) - x0) / (x1 - x0)
        fy = (ymax - float(y)) / (ymax - ymin)
        return QtCore.QPointF(rect.left() + fx * rect.width(), rect.top() + fy * rect.height())

    def _molecule_objects(self):
        objs = []
        for entry in self._molecule_state or []:
            try:
                objs.append(Molecule.from_dict(entry))
            except Exception:
                continue
        return objs

    def _store_molecule_objects(self, molecules):
        self._molecule_state = [mol.to_dict() for mol in (molecules or [])]
        self._update_rendered_pixmap()

    def _molecule_hit(self, pos: QtCore.QPointF):
        if not self._show_molecules or not self._molecule_state:
            return None
        best = None
        best_d2 = None
        molecules = self._molecule_objects()
        for mol_idx, mol in reversed(list(enumerate(molecules))):
            try:
                coords = mol.get_transformed_coordinates()
            except Exception:
                continue
            if coords is None or len(coords) == 0:
                continue
            for atom_idx, coord in enumerate(coords):
                try:
                    if not self._show_hydrogens and atom_idx < len(mol.elements) and str(mol.elements[atom_idx]).strip().upper() == "H":
                        continue
                    local = self._data_to_local(coord[0], coord[1])
                    if local is None:
                        continue
                    dx = local.x() - pos.x()
                    dy = local.y() - pos.y()
                    d2 = dx * dx + dy * dy
                    if best_d2 is None or d2 < best_d2:
                        best_d2 = d2
                        best = mol_idx
                except Exception:
                    continue
        if best is not None and best_d2 is not None and best_d2 <= 18.0 * 18.0:
            return best
        return None

    def _show_molecule_menu(self, screen_pos, mol_idx: int):
        molecules = self._molecule_objects()
        if not (0 <= mol_idx < len(molecules)):
            return
        mol = molecules[mol_idx]
        menu = QtWidgets.QMenu()
        props_act = menu.addAction("Edit molecule...")
        show_h_act = menu.addAction("Show hydrogens")
        show_h_act.setCheckable(True)
        show_h_act.setChecked(self._show_hydrogens)
        palette_menu = menu.addMenu("Atom palette")
        palette_actions = {}
        current_pal = (self._molecule_palette or "pymol").lower()
        for pal in available_atom_palettes():
            act = palette_menu.addAction(pal.title())
            act.setCheckable(True)
            act.setChecked(pal == current_pal)
            palette_actions[act] = pal
        menu.addSeparator()
        dup_act = menu.addAction("Duplicate")
        del_act = menu.addAction("Delete")
        clear_act = menu.addAction("Clear all molecules")
        undo_act = menu.addAction("Undo last molecule change")
        action = menu.exec_(screen_pos)
        if action is None:
            return
        if action == props_act:
            self._open_molecule_properties(mol_idx)
            return
        if action == show_h_act:
            self._show_hydrogens = show_h_act.isChecked()
            self._update_rendered_pixmap()
            if self._parent_window is not None:
                self._parent_window._push_undo_state()
            return
        if action in palette_actions:
            self._molecule_palette = palette_actions[action]
            self._update_rendered_pixmap()
            if self._parent_window is not None:
                self._parent_window._persist_item_molecules(self)
                self._parent_window._push_undo_state()
            return
        if action == dup_act:
            self._push_molecule_snapshot()
            clone = mol.copy()
            clone.offset = clone.offset + np.array([0.6, 0.6, 0.0], dtype=float)
            molecules.append(clone)
            self._store_molecule_objects(molecules)
        elif action == del_act:
            self._push_molecule_snapshot()
            del molecules[mol_idx]
            self._store_molecule_objects(molecules)
        elif action == clear_act:
            self._push_molecule_snapshot()
            self.clear_molecules()
        elif action == undo_act:
            self.undo_last_molecule_change()
        else:
            return
        if self._parent_window is not None:
            self._parent_window._persist_item_molecules(self)
            self._parent_window._push_undo_state()

    def _open_molecule_properties(self, mol_idx: int):
        molecules = self._molecule_objects()
        if not (0 <= mol_idx < len(molecules)):
            return
        mol = molecules[mol_idx]
        original = mol.to_dict()
        overlay_settings = {
            "palette": self._molecule_palette,
            "show_hydrogens": bool(self._show_hydrogens),
            "show_shadows_available": False,
            "show_hydrogens_available": True,
            "palette_available": True,
        }
        def _apply():
            updated = self._molecule_objects()
            if 0 <= mol_idx < len(updated):
                updated[mol_idx] = mol
                self._store_molecule_objects(updated)
                self._molecule_palette = str(overlay_settings.get("palette", self._molecule_palette or "pymol")).lower()
                self._show_hydrogens = bool(overlay_settings.get("show_hydrogens", self._show_hydrogens))
                if self._parent_window is not None:
                    self._parent_window._persist_item_molecules(self)
        dlg = MoleculePropertiesDialog(mol, parent=self._parent_window, callback=_apply, overlay_settings=overlay_settings)
        self._molecule_props_dialog = dlg
        dlg.finished.connect(lambda _res: self._finalize_molecule_dialog_change(original, mol_idx, mol))
        dlg.show()

    def _finalize_molecule_dialog_change(self, original_state: dict, mol_idx: int, edited_mol: Molecule):
        try:
            new_state = edited_mol.to_dict()
        except Exception:
            return
        if new_state == original_state:
            return
        self._push_molecule_snapshot()
        molecules = self._molecule_objects()
        if 0 <= mol_idx < len(molecules):
            molecules[mol_idx] = edited_mol
            self._store_molecule_objects(molecules)
            if self._parent_window is not None:
                self._parent_window._persist_item_molecules(self)
                self._parent_window._push_undo_state()

    def set_text_color_override(self, color: QtGui.QColor | None):
        self._text_color_override = color
        self._update_rendered_pixmap()

    def _copy_selected_svg(self):
        items = []
        try:
            if self.scene() is not None:
                items = [i for i in self.scene().selectedItems() if isinstance(i, CanvasImageItem)]
        except Exception:
            items = []
        if not items:
            return
        if len(items) == 1:
            items[0]._copy_svg_to_clipboard()
            return
        view = self._first_canvas_view()
        if view is None:
            return
        svg_bytes = view._compose_svg_bytes(items)
        if svg_bytes:
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", svg_bytes)
            QtWidgets.QApplication.clipboard().setMimeData(mime)

    def _first_canvas_view(self):
        try:
            if self.scene() is None:
                return None
            views = self.scene().views()
            if not views:
                return None
            return views[0]
        except Exception:
            return None

    def _render_vector_figure(self):
        if self._arr is None:
            return None
        width = max(2, int(round(self._tile_total_width())))
        height = max(2, int(round(self._tile_total_height())))
        metadata_height = self._metadata_bar_height() if self._metadata_bar_visible and (self._metadata_left_text or self._metadata_right_text) else 0.0
        text_scale = self._effective_text_scale()
        frame_color = self._frame_color.name() if isinstance(self._frame_color, QtGui.QColor) else "#070707"
        if self._text_color_override is not None and self._text_color_override.isValid():
            text_color = self._text_color_override.name()
        else:
            text_color = _text_color_for_frame(frame_color)
        show_overlay_main = self._show_overlay_main and not self._metadata_bar_visible
        show_overlay_file = self._show_overlay_file and not self._metadata_bar_visible
        scale_spec = self._scale_bar_spec()
        scale_length = scale_spec[0] if scale_spec else None
        scale_width = scale_spec[1] if scale_spec else None
        return render_tile_figure_mpl(
            self._arr,
            cmap=self._cmap,
            vmin=self._vmin,
            vmax=self._vmax,
            title=self._title,
            colorbar_label=self._colorbar_label,
            width_px=width,
            height_px=height,
            dpi=self._full_dpi,
            show_colorbar=self._show_colorbar,
            show_colorbar_ticks=self._show_colorbar_ticks,
            show_title=self._show_title,
            show_metadata=self._metadata_bar_visible and bool(self._metadata_left_text or self._metadata_right_text),
            metadata_left=self._metadata_left_text,
            metadata_right=self._metadata_right_text,
            show_overlay_main=show_overlay_main,
            overlay_main=self._overlay_main_text,
            show_overlay_file=show_overlay_file,
            overlay_file=self._overlay_file_text,
            cbar_position=self._colorbar_mode,
            metadata_height=metadata_height,
            frame_color=frame_color,
            text_scale=text_scale,
            text_color=text_color,
            show_scale_bar=self._show_scale_bar,
            scale_bar_length=scale_length,
            scale_bar_unit=self._axis_unit,
            scale_bar_width=scale_width,
            extent=self._extent,
            show_molecules=self._show_molecules,
            molecules=self._molecule_state,
            molecule_palette=self._molecule_palette,
            show_hydrogens=self._show_hydrogens,
        )

    def _copy_svg_to_clipboard(self):
        try:
            fig = self._render_vector_figure()
            if fig is None:
                return
            buf = io.BytesIO()
            with matplotlib.rc_context({'svg.fonttype': 'none'}):
                fig.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.02)
            svg_bytes = buf.getvalue()
            mime = QtCore.QMimeData()
            mime.setData("image/svg+xml", svg_bytes)
            QtWidgets.QApplication.clipboard().setMimeData(mime)
        except Exception:
            pass

    def _save_vector_to_file(self, fmt: str):
        fmt = (fmt or "").strip().lower()
        if fmt not in ("svg", "pdf"):
            return
        try:
            title = self._title or "view"
            default = f"{title}.{fmt}"
            label = "SVG Files (*.svg)" if fmt == "svg" else "PDF Files (*.pdf)"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(None, "Save view", default, label)
            if not path:
                return
            if not path.lower().endswith(f".{fmt}"):
                path = f"{path}.{fmt}"
            fig = self._render_vector_figure()
            if fig is None:
                return
            if fmt == 'svg':
                with matplotlib.rc_context({'svg.fonttype': 'none'}):
                    fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            else:
                fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            try:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))
            except Exception:
                pass
        except Exception:
            QtWidgets.QMessageBox.warning(None, "Save view", "Unable to save vector image.")

    @property
    def data_array(self) -> np.ndarray:
        return self._arr
