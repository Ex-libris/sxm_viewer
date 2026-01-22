"""Canvas items and context menus for the scientific canvas."""
from __future__ import annotations

import io

from ..._shared import QtCore, QtGui, QtWidgets, np, matplotlib
from .canvas_rendering import render_tile_mpl, render_tile_figure_mpl, _text_color_for_frame


def _append_canvas_menu_actions(menu: QtWidgets.QMenu, parent, view):
    actions = {}
    if parent is None or view is None:
        return actions

    actions["align_selected"] = menu.addAction("Align selected")
    actions["align_by_channel"] = menu.addAction("Align by channel")
    actions["reset_alignment"] = menu.addAction("Reset alignment")
    menu.addSeparator()

    actions["sync_ranges"] = menu.addAction("Sync ranges")
    actions["sync_ranges"].setCheckable(True)
    actions["sync_ranges"].setChecked(bool(getattr(parent, "_sync_colorbars", False)))

    actions["sync_colors_by_channel"] = menu.addAction("Sync colors by channel")
    actions["sync_colors_by_channel"].setCheckable(True)
    actions["sync_colors_by_channel"].setChecked(bool(getattr(parent, "_sync_by_channel", False)))

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
        self.setAcceptedMouseButtons(QtCore.Qt.LeftButton | QtCore.Qt.RightButton)
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
        self._full_dpi = 200
        self._fast_dpi = 96
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
        self._metadata_file_visible = False
        self._metadata_left_text = ""
        self._metadata_right_text = ""
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

    def hoverEnterEvent(self, event):
        if self._resize_handle_rect().contains(event.pos()):
            self.setCursor(QtCore.Qt.SizeFDiagCursor)
        else:
            self.setCursor(QtCore.Qt.OpenHandCursor)

    def hoverLeaveEvent(self, event):
        self.setCursor(QtCore.Qt.ArrowCursor)

    def hoverMoveEvent(self, event):
        if self._resize_handle_rect().contains(event.pos()):
            self.setCursor(QtCore.Qt.SizeFDiagCursor)
        else:
            self.setCursor(QtCore.Qt.OpenHandCursor)
        super().hoverMoveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.RightButton:
            if not self.isSelected():
                self.setSelected(True)
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
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        # Alt+drag to duplicate
        if event.modifiers() & QtCore.Qt.AltModifier and not hasattr(self, '_alt_duplicated'):
            if self._parent_window:
                self._parent_window._on_duplicate_item()
                self._alt_duplicated = True
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
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if hasattr(self, '_alt_duplicated'):
            delattr(self, '_alt_duplicated')
        
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
                self._parent_window._push_undo_state()
            return
        super().mouseReleaseEvent(event)
        if self._parent_window is not None:
            self._parent_window._push_undo_state()

    def contextMenuEvent(self, event):
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
        if action == canvas_actions.get("align_selected"):
            parent._on_align_selected()
        elif action == canvas_actions.get("align_by_channel"):
            parent._on_align_by_channels()
        elif action == canvas_actions.get("reset_alignment"):
            parent._reset_locked_alignment()
        elif action == canvas_actions.get("sync_ranges"):
            checked = canvas_actions["sync_ranges"].isChecked()
            if hasattr(parent, "sync_cbar_check"):
                parent.sync_cbar_check.setChecked(checked)
            else:
                parent._on_sync_colorbars_toggled(checked)
        elif action == canvas_actions.get("sync_colors_by_channel"):
            checked = canvas_actions["sync_colors_by_channel"].isChecked()
            if hasattr(parent, "sync_by_channel_check"):
                parent.sync_by_channel_check.setChecked(checked)
            else:
                parent._on_sync_by_channel_toggled(checked)
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

    def set_metadata_file_visible(self, visible: bool):
        self._metadata_file_visible = bool(visible)
        self._refresh_metadata_text()
        self._update_rendered_pixmap()

    def _refresh_metadata_text(self):
        right_parts = []
        if self._axis_unit:
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
            "kind": self._kind,
            "text_scale": self._fixed_text_scale_value if self._use_fixed_text_scale else None,
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
