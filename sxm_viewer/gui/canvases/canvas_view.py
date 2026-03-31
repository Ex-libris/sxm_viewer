"""Canvas view, drag/drop, and export helpers."""
from __future__ import annotations

import io
import json
import re
from pathlib import Path

import matplotlib
from mpl_toolkits.axes_grid1 import make_axes_locatable
from ..._shared import QtCore, QtGui, QtWidgets, np
from ..thumbnail_render import array_to_qimage
from .canvas_items import CanvasImageItem, AlignmentGuide, RubberBandSelection, _append_canvas_menu_actions

_CANVAS_MIME = "application/x-sxm-view"


class CanvasGraphicsView(QtWidgets.QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragMode(QtWidgets.QGraphicsView.NoDrag)
        self.setTransformationAnchor(QtWidgets.QGraphicsView.AnchorUnderMouse)
        self.setRenderHints(QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
        self._grid_size = 20
        self._show_grid = False
        self._snap_to_grid = False
        self._panning = False
        self._last_pan_pos = None
        self._rubber_band = None
        self._rubber_band_origin = None
        self._alignment_guides = []
        self._snap_distance = 8
        self._show_alignment_guides = True
        self.set_background_color(QtGui.QColor(30, 30, 30))

    def set_grid_size(self, size: int):
        self._grid_size = max(10, size)
        self.viewport().update()

    def set_show_grid(self, show: bool):
        self._show_grid = show
        self.viewport().update()

    def set_snap_to_grid(self, snap: bool):
        self._snap_to_grid = snap

    def set_background_color(self, color: QtGui.QColor):
        self.setBackgroundBrush(QtGui.QBrush(color))

    def clear_alignment_guides(self):
        for guide in self._alignment_guides:
            if guide.scene():
                self.scene().removeItem(guide)
        self._alignment_guides.clear()

    def show_alignment_guides(self, item, items):
        if not self._show_alignment_guides:
            return

        self.clear_alignment_guides()
        item_rect = item.sceneBoundingRect()
        threshold = self._snap_distance

        for other in items:
            if other == item or not isinstance(other, CanvasImageItem):
                continue

            other_rect = other.sceneBoundingRect()

            if abs(item_rect.left() - other_rect.left()) < threshold:
                x = other_rect.left()
                y1 = min(item_rect.top(), other_rect.top())
                y2 = max(item_rect.bottom(), other_rect.bottom())
                guide = AlignmentGuide(x, y1, x, y2)
                self.scene().addItem(guide)
                self._alignment_guides.append(guide)

            if abs(item_rect.right() - other_rect.right()) < threshold:
                x = other_rect.right()
                y1 = min(item_rect.top(), other_rect.top())
                y2 = max(item_rect.bottom(), other_rect.bottom())
                guide = AlignmentGuide(x, y1, x, y2)
                self.scene().addItem(guide)
                self._alignment_guides.append(guide)

            center_x = item_rect.center().x()
            other_center_x = other_rect.center().x()
            if abs(center_x - other_center_x) < threshold:
                x = other_center_x
                y1 = min(item_rect.top(), other_rect.top())
                y2 = max(item_rect.bottom(), other_rect.bottom())
                guide = AlignmentGuide(x, y1, x, y2)
                self.scene().addItem(guide)
                self._alignment_guides.append(guide)

            if abs(item_rect.top() - other_rect.top()) < threshold:
                y = other_rect.top()
                x1 = min(item_rect.left(), other_rect.left())
                x2 = max(item_rect.right(), other_rect.right())
                guide = AlignmentGuide(x1, y, x2, y)
                self.scene().addItem(guide)
                self._alignment_guides.append(guide)

            if abs(item_rect.bottom() - other_rect.bottom()) < threshold:
                y = other_rect.bottom()
                x1 = min(item_rect.left(), other_rect.left())
                x2 = max(item_rect.right(), other_rect.right())
                guide = AlignmentGuide(x1, y, x2, y)
                self.scene().addItem(guide)
                self._alignment_guides.append(guide)

            center_y = item_rect.center().y()
            other_center_y = other_rect.center().y()
            if abs(center_y - other_center_y) < threshold:
                y = other_center_y
                x1 = min(item_rect.left(), other_rect.left())
                x2 = max(item_rect.right(), other_rect.right())
                guide = AlignmentGuide(x1, y, x2, y)
                self.scene().addItem(guide)
                self._alignment_guides.append(guide)

    def snap_to_items(self, item, items):
        item_rect = item.sceneBoundingRect()
        threshold = self._snap_distance
        snap_x = None
        snap_y = None

        for other in items:
            if other == item or not isinstance(other, CanvasImageItem):
                continue

            other_rect = other.sceneBoundingRect()

            if abs(item_rect.left() - other_rect.left()) < threshold:
                snap_x = other_rect.left()
            elif abs(item_rect.right() - other_rect.right()) < threshold:
                snap_x = other_rect.right() - item_rect.width()
            elif abs(item_rect.center().x() - other_rect.center().x()) < threshold:
                snap_x = other_rect.center().x() - item_rect.width() / 2

            if abs(item_rect.top() - other_rect.top()) < threshold:
                snap_y = other_rect.top()
            elif abs(item_rect.bottom() - other_rect.bottom()) < threshold:
                snap_y = other_rect.bottom() - item_rect.height()
            elif abs(item_rect.center().y() - other_rect.center().y()) < threshold:
                snap_y = other_rect.center().y() - item_rect.height() / 2

        if snap_x is not None or snap_y is not None:
            pos = item.pos()
            if snap_x is not None:
                pos.setX(snap_x)
            if snap_y is not None:
                pos.setY(snap_y)
            item.setPos(pos)

    def wheelEvent(self, event):
        if event.modifiers() & QtCore.Qt.ControlModifier:
            delta = event.angleDelta().y()
            if delta:
                factor = 1.15 if delta > 0 else 1 / 1.15
                self.scale(factor, factor)
            event.accept()
            return
        super().wheelEvent(event)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            scene_pos = self.mapToScene(event.pos())
            item = self.scene().itemAt(scene_pos, QtGui.QTransform())
            if item is None:
                if not (event.modifiers() & QtCore.Qt.ControlModifier):
                    self.scene().clearSelection()
                self._rubber_band_origin = scene_pos
                self._rubber_band = RubberBandSelection()
                self._rubber_band.setRect(QtCore.QRectF(self._rubber_band_origin, QtCore.QSizeF(0, 0)))
                self.scene().addItem(self._rubber_band)
                event.accept()
                return
            elif not isinstance(item, CanvasImageItem):
                self._panning = True
                self._last_pan_pos = event.pos()
                self.setCursor(QtCore.Qt.ClosedHandCursor)
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._rubber_band is not None and self._rubber_band_origin is not None:
            current_pos = self.mapToScene(event.pos())
            rect = QtCore.QRectF(self._rubber_band_origin, current_pos).normalized()
            self._rubber_band.setRect(rect)
            path = QtGui.QPainterPath()
            path.addRect(rect)
            self.scene().setSelectionArea(path)
            event.accept()
            return

        if self._panning and self._last_pan_pos is not None:
            delta = event.pos() - self._last_pan_pos
            self._last_pan_pos = event.pos()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - delta.y())
            event.accept()
            return

        if event.buttons() & QtCore.Qt.LeftButton:
            item = self.itemAt(event.pos())
            if item and isinstance(item, CanvasImageItem):
                items = [i for i in self.scene().items() if isinstance(i, CanvasImageItem)]
                self.show_alignment_guides(item, items)

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._rubber_band is not None:
            if self._rubber_band.scene():
                self.scene().removeItem(self._rubber_band)
            self._rubber_band = None
            self._rubber_band_origin = None
            event.accept()
            return

        if event.button() == QtCore.Qt.LeftButton:
            if self._panning:
                self._panning = False
                self._last_pan_pos = None
                self.setCursor(QtCore.Qt.ArrowCursor)
                event.accept()
                return

            for item in self.scene().selectedItems():
                if isinstance(item, CanvasImageItem):
                    items = [i for i in self.scene().items() if isinstance(i, CanvasImageItem)]
                    self.snap_to_items(item, items)
            self.clear_alignment_guides()

        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        parent = self.parent()
        while parent is not None and not hasattr(parent, "_handle_canvas_key"):
            parent = parent.parent()
        if parent is not None:
            if parent._handle_canvas_key(event):
                return
        key = event.key()
        mods = event.modifiers()

        if key in (QtCore.Qt.Key_Left, QtCore.Qt.Key_Right, QtCore.Qt.Key_Up, QtCore.Qt.Key_Down):
            distance = 10 if mods & QtCore.Qt.ShiftModifier else 1
            dx = dy = 0
            if key == QtCore.Qt.Key_Left:
                dx = -distance
            elif key == QtCore.Qt.Key_Right:
                dx = distance
            elif key == QtCore.Qt.Key_Up:
                dy = -distance
            elif key == QtCore.Qt.Key_Down:
                dy = distance
            for item in self.scene().selectedItems():
                if isinstance(item, CanvasImageItem):
                    item.moveBy(dx, dy)
            event.accept()
            return

        if mods & QtCore.Qt.ControlModifier and key == QtCore.Qt.Key_A:
            for item in self.scene().items():
                if isinstance(item, CanvasImageItem):
                    item.setSelected(True)
            event.accept()
            return

        if key == QtCore.Qt.Key_Escape:
            self.scene().clearSelection()
            event.accept()
            return

        if key in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
            parent = self.parent()
            while parent is not None and not hasattr(parent, "_handle_canvas_key"):
                parent = parent.parent()
            if parent is not None:
                if parent._handle_canvas_key(event):
                    return

        super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        item = self.itemAt(event.pos())
        if item is not None and isinstance(item, CanvasImageItem):
            super().contextMenuEvent(event)
            return

        menu = QtWidgets.QMenu()
        select_all = menu.addAction("Select All")
        deselect_all = menu.addAction("Deselect All")
        menu.addSeparator()
        copy_svg = menu.addAction("Copy selected as SVG (vector)")
        save_svg = menu.addAction("Save selected as SVG...")
        save_pdf = menu.addAction("Save selected as PDF...")
        menu.addSeparator()
        zoom_in = menu.addAction("Zoom In")
        zoom_out = menu.addAction("Zoom Out")
        zoom_reset = menu.addAction("Reset Zoom")
        menu.addSeparator()
        fit_view = menu.addAction("Fit All in View")

        parent = self.parent()
        while parent is not None and not hasattr(parent, "_on_align_selected"):
            parent = parent.parent()
        canvas_actions = {}
        if parent is not None:
            menu.addSeparator()
            canvas_actions = _append_canvas_menu_actions(menu, parent, self)

        selected_items = [i for i in self.scene().selectedItems() if isinstance(i, CanvasImageItem)]
        has_selection = bool(selected_items)
        copy_svg.setEnabled(has_selection)
        save_svg.setEnabled(has_selection)
        save_pdf.setEnabled(has_selection)
        action = menu.exec_(event.globalPos())

        if action == select_all:
            for item in self.scene().items():
                if isinstance(item, CanvasImageItem):
                    item.setSelected(True)
        elif action == deselect_all:
            self.scene().clearSelection()
        elif action in (copy_svg, save_svg, save_pdf):
            self._export_selected_vectors(action, selected_items, copy_svg, save_svg, save_pdf)
        elif action == zoom_in:
            self.scale(1.15, 1.15)
        elif action == zoom_out:
            self.scale(1/1.15, 1/1.15)
        elif action == zoom_reset:
            self.resetTransform()
        elif action == fit_view:
            self.fitInView(self.scene().itemsBoundingRect(), QtCore.Qt.KeepAspectRatio)
        elif parent is not None:
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
                if hasattr(parent, "sync_cbar_check"):
                    parent.sync_cbar_check.setChecked(canvas_actions["sync_ranges"].isChecked())
                else:
                    parent._on_sync_colorbars_toggled(canvas_actions["sync_ranges"].isChecked())
            elif action == canvas_actions.get("copy_cmap"):
                parent._on_copy_cmap()
            elif action == canvas_actions.get("sync_colors_by_channel"):
                if hasattr(parent, "sync_by_channel_check"):
                    parent.sync_by_channel_check.setChecked(canvas_actions["sync_colors_by_channel"].isChecked())
                else:
                    parent._on_sync_by_channel_toggled(canvas_actions["sync_colors_by_channel"].isChecked())
            elif action in cmap_actions:
                parent._on_apply_cmap_to_selected(cmap_actions[action])
            elif action == canvas_actions.get("overlay_info"):
                if hasattr(parent, "overlay_info_check"):
                    parent.overlay_info_check.setChecked(canvas_actions["overlay_info"].isChecked())
                else:
                    parent._on_overlay_info_toggled(canvas_actions["overlay_info"].isChecked())
            elif action == canvas_actions.get("overlay_file"):
                if hasattr(parent, "overlay_file_check"):
                    parent.overlay_file_check.setChecked(canvas_actions["overlay_file"].isChecked())
                else:
                    parent._on_overlay_file_toggled(canvas_actions["overlay_file"].isChecked())
            elif action == canvas_actions.get("show_grid"):
                if hasattr(parent, "show_grid_check"):
                    parent.show_grid_check.setChecked(canvas_actions["show_grid"].isChecked())
                else:
                    self.set_show_grid(canvas_actions["show_grid"].isChecked())
            elif action == canvas_actions.get("snap_grid"):
                if hasattr(parent, "snap_grid_check"):
                    parent.snap_grid_check.setChecked(canvas_actions["snap_grid"].isChecked())
                else:
                    self.set_snap_to_grid(canvas_actions["snap_grid"].isChecked())
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
        elif canvas_actions:
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
                    self.set_show_grid(checked)
            elif action == canvas_actions.get("snap_grid"):
                checked = canvas_actions["snap_grid"].isChecked()
                if hasattr(parent, "snap_grid_check"):
                    parent.snap_grid_check.setChecked(checked)
                else:
                    self.set_snap_to_grid(checked)
            elif action == canvas_actions.get("canvas_color"):
                parent._on_canvas_color_clicked()
            elif action == canvas_actions.get("show_molecules"):
                checked = canvas_actions["show_molecules"].isChecked()
                parent._on_canvas_show_molecules_toggled(checked)
            elif action == canvas_actions.get("load_molecule"):
                parent._on_canvas_load_molecule()
            elif action == canvas_actions.get("clear_molecules"):
                parent._on_canvas_clear_molecules()
            elif action == canvas_actions.get("layout_2x2"):
                parent._apply_layout("2x2")
            elif action == canvas_actions.get("layout_1x3"):
                parent._apply_layout("1x3")
            elif action == canvas_actions.get("layout_3x1"):
                parent._apply_layout("3x1")
        event.accept()

    def _export_selected_vectors(self, action, items, copy_svg, save_svg, save_pdf):
        if not items:
            return
        if action == copy_svg:
            if len(items) == 1:
                items[0]._copy_svg_to_clipboard()
                return
            svg_bytes = self._compose_svg_bytes(items)
            if svg_bytes:
                mime = QtCore.QMimeData()
                mime.setData("image/svg+xml", svg_bytes)
                QtWidgets.QApplication.clipboard().setMimeData(mime)
            return
        fmt = "svg" if action == save_svg else "pdf"
        if len(items) == 1:
            items[0]._save_vector_to_file(fmt)
            return
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select export folder")
        if not folder:
            return
        for idx, item in enumerate(items, 1):
            title = item._title or "view"
            safe = "".join(c for c in title if c.isalnum() or c in (" ", "_", "-")).strip()
            if not safe:
                safe = f"view_{idx}"
            filename = f"{safe}_{idx}.{fmt}"
            path = str(Path(folder) / filename)
            try:
                fig = item._render_vector_figure()
                if fig is None:
                    continue
                if fmt == 'svg':
                    with matplotlib.rc_context({'svg.fonttype': 'none'}):
                        fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
                else:
                    fig.savefig(path, format=fmt, bbox_inches="tight", pad_inches=0.02)
            except Exception:
                continue

    def _compose_svg_bytes(self, items):
        if not items:
            return None
        rect = None
        for item in items:
            try:
                item_rect = item.sceneBoundingRect()
            except Exception:
                continue
            rect = item_rect if rect is None else rect.united(item_rect)
        if rect is None:
            return None
        width = float(rect.width())
        height = float(rect.height())
        if width <= 1 or height <= 1:
            return None
        svg_parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width:.2f}" height="{height:.2f}" viewBox="0 0 {width:.2f} {height:.2f}">']
        for idx, item in enumerate(sorted(items, key=lambda i: i.zValue())):
            try:
                fig = item._render_vector_figure()
                if fig is None:
                    continue
                buf = io.BytesIO()
                with matplotlib.rc_context({'svg.fonttype': 'none'}):
                    fig.savefig(buf, format="svg")
                svg = buf.getvalue().decode("utf-8", errors="ignore")
                svg = re.sub(r"<\?xml[^>]*>\s*", "", svg, flags=re.IGNORECASE)
                svg = re.sub(r"<!DOCTYPE[^>]*>\s*", "", svg, flags=re.IGNORECASE)
                match = re.search(r"<svg[^>]*>", svg, flags=re.IGNORECASE)
                if not match:
                    continue
                open_tag = match.group(0)
                inner = svg[match.end():]
                inner = re.sub(r"</svg>\s*$", "", inner, flags=re.IGNORECASE)
                view_box = None
                vb_match = re.search(r'viewBox="([^"]+)"', open_tag)
                if vb_match:
                    view_box = vb_match.group(1)
                prefix = f"item{idx}_"
                inner = re.sub(r'id="([^"]+)"', lambda m: f'id="{prefix}{m.group(1)}"', inner)
                inner = re.sub(r'url\(#([^\)]+)\)', lambda m: f'url(#{prefix}{m.group(1)})', inner)
                inner = re.sub(r'href="#([^"]+)"', lambda m: f'href="#{prefix}{m.group(1)}"', inner)
                inner = re.sub(r'xlink:href="#([^"]+)"', lambda m: f'xlink:href="#{prefix}{m.group(1)}"', inner)
                item_rect = item.sceneBoundingRect()
                x = float(item_rect.left() - rect.left())
                y = float(item_rect.top() - rect.top())
                w = float(item_rect.width())
                h = float(item_rect.height())
                vb_attr = f' viewBox="{view_box}"' if view_box else ""
                svg_parts.append(
                    f'<svg x="{x:.2f}" y="{y:.2f}" width="{w:.2f}" height="{h:.2f}"{vb_attr}>'
                    f"{inner}</svg>"
                )
            except Exception:
                continue
        svg_parts.append("</svg>")
        return "".join(svg_parts).encode("utf-8")

    def drawBackground(self, painter, rect):
        super().drawBackground(painter, rect)
        if not self._show_grid:
            return
        painter.setPen(QtGui.QPen(QtGui.QColor(50, 50, 50), 0))
        left = int(rect.left()) - (int(rect.left()) % self._grid_size)
        top = int(rect.top()) - (int(rect.top()) % self._grid_size)
        for x in range(left, int(rect.right()), self._grid_size):
            painter.drawLine(x, int(rect.top()), x, int(rect.bottom()))
        for y in range(top, int(rect.bottom()), self._grid_size):
            painter.drawLine(int(rect.left()), y, int(rect.right()), y)

    def dragEnterEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls() or mime.hasFormat(_CANVAS_MIME):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        mime = event.mimeData()
        if mime.hasUrls() or mime.hasFormat(_CANVAS_MIME):
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        parent = self.parent()
        while parent is not None and not hasattr(parent, "handle_drop"):
            parent = parent.parent()
        if parent is None:
            return super().dropEvent(event)
        mime = event.mimeData()
        payloads = []
        if mime.hasFormat(_CANVAS_MIME):
            try:
                data = bytes(mime.data(_CANVAS_MIME)).decode("utf-8")
                payload = json.loads(data)
                if isinstance(payload, dict) and payload.get("items"):
                    for item in payload.get("items") or []:
                        if item:
                            payloads.append({
                                "file_path": item,
                                "cmap": payload.get("cmap"),
                                "channel_index": payload.get("channel_index"),
                            })
                else:
                    payloads.append(payload)
            except Exception:
                payloads = []
        paths = []
        if mime.hasUrls():
            for url in mime.urls():
                if url.isLocalFile():
                    paths.append(url.toLocalFile())
        if payloads or paths:
            try:
                parent.handle_drop(payloads, paths)
            except Exception as exc:
                try:
                    QtWidgets.QMessageBox.critical(self, "Canvas drop", f"Unable to add selection: {exc}")
                except Exception:
                    pass
            event.acceptProposedAction()
            return
        super().dropEvent(event)
