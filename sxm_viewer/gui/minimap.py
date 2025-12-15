"""Mini-map widget showing SXM frames relative positions."""
from __future__ import annotations

from .._shared import *


class FrameMiniMap(QtWidgets.QWidget):
    entryClicked = QtCore.pyqtSignal(object)
    entryShiftClicked = QtCore.pyqtSignal(object)
    zoomChanged = QtCore.pyqtSignal(float)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(240)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.entries = []
        self.active_key = None
        self.range_nm = 2000.0  # total width of map (2 um)
        self.zoom_factor = 1.0
        self._min_zoom = 0.01
        self._max_zoom = 10000.0
        self._poly_map = []
        self._hidden_keys = set()
        self._entry_pixmaps = {}
        self.show_real_images = False
        self._hover_key = None
        self._pan_center_nm = QtCore.QPointF(0.0, 0.0)
        self._panning = False
        self._last_drag_pos = None
        self._current_scale = 1.0

    def set_entries(self, entries):
        self.entries = entries or []
        self._poly_map = []
        self.update()

    def set_hidden_entries(self, keys):
        self._hidden_keys = set(keys or [])
        self.update()

    def hide_entry(self, key):
        if key is None:
            return
        self._hidden_keys.add(str(key))
        self.update()

    def clear_hidden_entries(self):
        if not self._hidden_keys:
            return
        self._hidden_keys.clear()
        self.update()

    def set_active_key(self, key):
        if self.active_key == key:
            return
        self.active_key = key
        self.update()

    def set_real_view_enabled(self, enabled: bool):
        enabled = bool(enabled)
        if self.show_real_images == enabled:
            return
        self.show_real_images = enabled
        self.update()

    def set_entry_pixmaps(self, mapping):
        self._entry_pixmaps = dict(mapping or {})
        if self.show_real_images:
            self.update()

    def _entry_area(self, entry):
        try:
            return abs(float(entry.get('x_range_nm', 0.0)) * float(entry.get('y_range_nm', 0.0)))
        except Exception:
            return float('inf')

    def _entry_color(self, entry, active):
        tag = entry.get('tag')
        if tag == 'constant-height':
            base = QtGui.QColor(76, 214, 136)
        elif tag == 'constant-current':
            base = QtGui.QColor(92, 148, 255)
        else:
            base = QtGui.QColor(200, 200, 220)
        color = QtGui.QColor(base)
        if not active:
            color.setAlpha(220)
        return color

    def _view_rect(self):
        return self.rect().adjusted(8, 8, -8, -8)

    def _scale_for_zoom(self, zoom):
        rect = self._view_rect()
        visible_nm = self.range_nm / max(1e-6, zoom)
        if visible_nm <= 0 or rect.width() <= 0:
            return 1.0
        return rect.width() / visible_nm

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        rect = self._view_rect()
        painter.fillRect(rect, QtGui.QColor(16, 20, 28))
        painter.setPen(QtGui.QPen(QtGui.QColor(90, 100, 120), 1, QtCore.Qt.DashLine))
        painter.drawRect(rect)
        # axes
        center_x = rect.center().x()
        center_y = rect.center().y()
        painter.drawLine(QtCore.QLineF(center_x, rect.top(), center_x, rect.bottom()))
        painter.drawLine(QtCore.QLineF(rect.left(), center_y, rect.right(), center_y))
        painter.setPen(QtGui.QPen(QtGui.QColor(60, 70, 85), 1))
        for frac in (-0.5, 0.5):
            painter.drawLine(QtCore.QLineF(rect.left(),
                                           center_y + frac * rect.height(),
                                           rect.right(),
                                           center_y + frac * rect.height()))
            painter.drawLine(QtCore.QLineF(rect.center().x() + frac * rect.width(),
                                           rect.top(),
                                           rect.center().x() + frac * rect.width(),
                                           rect.bottom()))
        if not self.entries:
            painter.end()
            return
        scale = self._scale_for_zoom(self.zoom_factor)
        self._current_scale = scale if scale > 0 else 1.0
        self._poly_map = []
        ordered = sorted(self.entries, key=self._entry_area)
        for entry in ordered:
            key = entry.get('key')
            if key in self._hidden_keys:
                continue
            path = self._draw_entry(painter, rect, scale, entry, entry.get('key') == self.active_key)
            if path is not None:
                self._poly_map.append((key, path, entry))
        painter.end()

    def _draw_entry(self, painter, rect, scale, entry, active):
        cx = entry.get('cx_nm'); cy = entry.get('cy_nm')
        width = entry.get('x_range_nm'); height = entry.get('y_range_nm')
        angle = entry.get('angle_deg', 0.0)
        if None in (cx, cy, width, height):
            return
        half_w = width / 2.0
        half_h = height / 2.0
        pts = [
            QtCore.QPointF(-half_w, -half_h),
            QtCore.QPointF(half_w, -half_h),
            QtCore.QPointF(half_w, half_h),
            QtCore.QPointF(-half_w, half_h),
        ]
        transform = QtGui.QTransform()
        transform.rotate(-angle)
        transformed = [transform.map(pt) for pt in pts]
        offset_x = rect.center().x() + (cx - self._pan_center_nm.x()) * scale
        offset_y = rect.center().y() - (cy - self._pan_center_nm.y()) * scale
        qpoints = []
        for pt in transformed:
            qpoints.append(QtCore.QPointF(offset_x + pt.x() * scale,
                                          offset_y - pt.y() * scale))
        poly = QtGui.QPolygonF(qpoints)
        if self.show_real_images:
            pix = self._entry_pixmaps.get(entry.get('key'))
            if pix is not None:
                painter.save()
                painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
                painter.translate(offset_x, offset_y)
                painter.scale(scale, -scale)
                painter.rotate(-angle)
                target = QtCore.QRectF(-half_w, -half_h, width, height)
                source = QtCore.QRectF(pix.rect())
                painter.drawPixmap(target, pix, source)
                painter.restore()
        pen = QPen(QtGui.QColor(255, 255, 255, 220 if active else 120))
        pen.setWidth(2 if active else 1)
        painter.setPen(pen)
        color = self._entry_color(entry, active)
        pen.setColor(color)
        painter.setPen(pen)
        alpha = 30 if self.show_real_images else (90 if active else 40)
        brush = QBrush(QtGui.QColor(color.red(), color.green(), color.blue(), alpha))
        painter.setBrush(brush)
        painter.drawPolygon(poly)
        path = QtGui.QPainterPath()
        path.addPolygon(poly)
        return path

    def _world_from_pos(self, pos, scale):
        rect = self._view_rect()
        if rect.width() <= 0 or rect.height() <= 0 or scale == 0:
            return None, None
        world_x = self._pan_center_nm.x() + (pos.x() - rect.center().x()) / scale
        world_y = self._pan_center_nm.y() - (pos.y() - rect.center().y()) / scale
        return world_x, world_y

    def _set_pan_center(self, x_nm, y_nm):
        self._pan_center_nm = QtCore.QPointF(float(x_nm), float(y_nm))

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        if delta == 0:
            delta = event.pixelDelta().y()
        if not delta:
            super().wheelEvent(event)
            return

        rect = self._view_rect()
        anchor = event.pos() if rect.contains(event.pos()) else None
        scale_before = self._current_scale if self._current_scale > 0 else self._scale_for_zoom(self.zoom_factor)
        world_anchor = None
        if anchor is not None:
            world_anchor = self._world_from_pos(anchor, scale_before)
        factor = 1.2 if delta > 0 else (1.0 / 1.2)
        new_zoom = self.zoom_factor * factor
        self.set_zoom_factor(new_zoom)
        if anchor is not None and world_anchor and None not in world_anchor:
            scale_after = self._scale_for_zoom(self.zoom_factor)
            center = rect.center()
            world_x, world_y = world_anchor
            px = anchor.x()
            py = anchor.y()
            new_pan_x = world_x - (px - center.x()) / max(1e-6, scale_after)
            new_pan_y = world_y - (center.y() - py) / max(1e-6, scale_after)
            self._set_pan_center(new_pan_x, new_pan_y)
            self.update()
        event.accept()

    def set_zoom_factor(self, factor: float):
        new_factor = float(np.clip(factor, self._min_zoom, self._max_zoom))
        if abs(new_factor - self.zoom_factor) < 1e-6:
            return
        self.zoom_factor = new_factor
        self.update()
        try:
            self.zoomChanged.emit(self.zoom_factor)
        except Exception:
            pass

    def reset_pan(self):
        self._pan_center_nm = QtCore.QPointF(0.0, 0.0)
        self.update()

    def _entry_at_pos(self, pos):
        for key, path, entry in reversed(self._poly_map):
            if path.contains(pos):
                return key, entry
        return None, None

    def mouseMoveEvent(self, event):
        if self._panning and self._last_drag_pos is not None:
            delta = event.pos() - self._last_drag_pos
            self._last_drag_pos = QtCore.QPointF(event.pos())
            if self._current_scale != 0:
                dx_nm = delta.x() / self._current_scale
                dy_nm = delta.y() / self._current_scale
                self._pan_center_nm.setX(self._pan_center_nm.x() - dx_nm)
                self._pan_center_nm.setY(self._pan_center_nm.y() + dy_nm)
                self.update()
            event.accept()
            return
        key, entry = self._entry_at_pos(event.pos())
        if key != self._hover_key:
            self._hover_key = key
            if entry:
                QtWidgets.QToolTip.showText(event.globalPos(), Path(entry.get('key', '')).name)
            else:
                QtWidgets.QToolTip.hideText()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        self._panning = False
        self._last_drag_pos = None
        self.unsetCursor()
        self._hover_key = None
        QtWidgets.QToolTip.hideText()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            key, _ = self._entry_at_pos(event.pos())
            if key:
                if event.modifiers() & QtCore.Qt.ShiftModifier:
                    self.entryShiftClicked.emit(key)
                else:
                    self.entryClicked.emit(key)
                event.accept()
                return
            self._panning = True
            self._last_drag_pos = QtCore.QPointF(event.pos())
            self.setCursor(QtCore.Qt.ClosedHandCursor)
            event.accept()
            return
        if event.button() in (QtCore.Qt.MiddleButton, QtCore.Qt.RightButton):
            self._panning = True
            self._last_drag_pos = QtCore.QPointF(event.pos())
            self.setCursor(QtCore.Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self._panning:
            self._panning = False
            self._last_drag_pos = None
            self.unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)

