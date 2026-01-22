"""Canvas save/load/export helpers."""
from __future__ import annotations

import json
from pathlib import Path

from ..._shared import QtCore, QtGui, QtWidgets
from ..constants import CANVAS_DROP_OFFSET
from .canvas_items import CanvasImageItem


def export_image(window):
    path, _ = QtWidgets.QFileDialog.getSaveFileName(
        window,
        "Export Canvas",
        "canvas_export.png",
        "PNG Image (*.png);;JPEG Image (*.jpg);;PDF Document (*.pdf)",
    )
    if not path:
        return
    rect = window.scene.itemsBoundingRect()
    if rect.isEmpty():
        QtWidgets.QMessageBox.warning(window, "Export", "No items to export")
        return
    padding = 20
    rect = rect.adjusted(-padding, -padding, padding, padding)
    dpi_scale = 3
    image = QtGui.QImage(
        int(rect.width() * dpi_scale),
        int(rect.height() * dpi_scale),
        QtGui.QImage.Format_ARGB32,
    )
    image.fill(QtCore.Qt.white)
    painter = QtGui.QPainter(image)
    painter.setRenderHints(
        QtGui.QPainter.Antialiasing |
        QtGui.QPainter.TextAntialiasing |
        QtGui.QPainter.SmoothPixmapTransform
    )
    painter.scale(dpi_scale, dpi_scale)
    window.scene.render(painter, QtCore.QRectF(), rect)
    painter.end()
    if image.save(path):
        window.status_label.setText(f"Exported to {Path(path).name}")
    else:
        QtWidgets.QMessageBox.warning(window, "Export", f"Failed to save {path}")


def save_canvas(window):
    path, _ = QtWidgets.QFileDialog.getSaveFileName(window, "Save canvas", "canvas.json", "JSON Files (*.json)")
    if not path:
        return
    items = [item.to_state() for item in window.scene.items() if isinstance(item, CanvasImageItem)]
    payload = {"version": 1, "items": items}
    try:
        Path(path).write_text(json.dumps(payload, indent=2))
        window.status_label.setText(f"Saved canvas to {Path(path).name}")
    except Exception as exc:
        QtWidgets.QMessageBox.warning(window, "Save canvas", f"Unable to save: {exc}")


def load_canvas(window):
    path, _ = QtWidgets.QFileDialog.getOpenFileName(window, "Load canvas", "", "JSON Files (*.json)")
    if not path:
        return
    try:
        payload = json.loads(Path(path).read_text())
    except Exception as exc:
        QtWidgets.QMessageBox.warning(window, "Load canvas", f"Unable to load: {exc}")
        return
    window.scene.clear()
    window._drop_offset = QtCore.QPointF(*CANVAS_DROP_OFFSET)
    for state in payload.get("items", []):
        file_path = state.get("file_path")
        channel_idx = state.get("channel_index")
        if not file_path or channel_idx is None:
            continue
        item = window._add_view_from_header(Path(file_path), int(channel_idx), cmap_override=state.get("cmap"))
        if item:
            item.apply_state(state)
    window.status_label.setText(f"Loaded canvas from {Path(path).name}")
    window._push_undo_state()



