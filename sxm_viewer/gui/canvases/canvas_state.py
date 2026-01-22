"""Undo/redo and selection helpers for the canvas."""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtCore
from ..constants import CANVAS_DROP_OFFSET
from .canvas_items import CanvasImageItem


def canvas_items(scene):
    return [item for item in scene.items() if isinstance(item, CanvasImageItem)]


def selected_canvas_items(scene):
    return [item for item in scene.selectedItems() if isinstance(item, CanvasImageItem)]


def capture_state(window):
    return [item.to_state() for item in canvas_items(window.scene)]


def restore_state(window, state):
    window._restoring = True
    window.scene.clear()
    window._drop_offset = QtCore.QPointF(*CANVAS_DROP_OFFSET)
    for item_state in state:
        file_path = item_state.get("file_path")
        channel_idx = item_state.get("channel_index")
        if not file_path or channel_idx is None:
            continue
        item = window._add_view_from_header(Path(file_path), int(channel_idx), cmap_override=item_state.get("cmap"))
        if item:
            item.apply_state(item_state)
    window._restoring = False


def push_undo_state(window):
    if window._restoring:
        return
    state = capture_state(window)
    if window._undo_index >= 0 and window._undo_index < len(window._undo_stack) - 1:
        window._undo_stack = window._undo_stack[: window._undo_index + 1]
    window._undo_stack.append(state)
    window._undo_index = len(window._undo_stack) - 1


def undo(window):
    if window._undo_index <= 0:
        return
    window._undo_index -= 1
    restore_state(window, window._undo_stack[window._undo_index])


def redo(window):
    if window._undo_index >= len(window._undo_stack) - 1:
        return
    window._undo_index += 1
    restore_state(window, window._undo_stack[window._undo_index])


def delete_selected(window):
    selected = selected_canvas_items(window.scene)
    if not selected:
        return
    for item in selected:
        window.scene.removeItem(item)
    window._selected_item = None
    window._on_selection_changed()
    push_undo_state(window)



