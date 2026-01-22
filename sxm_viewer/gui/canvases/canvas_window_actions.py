"""Action helpers for the canvas window."""
from __future__ import annotations

from . import canvas_io, canvas_layout


def on_export_image(window):
    return canvas_io.export_image(window)


def on_save_canvas(window):
    return canvas_io.save_canvas(window)


def on_load_canvas(window):
    return canvas_io.load_canvas(window)


def arrange_by_kind(window, groups: list[dict]):
    canvas_layout.arrange_by_kind(window, groups)
    window._push_undo_state()


def apply_layout(window, layout_type: str):
    canvas_layout.apply_layout(window, layout_type)
    window._push_undo_state()


def place_item(window, item):
    return canvas_layout.place_item(window, item)



