"""Layout helpers for arranging canvas items."""
from __future__ import annotations

from ..._shared import QtCore
from ..constants import CANVAS_ALIGN_GAP, CANVAS_ALIGN_MARGIN, CANVAS_DROP_OFFSET_STEP
from .canvas_items import CanvasImageItem


def arrange_by_kind(window, groups: list[dict]):
    kinds = ["topo", "current", "df"]
    if not groups:
        return
    gap_x = CANVAS_ALIGN_GAP
    gap_y = CANVAS_ALIGN_GAP
    margin = CANVAS_ALIGN_MARGIN
    col_widths = []
    for group in groups:
        width = 0.0
        for item in group.values():
            width = max(width, item.boundingRect().width())
        col_widths.append(max(width, 200.0))
    row_heights = []
    for kind in kinds:
        height = 0.0
        for group in groups:
            item = group.get(kind)
            if item is None:
                continue
            height = max(height, item.boundingRect().height())
        row_heights.append(max(height, 0.0))
    for col_idx, group in enumerate(groups):
        x = margin + sum(col_widths[:col_idx]) + gap_x * col_idx
        for row_idx, kind in enumerate(kinds):
            item = group.get(kind)
            if item is None:
                continue
            y = margin + sum(row_heights[:row_idx]) + gap_y * row_idx
            item.setPos(x, y)


def apply_layout(window, layout_type: str):
    items = [i for i in window.scene.items() if isinstance(i, CanvasImageItem)]
    if not items:
        return
    margin = CANVAS_ALIGN_MARGIN
    item_width = max(200, int(items[0].boundingRect().width()))
    item_height = max(200, int(items[0].boundingRect().height()))
    if layout_type == "2x2":
        for i, item in enumerate(items[:4]):
            row = i // 2
            col = i % 2
            x = margin + col * (item_width + margin)
            y = margin + row * (item_height + margin)
            item.setPos(x, y)
    elif layout_type == "1x3":
        for i, item in enumerate(items[:3]):
            x = margin + i * (item_width + margin)
            y = margin
            item.setPos(x, y)
    elif layout_type == "3x1":
        for i, item in enumerate(items[:3]):
            x = margin
            y = margin + i * (item_height + margin)
            item.setPos(x, y)


def place_item(window, item: CanvasImageItem):
    item.setPos(window._drop_offset)
    window._drop_offset = window._drop_offset + QtCore.QPointF(
        CANVAS_DROP_OFFSET_STEP,
        CANVAS_DROP_OFFSET_STEP,
    )



