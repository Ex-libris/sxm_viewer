"""Enhanced canvas window with modern UI/UX and polished aesthetics."""
from __future__ import annotations

import json
from pathlib import Path
from typing import TYPE_CHECKING

from ..._shared import QtCore, QtGui, QtWidgets, colormaps, np
from ..thumbnail_render import _colormap_icon
from ...data.io import parse_header
from ...processing.detection import _find_topography_channel

if TYPE_CHECKING:
    from typing import Optional

_CANVAS_MIME = "application/x-sxm-view"


def _safe_float(text):
    try:
        return float(text)
    except Exception:
        return None


def _format_colorbar_value(value):
    magnitude = abs(value)
    if magnitude == 0:
        return "0"
    if magnitude < 0.01 or magnitude >= 1000:
        return f"{value:.2e}"
    if magnitude < 1:
        return f"{value:.3f}"
    return f"{value:.2f}"

def _normalize_cbar_label(label: str) -> str:
    if not label:
        return ""
    lbl = label.strip()
    low = lbl.lower()
    if low.startswith("df"):
        return f"Δf{lbl[2:]}"
    if low.startswith("delta f"):
        return f"Δf{lbl[6:]}"
    return lbl


def _normalized_value(norm, value):
    try:
        norm_val = norm(value)
        norm_val = float(norm_val)
    except Exception:
        return 0.5
    return float(np.clip(norm_val, 0.0, 1.0))


def _text_color_for_frame(frame_color):
    color = QtGui.QColor(frame_color or "#070707")
    if not color.isValid():
        color = QtGui.QColor("#070707")
    lum = (0.299 * color.redF()) + (0.587 * color.greenF()) + (0.114 * color.blueF())
    return "#101010" if lum > 0.55 else "#f5f5f5"


def _annotate_colorbar(cb, vmin, vmax, scale, orientation, show_ticks, text_color):
    if cb is None or vmin is None or vmax is None:
        return
    if not show_ticks:
        axis = cb.ax.xaxis if orientation == "horizontal" else cb.ax.yaxis
        axis.set_ticks([])
        axis.set_ticklabels([])
        axis.set_tick_params(length=0)
        for spine in cb.ax.spines.values():
            spine.set_visible(False)
        return
    ticks = [float(vmin), float(vmax)]
    cb.set_ticks(ticks)
    axis = cb.ax.xaxis if orientation == "horizontal" else cb.ax.yaxis
    labels = [_format_colorbar_value(val) for val in ticks]
    axis.set_ticklabels(labels)
    label_size = max(6.0, 8.0 * scale)
    axis.set_tick_params(labelsize=label_size, length=4, colors=text_color, width=0.8)
    for label in axis.get_ticklabels():
        label.set_color(text_color)
        label.set_fontsize(label_size)
    for spine in cb.ax.spines.values():
        spine.set_visible(False)


def render_tile_mpl(
    data,
    *,
    cmap,
    vmin,
    vmax,
    title,
    colorbar_label,
    width_px,
    height_px,
    dpi=200,
    show_colorbar=True,
    show_colorbar_ticks=True,
    show_title=True,
    show_metadata=True,
    metadata_left="",
    metadata_right="",
    show_overlay_main=False,
    overlay_main="",
    show_overlay_file=False,
    overlay_file="",
    cbar_position="bottom",
    metadata_height=0.0,
    frame_color="#070707",
    text_scale=None,
    text_color=None,
):
    """Render a canvas tile through Matplotlib, including annotations."""
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg

    width_px = max(2, int(round(width_px)))
    height_px = max(2, int(round(height_px)))
    normalized_position = (cbar_position or "bottom").lower()
    if normalized_position == "hidden":
        normalized_position = "none"
    if normalized_position not in ("bottom", "top", "left", "right", "inset", "none"):
        normalized_position = "bottom"
    rendered_colorbar = show_colorbar and normalized_position != "none"

    metadata_ratio = 0.0
    if show_metadata and metadata_height > 0:
        metadata_ratio = min(0.35, metadata_height / height_px)
    bottom_margin = 0.02 + metadata_ratio
    if text_scale is None:
        text_scale = 1.0
    text_scale = max(0.2, min(2.4, float(text_scale)))
    text_color = text_color or _text_color_for_frame(frame_color)
    cbar_label_text = _normalize_cbar_label(colorbar_label or "")

    min_title_px = 10.0
    min_tick_px = 9.0
    min_label_px = 9.0
    min_overlay_px = 8.5
    min_meta_px = 8.5

    title_fs = max(min_title_px, 11.0 * text_scale)
    tick_fs = max(min_tick_px, 9.5 * text_scale)
    label_fs = max(min_label_px, 9.5 * text_scale)
    overlay_fs = max(min_overlay_px, 8.5 * text_scale)
    meta_fs = max(min_meta_px, 8.5 * text_scale)
    tick_pad = max(2.0, 0.25 * tick_fs)

    extra_tick_margin = 0.0
    if rendered_colorbar and show_colorbar_ticks and normalized_position in ("top", "bottom"):
        extra_tick_margin = min(0.18, max(0.06, (tick_fs * 2.2) / max(1.0, height_px)))
    top_margin = 0.98 - (extra_tick_margin if normalized_position == "top" else 0.0)
    fig = Figure(figsize=(width_px / dpi, height_px / dpi), dpi=dpi, facecolor=frame_color or "#070707")
    fig.subplots_adjust(
        left=0.0,
        right=1.0,
        top=top_margin,
        bottom=bottom_margin + (extra_tick_margin if normalized_position == "bottom" else 0.0),
    )
    canvas = FigureCanvasAgg(fig)

    ax = None
    cax = None
    orientation = "horizontal"
    if rendered_colorbar and normalized_position in ("top", "bottom"):
        cbar_height = max(10.0, 1.5 * tick_fs)
        cbar_ratio = min(0.45, max(0.08, cbar_height / max(1.0, height_px)))
        ratios = [cbar_ratio, 1] if normalized_position == "top" else [1, cbar_ratio]
        gs = fig.add_gridspec(2, 1, height_ratios=ratios, hspace=0.04)
        if normalized_position == "top":
            cax = fig.add_subplot(gs[0])
            ax = fig.add_subplot(gs[1])
        else:
            ax = fig.add_subplot(gs[0])
            cax = fig.add_subplot(gs[1])
        orientation = "horizontal"
    elif rendered_colorbar and normalized_position in ("left", "right"):
        cbar_width = max(10.0, 1.5 * tick_fs)
        cbar_ratio = min(0.45, max(0.08, cbar_width / max(1.0, width_px)))
        ratios = [cbar_ratio, 1] if normalized_position == "left" else [1, cbar_ratio]
        gs = fig.add_gridspec(1, 2, width_ratios=ratios, wspace=0.04)
        if normalized_position == "left":
            cax = fig.add_subplot(gs[0])
            ax = fig.add_subplot(gs[1])
        else:
            ax = fig.add_subplot(gs[0])
            cax = fig.add_subplot(gs[1])
        orientation = "vertical"
    else:
        ax = fig.add_subplot(1, 1, 1)

    try:
        cmap_obj = colormaps.get(cmap) if cmap else colormaps.get("viridis")
    except Exception:
        cmap_obj = colormaps.get("viridis")

    im = ax.imshow(
        data,
        cmap=cmap_obj,
        vmin=vmin,
        vmax=vmax,
        origin="lower",
        interpolation="nearest",
    )

    actual_vmin = None
    actual_vmax = None
    if vmin is not None and vmax is not None:
        actual_vmin = float(vmin)
        actual_vmax = float(vmax)
    else:
        try:
            actual_vmin = float(np.nanmin(data))
            actual_vmax = float(np.nanmax(data))
        except Exception:
            actual_vmin = None
            actual_vmax = None

    ax.set_facecolor("#070707")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    if show_title and title:
        ax.set_title(title, fontsize=title_fs, color=text_color, pad=6)

    cb = None
    if rendered_colorbar and normalized_position in ("top", "bottom", "left", "right"):
        cb = fig.colorbar(im, cax=cax, orientation=orientation)
        cb.set_label("")  # place label manually inside
        cb.ax.tick_params(labelsize=tick_fs, length=3, width=0.7, colors=text_color, pad=tick_pad)
        cb.outline.set_edgecolor("#797979")
        cb.outline.set_linewidth(0.6)
        if orientation == "horizontal":
            cb.ax.xaxis.set_ticks_position("bottom")
            cb.ax.xaxis.set_label_position("bottom")
            cb.ax.text(0.5, 0.5, cbar_label_text, color=text_color, fontsize=label_fs, ha="center", va="center", transform=cb.ax.transAxes)
        else:
            cb.ax.yaxis.set_ticks_position("left")
            cb.ax.yaxis.set_label_position("left")
            cb.ax.text(0.5, 0.5, cbar_label_text, color=text_color, fontsize=label_fs, ha="center", va="center", rotation=90, rotation_mode="anchor", transform=cb.ax.transAxes)
    elif rendered_colorbar and normalized_position == "inset":
        inset_pos = [0.62, bottom_margin + 0.02, 0.3, 0.035]
        cax = fig.add_axes(inset_pos)
        cb = fig.colorbar(im, cax=cax, orientation="horizontal")
        cb.set_label("")
        cb.ax.tick_params(labelsize=tick_fs, length=3, width=0.5, colors=text_color, pad=tick_pad)
        cb.outline.set_edgecolor("#797979")
        cb.outline.set_linewidth(0.6)
        cb.ax.text(0.5, 0.5, cbar_label_text, color=text_color, fontsize=label_fs, ha="center", va="center", transform=cb.ax.transAxes)

    if cb is not None:
        annotate_orientation = orientation if normalized_position != "inset" else "horizontal"
        _annotate_colorbar(
            cb,
            actual_vmin,
            actual_vmax,
            tick_fs / 8.0,
            annotate_orientation,
            show_colorbar_ticks,
            text_color,
        )

    if ax is not None and cax is not None and normalized_position in ("top", "bottom"):
        ax_pos = ax.get_position()
        cax_pos = cax.get_position()
        cax.set_position([ax_pos.x0, cax_pos.y0, ax_pos.width, cax_pos.height])
    if ax is not None and cax is not None and normalized_position in ("left", "right"):
        ax_pos = ax.get_position()
        cax_pos = cax.get_position()
        cax.set_position([cax_pos.x0, ax_pos.y0, cax_pos.width, ax_pos.height])

    overlay_lines = []
    if show_overlay_main and overlay_main:
        overlay_lines.append(overlay_main)
    if show_overlay_file and overlay_file:
        overlay_lines.append(overlay_file)
    if overlay_lines:
        overlay_face = "#0b1424" if text_color == "#f5f5f5" else "#f5f5f5"
        ax.text(
            0.02,
            0.96,
            "\n".join(overlay_lines),
            fontsize=overlay_fs,
            color=text_color,
            weight="bold",
            ha="left",
            va="top",
            transform=ax.transAxes,
            bbox=dict(facecolor=overlay_face, alpha=0.75, edgecolor="none", boxstyle="round,pad=0.2"),
        )

    if show_metadata and (metadata_left or metadata_right):
        text_y = bottom_margin / 2
        meta_face = "#050505" if text_color == "#f5f5f5" else "#f5f5f5"
        bbox = dict(facecolor=meta_face, alpha=0.9, edgecolor="none", boxstyle="round,pad=0.2")
        font_size = meta_fs
        if metadata_left:
            fig.text(
                0.02,
                text_y,
                metadata_left,
                fontsize=font_size,
                color=text_color,
                ha="left",
                va="center",
                bbox=bbox,
            )
        if metadata_right:
            fig.text(
                0.98,
                text_y,
                metadata_right,
                fontsize=font_size,
                color=text_color,
                ha="right",
                va="center",
                bbox=bbox,
            )

    canvas.draw()
    buf = np.asarray(canvas.buffer_rgba())
    h, w, _ = buf.shape
    qimg = QtGui.QImage(buf.data, w, h, QtGui.QImage.Format_RGBA8888)
    return QtGui.QPixmap.fromImage(qimg.copy())


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
        self._show_title = True
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
        candidates = [0.1, 0.2, 0.5, 1, 2, 3, 5, 10, 20, 50, 100, 200, 500, 1000, 2000, 5000]
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
        text_color = _text_color_for_frame(frame_color)
        show_overlay_main = self._show_overlay_main and not self._metadata_bar_visible
        show_overlay_file = self._show_overlay_file and not self._metadata_bar_visible
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

        action = menu.exec_(event.screenPos())

        if action is not None:
            if action == duplicate_action:
                if self._parent_window:
                    self._parent_window._on_duplicate_item()
            elif action == delete_action:
                if self._parent_window:
                    self._parent_window._on_remove_item()
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

    def set_show_title(self, show: bool):
        self._show_title = show
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
            "show_title": self._show_title,
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
        self.set_show_title(state.get("show_title", True))
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
            self._fixed_text_scale_value = max(0.2, min(2.4, float(ts)))
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

    @property
    def data_array(self) -> np.ndarray:
        return self._arr


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

        action = menu.exec_(event.globalPos())

        if action == select_all:
            for item in self.scene().items():
                if isinstance(item, CanvasImageItem):
                    item.setSelected(True)
        elif action == deselect_all:
            self.scene().clearSelection()
        elif action == zoom_in:
            self.scale(1.15, 1.15)
        elif action == zoom_out:
            self.scale(1/1.15, 1/1.15)
        elif action == zoom_reset:
            self.resetTransform()
        elif action == fit_view:
            self.fitInView(self.scene().itemsBoundingRect(), QtCore.Qt.KeepAspectRatio)
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
            elif action == canvas_actions.get("layout_2x2"):
                parent._apply_layout("2x2")
            elif action == canvas_actions.get("layout_1x3"):
                parent._apply_layout("1x3")
            elif action == canvas_actions.get("layout_3x1"):
                parent._apply_layout("3x1")

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
                payloads.append(json.loads(data))
            except Exception:
                payloads = []
        paths = []
        if mime.hasUrls():
            for url in mime.urls():
                if url.isLocalFile():
                    paths.append(url.toLocalFile())
        if payloads or paths:
            parent.handle_drop(payloads, paths)
            event.acceptProposedAction()
            return
        super().dropEvent(event)


class ExperimentalCanvasWindow(QtWidgets.QDialog):
    def __init__(self, viewer, parent=None):
        super().__init__(parent)
        self.viewer = viewer
        self.setWindowTitle("Enhanced Scientific Canvas")
        self.resize(1400, 900)
        self._drop_offset = QtCore.QPointF(40.0, 40.0)
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
        self._global_show_title = True
        self._global_show_colorbar = True
        self._global_show_colorbar_ticks = True
        self._metadata_bar_default = True
        self._colorbar_mode = "bottom"
        self._undo_stack = []
        self._undo_index = -1
        self._file_scale_bars = {}
        self._restoring = False

        # Apply modern styling
        self._apply_styles()

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Create toolbar and view first
        self.scene = QtWidgets.QGraphicsScene(self)
        self.view = CanvasGraphicsView(self)
        self.view.setScene(self.scene)

        # Build UI
        toolbar_widget = self._build_toolbar()
        main_layout.addWidget(toolbar_widget)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.addWidget(self.view)
        splitter.addWidget(self._build_inspector())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([900, 420])
        main_layout.addWidget(splitter, 1)

        self.status_label = QtWidgets.QLabel("Ready")
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #2d2d2d;
                color: #ffffff;
                padding: 8px 16px;
                border-top: 2px solid #4a5568;
                font-size: 12px;
                font-weight: 500;
            }
        """)
        main_layout.addWidget(self.status_label)

        self.scene.selectionChanged.connect(self._on_selection_changed)
        self._push_undo_state()

    def _create_icon_button(self, text: str, icon_text: str = "", tooltip: str = "") -> QtWidgets.QPushButton:
        """Create a button with optional icon."""
        display_text = f"{icon_text} {text}" if icon_text else text
        btn = QtWidgets.QPushButton(display_text)
        if tooltip:
            btn.setToolTip(tooltip)
        return btn

    def _create_toolbar_section(self, title: str, widgets: list) -> QtWidgets.QWidget:
        """Create a visually grouped toolbar section."""
        section = QtWidgets.QWidget()
        section.setStyleSheet("""
            QWidget {
                background-color: #2d3748;
                border-radius: 8px;
                padding: 4px;
            }
        """)
        layout = QtWidgets.QHBoxLayout(section)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(4)
        
        if title:
            label = QtWidgets.QLabel(title)
            label.setStyleSheet("font-weight: bold; color: #a0aec0; font-size: 11px;")
            layout.addWidget(label)
        
        for widget in widgets:
            layout.addWidget(widget)
        
        return section

    def _build_toolbar(self):
        """Build organized toolbar with clear visual grouping."""

        toolbar_widget = QtWidgets.QWidget()
        toolbar_widget.setStyleSheet("""
            QWidget {
                background-color: #2d2d2d;
                border-bottom: 2px solid #4a5568;
            }
        """)
        toolbar_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)

        main_layout = QtWidgets.QVBoxLayout(toolbar_widget)
        main_layout.setContentsMargins(8, 6, 8, 6)
        main_layout.setSpacing(6)

        row1 = QtWidgets.QHBoxLayout()
        row1.setSpacing(8)

        file_group, file_layout = self._create_toolbar_group("FILE")
        self.save_btn = QtWidgets.QPushButton("Save")
        self.save_btn.setToolTip("Save canvas layout (Ctrl+S)")
        self.load_btn = QtWidgets.QPushButton("Load")
        self.load_btn.setToolTip("Load canvas layout")
        self.export_btn = QtWidgets.QPushButton("Export")
        self.export_btn.setToolTip("Export canvas as image")
        file_layout.addWidget(self.save_btn)
        file_layout.addWidget(self.load_btn)
        file_layout.addWidget(self.export_btn)
        row1.addWidget(file_group)

        row1.addWidget(self._create_separator())

        layout_group, layout_layout = self._create_toolbar_group("LAYOUT")
        self.layout_2x2_btn = QtWidgets.QPushButton("2x2")
        self.layout_2x2_btn.setToolTip("Arrange in 2x2 grid")
        self.layout_1x3_btn = QtWidgets.QPushButton("1x3")
        self.layout_1x3_btn.setToolTip("Arrange in 1x3 row")
        self.layout_3x1_btn = QtWidgets.QPushButton("3x1")
        self.layout_3x1_btn.setToolTip("Arrange in 3x1 column")
        layout_layout.addWidget(self.layout_2x2_btn)
        layout_layout.addWidget(self.layout_1x3_btn)
        layout_layout.addWidget(self.layout_3x1_btn)
        row1.addWidget(layout_group)

        row1.addWidget(self._create_separator())

        view_group, view_layout = self._create_toolbar_group("VIEW")
        self.show_grid_check = QtWidgets.QCheckBox("Grid")
        self.show_grid_check.setToolTip("Show alignment grid")
        self.snap_grid_check = QtWidgets.QCheckBox("Snap")
        self.snap_grid_check.setToolTip("Snap items to grid")
        self.canvas_color_btn = QtWidgets.QPushButton("Color")
        self.canvas_color_btn.setToolTip("Change canvas background color")
        self.canvas_color_btn.setMaximumWidth(60)
        view_layout.addWidget(self.show_grid_check)
        view_layout.addWidget(self.snap_grid_check)
        view_layout.addWidget(self.canvas_color_btn)
        row1.addWidget(view_group)

        row1.addStretch(1)
        main_layout.addLayout(row1)

        row2 = QtWidgets.QHBoxLayout()
        row2.setSpacing(8)

        annotation_group, annotation_layout = self._create_toolbar_group("ANNOTATE")
        self.show_title_check = QtWidgets.QCheckBox("Title")
        self.show_title_check.setToolTip("Show titles on all tiles")
        self.show_title_check.setChecked(self._global_show_title)
        self.show_colorbar_check = QtWidgets.QCheckBox("Colorbar")
        self.show_colorbar_check.setToolTip("Show colorbars on all tiles")
        self.show_colorbar_check.setChecked(self._global_show_colorbar)
        annotation_layout.addWidget(self.show_title_check)
        annotation_layout.addWidget(self.show_colorbar_check)
        row2.addWidget(annotation_group)

        row2.addWidget(self._create_separator())

        sync_group, sync_layout = self._create_toolbar_group("SYNC")
        self.sync_cbar_check = QtWidgets.QCheckBox("Ranges")
        self.sync_cbar_check.setToolTip("Sync color ranges across items")
        self.sync_by_channel_check = QtWidgets.QCheckBox("Colors by channel")
        self.sync_by_channel_check.setChecked(self._sync_by_channel)
        self.sync_by_channel_check.setToolTip("Sync colormaps by channel type")
        sync_layout.addWidget(self.sync_cbar_check)
        sync_layout.addWidget(self.sync_by_channel_check)
        row2.addWidget(sync_group)

        row2.addWidget(self._create_separator())

        overlay_group, overlay_layout = self._create_toolbar_group("OVERLAY")
        self.overlay_info_check = QtWidgets.QCheckBox("Channel/Date")
        self.overlay_info_check.setChecked(self._show_overlay_info)
        self.overlay_info_check.setToolTip("Show channel and date overlay")
        self.overlay_file_check = QtWidgets.QCheckBox("Filename")
        self.overlay_file_check.setChecked(self._show_overlay_file)
        self.overlay_file_check.setToolTip("Show filename overlay")
        overlay_layout.addWidget(self.overlay_info_check)
        overlay_layout.addWidget(self.overlay_file_check)
        row2.addWidget(overlay_group)

        row2.addWidget(self._create_separator())

        colorbar_group, colorbar_layout = self._create_toolbar_group("COLORBAR")
        self.colorbar_ticks_check = QtWidgets.QCheckBox("Ticks")
        self.colorbar_ticks_check.setToolTip("Show min/max ticks on colorbars")
        self.colorbar_ticks_check.setChecked(self._global_show_colorbar_ticks)
        self.colorbar_mode_combo = QtWidgets.QComboBox()
        for label in ("Bottom", "Top", "Right", "Left", "Inset", "Hidden"):
            self.colorbar_mode_combo.addItem(label)
        colorbar_layout.addWidget(self.colorbar_ticks_check)
        colorbar_layout.addWidget(QtWidgets.QLabel("Position"))
        colorbar_layout.addWidget(self.colorbar_mode_combo)
        row2.addWidget(colorbar_group)
        row2.addWidget(self._create_separator())

        align_group, align_layout = self._create_toolbar_group("ALIGN")
        self.align_btn = QtWidgets.QPushButton("Align selected")
        self.align_btn.setToolTip("Align selected items")
        self.align_channels_btn = QtWidgets.QPushButton("Align by channel")
        self.align_channels_btn.setToolTip("Align all items by channel type")
        self.reset_alignment_btn = QtWidgets.QPushButton("Reset alignment")
        self.reset_alignment_btn.setToolTip("Unlock aligned sizes and text scaling")
        align_layout.addWidget(self.align_btn)
        align_layout.addWidget(self.align_channels_btn)
        align_layout.addWidget(self.reset_alignment_btn)
        row2.addWidget(align_group)

        row2.addStretch(1)
        main_layout.addLayout(row2)

        self.save_btn.clicked.connect(self._on_save_canvas)
        self.load_btn.clicked.connect(self._on_load_canvas)
        self.export_btn.clicked.connect(self._on_export_image)
        self.layout_2x2_btn.clicked.connect(lambda: self._apply_layout("2x2"))
        self.layout_1x3_btn.clicked.connect(lambda: self._apply_layout("1x3"))
        self.layout_3x1_btn.clicked.connect(lambda: self._apply_layout("3x1"))
        self.show_grid_check.toggled.connect(self.view.set_show_grid)
        self.snap_grid_check.toggled.connect(self.view.set_snap_to_grid)
        self.sync_cbar_check.toggled.connect(self._on_sync_colorbars_toggled)
        self.sync_by_channel_check.toggled.connect(self._on_sync_by_channel_toggled)
        self.overlay_info_check.toggled.connect(self._on_overlay_info_toggled)
        self.overlay_file_check.toggled.connect(self._on_overlay_file_toggled)
        self.colorbar_ticks_check.toggled.connect(self._on_global_show_colorbar_ticks_toggled)
        self.colorbar_mode_combo.currentTextChanged.connect(self._on_colorbar_position_changed)
        self.canvas_color_btn.clicked.connect(self._on_canvas_color_clicked)
        self.align_btn.clicked.connect(self._on_align_selected)
        self.align_channels_btn.clicked.connect(self._on_align_by_channels)
        self.reset_alignment_btn.clicked.connect(self._reset_locked_alignment)

        self._on_overlay_info_toggled(self.overlay_info_check.isChecked())
        self._on_overlay_file_toggled(self.overlay_file_check.isChecked())
        self.colorbar_mode_combo.setCurrentText(self._colorbar_mode.capitalize())

        return toolbar_widget

    def _create_toolbar_group(self, title):
        """Create a visually grouped section in toolbar."""
        container = QtWidgets.QWidget()
        container_layout = QtWidgets.QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(2)

        label = QtWidgets.QLabel(title)
        label.setStyleSheet("""
            QLabel {
                color: #9ca3af;
                font-size: 10px;
                font-weight: bold;
                padding: 2px 8px 0px 8px;
            }
        """)
        label.setAlignment(QtCore.Qt.AlignLeft)
        container_layout.addWidget(label)

        group = QtWidgets.QWidget()
        group.setStyleSheet("""
            QWidget {
                background-color: #3a3a3a;
                border: 1px solid #4a5568;
                border-radius: 4px;
            }
        """)
        group_layout = QtWidgets.QHBoxLayout(group)
        group_layout.setContentsMargins(8, 4, 8, 4)
        group_layout.setSpacing(4)
        container_layout.addWidget(group)
        return container, group_layout

    def _create_separator(self):
        """Create a vertical separator line."""
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.VLine)
        separator.setFrameShadow(QtWidgets.QFrame.Sunken)
        separator.setStyleSheet("""
            QFrame {
                color: #4a5568;
                max-width: 1px;
            }
        """)
        return separator

    def _apply_styles(self):
        """Apply scientific GUI styling - high contrast, clear organization."""
        self.setStyleSheet("""
            /* Main Window */
            QDialog {
                background-color: #1a1a1a;
            }
            
            /* Primary Buttons (File operations) */
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, 
                                            stop:0 #4a5568, stop:1 #2d3748);
                color: #ffffff;
                border: 1px solid #2d3748;
                border-radius: 4px;
                padding: 6px 14px;
                font-weight: 500;
                font-size: 13px;
                min-height: 24px;
            }
            
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #5a67d8, stop:1 #4c51bf);
                border: 1px solid #5a67d8;
            }
            
            QPushButton:pressed {
                background: #4c51bf;
            }
            
            QPushButton:disabled {
                background: #2d3748;
                color: #6b7280;
            }
            
            /* Input Fields */
            QLineEdit, QComboBox {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 1px solid #4a5568;
                border-radius: 4px;
                padding: 6px 8px;
                font-size: 13px;
                selection-background-color: #5a67d8;
            }
            
            QLineEdit:focus, QComboBox:focus {
                border: 2px solid #5a67d8;
                background-color: #353535;
            }
            
            QLineEdit::placeholder {
                color: #6b7280;
            }
            
            /* Checkboxes */
            QCheckBox {
                color: #ffffff;
                spacing: 6px;
                font-size: 13px;
                padding: 2px;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #4a5568;
                background-color: #2d2d2d;
            }
            
            QCheckBox::indicator:hover {
                border-color: #5a67d8;
                background-color: #353535;
            }
            
            QCheckBox::indicator:checked {
                background-color: #5a67d8;
                border-color: #5a67d8;
                image: url(none);
            }
            
            QCheckBox::indicator:checked:hover {
                background-color: #6366f1;
                border-color: #6366f1;
            }
            
            /* Labels */
            QLabel {
                color: #e5e5e5;
                font-size: 13px;
            }
            
            /* Group Boxes */
            QGroupBox {
                color: #ffffff;
                font-weight: bold;
                font-size: 14px;
                border: 2px solid #4a5568;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 18px;
                background-color: #252525;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px;
                background-color: #252525;
            }
            
            /* Combo Box Drop Down */
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 6px solid #ffffff;
                margin-right: 6px;
            }
            
            QComboBox QAbstractItemView {
                background-color: #2d2d2d;
                color: #ffffff;
                selection-background-color: #5a67d8;
                border: 1px solid #4a5568;
            }
            
            /* Scroll Area */
            QScrollArea {
                border: none;
                background-color: #1a1a1a;
            }
            
            QScrollBar:vertical {
                background: #2d2d2d;
                width: 12px;
                border-radius: 6px;
            }
            
            QScrollBar::handle:vertical {
                background: #4a5568;
                border-radius: 6px;
                min-height: 20px;
            }
            
            QScrollBar::handle:vertical:hover {
                background: #5a67d8;
            }
        """)

    def _build_inspector(self):
        """Build inspector panel with better visual hierarchy."""
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        panel = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(panel)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        # Header with higher contrast
        header = QtWidgets.QLabel("SELECTED ITEM")
        header.setStyleSheet("""
            QLabel {
                color: #ffffff;
                font-size: 16px;
                font-weight: bold;
                padding-bottom: 8px;
                border-bottom: 2px solid #5a67d8;
            }
        """)
        layout.addWidget(header)

        # Item Info Group
        info_group = QtWidgets.QGroupBox("Item Info")
        info_layout = QtWidgets.QFormLayout()
        info_layout.setLabelAlignment(QtCore.Qt.AlignRight)
        info_layout.setVerticalSpacing(10)
        info_layout.setHorizontalSpacing(12)

        label_style = "QLabel { color: #d0d0d0; font-weight: 500; }"

        file_label_text = QtWidgets.QLabel("File:")
        file_label_text.setStyleSheet(label_style)
        self.file_label = QtWidgets.QLabel("-")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet("QLabel { color: #ffffff; }")
        info_layout.addRow(file_label_text, self.file_label)

        channel_label_text = QtWidgets.QLabel("Channel:")
        channel_label_text.setStyleSheet(label_style)
        self.channel_label = QtWidgets.QLabel("-")
        self.channel_label.setStyleSheet("QLabel { color: #ffffff; }")
        info_layout.addRow(channel_label_text, self.channel_label)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # Appearance Group
        appearance_group = QtWidgets.QGroupBox("Appearance")
        appearance_layout = QtWidgets.QFormLayout()
        appearance_layout.setLabelAlignment(QtCore.Qt.AlignRight)
        appearance_layout.setVerticalSpacing(10)
        appearance_layout.setHorizontalSpacing(12)

        colorbar_label = QtWidgets.QLabel("Label:")
        colorbar_label.setStyleSheet(label_style)
        self.colorbar_edit = QtWidgets.QLineEdit()
        self.colorbar_edit.setPlaceholderText("Enter label...")
        appearance_layout.addRow(colorbar_label, self.colorbar_edit)

        text_scale_label = QtWidgets.QLabel("Text Scale:")
        text_scale_label.setStyleSheet(label_style)
        self.text_scale_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.text_scale_slider.setMinimum(20)
        self.text_scale_slider.setMaximum(240)
        self.text_scale_slider.setValue(100)
        self.text_scale_slider.setEnabled(False)
        appearance_layout.addRow(text_scale_label, self.text_scale_slider)

        appearance_group.setLayout(appearance_layout)
        layout.addWidget(appearance_group)

        # Colormap Group
        colormap_group = QtWidgets.QGroupBox("Colormap")
        colormap_layout = QtWidgets.QVBoxLayout()
        colormap_layout.setSpacing(10)

        self.cmap_combo = QtWidgets.QComboBox()
        try:
            cmap_list = sorted(colormaps.keys())
        except Exception:
            cmap_list = ["viridis", "plasma", "inferno", "magma", "cividis"]
        for name in cmap_list:
            try:
                icon = _colormap_icon(name, width=96, height=14)
            except Exception:
                icon = QtGui.QIcon()
            self.cmap_combo.addItem(icon, name)
        colormap_layout.addWidget(self.cmap_combo)

        range_label = QtWidgets.QLabel("Range:")
        range_label.setStyleSheet("QLabel { color: #d0d0d0; font-weight: 500; }")
        colormap_layout.addWidget(range_label)

        range_container = QtWidgets.QWidget()
        range_layout = QtWidgets.QHBoxLayout(range_container)
        range_layout.setContentsMargins(0, 0, 0, 0)
        range_layout.setSpacing(8)

        self.vmin_edit = QtWidgets.QLineEdit()
        self.vmin_edit.setPlaceholderText("min")
        range_layout.addWidget(self.vmin_edit)

        to_label = QtWidgets.QLabel("to")
        to_label.setStyleSheet("QLabel { color: #9ca3af; }")
        range_layout.addWidget(to_label)

        self.vmax_edit = QtWidgets.QLineEdit()
        self.vmax_edit.setPlaceholderText("max")
        range_layout.addWidget(self.vmax_edit)

        colormap_layout.addWidget(range_container)

        self.auto_range_btn = QtWidgets.QPushButton("Auto Range")
        self.auto_range_btn.setToolTip("Reset to automatic range")
        self.copy_range_btn = QtWidgets.QPushButton("Copy Range to Selected")
        self.copy_range_btn.setToolTip("Copy range to other selected items")

        colormap_layout.addWidget(self.auto_range_btn)
        colormap_layout.addWidget(self.copy_range_btn)
        colormap_group.setLayout(colormap_layout)
        layout.addWidget(colormap_group)

        # Statistics Group
        stats_group = QtWidgets.QGroupBox("Statistics")
        stats_layout = QtWidgets.QVBoxLayout()
        self.stats_label = QtWidgets.QLabel("-")
        self.stats_label.setWordWrap(True)
        self.stats_label.setStyleSheet("""
            QLabel {
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 12px;
                color: #ffffff;
                background-color: #2d2d2d;
                padding: 8px;
                border-radius: 4px;
            }
        """)
        stats_layout.addWidget(self.stats_label)
        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        # Actions Group
        actions_group = QtWidgets.QGroupBox("Actions")
        actions_layout = QtWidgets.QVBoxLayout()
        actions_layout.setSpacing(8)

        self.duplicate_btn = QtWidgets.QPushButton("Duplicate")
        self.duplicate_btn.setToolTip("Duplicate selected item (Ctrl+D)")

        self.remove_btn = QtWidgets.QPushButton("Remove Item")
        self.remove_btn.setToolTip("Remove selected item (Delete)")
        self.remove_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #dc2626, stop:1 #b91c1c);
                color: #ffffff;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #ef4444, stop:1 #dc2626);
            }
            QPushButton:pressed {
                background: #b91c1c;
            }
        """)

        actions_layout.addWidget(self.duplicate_btn)
        actions_layout.addWidget(self.remove_btn)
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)

        layout.addStretch(1)

        # Connect signals
        self.colorbar_edit.editingFinished.connect(self._on_colorbar_changed)
        self.text_scale_slider.valueChanged.connect(self._on_text_scale_changed)
        self.cmap_combo.currentTextChanged.connect(self._on_cmap_changed)
        self.vmin_edit.editingFinished.connect(self._on_range_changed)
        self.vmax_edit.editingFinished.connect(self._on_range_changed)
        self.auto_range_btn.clicked.connect(self._on_auto_range)
        self.copy_range_btn.clicked.connect(self._on_copy_range)
        self.show_title_check.toggled.connect(self._on_global_show_title_toggled)
        self.show_colorbar_check.toggled.connect(self._on_global_show_colorbar_toggled)
        self.duplicate_btn.clicked.connect(self._on_duplicate_item)
        self.remove_btn.clicked.connect(self._on_remove_item)

        scroll.setWidget(panel)
        self._set_inspector_enabled(False)
        return scroll

    def _hline(self):
        line = QtWidgets.QFrame()
        line.setFrameShape(QtWidgets.QFrame.HLine)
        line.setFrameShadow(QtWidgets.QFrame.Sunken)
        return line

    def _set_inspector_enabled(self, enabled: bool):
        for widget in (
            self.colorbar_edit,
            self.text_scale_slider,
            self.cmap_combo,
            self.vmin_edit,
            self.vmax_edit,
            self.auto_range_btn,
            self.copy_range_btn,
            self.duplicate_btn,
            self.remove_btn,
        ):
            widget.setEnabled(enabled)

    def _on_selection_changed(self):
        selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        item = selected[0] if selected else None
        self._selected_item = item
        if item is None:
            self.file_label.setText("-")
            self.channel_label.setText("-")
            self.colorbar_edit.setText("")
            self.text_scale_slider.setValue(100)
            self.vmin_edit.setText("")
            self.vmax_edit.setText("")
            self.stats_label.setText("-")
            self._set_inspector_enabled(False)
            return
        self._set_inspector_enabled(True)
        self.file_label.setText(Path(item.file_path).name)
        self.channel_label.setText(str(item.channel_index))
        self.colorbar_edit.setText(item.colorbar_label)
        try:
            self.text_scale_slider.blockSignals(True)
            self.text_scale_slider.setValue(int(round((item._fixed_text_scale_value if item._use_fixed_text_scale else 1.0) * 100)))
        finally:
            self.text_scale_slider.blockSignals(False)
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
        if self._selected_item is None:
            return
        scale = max(0.4, min(2.4, value / 100.0))
        self._selected_item._fixed_text_scale_value = scale
        self._selected_item._use_fixed_text_scale = True
        # Clear any alignment-locked text scale so the slider takes effect.
        self._selected_item.set_locked_text_scale(None)
        self._selected_item._update_rendered_pixmap()
        self._push_undo_state()

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

    def _on_copy_range(self):
        if self._selected_item is None:
            return
        vmin = self._selected_item.vmin
        vmax = self._selected_item.vmax
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem) and item.isSelected():
                item.set_range(vmin, vmax)
        self._push_undo_state()

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

    def _on_overlay_file_toggled(self, checked: bool):
        self._show_overlay_file = bool(checked)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_overlay(self._show_overlay_info, self._show_overlay_file)
                item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
                item.set_metadata_file_visible(self._show_overlay_file)

    def _metadata_bar_visible_default(self) -> bool:
        return bool(self._metadata_bar_default)

    def _on_colorbar_position_changed(self, text: str):
        mode = text.lower()
        if mode == "hidden":
            mode = "none"
        mode = mode if mode in ("bottom", "top", "left", "right", "inset", "none") else "bottom"
        self._colorbar_mode = mode
        self._apply_colorbar_mode_to_all(mode)

    def _apply_colorbar_mode_to_all(self, mode: str):
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_colorbar_mode(mode)
        self.status_label.setText(f"Colorbar mode: {mode.capitalize()}")

    def _on_global_show_title_toggled(self, checked: bool):
        self._apply_global_show_title(checked)

    def _on_global_show_colorbar_toggled(self, checked: bool):
        self._apply_global_show_colorbar(checked)

    def _on_global_show_colorbar_ticks_toggled(self, checked: bool):
        self._apply_global_show_colorbar_ticks(checked)

    def _apply_global_show_title(self, show: bool):
        self._global_show_title = bool(show)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_title(self._global_show_title)

    def _apply_global_show_colorbar(self, show: bool):
        self._global_show_colorbar = bool(show)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_colorbar(self._global_show_colorbar)

    def _apply_global_show_colorbar_ticks(self, show: bool):
        self._global_show_colorbar_ticks = bool(show)
        for item in self.scene.items():
            if isinstance(item, CanvasImageItem):
                item.set_show_colorbar_ticks(self._global_show_colorbar_ticks)

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
        groups = []
        for payload in payloads:
            file_path = payload.get("file_path")
            cmap = payload.get("cmap")
            if file_path:
                try:
                    group = self._add_kind_views_for_header(Path(file_path), cmap_override=cmap)
                    if group:
                        groups.append(group)
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
        if channel_idx < 0 or channel_idx >= len(fds):
            return None
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
        item.set_show_overlay(self._show_overlay_info, self._show_overlay_file)
        item.set_metadata_bar_visible(False if self._show_overlay_info else self._metadata_bar_visible_default())
        item.set_show_colorbar_ticks(self._global_show_colorbar_ticks)
        item.set_parent_window(self)
        if file_key not in self._file_scale_bars:
            self._file_scale_bars[file_key] = item._scale_bar_spec()[0] if item._scale_bar_spec() else None
        item.set_scale_bar_length(self._file_scale_bars.get(file_key))
        item.set_frame_color(self.view.backgroundBrush().color())
        if place:
            self._place_item(item)
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
        kinds = ["topo", "current", "df"]
        if not groups:
            return
        gap_x = 24.0
        gap_y = 24.0
        margin = 20.0
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
        self._push_undo_state()

    def _on_align_selected(self):
        selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        if len(selected) < 2:
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
            return
        kinds = ["topo", "current", "df"]
        columns = []
        for file_path, group in groups.items():
            min_x = min((item.pos().x() for item in group.values()), default=0.0)
            columns.append((min_x, file_path, group))
        columns.sort(key=lambda entry: entry[0])
        margin = 20.0
        gap_x = 24.0
        gap_y = 24.0
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
        items = [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        if not items:
            return
        margin = 20
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
        self._push_undo_state()

    def _on_export_image(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export Canvas",
            "canvas_export.png",
            "PNG Image (*.png);;JPEG Image (*.jpg);;PDF Document (*.pdf)",
        )
        if not path:
            return
        rect = self.scene.itemsBoundingRect()
        if rect.isEmpty():
            QtWidgets.QMessageBox.warning(self, "Export", "No items to export")
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
        self.scene.render(painter, QtCore.QRectF(), rect)
        painter.end()
        if image.save(path):
            self.status_label.setText(f"Exported to {Path(path).name}")
        else:
            QtWidgets.QMessageBox.warning(self, "Export", f"Failed to save {path}")

    def _on_save_canvas(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save canvas", "canvas.json", "JSON Files (*.json)")
        if not path:
            return
        items = [item.to_state() for item in self.scene.items() if isinstance(item, CanvasImageItem)]
        payload = {"version": 1, "items": items}
        try:
            Path(path).write_text(json.dumps(payload, indent=2))
            self.status_label.setText(f"Saved canvas to {Path(path).name}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save canvas", f"Unable to save: {exc}")

    def _on_load_canvas(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load canvas", "", "JSON Files (*.json)")
        if not path:
            return
        try:
            payload = json.loads(Path(path).read_text())
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load canvas", f"Unable to load: {exc}")
            return
        self.scene.clear()
        self._drop_offset = QtCore.QPointF(40.0, 40.0)
        for state in payload.get("items", []):
            file_path = state.get("file_path")
            channel_idx = state.get("channel_index")
            if not file_path or channel_idx is None:
                continue
            item = self._add_view_from_header(Path(file_path), int(channel_idx), cmap_override=state.get("cmap"))
            if item:
                item.apply_state(state)
        self.status_label.setText(f"Loaded canvas from {Path(path).name}")
        self._push_undo_state()

    def _delete_selected(self):
        selected = [i for i in self.scene.selectedItems() if isinstance(i, CanvasImageItem)]
        if not selected:
            return
        for item in selected:
            self.scene.removeItem(item)
        self._selected_item = None
        self._on_selection_changed()
        self._push_undo_state()

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
        items = [i for i in self.scene.items() if isinstance(i, CanvasImageItem)]
        state = []
        for item in items:
            state.append(item.to_state())
        return state

    def _restore_state(self, state):
        self._restoring = True
        self.scene.clear()
        self._drop_offset = QtCore.QPointF(40.0, 40.0)
        for item_state in state:
            file_path = item_state.get("file_path")
            channel_idx = item_state.get("channel_index")
            if not file_path or channel_idx is None:
                continue
            item = self._add_view_from_header(Path(file_path), int(channel_idx), cmap_override=item_state.get("cmap"))
            if item:
                item.apply_state(item_state)
        self._restoring = False

    def _push_undo_state(self):
        if self._restoring:
            return
        state = self._capture_state()
        if self._undo_index >= 0 and self._undo_index < len(self._undo_stack) - 1:
            self._undo_stack = self._undo_stack[: self._undo_index + 1]
        self._undo_stack.append(state)
        self._undo_index = len(self._undo_stack) - 1

    def _undo(self):
        if self._undo_index <= 0:
            return
        self._undo_index -= 1
        self._restore_state(self._undo_stack[self._undo_index])

    def _redo(self):
        if self._undo_index >= len(self._undo_stack) - 1:
            return
        self._undo_index += 1
        self._restore_state(self._undo_stack[self._undo_index])

    def _place_item(self, item: CanvasImageItem):
        item.setPos(self._drop_offset)
        self._drop_offset = QtCore.QPointF(self._drop_offset.x() + 30.0, self._drop_offset.y() + 30.0)












