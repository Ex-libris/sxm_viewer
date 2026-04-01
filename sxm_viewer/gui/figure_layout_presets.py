"""Shared figure-layout presets for publication and slide export.

These helpers keep the common journal/slide size logic in one place so plot
windows can offer the same quick presets and vector-first export actions.
"""
from __future__ import annotations

from dataclasses import dataclass

from .._shared import QtCore, QtGui, QtWidgets, Path, io, matplotlib


@dataclass(frozen=True)
class FigureLayoutPreset:
    """Compact figure sizing/typography preset for plot-style windows."""

    key: str
    label: str
    width_mm: float
    height_mm: float
    font_family: str = "Arial"
    font_scale: float = 1.0
    legend_font_pt: float = 8.0
    line_width: float = 1.5


_FIGURE_PRESETS: tuple[FigureLayoutPreset, ...] = (
    FigureLayoutPreset("interactive", "Interactive", 152.4, 101.6, "sans-serif", 1.0, 8.0, 1.6),
    FigureLayoutPreset("journal_88_square", "Journal 1-col square (88 mm)", 88.0, 88.0, "Arial", 0.78, 6.0, 1.0),
    FigureLayoutPreset("journal_85_square", "Journal 1-col square (85 mm)", 85.0, 85.0, "Arial", 0.76, 5.8, 0.95),
    FigureLayoutPreset("journal_114_square", "Journal 1.5-col square (114 mm)", 114.0, 114.0, "Arial", 0.84, 6.4, 1.05),
    FigureLayoutPreset("journal_174_square", "Journal 2-col square (174 mm)", 174.0, 174.0, "Arial", 0.92, 6.8, 1.15),
    FigureLayoutPreset("slide_square", "Slides square (127 mm)", 127.0, 127.0, "Arial", 1.05, 8.5, 1.8),
)


def iter_figure_layout_presets() -> tuple[FigureLayoutPreset, ...]:
    """Return the ordered set of shared plot presets."""
    return _FIGURE_PRESETS


def get_figure_layout_preset(key: str | None) -> FigureLayoutPreset:
    """Resolve a preset key, falling back to the interactive preset."""
    wanted = str(key or "").strip()
    for preset in _FIGURE_PRESETS:
        if preset.key == wanted:
            return preset
    return _FIGURE_PRESETS[0]


def apply_figure_layout(fig, preset: FigureLayoutPreset) -> None:
    """Apply the physical size of a preset to a Matplotlib figure."""
    width_in = max(0.5, float(preset.width_mm) / 25.4)
    height_in = max(0.5, float(preset.height_mm) / 25.4)
    fig.set_size_inches(width_in, height_in, forward=True)


def preset_pixel_size(widget, preset: FigureLayoutPreset, *, max_fraction: float = 0.72) -> tuple[int, int]:
    """Convert a physical preset size into an on-screen pixel target."""
    dpi_x = dpi_y = 96.0
    screen = None
    try:
        handle = widget.windowHandle() if widget is not None else None
        screen = handle.screen() if handle is not None else None
    except Exception:
        screen = None
    if screen is None:
        try:
            screen = QtWidgets.QApplication.primaryScreen()
        except Exception:
            screen = None
    if screen is not None:
        try:
            dpi_x = float(screen.logicalDotsPerInchX() or dpi_x)
            dpi_y = float(screen.logicalDotsPerInchY() or dpi_y)
        except Exception:
            pass
        try:
            geom = screen.availableGeometry()
            max_w = int(max(240, geom.width() * float(max_fraction)))
            max_h = int(max(240, geom.height() * float(max_fraction)))
        except Exception:
            max_w = max_h = 10_000
    else:
        max_w = max_h = 10_000
    width_px = int(round((float(preset.width_mm) / 25.4) * dpi_x))
    height_px = int(round((float(preset.height_mm) / 25.4) * dpi_y))
    if preset.key != "interactive":
        side = max(width_px, height_px)
        width_px = side
        height_px = side
    width_px = max(220, min(width_px, max_w))
    height_px = max(220, min(height_px, max_h))
    return width_px, height_px


def apply_canvas_widget_preset(canvas, preset: FigureLayoutPreset, width_px: int, height_px: int) -> None:
    """Apply widget-level size constraints so square presets stay square on screen."""
    if canvas is None:
        return
    if preset.key == "interactive":
        try:
            canvas.setMinimumSize(0, 0)
        except Exception:
            pass
        try:
            canvas.setMaximumSize(QtCore.QSize(16777215, 16777215))
        except Exception:
            pass
        try:
            canvas.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        except Exception:
            pass
        return
    try:
        canvas.setMinimumSize(int(width_px), int(height_px))
    except Exception:
        pass
    try:
        canvas.setMaximumSize(QtCore.QSize(int(width_px), int(height_px)))
    except Exception:
        pass
    try:
        canvas.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
    except Exception:
        pass
    try:
        canvas.resize(int(width_px), int(height_px))
    except Exception:
        pass


def copy_figure_to_clipboard(parent, fig, fmt: str = "png", *, dpi: int = 300) -> None:
    """Copy a Matplotlib figure to the clipboard, preserving text in SVG."""
    fmt = str(fmt or "png").lower()
    buf = io.BytesIO()
    if fmt == "svg":
        with matplotlib.rc_context({"svg.fonttype": "none"}):
            fig.savefig(buf, format="svg", bbox_inches="tight", pad_inches=0.02)
        mime = QtCore.QMimeData()
        data = buf.getvalue()
        mime.setData("image/svg+xml", data)
        try:
            mime.setText(data.decode("utf-8"))
        except Exception:
            pass
        QtWidgets.QApplication.clipboard().setMimeData(mime)
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), "Plot copied as SVG", parent)
        return
    fig.savefig(buf, format="png", dpi=int(dpi), bbox_inches="tight", pad_inches=0.02)
    image = QtGui.QImage.fromData(buf.getvalue(), "PNG")
    QtWidgets.QApplication.clipboard().setImage(image)
    QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), f"Plot copied as PNG ({int(dpi)} dpi)", parent)


def save_figure_with_dialog(parent, fig, *, default_stem: str, fmt: str = "png", dpi: int = 300) -> None:
    """Save a Matplotlib figure using a format-aware file dialog."""
    fmt = str(fmt or "png").lower()
    if fmt == "svg":
        filt = "SVG Files (*.svg)"
        default_name = f"{default_stem}.svg"
    elif fmt == "pdf":
        filt = "PDF Files (*.pdf)"
        default_name = f"{default_stem}.pdf"
    else:
        filt = "PNG Files (*.png)"
        default_name = f"{default_stem}.png"
    path, _ = QtWidgets.QFileDialog.getSaveFileName(parent, "Save plot", default_name, filt)
    if not path:
        return
    try:
        if fmt == "svg":
            with matplotlib.rc_context({"svg.fonttype": "none"}):
                fig.savefig(path, format="svg", bbox_inches="tight", pad_inches=0.02)
        elif fmt == "pdf":
            with matplotlib.rc_context({"pdf.fonttype": 42, "ps.fonttype": 42}):
                fig.savefig(path, format="pdf", bbox_inches="tight", pad_inches=0.02)
        else:
            fig.savefig(path, format="png", dpi=int(dpi), bbox_inches="tight", pad_inches=0.02)
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), f"Saved {Path(path).name}", parent)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(parent, "Save plot", str(exc))
