"""Spectroscopy-related helpers for the main window."""
from __future__ import annotations

from pathlib import Path
from PyQt5 import QtCore, QtWidgets

from .spectroscopy import browser as spectro_browser
from .spectroscopy import overlays as spectro_overlays


def header_extent(viewer, header):
    """
    Return extent [x0, x1, y1, y0] in same convention used elsewhere.
    Fallback to unit square if header keys are missing.
    """
    try:
        # Prefer explicit scan range/center keys; be permissive with key names.
        xr = header.get("XScanRange", header.get("XRange", header.get("ScanRange", 0.0)))
        yr = header.get("YScanRange", header.get("YRange", header.get("ScanRange", 0.0)))
        x_range = float(xr or 0.0)
        y_range = float(yr or 0.0)
        cx_keys = ["xCenter", "XCenter", "XOffset", "OffsetX", "XPosition", "XPos"]
        cy_keys = ["yCenter", "YCenter", "YOffset", "OffsetY", "YPosition", "YPos"]
        x_center = 0.0
        y_center = 0.0
        for k in cx_keys:
            if k in header and header.get(k) not in (None, ""):
                x_center = float(header.get(k))
                break
        for k in cy_keys:
            if k in header and header.get(k) not in (None, ""):
                y_center = float(header.get(k))
                break
        if x_range == 0.0 or y_range == 0.0:
            # fall back to simple unit square with correct orientation
            return [0.0, 1.0, 0.0, 1.0]
        if x_center == 0.0 and y_center == 0.0:
            # assume centered scan if center missing
            x_center = 0.5 * x_range
            y_center = 0.5 * y_range
        x0 = x_center - 0.5 * x_range
        x1 = x_center + 0.5 * x_range
        y0 = y_center - 0.5 * y_range
        y1 = y_center + 0.5 * y_range
        return [x0, x1, y1, y0]
    except Exception:
        return [0.0, 1.0, 0.0, 1.0]


def display_extent(viewer, extent, header=None):
    if not extent:
        return extent
    if not getattr(viewer, "relative_axes", False):
        return extent
    try:
        x0, x1, y1, y0 = extent
        xr = abs(float(x1) - float(x0))
        yr = abs(float(y0) - float(y1))
        if abs(xr) <= 1e-12 or abs(yr) <= 1e-12:
            if header:
                xr = header.get("XScanRange", header.get("XRange"))
                yr = header.get("YScanRange", header.get("YRange"))
        xr = float(xr)
        yr = float(yr)
        if xr <= 0 or yr <= 0:
            xr = max(xr, 1.0)
            yr = max(yr, 1.0)
        return [0.0, xr, 0.0, yr]
    except Exception:
        return extent


def spectros_near_thumb_pos(viewer, file_key: str, header: dict, thumb_pos_px: QtCore.QPoint, thumb_dims):
    return spectro_overlays._spectros_near_thumb_pos(viewer, file_key, header, thumb_pos_px, thumb_dims)


def open_single_spectro_popup(viewer, spectro):
    """Hook to open an existing single-spectroscopy popup. Minimal fallback: log."""
    try:
        # Prefer the main spectroscopy popup handler (matrix or single).
        if hasattr(viewer, "_open_spectroscopy_popup"):
            viewer._open_spectroscopy_popup(spectro)
        elif hasattr(viewer, "_show_spectro_popup"):
            viewer._show_spectro_popup(spectro)
        else:
            QtWidgets.QMessageBox.information(
                viewer, "Spectro", f"Spectroscopy at {spectro.get('x')}/{spectro.get('y')}"
            )
    except Exception:
        pass


def open_spectro_summary_for_file(viewer, file_key, show_mode="single", quiet=False):
    entries = viewer.spectros_by_image.get(str(file_key), []) or []
    if show_mode == "single":
        entries = [s for s in entries if s.get("matrix_index") is None]
    elif show_mode == "matrix":
        entries = [s for s in entries if s.get("matrix_index") is not None]
    if not entries:
        if not quiet:
            QtWidgets.QMessageBox.information(viewer, "Spectroscopy", "No spectroscopies found for this file.")
        return

    # Use the modern SpectroscopyCompareDialog (table view) instead of the old summary
    from .detail_panels import SpectroscopyCompareDialog
    dlg = SpectroscopyCompareDialog(entries, parent=viewer)
    dlg.setWindowTitle(f"Spectroscopy: {Path(file_key).name}")
    try:
        dlg.setWindowModality(QtCore.Qt.NonModal)
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.move(viewer._next_popup_pos())
    except Exception:
        pass
    dlg.show()
    if hasattr(viewer, '_popup_refs'):
        viewer._popup_refs.append(dlg)
        try:
            dlg.finished.connect(lambda _: viewer._popup_refs.remove(dlg) if dlg in viewer._popup_refs else None)
        except Exception:
            pass
    controller = getattr(viewer, "quick_crop_controller", None)
    if controller:
        try:
            dlg.finished.connect(lambda _=None, c=controller: c.update_popup_actions())
        except Exception:
            pass
        controller.update_popup_actions()


def ensure_spectro_dock(viewer):
    return spectro_browser._ensure_spectro_dock(viewer)


def open_spectro_browser(viewer, entries=None):
    return spectro_browser.open_spectro_browser(viewer, entries=entries)


def filter_spectro_browser(viewer):
    return spectro_browser._filter_spectro_browser(viewer)


def on_spectro_browser_selection(viewer, current, _prev):
    return spectro_browser._on_spectro_browser_selection(viewer, current, _prev)


def update_spectro_stats_label(viewer, stats=None):
    return spectro_browser._update_spectro_stats_label(viewer, stats=stats)
