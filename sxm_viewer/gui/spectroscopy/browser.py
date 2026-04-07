"""Spectroscopy browser helpers for SXMGridViewer."""
from __future__ import annotations

from ..._shared import (
    QtCore,
    QtGui,
    QtWidgets,
    QIcon,
    QPixmap,
    QImage,
    QPainter,
    QPen,
    QBrush,
    FigureCanvas,
    Figure,
    Line2D,
    colormaps,
    np,
    Path,
    defaultdict,
    OrderedDict,
    datetime,
    hashlib,
    itertools,
    io,
    json,
    math,
    os,
    sys,
    threading,
    _scipy_ndimage,
    log_status,
    matplotlib,
)
from PyQt5.QtWidgets import QLabel, QListWidget, QListWidgetItem

def _update_spectro_stats_label(viewer, stats=None):
    if not hasattr(viewer, 'spectro_stats_label'):
        return
    thumb_markers = bool(getattr(viewer, 'show_spectra', True))
    preview_markers = bool(getattr(viewer, 'show_preview_spectra', thumb_markers))
    miniatures = bool(getattr(viewer, 'show_spectro_miniatures', False))
    shared_repeats = bool(getattr(viewer, "spectro_share_overlapping_repeats", False))
    mode_text = (
        f"Thumbnail markers {'On' if thumb_markers else 'Off'} | "
        f"Preview {'On' if preview_markers else 'Off'} | "
        f"Miniatures {'On' if miniatures else 'Off'} | "
        f"Assignment {'Shared repeats' if shared_repeats else 'Nearest image'}"
    )
    assignment_tip = (
        "Assignment mode: shared repeats. Spectra inside several near-identical overlapping scans appear on all of those repeat images."
        if shared_repeats
        else "Assignment mode: nearest image. Each spectrum is attached to one best-matching image only."
    )
    if getattr(viewer, '_spectros_loading', False):
        viewer.spectro_stats_label.setText(f"Spectroscopy loading...\n{mode_text}")
        viewer.spectro_stats_label.setToolTip(
            "Spectroscopy files are being scanned now. "
            "Thumbnail markers draw clickable points on image thumbnails. "
            "Preview markers draw the same points in the preview panel. "
            "Miniatures add separate spectroscopy cards into the thumbnail stream. "
            + assignment_tip
        )
        return
    if getattr(viewer, '_spectros_pending', False) and not getattr(viewer, '_spectros_loaded', False):
        viewer.spectro_stats_label.setText(f"Spectroscopy pending load\n{mode_text}")
        viewer.spectro_stats_label.setToolTip(
            "Spectroscopy scanning is deferred until a browser or visible spectroscopy mode needs it. "
            "Thumbnail markers draw clickable points on image thumbnails. "
            "Preview markers draw the same points in the preview panel. "
            "Miniatures add separate spectroscopy cards into the thumbnail stream. "
            + assignment_tip
        )
        return
    total = len(getattr(viewer, 'spectros', []) or [])
    single_count = sum(1 for s in getattr(viewer, 'spectros', []) if s.get('matrix_index') is None)
    xy_stack_count = len({
        str(s.get("xy_stack_key"))
        for s in (getattr(viewer, "spectros", []) or [])
        if s.get("xy_stack_key") and int(s.get("xy_stack_count") or 0) > 1
    })
    if stats:
        total = stats.get('total_specs', total)
        single_count = stats.get('single_entries', single_count)
    matrix_datasets = getattr(viewer, 'matrix_datasets', {}) or {}
    matrix_count = len(matrix_datasets)
    sample_ds = next(iter(matrix_datasets.values()), None)
    matrix_desc = ""
    if sample_ds:
        matrix_desc = f" ({sample_ds.cols}x{sample_ds.rows})"
    elif matrix_count == 0:
        matrix_desc = ""
    viewer.spectro_stats_label.setText(
        f"Spectra {total} | Single {single_count} | XY stacks {xy_stack_count} | Matrix {matrix_count}{matrix_desc}\n{mode_text}"
    )
    viewer.spectro_stats_label.setToolTip(
        f"Loaded spectroscopy entries: {total}. Single traces: {single_count}. "
        f"Same-XY stacks: {xy_stack_count}. "
        f"Matrix datasets: {matrix_count}{matrix_desc}. "
        "Thumbnail markers draw clickable points on image thumbnails. "
        "Preview markers draw the same points in the preview panel. "
        "Miniatures add separate spectroscopy cards into the thumbnail stream. "
        + assignment_tip
    )


def _ensure_spectro_dock(viewer):
    if viewer.spectro_dock:
        return
    dock = QtWidgets.QDockWidget("Spectro Browser", viewer)
    dock.setFloating(True)
    container = QtWidgets.QWidget(dock)
    v = QtWidgets.QVBoxLayout(container); v.setContentsMargins(6,6,6,6); v.setSpacing(6)
    viewer.spectro_search = QtWidgets.QLineEdit()
    viewer.spectro_search.setPlaceholderText("Search spectra (file/pos)")
    v.addWidget(viewer.spectro_search)
    viewer.spectro_list = QListWidget()
    v.addWidget(viewer.spectro_list, 1)
    viewer.spectro_preview_lbl = QLabel("Select a spectroscopy")
    viewer.spectro_preview_lbl.setAlignment(QtCore.Qt.AlignCenter)
    viewer.spectro_preview_lbl.setMinimumHeight(120)
    viewer.spectro_preview_lbl.setStyleSheet("QLabel { color: #999; }")
    v.addWidget(viewer.spectro_preview_lbl)
    container.setLayout(v)
    dock.setWidget(container)
    viewer.spectro_dock = dock
    viewer.spectro_search.textChanged.connect(viewer._filter_spectro_browser)
    viewer.spectro_list.currentItemChanged.connect(viewer._on_spectro_browser_selection)


def open_spectro_browser(viewer, entries=None):
    viewer._ensure_spectro_dock()
    if entries is None:
        entries = list(viewer.spectros or [])
    viewer._spectro_browser_entries = list(entries)
    viewer._filter_spectro_browser()
    viewer.spectro_dock.show()
    viewer.spectro_dock.raise_()


def _filter_spectro_browser(viewer):
    if not hasattr(viewer, 'spectro_list'):
        return
    txt = viewer.spectro_search.text().strip().lower() if hasattr(viewer, 'spectro_search') else ''
    viewer.spectro_list.clear()
    for idx, s in enumerate(viewer._spectro_browser_entries):
        name = Path(s.get('path','')).name.lower()
        pos = ""
        try:
            if s.get('x') is not None and s.get('y') is not None:
                pos = f"{float(s.get('x')):.1f}/{float(s.get('y')):.1f}"
        except Exception:
            pos = ""
        stack = str(s.get("xy_stack_display") or "").strip()
        stack_suffix = f" [{stack}]" if stack else ""
        label = f"{idx+1}. {name} {pos}{stack_suffix}"
        if txt and txt not in label.lower():
            continue
        item = QListWidgetItem(label)
        item.setData(QtCore.Qt.UserRole, s)
        viewer.spectro_list.addItem(item)


def _on_spectro_browser_selection(viewer, current, _prev):
    if not current:
        if hasattr(viewer, '_highlight_spectrum_entry'):
            viewer._highlight_spectrum_entry(None)
        return
    spec = current.data(QtCore.Qt.UserRole)
    if spec is None:
        if hasattr(viewer, '_highlight_spectrum_entry'):
            viewer._highlight_spectrum_entry(None)
        return
    try:
        x = spec.get('x'); y = spec.get('y')
        lines = [Path(spec.get('path','')).name, f"({x},{y})"]
        summary = str(spec.get("xy_stack_summary") or "").strip()
        if summary:
            lines.append(summary)
        viewer.spectro_preview_lbl.setText("\n".join(lines))
    except Exception:
        viewer.spectro_preview_lbl.setText(Path(spec.get('path','')).name)
    try:
        image_key = spec.get('image_key')
        shared_keys = [str(key) for key in (spec.get("shared_image_keys") or []) if key]
        current_preview = str(viewer.last_preview[0]) if getattr(viewer, "last_preview", None) else ""
        if current_preview and current_preview in shared_keys:
            image_key = current_preview
        elif not image_key and shared_keys:
            image_key = shared_keys[0]
        if image_key and image_key in viewer._thumb_labels:
            viewer.selected_file_for_thumbs = image_key
            viewer._refresh_thumb_selection_styles()
    except Exception:
        pass
    try:
        if hasattr(viewer, '_highlight_spectrum_entry'):
            viewer._highlight_spectrum_entry(spec)
    except Exception:
        pass
    try:
        if hasattr(viewer, '_show_spectro_popup'):
            viewer._show_spectro_popup(spec)
    except Exception:
        pass
__all__ = [
    "_update_spectro_stats_label",
    "_ensure_spectro_dock",
    "open_spectro_browser",
    "_filter_spectro_browser",
    "_on_spectro_browser_selection",
]




