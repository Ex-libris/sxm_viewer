"""Thumbnail UI helpers for SXMGridViewer."""
from __future__ import annotations

import re
import sip

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
from ...config import save_config


def _safe_set_property(widget, name, value):
    try:
        widget.setProperty(name, value)
        return True
    except RuntimeError:
        return False

def _thumb_dimensions(viewer):
    """Return (width, height) for thumbnails preserving 4:3 aspect ratio."""
    w = int(max(64, min(360, getattr(viewer, 'thumb_size_px', 160))))
    h = int(max(48, round(w * 0.75)))
    return w, h


def _resize_thumbnail_scale(viewer, delta_px):
    new_w = int(max(64, min(360, viewer.thumb_size_px + delta_px)))
    if new_w == viewer.thumb_size_px:
        return
    viewer.thumb_size_px = new_w
    viewer.config['thumb_size_px'] = new_w
    save_config(viewer.config)
    viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())


def clear_thumbs(viewer):
    while viewer.thumb_layout.count():
        item = viewer.thumb_layout.takeAt(0); w = item.widget()
        if w: w.setParent(None)
    viewer.thumb_widgets = {}
    viewer._thumb_labels = {}
    viewer._thumb_meta = {}
    viewer._thumb_loaded = set()
    viewer._thumb_inflight = set()
    viewer._thumb_card_height = None


def populate_thumbnails_for_channel(viewer, channel_idx:int):
    viewer.clear_thumbs()
    thumb_w, thumb_h = viewer._thumb_dimensions()
    # Compute number of columns responsively based on available viewport width so the
    # thumbnail grid reflows when the splitter or window is resized.
    try:
        vp = getattr(viewer, '_thumb_viewport', None)
        avail_w = vp.width() if vp is not None else (viewer.thumb_container.width() if hasattr(viewer, 'thumb_container') else 800)
    except Exception:
        avail_w = 800
    # estimate per-card width including margins and label area
    card_w = thumb_w + 24
    max_cols = max(1, min(12, int(avail_w / card_w)))
    row = 0; col = 0
    cmap_name = viewer.thumb_cmap_combo.currentText() or viewer.thumb_cmap
    viewer._thumb_generation += 1
    generation = viewer._thumb_generation
    files_iter = list(viewer.files)
    try:
        viewer.thumb_grid_columns = max_cols
    except Exception:
        viewer.thumb_grid_columns = 1

    filt = (viewer.thumb_filter_combo.currentText() if hasattr(viewer, 'thumb_filter_combo') else 'All')
    if filt and filt != 'All':
        matrix_set = set(getattr(viewer, 'files_with_matrix', set()) or [])
        def include(path_str):
            tag = (viewer.tags.get(path_str, {}) or {}).get('tag', None)
            if filt == 'Constant height':
                return tag == 'constant-height'
            if filt == 'Constant current':
                return tag == 'constant-current'
            if filt == 'Untagged':
                return tag is None
            if filt == 'Matrix datasets':
                return path_str in matrix_set
            return True
        files_iter = [t for t in files_iter if include(str(t))]

    sort_mode = (viewer.thumb_sort_combo.currentText() if hasattr(viewer, 'thumb_sort_combo') else 'Name (A?Z)')
    if sort_mode.startswith('Name'):
        def _natural_key(name: str):
            parts = re.split(r"(\\d+)", name)
            key = []
            for part in parts:
                if part.isdigit():
                    try:
                        key.append(int(part))
                    except Exception:
                        key.append(part)
                else:
                    key.append(part.lower())
            return key
        files_iter.sort(key=lambda p: _natural_key(Path(p).name))
    elif 'Date (new' in sort_mode or 'Date (old' in sort_mode:
        rev = ('new' in sort_mode)
        def sort_key_date(p):
            hdr = viewer.headers.get(str(p), (None, None))[0]
            return viewer._parse_header_datetime(hdr, path=p)
        files_iter.sort(key=sort_key_date, reverse=rev)
    elif sort_mode.startswith('Tag'):
        order = {'constant-height': 0, 'constant-current': 1, None: 2}
        files_iter.sort(key=lambda p: (order.get((viewer.tags.get(str(p), {}) or {}).get('tag', None), 2), Path(p).name.lower()))

    viewer.current_thumb_files = [str(f) for f in files_iter]
    viewer._thumb_meta = {}
    # approximate per-card height: thumb + label + padding
    viewer._thumb_card_height = thumb_h + 48
    for i, t in enumerate(files_iter):
        key = str(t)
        if key not in viewer.headers:
            continue
        header, fds = viewer.headers[key]
        lbl = QtWidgets.QLabel()
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setProperty("file_path", key)
        lbl.setProperty("channel_index", int(channel_idx))
        lbl.setProperty("spec_markers", [])
        lbl.setProperty("thumb_dims", (thumb_w, thumb_h))
        lbl.setProperty("drag_start", None)
        lbl.setProperty("dragging", False)
        placeholder = QtGui.QPixmap(thumb_w, thumb_h)
        placeholder.fill(QtGui.QColor('#0b0b12'))
        lbl.setPixmap(placeholder)
        lbl.setMouseTracking(True)
        lbl.mousePressEvent = viewer._make_thumb_press_handler(lbl)
        lbl.mouseReleaseEvent = viewer._make_thumb_release_handler(lbl)
        lbl.mouseMoveEvent = viewer._make_thumb_move_handler(lbl)
        lbl.mouseDoubleClickEvent = viewer._make_thumb_double_handler(lbl)
        lbl.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        lbl.customContextMenuRequested.connect(lambda pos, lb=lbl: viewer._on_thumb_context_menu(lb, pos))
        vbox = QtWidgets.QVBoxLayout(); vbox.setContentsMargins(0,0,0,0); vbox.setSpacing(2)
        card = QtWidgets.QFrame(); card.setFrameShape(QtWidgets.QFrame.StyledPanel); card.setLineWidth(0)
        card_layout = QtWidgets.QVBoxLayout(card); card_layout.setContentsMargins(4,4,4,4); card_layout.setSpacing(4)
        vbox.addWidget(lbl)
        cap = QtWidgets.QLabel(Path(t).name); cap.setAlignment(QtCore.Qt.AlignCenter); cap.setMaximumHeight(18)
        cap.setFont(QtGui.QFont("Segoe UI", 9)); vbox.addWidget(cap)
        cap.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        cap.customContextMenuRequested.connect(lambda pos, lb=lbl: viewer._on_thumb_context_menu(lb, pos))
        card_layout.addLayout(vbox)
        viewer.thumb_layout.addWidget(card, row, col)
        viewer.thumb_widgets[key] = card
        viewer._thumb_labels[key] = lbl
        try:
            if key in getattr(viewer, 'thumb_multi_select', set()):
                card.setStyleSheet("QFrame { border: 2px solid #a36bff; border-radius: 10px; background-color: rgba(163,107,255,40); }")
            elif key == str(getattr(viewer, 'selected_file_for_thumbs', None)):
                card.setStyleSheet("QFrame { border: 2px solid #5f8dd3; border-radius: 10px; background-color: rgba(95,141,211,40); }")
            else:
                card.setStyleSheet("QFrame { border: 1px solid rgba(255,255,255,30); border-radius: 10px; background-color: transparent; }")
        except Exception:
            pass

        if fds and 0 <= channel_idx < len(fds):
            fd = fds[channel_idx]
            base_pix = None
            data_key = None
            try:
                data_key = viewer._thumbnail_data_key(key, channel_idx, fd, thumb_w, thumb_h)
            except Exception:
                data_key = None
            if data_key:
                base_pix = viewer.thumb_cache.get((data_key, cmap_name))
            if base_pix is not None:
                pix = base_pix.copy()
                crop_info = None
                try:
                    with viewer._thumb_data_lock:
                        crop_info = viewer._thumb_crop_cache.get(data_key)
                except Exception:
                    crop_info = None
                markers = viewer._decorate_thumbnail_pixmap(pix, key, channel_idx, header, fds, thumb_crop=crop_info)
                lbl.setPixmap(pix)
                lbl.setProperty("spec_markers", markers)
                try:
                    lbl.setProperty("thumb_crop", crop_info)
                except Exception:
                    pass
                viewer._thumb_loaded.add(key)
            else:
                lbl.setProperty("spec_markers", [])
            viewer._thumb_meta[key] = (channel_idx, header, fd, thumb_w, thumb_h, cmap_name, generation)
        else:
            blank = QtGui.QPixmap(thumb_w, thumb_h)
            blank.fill(QtGui.QColor('black'))
            lbl.setPixmap(blank)
            lbl.setProperty("spec_markers", [])

        col += 1
        if col >= max_cols:
            col = 0; row += 1
    # kick off initial batch for visible thumbs
    try:
        viewer._request_visible_thumbs()
    except Exception:
        pass
    viewer._refresh_frame_map_pixmaps()

def on_thumb_sort_changed(viewer, idx):
    try:
        viewer.config['thumb_sort'] = viewer.thumb_sort_combo.currentText(); save_config(viewer.config)
    except Exception:
        pass
    viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())


def on_thumb_filter_changed(viewer, idx):
    try:
        viewer.config['thumb_filter'] = viewer.thumb_filter_combo.currentText(); save_config(viewer.config)
    except Exception:
        pass
    viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())


def _thumbnail_pixmap_for_file(viewer, file_key, channel_idx, width, height, cmap_name):
    if not file_key:
        return None
    header, fds = viewer.headers.get(str(file_key), (None, None))
    if not header or not fds:
        return None
    if channel_idx < 0 or channel_idx >= len(fds):
        if not fds:
            return None
        channel_idx = min(max(channel_idx, 0), len(fds) - 1)
    fd = fds[channel_idx]
    data_key = None
    try:
        data_key = viewer._thumbnail_data_key(str(file_key), channel_idx, fd, width, height)
    except Exception:
        data_key = None
    if data_key is not None:
        cache_key = ('frame', data_key, cmap_name)
        pix = viewer._frame_real_pixmap_cache.get(cache_key)
        if pix is not None:
            return pix
        base_pix = viewer.thumb_cache.get((data_key, cmap_name))
        if base_pix is not None:
            viewer._frame_real_pixmap_cache[cache_key] = base_pix
            return base_pix
        prefix = data_key[:4]
        for key, cmap in list(viewer.thumb_cache.keys()):
            if cmap != cmap_name:
                continue
            if key[:4] == prefix:
                base_pix = viewer.thumb_cache.get((key, cmap_name))
                if base_pix is not None:
                    viewer._frame_real_pixmap_cache[cache_key] = base_pix
                    return base_pix
    try:
        data_key, arr = viewer._get_thumbnail_array(str(file_key), channel_idx, header, fd, width, height)
    except Exception:
        return None
    cache_key = ('frame', data_key, cmap_name)
    pix = viewer._frame_real_pixmap_cache.get(cache_key)
    if pix is None:
        try:
            qimg = array_to_qimage(arr, cmap_name=cmap_name)
            pix = QtGui.QPixmap.fromImage(qimg)
            viewer._frame_real_pixmap_cache[cache_key] = pix
        except Exception:
            pix = None
    return pix


def _refresh_thumb_selection_styles(viewer):
    sel = str(getattr(viewer, 'selected_file_for_thumbs', '') or '')
    multi = getattr(viewer, 'thumb_multi_select', set())
    for fp, w in list(getattr(viewer, 'thumb_widgets', {}).items()):
        try:
            if str(fp) in multi:
                w.setStyleSheet("QFrame { border: 2px solid #a36bff; border-radius: 10px; background-color: rgba(163,107,255,40); }")
            elif str(fp) == sel and sel:
                w.setStyleSheet("QFrame { border: 2px solid #5f8dd3; border-radius: 10px; background-color: rgba(95,141,211,40); }")
            else:
                w.setStyleSheet("QFrame { border: 1px solid rgba(255,255,255,30); border-radius: 10px; background-color: transparent; }")
        except Exception:
            continue


def _handle_thumb_click(viewer, label_widget, event):
    if event.button() != QtCore.Qt.LeftButton:
        return
    if viewer._handle_spec_marker_click(label_widget, event):
        return
    if getattr(viewer, '_highlighted_spec', None):
        try:
            viewer._highlight_spectrum_entry(None)
        except Exception:
            pass
    fp = label_widget.property("file_path")
    ch_idx = int(label_widget.property("channel_index"))
    mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
    
    if mods & QtCore.Qt.ShiftModifier:
        if not hasattr(viewer, 'thumb_multi_select') or viewer.thumb_multi_select is None:
            viewer.thumb_multi_select = set()
        
        anchor = getattr(viewer, 'last_thumb_anchor', None)
        if not anchor and getattr(viewer, 'selected_file_for_thumbs', None):
            anchor = str(viewer.selected_file_for_thumbs)
        if not anchor:
            anchor = str(fp)
            
        current_files = getattr(viewer, 'current_thumb_files', [])
        if str(anchor) in current_files and str(fp) in current_files:
            idx1 = current_files.index(str(anchor))
            idx2 = current_files.index(str(fp))
            start, end = min(idx1, idx2), max(idx1, idx2)
            subset = current_files[start : end+1]
            if mods & QtCore.Qt.ControlModifier:
                viewer.thumb_multi_select.update(subset)
            else:
                viewer.thumb_multi_select = set(subset)
        else:
            viewer.thumb_multi_select.add(str(fp))
        viewer._refresh_thumb_selection_styles()
        return

    if mods & QtCore.Qt.ControlModifier:
        viewer._toggle_thumb_multi_selection(fp)
        viewer.last_thumb_anchor = str(fp)
        return

    viewer._clear_thumb_multi_selection(update_styles=False)
    viewer.on_thumbnail_clicked(fp, ch_idx)
    viewer.last_thumb_anchor = str(fp)
    try:
        if not viewer.show_spectra:
            return
        entries = viewer.spectros_by_image.get(str(fp), [])
        if not entries:
            return
        matrix_specs = [s for s in entries if s.get('matrix_index') is not None and 'matrix' in Path(s.get('path','')).name.lower()]
        if matrix_specs:
            viewer._open_matrix_explorer_for_file(str(fp))
            return
        viewer._open_spectro_summary_for_file(fp, show_mode="single", quiet=True)
    except Exception:
        pass


def _make_thumb_press_handler(viewer, label_widget):
    def handler(event):
        if event.button() != QtCore.Qt.LeftButton:
            return
        if not _safe_set_property(label_widget, "drag_start", event.pos()):
            return
        _safe_set_property(label_widget, "dragging", False)
        QtWidgets.QLabel.mousePressEvent(label_widget, event)
    return handler


def _make_thumb_release_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        dragging = bool(label_widget.property("dragging"))
        if not _safe_set_property(label_widget, "drag_start", None):
            return
        _safe_set_property(label_widget, "dragging", False)
        if dragging:
            return
        _handle_thumb_click(viewer, label_widget, event)
    return handler


def _make_thumb_move_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        dragging = bool(label_widget.property("dragging"))
        start = label_widget.property("drag_start")
        if start is not None and event.buttons() & QtCore.Qt.LeftButton and not dragging:
            if (event.pos() - start).manhattanLength() >= 10:
                fp = label_widget.property("file_path")
                ch_idx = int(label_widget.property("channel_index"))
                cmap = viewer.preview_cmap_combo.currentText() or viewer.preview_cmap
                selected = set(getattr(viewer, "thumb_multi_select", set()) or [])
                # Always include the current drag origin in the payload and respect any existing selection,
                # so users don't need to keep modifiers pressed while starting the drag.
                if selected:
                    selected.add(str(fp))
                    payload = {
                        "items": sorted(selected),
                        "channel_index": ch_idx,
                        "cmap": cmap,
                    }
                else:
                    payload = {
                        "file_path": fp,
                        "channel_index": ch_idx,
                        "cmap": cmap,
                    }
                if hasattr(viewer, "_ensure_canvas_for_drag"):
                    viewer._ensure_canvas_for_drag()
                drag_parent = label_widget
                try:
                    if isinstance(viewer, QtWidgets.QWidget):
                        drag_parent = viewer
                except Exception:
                    pass
                if sip.isdeleted(drag_parent):
                    drag_parent = label_widget
                drag = QtGui.QDrag(drag_parent)
                mime = QtCore.QMimeData()
                try:
                    mime.setData("application/x-sxm-view", json.dumps(payload).encode("utf-8"))
                except Exception:
                    pass
                pix = label_widget.pixmap()
                if pix is not None:
                    mime.setImageData(pix.toImage())
                    drag.setPixmap(pix)
                drag.setMimeData(mime)
                if not _safe_set_property(label_widget, "dragging", True):
                    return
                drag.exec_(QtCore.Qt.CopyAction)
                _safe_set_property(label_widget, "dragging", False)
                _safe_set_property(label_widget, "drag_start", None)
                return
        if not viewer._handle_spec_hover(label_widget, event):
            QtWidgets.QLabel.mouseMoveEvent(label_widget, event)
    return handler


def _make_thumb_double_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        if event.button() != QtCore.Qt.LeftButton:
            return
        fp = label_widget.property("file_path")
        ch_idx = int(label_widget.property("channel_index") or 0)
        try:
            viewer.on_thumbnail_double_clicked(fp, ch_idx)
        except Exception:
            pass
    return handler

# ---------- thumbnail clicked -> preview + inspector populate ----------

def _toggle_thumb_multi_selection(viewer, file_path):
    path = str(file_path)
    if not hasattr(viewer, 'thumb_multi_select'):
        viewer.thumb_multi_select = set()
    if path in viewer.thumb_multi_select:
        viewer.thumb_multi_select.remove(path)
    else:
        viewer.thumb_multi_select.add(path)
    viewer._refresh_thumb_selection_styles()


def _clear_thumb_multi_selection(viewer, update_styles=True):
    viewer.thumb_multi_select = set()
    if update_styles:
        viewer._refresh_thumb_selection_styles()


def on_thumb_cmap_changed(viewer, idx):
    viewer.thumb_cmap = viewer.thumb_cmap_combo.currentText(); viewer.config['thumbnail_cmap'] = viewer.thumb_cmap; save_config(viewer.config)
    viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
__all__ = [
    "_thumb_dimensions",
    "_resize_thumbnail_scale",
    "clear_thumbs",
    "populate_thumbnails_for_channel",
    "on_thumb_sort_changed",
    "on_thumb_filter_changed",
    "_thumbnail_pixmap_for_file",
    "_refresh_thumb_selection_styles",
    "_handle_thumb_click",
    "_make_thumb_press_handler",
    "_make_thumb_release_handler",
    "_make_thumb_move_handler",
    "_make_thumb_double_handler",
    "_toggle_thumb_multi_selection",
    "_clear_thumb_multi_selection",
    "on_thumb_cmap_changed",
]
