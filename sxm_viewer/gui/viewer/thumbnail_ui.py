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
from ..palettes import get_color_cycle, DEFAULT_COLOR_CYCLE


def _safe_set_property(widget, name, value):
    try:
        widget.setProperty(name, value)
        return True
    except RuntimeError:
        return False


def _thumbnail_channel_for_key(viewer, file_key, default_channel_idx):
    try:
        key = str(file_key)
    except Exception:
        key = file_key
    try:
        processed = getattr(viewer, "_processed_views", {}) or {}
        payload = processed.get(key)
        if payload and "channel_idx" in payload and bool(payload.get("lock_channel", True)):
            return int(payload.get("channel_idx"))
    except Exception:
        pass
    try:
        return int(default_channel_idx)
    except Exception:
        return 0


def _thumbnail_cmap_for_key(viewer, file_key, channel_idx, default_cmap):
    try:
        return viewer._thumbnail_cmap_override(file_key, channel_idx, default_cmap)
    except Exception:
        return default_cmap


def _thumbnail_clim_for_key(viewer, file_key, channel_idx, default_clim=None):
    try:
        return viewer._thumbnail_clim_override(file_key, channel_idx, default_clim)
    except Exception:
        return default_clim


def _thumb_render_cache_key(data_key, cmap_name, clim):
    return (data_key, cmap_name, clim)


def _ensure_thumb_click_timer(viewer):
    timer = getattr(viewer, "_thumb_click_timer", None)
    if timer is not None:
        return timer
    parent = viewer if isinstance(viewer, QtCore.QObject) else None
    timer = QtCore.QTimer(parent)
    timer.setSingleShot(True)
    timer.timeout.connect(lambda v=viewer: _flush_pending_thumb_click(v))
    viewer._thumb_click_timer = timer
    return timer


def _cancel_pending_thumb_click(viewer):
    timer = getattr(viewer, "_thumb_click_timer", None)
    if timer is not None:
        try:
            timer.stop()
        except Exception:
            pass
    viewer._pending_thumb_click = None


def _run_plain_thumb_click(viewer, fp, ch_idx):
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


def _flush_pending_thumb_click(viewer):
    payload = getattr(viewer, "_pending_thumb_click", None)
    viewer._pending_thumb_click = None
    if not payload:
        return
    fp = payload.get("file_path")
    try:
        ch_idx = int(payload.get("channel_index") or 0)
    except Exception:
        ch_idx = 0
    _run_plain_thumb_click(viewer, fp, ch_idx)


def _schedule_plain_thumb_click(viewer, label_widget, fp, ch_idx):
    _cancel_pending_thumb_click(viewer)
    viewer._clear_thumb_multi_selection(update_styles=False)
    viewer.selected_file_for_thumbs = str(fp)
    try:
        viewer._refresh_thumb_selection_styles()
    except Exception:
        pass
    viewer.last_thumb_anchor = str(fp)
    viewer._pending_thumb_click = {
        "file_path": fp,
        "channel_index": int(ch_idx),
        "label_widget": label_widget,
    }
    interval = max(0, int(QtWidgets.QApplication.doubleClickInterval()))
    _ensure_thumb_click_timer(viewer).start(interval)

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
    viewer.spectro_thumb_widgets = {}
    viewer._spectro_thumb_labels = {}
    viewer.current_spectro_thumb_files = []
    viewer.current_thumbnail_entries = []
    viewer.current_thumbnail_kind_by_key = {}
    viewer._thumb_meta = {}
    viewer._thumb_loaded = set()
    viewer._thumb_inflight = set()
    viewer._thumb_card_height = None


def _spectro_thumb_selection_key(viewer, path):
    try:
        return str(Path(path).resolve())
    except Exception:
        return str(path)


def _spectro_available_channels(spec):
    channels = list((spec.get("channels") or {}).keys())
    if not channels:
        channels = list(spec.get("available_channels") or [])
    return [str(ch) for ch in channels if str(ch).strip()]


def _spectro_display_channel(viewer, spec):
    path = _spectro_thumb_selection_key(viewer, spec.get("path", ""))
    override = getattr(viewer, "spectro_thumb_channel_by_path", {}).get(path, "")
    available = set(_spectro_available_channels(spec))
    if override and override in available:
        return override
    default = getattr(viewer, "spectro_miniature_default_channel", "")
    if default and default in available:
        return default
    channels = _spectro_available_channels(spec)
    return channels[0] if channels else ""


def _spectro_entry_time(viewer, spec):
    try:
        ts = spec.get("display_time")
        if ts is not None:
            if hasattr(ts, "timestamp"):
                try:
                    return float(ts.timestamp())
                except Exception:
                    pass
            try:
                return float(ts)
            except Exception:
                pass
    except Exception:
        pass
    try:
        ts = spec.get("time")
        if ts is not None:
            if hasattr(ts, "timestamp"):
                try:
                    return float(ts.timestamp())
                except Exception:
                    pass
            try:
                return float(ts)
            except Exception:
                pass
    except Exception:
        pass
    try:
        mt = spec.get("file_mtime")
        if mt is not None:
            if hasattr(mt, "timestamp"):
                try:
                    return float(mt.timestamp())
                except Exception:
                    pass
            try:
                return float(mt)
            except Exception:
                pass
    except Exception:
        pass
    try:
        image_key = str(spec.get("image_key") or "")
        if image_key and getattr(viewer, "image_time_index", None):
            img_ts = viewer.image_time_index.get(image_key)
            if img_ts is not None:
                if hasattr(img_ts, "timestamp"):
                    try:
                        return float(img_ts.timestamp())
                    except Exception:
                        pass
                try:
                    return float(img_ts)
                except Exception:
                    pass
    except Exception:
        pass
    try:
        return float(Path(spec.get("path", "")).stat().st_mtime)
    except Exception:
        return 0.0


def _refresh_spectro_thumb_selection_styles(viewer):
    sel = str(getattr(viewer, "selected_spectro_thumb_file", "") or "")
    multi = getattr(viewer, "spectro_thumb_multi_select", set())
    for fp, w in list(getattr(viewer, "spectro_thumb_widgets", {}).items()):
        try:
            if str(fp) in multi:
                w.setStyleSheet("QFrame { border: 2px solid #ff9c3a; border-radius: 10px; background-color: rgba(255,156,58,42); }")
            elif str(fp) == sel and sel:
                w.setStyleSheet("QFrame { border: 2px solid #5f8dd3; border-radius: 10px; background-color: rgba(95,141,211,36); }")
            else:
                w.setStyleSheet("QFrame { border: 1px solid rgba(255,184,77,170); border-radius: 10px; background-color: rgba(255,184,77,22); }")
        except Exception:
            continue


def _combined_thumbnail_order(viewer):
    entries = list(getattr(viewer, "current_thumbnail_entries", []) or [])
    if entries:
        return [str(item.get("key", "")) for item in entries if str(item.get("key", ""))]
    combined = []
    combined.extend([str(fp) for fp in list(getattr(viewer, "current_thumb_files", []) or []) if str(fp)])
    combined.extend([str(fp) for fp in list(getattr(viewer, "current_spectro_thumb_files", []) or []) if str(fp)])
    return combined


def _set_thumbnail_selection_by_order(viewer, ordered_keys, start_key, end_key, modifiers):
    if start_key not in ordered_keys or end_key not in ordered_keys:
        return False
    idx1 = ordered_keys.index(start_key)
    idx2 = ordered_keys.index(end_key)
    start, end = min(idx1, idx2), max(idx1, idx2)
    subset = ordered_keys[start:end + 1]
    image_sel = set(getattr(viewer, "thumb_multi_select", set()) or [])
    spectro_sel = set(getattr(viewer, "spectro_thumb_multi_select", set()) or [])
    if modifiers & QtCore.Qt.ControlModifier:
        for key in subset:
            if key in image_sel:
                image_sel.remove(key)
            elif key in spectro_sel:
                spectro_sel.remove(key)
            else:
                kind = getattr(viewer, "current_thumbnail_kind_by_key", {}).get(key, "")
                if kind == "spectro":
                    spectro_sel.add(key)
                else:
                    image_sel.add(key)
    else:
        image_sel = set()
        spectro_sel = set()
        for key in subset:
            kind = getattr(viewer, "current_thumbnail_kind_by_key", {}).get(key, "")
            if kind == "spectro":
                spectro_sel.add(key)
            else:
                image_sel.add(key)
    viewer.thumb_multi_select = image_sel
    viewer.spectro_thumb_multi_select = spectro_sel
    viewer._refresh_thumb_selection_styles()
    viewer._refresh_spectro_thumb_selection_styles()
    return True


def _spectroscopy_miniature_pixmap(viewer, spec, width, height, channel_name=None):
    """Render a compact spectral preview for spectroscopy-only thumbnail cards."""
    try:
        cache_key = (
            str(spec.get("path") or ""),
            spec.get("file_mtime") or spec.get("time") or "",
            int(width),
            int(height),
            str(channel_name or ""),
            str(getattr(viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE)),
        )
        pix_cache = getattr(viewer, "_spectro_miniature_cache", None)
        if pix_cache is not None and cache_key in pix_cache:
            return QtGui.QPixmap(pix_cache[cache_key])
    except Exception:
        cache_key = None
        pix_cache = None
    if not (spec.get("channels") or {}) and hasattr(viewer, "hydrate_spectro_entry"):
        try:
            viewer.hydrate_spectro_entry(spec)
        except Exception:
            pass
    pix = QtGui.QPixmap(max(32, int(width)), max(32, int(height)))
    pix.fill(QtGui.QColor("#0d1220"))
    painter = QtGui.QPainter(pix)
    painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
    try:
        rect = QtCore.QRectF(8, 8, pix.width() - 16, pix.height() - 16)
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 35), 1.0))
        painter.setBrush(QtGui.QColor(20, 28, 44))
        painter.drawRoundedRect(rect, 8, 8)

        channels_map = spec.get("channels") or {}
        channel_keys = _spectro_available_channels(spec)
        if channel_name and channel_name in channels_map:
            channel_keys = [channel_name]
        elif channel_keys:
            channel_name = _spectro_display_channel(viewer, spec)
            if channel_name in channels_map:
                channel_keys = [channel_name]
        channels = [(name, channels_map[name]) for name in channel_keys if name in channels_map]
        if not channels:
            painter.setPen(QtGui.QPen(QtGui.QColor(220, 220, 220), 1.0))
            painter.drawText(rect, QtCore.Qt.AlignCenter, "No channels")
            return pix

        axes = spec.get("AxisChoices") or []
        axis_vals = None
        axis_label = spec.get("AxisLabel") or "Axis"
        axis_unit = spec.get("AxisUnit") or ""
        if axes:
            axis_vals = np.asarray((axes[0] or {}).get("values", []), dtype=float)
            axis_label = (axes[0] or {}).get("label") or axis_label
            axis_unit = (axes[0] or {}).get("unit") or axis_unit
        if axis_vals is None or axis_vals.size == 0:
            axis_vals = np.asarray(spec.get("V", []), dtype=float)
        if axis_vals.size < 2:
            painter.setPen(QtGui.QPen(QtGui.QColor(220, 220, 220), 1.0))
            painter.drawText(rect, QtCore.Qt.AlignCenter, "No axis")
            return pix

        margin = 10.0
        plot_rect = rect.adjusted(margin, 20.0, -margin, -20.0)
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 50), 1.0))
        painter.drawRect(plot_rect)

        cycle = get_color_cycle(getattr(viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE))
        title = Path(spec.get("path", "")).name
        painter.setPen(QtGui.QPen(QtGui.QColor(240, 240, 240), 1.0))
        title_font = QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold)
        painter.setFont(title_font)
        painter.drawText(QtCore.QRectF(rect.left() + 4, rect.top() + 2, rect.width() - 8, 14), QtCore.Qt.AlignCenter, title)
        stack_label = str(spec.get("xy_stack_display") or "").strip()
        if stack_label:
            badge_font = QtGui.QFont("Segoe UI", 7, QtGui.QFont.Bold)
            painter.setFont(badge_font)
            metrics = painter.fontMetrics()
            badge_w = max(20, metrics.horizontalAdvance(stack_label) + 10)
            badge_h = max(14, metrics.height() + 2)
            badge_rect = QtCore.QRectF(rect.right() - badge_w - 6, rect.top() + 4, badge_w, badge_h)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 238, 170), 1.0))
            painter.setBrush(QtGui.QColor(40, 30, 18, 220))
            painter.drawRoundedRect(badge_rect, 5, 5)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 228, 120), 1.0))
            painter.drawText(badge_rect, QtCore.Qt.AlignCenter, stack_label)

        x_vals = np.asarray(axis_vals, dtype=float)
        x_min = float(np.nanmin(x_vals))
        x_max = float(np.nanmax(x_vals))
        if not np.isfinite(x_min) or not np.isfinite(x_max) or x_max == x_min:
            x_min, x_max = 0.0, float(max(1, x_vals.size - 1))
            x_vals = np.linspace(x_min, x_max, x_vals.size)

        y_candidates = []
        for _, vals in channels[:1]:
            arr = np.asarray(vals, dtype=float)
            if arr.size:
                y_candidates.append(arr)
        if not y_candidates:
            painter.setPen(QtGui.QPen(QtGui.QColor(220, 220, 220), 1.0))
            painter.drawText(plot_rect, QtCore.Qt.AlignCenter, "No data")
            return pix
        y_all = np.concatenate([np.ravel(arr[np.isfinite(arr)]) for arr in y_candidates if np.isfinite(arr).any()]) if any(np.isfinite(arr).any() for arr in y_candidates) else np.asarray([])
        if y_all.size:
            y_min = float(np.nanmin(y_all))
            y_max = float(np.nanmax(y_all))
        else:
            y_min, y_max = -1.0, 1.0
        if not np.isfinite(y_min) or not np.isfinite(y_max) or y_max == y_min:
            y_min, y_max = -1.0, 1.0
        pad_y = 0.08 * max(1e-9, (y_max - y_min))
        y_min -= pad_y
        y_max += pad_y

        def _map_point(x, y):
            fx = 0.0 if x_max == x_min else (x - x_min) / (x_max - x_min)
            fy = 0.0 if y_max == y_min else (y - y_min) / (y_max - y_min)
            px = plot_rect.left() + fx * plot_rect.width()
            py = plot_rect.bottom() - fy * plot_rect.height()
            return QtCore.QPointF(px, py)

        for idx, (name, vals) in enumerate(channels[:1]):
            arr = np.asarray(vals, dtype=float)
            n = min(arr.size, x_vals.size)
            if n < 2:
                continue
            xs = x_vals[:n]
            ys = arr[:n]
            mask = np.isfinite(xs) & np.isfinite(ys)
            if mask.sum() < 2:
                continue
            color = QtGui.QColor(cycle[idx % len(cycle)])
            painter.setPen(QtGui.QPen(color, 1.7))
            pts = [_map_point(float(x), float(y)) for x, y in zip(xs[mask], ys[mask])]
            painter.drawPolyline(QtGui.QPolygonF(pts))

        footer = axis_label or "Axis"
        if channel_keys:
            footer = f"{footer} · {channel_keys[0]}"
        if axis_unit:
            footer = f"{footer} ({axis_unit})"
        painter.setPen(QtGui.QPen(QtGui.QColor(210, 210, 210, 180), 1.0))
        footer_font = QtGui.QFont("Segoe UI", 7)
        painter.setFont(footer_font)
        painter.drawText(QtCore.QRectF(rect.left() + 6, rect.bottom() - 16, rect.width() - 12, 12), QtCore.Qt.AlignLeft, footer)
    finally:
        painter.end()
    try:
        if pix_cache is not None and cache_key is not None:
            pix_cache[cache_key] = QtGui.QPixmap(pix)
            while len(pix_cache) > 96:
                pix_cache.popitem(last=False)
    except Exception:
        pass
    return pix


def populate_thumbnails_for_channel(viewer, channel_idx:int):
    if getattr(viewer, "show_spectro_miniatures", False) and not getattr(viewer, "_spectros_loaded", False):
        try:
            viewer.ensure_spectros_loaded(refresh=False)
        except Exception:
            pass
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
    cmap_name = getattr(viewer, "thumb_cmap", None) or viewer.thumb_cmap_combo.currentText()
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
    real_files_iter = [str(p) for p in files_iter if not viewer._is_processed_key(str(p))]
    processed_files_iter = [str(p) for p in files_iter if viewer._is_processed_key(str(p))]
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
        real_files_iter.sort(key=lambda p: _natural_key(Path(p).name))
    elif 'Date (new' in sort_mode or 'Date (old' in sort_mode:
        rev = ('new' in sort_mode)
        def sort_key_date(p):
            hdr = viewer.headers.get(str(p), (None, None))[0]
            return viewer._parse_header_datetime(hdr, path=p)
        real_files_iter.sort(key=sort_key_date, reverse=rev)
    elif sort_mode.startswith('Tag'):
        order = {'constant-height': 0, 'constant-current': 1, None: 2}
        real_files_iter.sort(key=lambda p: (order.get((viewer.tags.get(str(p), {}) or {}).get('tag', None), 2), Path(p).name.lower()))

    try:
        files_iter = viewer._ordered_virtual_thumbnail_files(real_files_iter, processed_files_iter)
    except Exception:
        files_iter = list(real_files_iter) + list(processed_files_iter)

    viewer.current_thumb_files = [str(f) for f in files_iter]
    viewer._thumb_meta = {}
    # approximate per-card height: thumb + label + padding
    viewer._thumb_card_height = thumb_h + 48
    if getattr(viewer, "show_spectro_miniatures", False):
        display_entries = []
        for t in files_iter:
            key = str(t)
            if key not in viewer.headers:
                continue
            header, fds = viewer.headers[key]
            try:
                item_time = viewer._parse_header_datetime(header, path=t)
            except Exception:
                item_time = 0.0
            display_entries.append({
                "kind": "image",
                "key": key,
                "path": t,
                "time": item_time,
                "header": header,
                "fds": fds,
            })
        spectro_entries = []
        seen_paths = set()
        for spec in list(getattr(viewer, "spectros", []) or []):
            try:
                key = str(Path(spec.get("path", "")).resolve()).lower()
            except Exception:
                key = str(spec.get("path", "")).lower()
            if not key or key in seen_paths:
                continue
            seen_paths.add(key)
            spectro_entries.append({
                "kind": "spectro",
                "key": str(spec.get("path", "")),
                "path": spec.get("path", ""),
                "time": _spectro_entry_time(viewer, spec),
                "spec": spec,
            })
        display_entries.extend(spectro_entries)
        display_entries.sort(key=lambda item: (
            item.get("time") or 0.0,
            0 if item.get("kind") == "image" else 1,
            Path(item.get("key", "")).name.lower(),
        ))
        viewer.current_thumbnail_entries = list(display_entries)
        viewer.current_thumbnail_kind_by_key = {str(item["key"]): str(item.get("kind") or "") for item in display_entries}
        if spectro_entries:
            viewer.current_spectro_thumb_files = [item["key"] for item in display_entries if item.get("kind") == "spectro"]
        if display_entries:
            viewer.current_thumb_files = [item["key"] for item in display_entries if item.get("kind") == "image"]
            for item in display_entries:
                if item.get("kind") == "image":
                    key = item["key"]
                    header = item["header"]
                    fds = item["fds"]
                    t = item["path"]
                    thumb_channel_idx = _thumbnail_channel_for_key(viewer, key, channel_idx)
                    entry_cmap_name = _thumbnail_cmap_for_key(viewer, key, thumb_channel_idx, cmap_name)
                    entry_clim = _thumbnail_clim_for_key(viewer, key, thumb_channel_idx)
                    lbl = QtWidgets.QLabel()
                    lbl.setAlignment(QtCore.Qt.AlignCenter)
                    lbl.setProperty("file_path", key)
                    lbl.setProperty("channel_index", int(thumb_channel_idx))
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
                    if fds and 0 <= thumb_channel_idx < len(fds):
                        fd = fds[thumb_channel_idx]
                        base_pix = None
                        data_key = None
                        try:
                            data_key = viewer._thumbnail_data_key(key, thumb_channel_idx, fd, thumb_w, thumb_h)
                        except Exception:
                            data_key = None
                        if data_key:
                            base_pix = viewer.thumb_cache.get(_thumb_render_cache_key(data_key, entry_cmap_name, entry_clim))
                        if base_pix is not None:
                            pix = base_pix.copy()
                            crop_info = None
                            try:
                                with viewer._thumb_data_lock:
                                    crop_info = viewer._thumb_crop_cache.get(data_key)
                            except Exception:
                                crop_info = None
                            markers = viewer._decorate_thumbnail_pixmap(pix, key, thumb_channel_idx, header, fds, thumb_crop=crop_info)
                            lbl.setPixmap(pix)
                            lbl.setProperty("spec_markers", markers)
                            try:
                                lbl.setProperty("thumb_crop", crop_info)
                            except Exception:
                                pass
                            viewer._thumb_loaded.add(key)
                        else:
                            lbl.setProperty("spec_markers", [])
                        viewer._thumb_meta[key] = (thumb_channel_idx, header, fd, thumb_w, thumb_h, entry_cmap_name, entry_clim, generation)
                    else:
                        blank = QtGui.QPixmap(thumb_w, thumb_h)
                        blank.fill(QtGui.QColor('black'))
                        lbl.setPixmap(blank)
                        lbl.setProperty("spec_markers", [])
                else:
                    spec = item["spec"]
                    key = item["key"]
                    card = QtWidgets.QFrame()
                    card.setFrameShape(QtWidgets.QFrame.StyledPanel)
                    card.setLineWidth(0)
                    card.setStyleSheet("QFrame { border: 1px solid rgba(255, 184, 77, 170); border-radius: 10px; background-color: rgba(255, 184, 77, 22); }")
                    vbox = QtWidgets.QVBoxLayout(card)
                    vbox.setContentsMargins(4, 4, 4, 4)
                    vbox.setSpacing(4)
                    lbl = QtWidgets.QLabel()
                    lbl.setAlignment(QtCore.Qt.AlignCenter)
                    lbl.setProperty("file_path", key)
                    lbl.setProperty("spectro_entry", spec)
                    lbl.setProperty("drag_start", None)
                    lbl.setProperty("dragging", False)
                    lbl.setProperty("spectro_channel", _spectro_display_channel(viewer, spec))
                    lbl.setPixmap(_spectroscopy_miniature_pixmap(viewer, spec, thumb_w, thumb_h, lbl.property("spectro_channel")))
                    lbl.setMouseTracking(True)
                    lbl.mousePressEvent = _make_spectro_thumb_press_handler(viewer, lbl)
                    lbl.mouseMoveEvent = _make_spectro_thumb_move_handler(viewer, lbl)
                    lbl.mouseReleaseEvent = _make_spectro_thumb_release_handler(viewer, lbl)
                    lbl.mouseDoubleClickEvent = _make_spectro_thumb_double_handler(viewer, lbl)
                    lbl.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
                    lbl.customContextMenuRequested.connect(lambda pos, lb=lbl: viewer._on_spectro_thumb_context_menu(lb, pos))
                    cap = QtWidgets.QLabel(Path(key).name or "Spectroscopy")
                    cap.setAlignment(QtCore.Qt.AlignCenter)
                    cap.setMaximumHeight(18)
                    cap.setFont(QtGui.QFont("Segoe UI", 9))
                    vbox.addWidget(lbl)
                    vbox.addWidget(cap)
                    viewer.thumb_layout.addWidget(card, row, col)
                    viewer.spectro_thumb_widgets[key] = card
                    viewer._spectro_thumb_labels[key] = lbl
                    if key in getattr(viewer, "spectro_thumb_multi_select", set()):
                        card.setStyleSheet("QFrame { border: 2px solid #ff9c3a; border-radius: 10px; background-color: rgba(255,156,58,42); }")
                # Advance the grid once per entry so images and spectra stay in the same
                # acquisition-order stream instead of collapsing into separate blocks.
                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1
            try:
                viewer._refresh_spectro_thumb_selection_styles()
            except Exception:
                pass
            try:
                viewer._refresh_thumb_selection_styles()
            except Exception:
                pass
            try:
                viewer._request_visible_thumbs()
            except Exception:
                pass
            viewer._refresh_frame_map_pixmaps()
            return

    for i, t in enumerate(files_iter):
        key = str(t)
        if key not in viewer.headers:
            continue
        header, fds = viewer.headers[key]
        thumb_channel_idx = _thumbnail_channel_for_key(viewer, key, channel_idx)
        entry_cmap_name = _thumbnail_cmap_for_key(viewer, key, thumb_channel_idx, cmap_name)
        entry_clim = _thumbnail_clim_for_key(viewer, key, thumb_channel_idx)
        lbl = QtWidgets.QLabel()
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setProperty("file_path", key)
        lbl.setProperty("channel_index", int(thumb_channel_idx))
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

        if fds and 0 <= thumb_channel_idx < len(fds):
            fd = fds[thumb_channel_idx]
            base_pix = None
            data_key = None
            try:
                data_key = viewer._thumbnail_data_key(key, thumb_channel_idx, fd, thumb_w, thumb_h)
            except Exception:
                data_key = None
            if data_key:
                base_pix = viewer.thumb_cache.get(_thumb_render_cache_key(data_key, entry_cmap_name, entry_clim))
            if base_pix is not None:
                pix = base_pix.copy()
                crop_info = None
                try:
                    with viewer._thumb_data_lock:
                        crop_info = viewer._thumb_crop_cache.get(data_key)
                except Exception:
                    crop_info = None
                markers = viewer._decorate_thumbnail_pixmap(pix, key, thumb_channel_idx, header, fds, thumb_crop=crop_info)
                lbl.setPixmap(pix)
                lbl.setProperty("spec_markers", markers)
                try:
                    lbl.setProperty("thumb_crop", crop_info)
                except Exception:
                    pass
                viewer._thumb_loaded.add(key)
            else:
                lbl.setProperty("spec_markers", [])
            viewer._thumb_meta[key] = (thumb_channel_idx, header, fd, thumb_w, thumb_h, entry_cmap_name, entry_clim, generation)
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
    cmap_name = _thumbnail_cmap_for_key(viewer, file_key, channel_idx, cmap_name)
    header, fds = viewer.headers.get(str(file_key), (None, None))
    if not header or not fds:
        return None
    if channel_idx < 0 or channel_idx >= len(fds):
        if not fds:
            return None
        channel_idx = min(max(channel_idx, 0), len(fds) - 1)
    fd = fds[channel_idx]
    clim = _thumbnail_clim_for_key(viewer, file_key, channel_idx)
    data_key = None
    try:
        data_key = viewer._thumbnail_data_key(str(file_key), channel_idx, fd, width, height)
    except Exception:
        data_key = None
    if data_key is not None:
        cache_key = ('frame', data_key, cmap_name, clim)
        pix = viewer._frame_real_pixmap_cache.get(cache_key)
        if pix is not None:
            return pix
        base_pix = viewer.thumb_cache.get(_thumb_render_cache_key(data_key, cmap_name, clim))
        if base_pix is not None:
            viewer._frame_real_pixmap_cache[cache_key] = base_pix
            return base_pix
        prefix = data_key[:4]
        for cache_entry in list(viewer.thumb_cache.keys()):
            try:
                key, cmap, key_clim = cache_entry
            except Exception:
                continue
            if cmap != cmap_name or key_clim != clim:
                continue
            if key[:4] == prefix:
                base_pix = viewer.thumb_cache.get(_thumb_render_cache_key(key, cmap_name, clim))
                if base_pix is not None:
                    viewer._frame_real_pixmap_cache[cache_key] = base_pix
                    return base_pix
    try:
        data_key, arr = viewer._get_thumbnail_array(str(file_key), channel_idx, header, fd, width, height)
    except Exception:
        return None
    cache_key = ('frame', data_key, cmap_name, clim)
    pix = viewer._frame_real_pixmap_cache.get(cache_key)
    if pix is None:
        try:
            vmin = vmax = None
            if clim is not None:
                vmin, vmax = clim
            qimg = array_to_qimage(arr, cmap_name=cmap_name, vmin=vmin, vmax=vmax)
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


def _toggle_spectro_thumb_multi_selection(viewer, file_path):
    path = _spectro_thumb_selection_key(viewer, file_path)
    if not hasattr(viewer, 'spectro_thumb_multi_select') or viewer.spectro_thumb_multi_select is None:
        viewer.spectro_thumb_multi_select = set()
    if path in viewer.spectro_thumb_multi_select:
        viewer.spectro_thumb_multi_select.remove(path)
    else:
        viewer.spectro_thumb_multi_select.add(path)
    viewer.selected_spectro_thumb_file = path
    _refresh_spectro_thumb_selection_styles(viewer)


def _clear_spectro_thumb_multi_selection(viewer, update_styles=True):
    viewer.spectro_thumb_multi_select = set()
    if update_styles:
        _refresh_spectro_thumb_selection_styles(viewer)


def _handle_thumb_click(viewer, label_widget, event):
    if event.button() != QtCore.Qt.LeftButton:
        return
    _cancel_pending_thumb_click(viewer)
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
        if not hasattr(viewer, 'spectro_thumb_multi_select') or viewer.spectro_thumb_multi_select is None:
            viewer.spectro_thumb_multi_select = set()
        anchor = getattr(viewer, 'last_thumb_anchor', None) or getattr(viewer, 'selected_file_for_thumbs', None) or str(fp)
        order = _combined_thumbnail_order(viewer)
        if _set_thumbnail_selection_by_order(viewer, order, str(anchor), str(fp), mods):
            viewer.last_thumb_anchor = str(fp)
        return

    if mods & QtCore.Qt.ControlModifier:
        viewer._toggle_thumb_multi_selection(fp)
        viewer.last_thumb_anchor = str(fp)
        return

    _schedule_plain_thumb_click(viewer, label_widget, fp, ch_idx)


def _make_thumb_press_handler(viewer, label_widget):
    def handler(event):
        if event.button() != QtCore.Qt.LeftButton:
            return
        if not _safe_set_property(label_widget, "drag_start", event.pos()):
            return
        _safe_set_property(label_widget, "press_modifiers", int(event.modifiers()))
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
            _safe_set_property(label_widget, "press_modifiers", None)
            return
        if bool(label_widget.property("skip_release_click")):
            _safe_set_property(label_widget, "skip_release_click", False)
            _safe_set_property(label_widget, "press_modifiers", None)
            return
        _handle_thumb_click(viewer, label_widget, event)
        _safe_set_property(label_widget, "press_modifiers", None)
    return handler


def _make_thumb_move_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        dragging = bool(label_widget.property("dragging"))
        start = label_widget.property("drag_start")
        if start is not None and event.buttons() & QtCore.Qt.LeftButton and not dragging:
            if (event.pos() - start).manhattanLength() >= 10:
                _safe_set_property(label_widget, "dragging", True)
                _start_thumb_collection_drag(viewer, label_widget)
                return
        if not viewer._handle_spec_hover(label_widget, event):
            QtWidgets.QLabel.mouseMoveEvent(label_widget, event)
    return handler


def _start_thumb_collection_drag(viewer, label_widget):
    """Start a lightweight drag from thumbnails for the collection tray."""
    try:
        file_path = str(label_widget.property("file_path") or "").strip()
        channel_idx = label_widget.property("channel_index")
        if not file_path or channel_idx is None:
            return
        try:
            channel_idx = int(channel_idx)
        except Exception:
            return
        selected = list(getattr(viewer, "_ordered_thumbnail_selection", lambda: [])() or [])
        if file_path in selected and selected:
            entries = [{"file_path": str(path), "channel_index": channel_idx} for path in selected if path]
        else:
            entries = [{"file_path": file_path, "channel_index": channel_idx}]
        mime = QtCore.QMimeData()
        mime.setData(
            "application/x-sxm-thumb-selection",
            json.dumps({"entries": entries, "drag_origin": "thumbnail_browser"}).encode("utf-8"),
        )
        drag = QtGui.QDrag(label_widget)
        drag.setMimeData(mime)
        pix = label_widget.pixmap()
        if pix is not None and not pix.isNull():
            drag.setPixmap(pix.scaled(120, 120, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        drag.exec_(QtCore.Qt.CopyAction)
    except Exception:
        pass


def _make_thumb_double_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        if event.button() != QtCore.Qt.LeftButton:
            return
        _cancel_pending_thumb_click(viewer)
        _safe_set_property(label_widget, "skip_release_click", True)
        fp = label_widget.property("file_path")
        ch_idx = int(label_widget.property("channel_index") or 0)
        try:
            viewer.on_thumbnail_double_clicked(fp, ch_idx)
        except Exception:
            pass
    return handler


def _handle_spectro_thumb_click(viewer, label_widget, event):
    if getattr(event, "button", None) and event.button() != QtCore.Qt.LeftButton:
        return False
    spec = label_widget.property("spectro_entry")
    if not spec:
        return False
    path = str(spec.get("path", "") or "")
    mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
    press_mods = label_widget.property("press_modifiers")
    if press_mods is not None:
        try:
            mods = QtCore.Qt.KeyboardModifiers(int(press_mods))
        except Exception:
            mods = press_mods
    if mods & QtCore.Qt.ShiftModifier:
        if not hasattr(viewer, 'spectro_thumb_multi_select') or viewer.spectro_thumb_multi_select is None:
            viewer.spectro_thumb_multi_select = set()
        anchor = getattr(viewer, 'last_spectro_thumb_anchor', None)
        if not anchor and getattr(viewer, 'selected_spectro_thumb_file', None):
            anchor = str(viewer.selected_spectro_thumb_file)
        if not anchor:
            anchor = path
        order = _combined_thumbnail_order(viewer)
        if _set_thumbnail_selection_by_order(viewer, order, str(anchor), path, mods):
            viewer.selected_spectro_thumb_file = path
            viewer.last_spectro_thumb_anchor = path
            return True
        viewer.spectro_thumb_multi_select.add(path)
        viewer.selected_spectro_thumb_file = path
        viewer.last_spectro_thumb_anchor = path
        _refresh_spectro_thumb_selection_styles(viewer)
        return True
    if mods & QtCore.Qt.ControlModifier:
        _toggle_spectro_thumb_multi_selection(viewer, path)
        viewer.last_spectro_thumb_anchor = path
        return True
    viewer.selected_spectro_thumb_file = path
    viewer.last_spectro_thumb_anchor = path
    _clear_spectro_thumb_multi_selection(viewer, update_styles=True)
    return True


def _make_spectro_thumb_press_handler(viewer, label_widget):
    def handler(event):
        if event.button() != QtCore.Qt.LeftButton:
            return
        if not _safe_set_property(label_widget, "drag_start", event.pos()):
            return
        _safe_set_property(label_widget, "press_modifiers", int(event.modifiers()))
        _safe_set_property(label_widget, "dragging", False)
    return handler


def _make_spectro_thumb_move_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        dragging = bool(label_widget.property("dragging"))
        start = label_widget.property("drag_start")
        if start is not None and event.buttons() & QtCore.Qt.LeftButton and not dragging:
            if (event.pos() - start).manhattanLength() >= 10:
                _safe_set_property(label_widget, "dragging", True)
                return
        if not viewer._handle_spec_hover(label_widget, event):
            QtWidgets.QLabel.mouseMoveEvent(label_widget, event)
    return handler


def _make_spectro_thumb_release_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        dragging = bool(label_widget.property("dragging"))
        if not _safe_set_property(label_widget, "drag_start", None):
            return
        _safe_set_property(label_widget, "dragging", False)
        if dragging:
            _safe_set_property(label_widget, "press_modifiers", None)
            return
        _handle_spectro_thumb_click(viewer, label_widget, event)
        _safe_set_property(label_widget, "press_modifiers", None)
    return handler


def _make_spectro_thumb_double_handler(viewer, label_widget):
    def handler(event):
        if sip.isdeleted(label_widget):
            return
        if event.button() != QtCore.Qt.LeftButton:
            return
        if _handle_spectro_thumb_click(viewer, label_widget, event):
            try:
                spec = label_widget.property("spectro_entry")
                if spec:
                    viewer._open_spectroscopy_popup(spec)
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
    cmap_name = viewer.thumb_cmap_combo.currentText()
    targets = list(getattr(viewer, "_ordered_thumbnail_selection", lambda: [])() or [])
    if not targets:
        current = str(getattr(viewer, "selected_file_for_thumbs", "") or "")
        if current:
            targets = [current]
    image_targets = [str(path) for path in targets if str(path) in getattr(viewer, "thumb_widgets", {})]
    if image_targets:
        try:
            viewer._set_thumbnail_entry_cmap(image_targets, cmap_name)
        except Exception:
            pass
        combo = getattr(viewer, "thumb_cmap_combo", None)
        if combo is not None:
            try:
                combo.blockSignals(True)
                combo.setCurrentText(str(cmap_name))
            finally:
                combo.blockSignals(False)
        return
    viewer.thumb_cmap = cmap_name
    viewer.config['thumbnail_cmap'] = viewer.thumb_cmap
    save_config(viewer.config)
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
    "_spectroscopy_miniature_pixmap",
    "_refresh_spectro_thumb_selection_styles",
    "_toggle_spectro_thumb_multi_selection",
    "_clear_spectro_thumb_multi_selection",
    "_make_spectro_thumb_press_handler",
    "_make_spectro_thumb_move_handler",
    "_make_spectro_thumb_release_handler",
    "_make_spectro_thumb_double_handler",
    "_toggle_thumb_multi_selection",
    "_clear_thumb_multi_selection",
    "on_thumb_cmap_changed",
]
