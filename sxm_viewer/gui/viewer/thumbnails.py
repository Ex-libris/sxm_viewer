"""Thumbnail helpers for SXMGridViewer."""
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
from ...data.io import (
    parse_header,
    read_channel_file,
    normalize_unit_and_data,
    _split_key_value,
    _coerce_value,
    _canonical_header_key,
    _parse_inline_channels,
    _trailing_digits,
    _load_ascii_grid,
    _load_binary_grid,
    _load_tokenized_grid,
    _load_binary_with_inference,
    _binary_dtype_candidates,
)
from ...processing.filters import _filter_signature
from ..thumbnail_render import detect_valid_scan_region

def _thumbnail_filter_signature(viewer, file_key):
    spec = viewer.thumbnail_filters.get(str(file_key))
    return _filter_signature(spec)


def _thumbnail_display_signature(viewer, file_key, channel_idx):
    try:
        adjust_spec = viewer._get_adjust_spec(file_key, channel_idx)
    except Exception:
        adjust_spec = None
    try:
        adjust_sig = json.dumps(adjust_spec, sort_keys=True, default=str) if adjust_spec else ""
    except Exception:
        adjust_sig = repr(adjust_spec)
    return (
        bool(getattr(viewer, "display_units_si", False)),
        bool(getattr(viewer, "display_units_relative", False)),
        adjust_sig,
    )


def _downsample_for_thumbnail(viewer, arr, thumb_w, thumb_h):
    arr = np.asarray(arr, dtype=float)
    if arr.size == 0:
        return arr
    h, w = arr.shape
    # Preserve aspect ratio when downsampling/cropping to avoid distortion.
    if h > thumb_h or w > thumb_w:
        scale = min(thumb_w / float(w), thumb_h / float(h))
        if scale >= 1.0:
            return arr
        target_w = max(1, int(round(w * scale)))
        target_h = max(1, int(round(h * scale)))
        ys = np.linspace(0, h - 1, target_h).astype(int)
        xs = np.linspace(0, w - 1, target_w).astype(int)
        return arr[np.ix_(ys, xs)]
    return arr


def _get_thumbnail_array(viewer, file_key, channel_idx, header, fd, thumb_w, thumb_h):
    filter_sig = viewer._thumbnail_filter_signature(file_key)
    display_sig = _thumbnail_display_signature(viewer, file_key, channel_idx)
    fname = fd.get("FileName")
    if not fname:
        raise ValueError("Missing FileName for channel")
    bin_path = Path(file_key).parent / fname
    try:
        bin_mtime = bin_path.stat().st_mtime
    except Exception:
        bin_mtime = 0.0
    data_key = (file_key, channel_idx, bin_mtime, filter_sig, display_sig, thumb_w, thumb_h)
    with viewer._thumb_data_lock:
        cached = viewer._thumb_data_cache.get(data_key)
    if cached is not None:
        return data_key, cached
    _, arr_conv = viewer._get_filtered_channel_array(file_key, channel_idx, header, fd)
    arr_use = np.asarray(arr_conv, dtype=float)
    try:
        base_extent = viewer._header_extent(header)
    except Exception:
        base_extent = None
    try:
        arr_use, _ = viewer._apply_adjustments_for_channel(file_key, channel_idx, arr_use, base_extent)
    except Exception:
        pass
    try:
        _, arr_use, _ = viewer._scale_unit_for_display(fd.get("PhysUnit", ""), arr_use)
    except Exception:
        arr_use = np.asarray(arr_use, dtype=float)
    crop_info = None
    try:
        if arr_use.ndim == 2:
            orig_rows, orig_cols = int(arr_use.shape[0]), int(arr_use.shape[1])
            region = detect_valid_scan_region(arr_use)
            if region:
                r0, r1 = region
                arr_use = arr_use[r0 : r1 + 1, :]
                crop_info = {
                    "r0": int(r0),
                    "r1": int(r1),
                    "orig_rows": orig_rows,
                    "orig_cols": orig_cols,
                }
    except Exception:
        arr_use = np.asarray(arr_use, dtype=float)
        crop_info = None
    thumb_arr = viewer._downsample_for_thumbnail(arr_use, thumb_w, thumb_h)
    with viewer._thumb_data_lock:
        viewer._thumb_data_cache[data_key] = thumb_arr
        if hasattr(viewer, "_thumb_crop_cache"):
            if crop_info is None:
                viewer._thumb_crop_cache.pop(data_key, None)
            else:
                viewer._thumb_crop_cache[data_key] = crop_info
    return data_key, thumb_arr


def _thumbnail_data_key(viewer, file_key, channel_idx, fd, thumb_w, thumb_h):
    filter_sig = viewer._thumbnail_filter_signature(file_key)
    display_sig = _thumbnail_display_signature(viewer, file_key, channel_idx)
    fname = fd.get("FileName")
    if not fname:
        raise ValueError("Missing FileName for channel")
    bin_path = Path(file_key).parent / fname
    try:
        bin_mtime = bin_path.stat().st_mtime
    except Exception:
        bin_mtime = 0.0
    return (file_key, channel_idx, bin_mtime, filter_sig, display_sig, thumb_w, thumb_h)


def _invalidate_thumbnail_cache(viewer, paths=None):
    if not paths:
        with viewer._thumb_data_lock:
            viewer._thumb_data_cache.clear()
            if hasattr(viewer, "_thumb_crop_cache"):
                viewer._thumb_crop_cache.clear()
        viewer.thumb_cache.clear()
        viewer._frame_real_pixmap_cache.clear()
        return
    path_set = {str(Path(p)) for p in paths}
    with viewer._thumb_data_lock:
        data_keys = [k for k in viewer._thumb_data_cache.keys() if k[0] in path_set]
        for k in data_keys:
            viewer._thumb_data_cache.pop(k, None)
        if hasattr(viewer, "_thumb_crop_cache"):
            crop_keys = [k for k in viewer._thumb_crop_cache.keys() if k[0] in path_set]
            for k in crop_keys:
                viewer._thumb_crop_cache.pop(k, None)
    pix_keys = [k for k in viewer.thumb_cache.keys() if k[0][0] in path_set]
    for k in pix_keys:
        viewer.thumb_cache.pop(k, None)
    viewer._frame_real_pixmap_cache.clear()
__all__ = [
    "_thumbnail_filter_signature",
    "_downsample_for_thumbnail",
    "_get_thumbnail_array",
    "_thumbnail_data_key",
    "_invalidate_thumbnail_cache",
]




