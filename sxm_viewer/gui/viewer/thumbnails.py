"""Thumbnail helpers for SXMGridViewer."""
from __future__ import annotations

from ..._shared import *
from ...data.io import *
from ...processing.filters import _filter_signature

def _thumbnail_filter_signature(viewer, file_key):
    spec = viewer.thumbnail_filters.get(str(file_key))
    return _filter_signature(spec)


def _downsample_for_thumbnail(viewer, arr, thumb_w, thumb_h):
    arr = np.asarray(arr, dtype=float)
    if arr.size == 0:
        return arr
    h, w = arr.shape
    if h > thumb_h or w > thumb_w:
        ys = np.linspace(0, h - 1, thumb_h).astype(int)
        xs = np.linspace(0, w - 1, thumb_w).astype(int)
        return arr[np.ix_(ys, xs)]
    return arr


def _get_thumbnail_array(viewer, file_key, channel_idx, header, fd, thumb_w, thumb_h):
    filter_sig = viewer._thumbnail_filter_signature(file_key)
    fname = fd.get("FileName")
    if not fname:
        raise ValueError("Missing FileName for channel")
    bin_path = Path(file_key).parent / fname
    try:
        bin_mtime = bin_path.stat().st_mtime
    except Exception:
        bin_mtime = 0.0
    data_key = (file_key, channel_idx, bin_mtime, filter_sig, thumb_w, thumb_h)
    with viewer._thumb_data_lock:
        cached = viewer._thumb_data_cache.get(data_key)
    if cached is not None:
        return data_key, cached
    _, arr_conv = viewer._get_filtered_channel_array(file_key, channel_idx, header, fd)
    thumb_arr = viewer._downsample_for_thumbnail(arr_conv, thumb_w, thumb_h)
    with viewer._thumb_data_lock:
        viewer._thumb_data_cache[data_key] = thumb_arr
    return data_key, thumb_arr


def _thumbnail_data_key(viewer, file_key, channel_idx, fd, thumb_w, thumb_h):
    filter_sig = viewer._thumbnail_filter_signature(file_key)
    fname = fd.get("FileName")
    if not fname:
        raise ValueError("Missing FileName for channel")
    bin_path = Path(file_key).parent / fname
    try:
        bin_mtime = bin_path.stat().st_mtime
    except Exception:
        bin_mtime = 0.0
    return (file_key, channel_idx, bin_mtime, filter_sig, thumb_w, thumb_h)


def _invalidate_thumbnail_cache(viewer, paths=None):
    if not paths:
        with viewer._thumb_data_lock:
            viewer._thumb_data_cache.clear()
        viewer.thumb_cache.clear()
        viewer._frame_real_pixmap_cache.clear()
        return
    path_set = {str(Path(p)) for p in paths}
    with viewer._thumb_data_lock:
        data_keys = [k for k in viewer._thumb_data_cache.keys() if k[0] in path_set]
        for k in data_keys:
            viewer._thumb_data_cache.pop(k, None)
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
