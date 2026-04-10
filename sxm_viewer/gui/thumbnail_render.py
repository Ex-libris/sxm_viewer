"""Thumbnail rendering, caching and export helpers."""
from __future__ import annotations

from .._shared import (
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
from ..config import (
    CONFIG_PATH,
    HEADER_CACHE_PATH,
    HEADER_CACHE_VERSION,
    CH_EQUALITY_TOL_NM,
    CH_SAMPLE_POINTS,
    CHANNEL_DATA_CACHE_LIMIT,
    FILTERED_CACHE_LIMIT,
    THUMB_DISK_CACHE_DIR,
    load_config,
    save_config,
    load_header_cache,
    save_header_cache,
)
from ..data.io import (
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
from ..processing.filters import (
    flatten_remove_median,
    subtract_best_fit_plane,
    subtract_2nd_order_plane,
    gaussian_filter_image,
    highpass_filter,
    FILTER_DEFINITIONS,
    _gaussian_available,
    _filter_signature,
)
from ..data.spectroscopy import (
    parse_spectroscopy_file,
    fit_parabola_bias,
    find_last_image_for_spec,
    _matrix_base_name,
    _rows_to_spec,
    _channel_labels,
    _clean_channel_label,
    _normalize_bias_axis,
    _extract_meta,
    _guess_index_from_name,
    _extract_section_value,
    _parse_section_metadata,
    _split_key_value,
    _split_tokens,
    _split_header_columns,
    _row_is_numeric,
    _normalize_meta_key,
    _coerce_value,
    _maybe_float,
    _maybe_int,
    _parse_datetime,
    _parse_date_and_time,
    _mtime,
    _read_text,
)


def array_to_qimage(arr, cmap_name='viridis', vmin=None, vmax=None, gamma=1.0):
    arr = np.asarray(arr, dtype=np.float64)
    invalid = ~np.isfinite(arr)
    try:
        if vmin is None:
            vmin = np.nanpercentile(arr, 1.0)
            vmax = np.nanpercentile(arr, 99.0)
    except Exception:
        vmin = float(np.nanmin(arr)); vmax = float(np.nanmax(arr))
    if vmin == vmax:
        vmin = float(np.nanmin(arr)); vmax = float(np.nanmax(arr))
    norm = (arr - vmin) / (vmax - vmin + 1e-30)
    norm = np.clip(norm, 0.0, 1.0) ** (1.0/gamma)
    if invalid.any():
        norm = np.array(norm, copy=True)
        norm[invalid] = 0.0
    cmap = colormaps.get_cmap(cmap_name)
    rgba = cmap(norm)
    if invalid.any():
        rgba = np.array(rgba, copy=True)
        rgba[invalid, 0:3] = 0.0
    rgba8 = (rgba * 255).astype(np.uint8)
    h,w = rgba8.shape[:2]
    img = QtGui.QImage(rgba8.data, w, h, rgba8.strides[0], QtGui.QImage.Format_RGBA8888)
    return img.copy()


# ---------- Background thumbnail helpers ----------
class _ThumbnailJobSignals(QtCore.QObject):
    finished = QtCore.pyqtSignal(str, int, object, object, str, object, int)
    failed = QtCore.pyqtSignal(str, int, str, int)


class _ThumbnailJob(QtCore.QRunnable):
    """
    Background task that builds a QImage for a thumbnail and passes it back
    to the GUI thread via signals.
    """
    def __init__(self, viewer, file_key, channel_idx, header, fd, thumb_w, thumb_h, cmap_name, clim, generation):
        super().__init__()
        self.viewer = viewer
        self.file_key = str(file_key)
        self.channel_idx = int(channel_idx)
        self.header = header
        self.fd = fd
        self.thumb_w = int(thumb_w)
        self.thumb_h = int(thumb_h)
        self.cmap_name = str(cmap_name)
        self.clim = clim
        self.generation = int(generation)
        self.signals = _ThumbnailJobSignals()

    def run(self):
        try:
            data_key, thumb_arr = self.viewer._get_thumbnail_array(
                self.file_key,
                self.channel_idx,
                self.header,
                self.fd,
                self.thumb_w,
                self.thumb_h,
            )
            vmin = vmax = None
            try:
                if self.clim is not None:
                    vmin, vmax = self.clim
            except Exception:
                vmin = vmax = None
            qimg = array_to_qimage(thumb_arr, cmap_name=self.cmap_name, vmin=vmin, vmax=vmax)
            self.signals.finished.emit(
                self.file_key,
                self.channel_idx,
                qimg,
                data_key,
                self.cmap_name,
                self.clim,
                self.generation,
            )
        except Exception as exc:
            self.signals.failed.emit(self.file_key, self.channel_idx, str(exc), self.generation)


# cache for generated icons to avoid regenerating
_CMAP_ICON_CACHE = {}
_SI_UNIT_MAP = {
    'pm': ('m', 1e-12),
    'nm': ('m', 1e-9),
    'um': ('m', 1e-6),
    'µm': ('m', 1e-6),
    'μm': ('m', 1e-6),
    'mm': ('m', 1e-3),
    'cm': ('m', 1e-2),
    'm': ('m', 1.0),
    'pa': ('A', 1e-12),
    'pA': ('A', 1e-12),
    'na': ('A', 1e-9),
    'nA': ('A', 1e-9),
    'ua': ('A', 1e-6),
    'uA': ('A', 1e-6),
    'µa': ('A', 1e-6),
    'µA': ('A', 1e-6),
    'μa': ('A', 1e-6),
    'μA': ('A', 1e-6),
    'ma': ('A', 1e-3),
    'mA': ('A', 1e-3),
    'a': ('A', 1.0),
    'mv': ('V', 1e-3),
    'mV': ('V', 1e-3),
    'kv': ('V', 1e3),
    'kV': ('V', 1e3),
    'v': ('V', 1.0),
    'V': ('V', 1.0),
    'hz': ('Hz', 1.0),
    'kHz': ('Hz', 1e3),
    'khz': ('Hz', 1e3),
    'mhz': ('Hz', 1e6),
    'MHz': ('Hz', 1e6),
    'ghz': ('Hz', 1e9),
    'GHz': ('Hz', 1e9),
}

def _colormap_icon(name: str, width: int = 96, height: int = 14) -> QIcon:
    """
    Return a QIcon showing a small horizontal gradient for the matplotlib colormap `name`.
    Caches icons for faster reuse.
    """
    key = (name, width, height)
    if key in _CMAP_ICON_CACHE:
        return _CMAP_ICON_CACHE[key]
    try:
        cmap = colormaps.get_cmap(name)
    except Exception:
        cmap = colormaps.get_cmap('viridis')
    grad = np.linspace(0.0, 1.0, width, dtype=np.float32)
    rgba = cmap(grad)
    rgba8 = (rgba * 255).astype(np.uint8)
    rgba8 = np.repeat(rgba8[np.newaxis, :, :], height, axis=0)
    rgba8 = np.ascontiguousarray(rgba8)
    h, w = rgba8.shape[:2]
    img = QImage(rgba8.data, w, h, rgba8.strides[0], QImage.Format_RGBA8888)
    pix = QPixmap.fromImage(img.copy())
    icon = QIcon(pix)
    _CMAP_ICON_CACHE[key] = icon
    return icon

# ---------------- Visualization & export helpers ----------------

def convert_to_si(arr, unit):
    """Convert numeric array values to SI units when possible."""
    if unit is None:
        return np.array(arr, dtype=float), None
    key = str(unit).strip()
    key_lower = key.lower()
    target = _SI_UNIT_MAP.get(key) or _SI_UNIT_MAP.get(key_lower)
    data = np.array(arr, dtype=float)
    if target:
        target_unit, factor = target
        return data * factor, target_unit
    return data, unit

def _unit_to_nm_factor(unit):
    """Return the conversion factor from the given unit string to nanometers."""
    if not unit:
        return 1.0
    u = str(unit).strip().lower()
    if not u:
        return 1.0
    if u in ('nm','nanometer','nanometre'):
        return 1.0
    if u in ('pm','picometer','picometre'):
        return 1e-3
    if u in ('µm','μm','um','micrometer','micrometre'):
        return 1e3
    if u in ('mm','millimeter','millimetre'):
        return 1e6
    if u in ('m','meter','metre'):
        return 1e9
    if u in ('ang','angstrom','ångstrom','ångström','å'):
        return 0.1
    if not u:
        return 1.0
    if u in ('nm','nanometer','nanometre'):
        return 1.0
    if u in ('pm','picometer','picometre'):
        return 1e-3
    if u in ('ï¿½m','um','micrometer','micrometre'):
        return 1e3
    if u in ('mm','millimeter','millimetre'):
        return 1e6
    if u in ('m','meter','metre'):
        return 1e9
    if u in ('ang','ï¿½ngstrï¿½m','angstom','ï¿½'):
        return 0.1
    return 1.0

def _value_in_nm(val, unit):
    """Convert a numeric value expressed in unit to nanometers."""
    try:
        if val is None:
            return None
        return float(val) * _unit_to_nm_factor(unit)
    except Exception:
        return None

def detect_valid_scan_region(arr, tolerance=1e-10):
    """
    Detect contiguous valid rows in a scan by finding where variation disappears (aborted/partial scans).
    Returns (first_valid_row, last_valid_row) or None if not found.
    """
    a = np.asarray(arr, dtype=float)
    if a.ndim != 2:
        return None
    rows, cols = a.shape
    if rows == 0 or cols == 0:
        return None
    first_valid = None
    for i in range(rows):
        row = a[i, :]
        finite = row[np.isfinite(row)]
        if finite.size < 2:
            continue
        if np.ptp(finite) > tolerance or np.std(finite) > tolerance:
            first_valid = i
            break
    if first_valid is None:
        return None
    last_valid = first_valid
    for i in range(first_valid + 1, rows):
        row = a[i, :]
        finite = row[np.isfinite(row)]
        if finite.size < 2:
            if i > first_valid + 5:
                break
            continue
        if np.ptp(finite) > tolerance or np.std(finite) > tolerance:
            last_valid = i
        else:
            break
    return (first_valid, last_valid)

def robust_limits(arr, low_pct=2.0, high_pct=98.0):
    """
    Return percentile-based intensity limits with automatic aborted-scan detection and optional flat suppression.
    """
    data_arr = np.asarray(arr, dtype=float)
    if data_arr.ndim == 2:
        region = detect_valid_scan_region(data_arr)
        if region:
            r0, r1 = region
            data_arr = data_arr[r0:r1 + 1, :]
    data = data_arr[np.isfinite(data_arr)]
    if data.size == 0:
        return None, None
    # Optionally trim a dominant flat bin
    try:
        hist, edges = np.histogram(data, bins=256)
        idx_max = int(np.argmax(hist))
        frac = hist[idx_max] / float(data.size)
        if frac > 0.7:
            lo_edge, hi_edge = edges[idx_max], edges[idx_max + 1]
            trimmed = data[(data < lo_edge) | (data > hi_edge)]
            if trimmed.size >= max(10, int(0.001 * data.size)):
                if trimmed.size > 100:
                    if np.std(trimmed) > 1e-12 and np.ptp(trimmed) > 1e-12:
                        data = trimmed
                else:
                    data = trimmed
    except Exception:
        pass
    low = max(0.0, min(low_pct, 100.0))
    high = max(low + 0.001, min(high_pct, 100.0))
    vmin = float(np.percentile(data, low))
    vmax = float(np.percentile(data, high))
    if vmin == vmax:
        vmax = vmin + 1e-12
    return vmin, vmax

def _interp_index(coord, start, end, size):
    """Interpolate a coordinate along the axis defined by start/end into pixel space."""
    if size <= 0 or start == end:
        return None
    lo = min(start, end)
    hi = max(start, end)
    if coord < lo or coord > hi:
        return None
    if end > start:
        t = (coord - start) / (end - start)
    else:
        t = (coord - end) / (start - end)
    return t * (size - 1)

def sample_array_value(arr, x, y, extent=None):
    """Sample array arr at physical coordinate (x,y), mapping via extent when provided."""
    arr = np.asarray(arr, dtype=float)
    if arr.size == 0 or x is None or y is None:
        return None
    h, w = arr.shape
    if extent is not None:
        xmin, xmax, ymin, ymax = extent
        col = _interp_index(x, xmin, xmax, w)
        row = _interp_index(y, ymin, ymax, h)
    else:
        if x < 0 or y < 0 or x > (w - 1) or y > (h - 1):
            return None
        col = x
        row = y
    if col is None or row is None:
        return None
    col = int(np.clip(round(col), 0, w - 1))
    row = int(np.clip(round(row), 0, h - 1))
    val = arr[row, col]
    if not np.isfinite(val):
        return None
    return float(val)

def apply_adjustment_spec(arr, extent, spec):
    """Apply crop/flip/rotate/clip/gamma adjustments described by spec to arr."""
    if spec is None:
        return np.array(arr, copy=True), extent
    result = np.array(arr, dtype=float, copy=True)
    out_extent = extent
    h, w = result.shape
    crop = spec.get('crop') or {}
    x0 = int(np.clip(crop.get('x0', 0), 0, max(0, w - 1)))
    x1 = int(np.clip(crop.get('x1', w), x0 + 1, w))
    y0 = int(np.clip(crop.get('y0', 0), 0, max(0, h - 1)))
    y1 = int(np.clip(crop.get('y1', h), y0 + 1, h))
    if (x0, x1, y0, y1) != (0, w, 0, h):
        result = result[y0:y1, x0:x1]
        if extent is not None:
            xmin, xmax, ymin, ymax = extent
            dx = (xmax - xmin) / float(w)
            dy = (ymax - ymin) / float(h)
            new_xmin = xmin + dx * x0
            new_xmax = xmin + dx * x1
            new_ymin = ymin + dy * y0
            new_ymax = ymin + dy * y1
            out_extent = [new_xmin, new_xmax, new_ymin, new_ymax]
    rot = float(spec.get('rotate', 0.0) or 0.0)
    flip_h = bool(spec.get('flip_h'))
    flip_v = bool(spec.get('flip_v'))
    if flip_h:
        result = np.flip(result, axis=1)
    if flip_v:
        result = np.flip(result, axis=0)
    if abs(rot) > 1e-3:
        if _scipy_ndimage is not None:
            result = _scipy_ndimage.rotate(result, rot, reshape=True, order=1, mode='constant', cval=np.nan)
        else:
            k = int(round(rot / 90.0)) % 4
            if k:
                result = np.rot90(result, k)
        if out_extent is not None:
            out_extent = _rotate_extent_box(out_extent, rot)
        result, out_extent = _trim_nan_border(result, out_extent)
    clip = spec.get('clip') or {}
    low_pct = clip.get('low')
    high_pct = clip.get('high')
    if low_pct is not None or high_pct is not None:
        finite = result[np.isfinite(result)]
        if finite.size:
            low_val = np.nanpercentile(finite, float(low_pct)) if low_pct is not None else np.nanmin(finite)
            high_val = np.nanpercentile(finite, float(high_pct)) if high_pct is not None else np.nanmax(finite)
            if high_val == low_val:
                high_val = low_val + 1e-12
            result = np.clip(result, low_val, high_val)
    gamma = float(spec.get('gamma', 1.0) or 1.0)
    if abs(gamma - 1.0) > 1e-3:
        finite = result[np.isfinite(result)]
        if finite.size:
            vmin = float(np.nanmin(finite))
            vmax = float(np.nanmax(finite))
            if vmax == vmin:
                vmax = vmin + 1e-12
            norm = np.clip((result - vmin) / (vmax - vmin), 0.0, 1.0)
            norm = norm ** gamma
            result = norm * (vmax - vmin) + vmin
    return result, out_extent


def _rotate_extent_box(extent, angle_deg):
    if extent is None:
        return None
    xmin, xmax, ymin, ymax = map(float, extent)
    cx = 0.5 * (xmin + xmax)
    cy = 0.5 * (ymin + ymax)
    rad = np.deg2rad(angle_deg)
    sin_t = np.sin(rad)
    cos_t = np.cos(rad)
    corners = [
        (xmin, ymin),
        (xmin, ymax),
        (xmax, ymin),
        (xmax, ymax),
    ]
    rx = []
    ry = []
    for x, y in corners:
        dx = x - cx
        dy = y - cy
        rx.append(cx + dx * cos_t - dy * sin_t)
        ry.append(cy + dx * sin_t + dy * cos_t)
    return [min(rx), max(rx), min(ry), max(ry)]


def _trim_nan_border(arr, extent):
    if arr.size == 0:
        return arr, extent
    mask = np.isfinite(arr)
    if not np.any(mask):
        return arr, extent
    rows = np.where(mask.any(axis=1))[0]
    cols = np.where(mask.any(axis=0))[0]
    r0, r1 = rows[0], rows[-1] + 1
    c0, c1 = cols[0], cols[-1] + 1
    if r0 == 0 and c0 == 0 and r1 == arr.shape[0] and c1 == arr.shape[1]:
        return arr, extent
    trimmed = arr[r0:r1, c0:c1]
    if extent is not None:
        xmin, xmax, ymin, ymax = map(float, extent)
        h, w = arr.shape
        dx = (xmax - xmin) / float(w)
        dy = (ymax - ymin) / float(h)
        new_xmin = xmin + dx * c0
        new_xmax = xmin + dx * c1
        new_ymin = ymin + dy * r0
        new_ymax = ymin + dy * r1
        extent = [new_xmin, new_xmax, new_ymin, new_ymax]
    return trimmed, extent

def save_wsxm_xyz(path, arr, x_vals, y_vals, name, z_unit="a.u.", z_scale=1.0):
    """Save arr as WSxM ASCII XYZ file (same structure as historical exports)."""
    arr = np.asarray(arr, dtype=float)
    if not np.any(np.isfinite(arr)):
        return
    os.makedirs(path, exist_ok=True)
    ny, nx = arr.shape
    z = np.array(arr, copy=True, dtype=float)
    z[~np.isfinite(z)] = 0.0
    z *= float(z_scale)
    x_vals = np.asarray(x_vals, dtype=float)
    y_vals = np.asarray(y_vals, dtype=float)
    if x_vals.size != nx:
        x_vals = np.arange(nx, dtype=float)
    if y_vals.size != ny:
        y_vals = np.arange(ny, dtype=float)
    fname = os.path.join(path, f"{name}.txt")
    with open(fname, "w") as f:
        f.write("WSxM file copyright UAM\n")
        f.write("WSxM ASCII XYZ file\n")
        f.write(f"X[nm]\t\tY[nm]\t\tZ[{z_unit}]\n\n")
        for iy, y in enumerate(y_vals):
            for ix, x in enumerate(x_vals):
                f.write(f"{x:.6f}\t{y:.6f}\t{z[iy, ix]:.7g}\n")


__all__ = [
    "array_to_qimage",
    "_ThumbnailJobSignals",
    "_ThumbnailJob",
    "_colormap_icon",
    "convert_to_si",
    "_unit_to_nm_factor",
    "_value_in_nm",
    "robust_limits",
    "_interp_index",
    "sample_array_value",
    "apply_adjustment_spec",
    "save_wsxm_xyz",
]



