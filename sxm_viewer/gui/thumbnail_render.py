"""Thumbnail rendering, caching and export helpers."""
from __future__ import annotations

import warnings

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
from .. import cmap_registry
from ..config import (
    CONFIG_PATH,
    HEADER_CACHE_PATH,
    HEADER_CACHE_VERSION,
    CH_EQUALITY_TOL_NM,
    CH_SAMPLE_POINTS,
    CHANNEL_DATA_CACHE_LIMIT,
    FILTERED_CACHE_LIMIT,
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


def _get_cached_colormap(name):
    """Shim kept for existing callers: resolution + the shared-object
    cache now live in ``sxm_viewer.cmap_registry.get_cmap`` (which keeps
    this module's original invariant: cached Colormap objects are shared
    and must never be mutated — no set_bad/set_over/set_under)."""
    return cmap_registry.get_cmap(name)


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
    # effective_cmap honors the "Full amber imagery" display override
    # (identity when off) — covers every Qt-pixmap surface built here.
    cmap = cmap_registry.effective_cmap(cmap_name)
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
        cmap = _get_cached_colormap(name)
    except Exception:
        cmap = _get_cached_colormap('viridis')
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
    # Per-row ptp/std computed once across the whole array (vectorized)
    # instead of one np.isfinite/np.ptp/np.std call per row - measured
    # 712ms across 208 real thumbnails from ~47k small per-row numpy calls,
    # each dominated by Python/dispatch overhead rather than actual work.
    # np.where(...,np.nan) replaces non-finite entries (NaN *and* +-inf, to
    # match np.isfinite's original semantics exactly) before the nan-aware
    # reductions, since nanmax/nanmin/nanstd only skip NaN, not inf.
    finite_mask = np.isfinite(a)
    finite_count = finite_mask.sum(axis=1)
    masked = np.where(finite_mask, a, np.nan)
    # All-NaN rows are common here (aborted-scan padding) and are already
    # excluded below via finite_count >= 2 regardless of what nanmax/nanmin/
    # nanstd compute for them - np.errstate doesn't cover their "All-NaN
    # slice"/"Degrees of freedom <= 0" RuntimeWarnings (raised via Python's
    # warnings module, not the floating-point error state), so silence
    # those specifically rather than let them spam real usage.
    with warnings.catch_warnings(), np.errstate(all='ignore'):
        warnings.filterwarnings("ignore", category=RuntimeWarning)
        row_ptp = np.nanmax(masked, axis=1) - np.nanmin(masked, axis=1)
        row_std = np.nanstd(masked, axis=1)
    has_variation = (finite_count >= 2) & ((row_ptp > tolerance) | (row_std > tolerance))

    valid_idx = np.flatnonzero(has_variation)
    if valid_idx.size == 0:
        return None
    first_valid = int(valid_idx[0])

    empty_row = finite_count < 2
    last_valid = first_valid
    for i in range(first_valid + 1, rows):
        if empty_row[i]:
            if i > first_valid + 5:
                break
            continue
        if has_variation[i]:
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


def largest_inscribed_rect(w, h, angle_deg):
    """Width/height of the largest axis-aligned rectangle (centered) that
    fits entirely inside a w x h rectangle rotated by angle_deg.

    Classic closed-form ("rotatedRectWithMaxArea"). Units follow the
    inputs (used here with display-pixel spans)."""
    w = float(w)
    h = float(h)
    if w <= 0.0 or h <= 0.0:
        return 0.0, 0.0
    angle = math.radians(abs(float(angle_deg))) % math.pi
    if angle > math.pi / 2.0:
        angle = math.pi - angle
    sin_a = math.sin(angle)
    cos_a = math.cos(angle)
    if sin_a <= 1e-12:
        return w, h
    width_is_longer = w >= h
    side_long, side_short = (w, h) if width_is_longer else (h, w)
    if side_short <= 2.0 * sin_a * cos_a * side_long or abs(sin_a - cos_a) < 1e-10:
        x = 0.5 * side_short
        if width_is_longer:
            wr, hr = x / sin_a, x / cos_a
        else:
            wr, hr = x / cos_a, x / sin_a
    else:
        cos_2a = cos_a * cos_a - sin_a * sin_a
        wr = (w * cos_a - h * sin_a) / cos_2a
        hr = (h * cos_a - w * sin_a) / cos_2a
    return max(0.0, wr), max(0.0, hr)


def resample_geometry(arr, extent, geom):
    """Apply the redesigned Crop/Rotate dialog's geometry (flips, free
    rotation, crop) by resampling in the rotated frame — the approach the
    quick-crop template already uses (`_extract_rotated_crop`), which never
    produces the NaN-padded bounding-box frames `scipy.ndimage.rotate(
    reshape=True)` does.

    Coordinate conventions ("display frame"):
    - The display array is ``flipud(raw)`` shown with ``origin='lower'``
      and extent ``(0, w, 0, h)`` in source-pixel units (y up).
    - ``geom['rotate']`` rotates the displayed image CCW about the display
      center; ``geom['crop_rect'] = (left, right, bottom, top)`` is an
      axis-aligned rect *in that rotated display frame*.
    - Callers keep the rect inside :func:`largest_inscribed_rect` so no
      output pixel samples outside the source (out-of-bounds samples
      become NaN so mistakes are visible, not silently wrong).

    Returns ``(out_arr, out_extent)`` with ``out_arr`` in raw row order
    (row 0 = same end as the input) and ``out_extent`` the axis-aligned
    physical rect (source-extent units) whose center is the crop-rect
    center mapped back through the inverse rotation — the same accepted
    convention as quick-crop virtual copies (the rotation angle itself is
    not representable in the header). ``out_extent`` is None when
    ``extent`` is None. Output pixel counts preserve the source pixel
    pitch."""
    raw = np.asarray(arr, dtype=float)
    if raw.ndim < 2 or raw.size == 0:
        return np.array(raw, copy=True), extent
    geom = geom or {}
    if geom.get('flip_h'):
        raw = np.flip(raw, axis=1)
    if geom.get('flip_v'):
        raw = np.flip(raw, axis=0)
    h, w = raw.shape[:2]
    angle = float(geom.get('rotate', 0.0) or 0.0)
    rect = geom.get('crop_rect')
    if rect is None:
        rw, rh = largest_inscribed_rect(w, h, angle)
        left = 0.5 * w - 0.5 * rw
        right = 0.5 * w + 0.5 * rw
        bottom = 0.5 * h - 0.5 * rh
        top = 0.5 * h + 0.5 * rh
    else:
        left, right, bottom, top = [float(v) for v in rect]
        if right < left:
            left, right = right, left
        if top < bottom:
            bottom, top = top, bottom
    rect_w = max(right - left, 1e-9)
    rect_h = max(top - bottom, 1e-9)
    nx = max(2, int(round(rect_w)))
    ny = max(2, int(round(rect_h)))

    disp = np.flipud(raw)
    xs = np.linspace(left, right, nx, dtype=np.float64)
    ys = np.linspace(bottom, top, ny, dtype=np.float64)
    gx, gy = np.meshgrid(xs, ys)
    if abs(angle) > 1e-9:
        # Display point -> source display point: undo the display rotation
        # about the display center.
        rad = math.radians(-angle)
        cos_t = math.cos(rad)
        sin_t = math.sin(rad)
        cx = 0.5 * w
        cy = 0.5 * h
        dx = gx - cx
        dy = gy - cy
        gx = cx + dx * cos_t - dy * sin_t
        gy = cy + dx * sin_t + dy * cos_t
    cols = np.clip(gx / float(w) * (w - 1), -1.0, float(w))
    rows = np.clip(gy / float(h) * (h - 1), -1.0, float(h))
    if _scipy_ndimage is not None:
        sampled = _scipy_ndimage.map_coordinates(
            disp, [rows, cols], order=1, mode='constant', cval=np.nan)
    else:
        ri = np.clip(np.rint(rows).astype(np.int64), 0, h - 1)
        ci = np.clip(np.rint(cols).astype(np.int64), 0, w - 1)
        sampled = disp[ri, ci]
        oob = (rows < -0.5) | (rows > h - 0.5) | (cols < -0.5) | (cols > w - 0.5)
        if np.any(oob):
            sampled = np.array(sampled, dtype=float, copy=True)
            sampled[oob] = np.nan
    out_disp = np.asarray(sampled, dtype=float).reshape((ny, nx))
    out_arr = np.flipud(out_disp).copy()

    out_extent = None
    if extent is not None:
        xmin, xmax, ymin, ymax = [float(v) for v in extent]
        pitch_x = (xmax - xmin) / float(w)
        pitch_y = (ymax - ymin) / float(h)
        rcx = 0.5 * (left + right)
        rcy = 0.5 * (bottom + top)
        if abs(angle) > 1e-9:
            rad = math.radians(-angle)
            cos_t = math.cos(rad)
            sin_t = math.sin(rad)
            dx = rcx - 0.5 * w
            dy = rcy - 0.5 * h
            rcx = 0.5 * w + dx * cos_t - dy * sin_t
            rcy = 0.5 * h + dx * sin_t + dy * cos_t
        phys_cx = xmin + rcx * pitch_x
        phys_cy = ymin + rcy * pitch_y
        half_w = 0.5 * rect_w * pitch_x
        half_h = 0.5 * rect_h * pitch_y
        out_extent = [phys_cx - half_w, phys_cx + half_w,
                      phys_cy - half_h, phys_cy + half_h]
    return out_arr, out_extent


def geometry_is_identity(geom, shape):
    """True when a dialog geometry spec would leave the image unchanged."""
    if not geom:
        return True
    if geom.get('flip_h') or geom.get('flip_v'):
        return False
    if abs(float(geom.get('rotate', 0.0) or 0.0)) > 1e-3:
        return False
    rect = geom.get('crop_rect')
    if rect is None:
        return True
    h, w = shape[:2]
    left, right, bottom, top = [float(v) for v in rect]
    return (abs(left) < 0.5 and abs(bottom) < 0.5
            and abs(right - w) < 0.5 and abs(top - h) < 0.5)


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



