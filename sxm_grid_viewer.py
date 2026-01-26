#!/usr/bin/env python3
"""
sxm_grid_viewer.py

Updated viewer:
 - Automatic detection of constant-height (CH) and constant-current (CC) files on folder load
 - Absolute z stored as integer picometers (pm), displayed in nm with 1 pm precision (3 decimals)
 - dz vs previous non-CH (last topo/CC before CH) and dz vs previous CH computed and shown
 - Thumbnail badges/borders indicating CH/CC
 - All previous features preserved (per-file inspector, Add channel view, persistent tags)

Heuristics for detection:
  - Header search for explicit 'constant' keywords (Mode, ScanMode, OperationMode, FeedbackState)
  - If not found: statistical test on topo-like channel: std < CH_STD_THRESHOLD_NM -> mark CH
  - CC detection mainly header-driven (presence of 'current' in Mode/Feedback or channel names)
"""
import sys, json, re, math, io, itertools, threading
from pathlib import Path
from collections import defaultdict, OrderedDict
import numpy as np
import hashlib
from datetime import datetime
import matplotlib
matplotlib.use("Agg")
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from matplotlib import colormaps
from PyQt5.QtGui import QIcon, QPixmap, QImage

from PyQt5 import QtWidgets, QtCore, QtGui
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# Faster colormap LUT cache for thumbnails (added)
_LUT_CACHE = {}

# ---------------- Config + thresholds ----------------
CONFIG_PATH = Path.home() / ".sxm_viewer_config.json"
HEADER_CACHE_PATH = Path.home() / ".sxm_viewer_header_cache.json"
CH_STD_THRESHOLD_NM = 0.02   # 0.02 nm = 20 pm; conservative threshold for CH detection by statistics
CHANNEL_DATA_CACHE_LIMIT = 24  # max channel arrays cached in-memory
FILTERED_CACHE_LIMIT = 32      # max filtered arrays cached in-memory
THUMB_DISK_CACHE_DIR = Path.home() / ".sxm_thumb_cache"

def load_config():
    try:
        s = CONFIG_PATH.read_text()
        return json.loads(s)
    except Exception:
        return {}

def save_config(cfg):
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2))
    except Exception:
        pass

def load_header_cache():
    """Load cached headers parsed in previous sessions."""
    try:
        s = HEADER_CACHE_PATH.read_text()
        return json.loads(s)
    except Exception:
        return {}

def save_header_cache(cache):
    """Persist header cache (used to speed up future loads)."""
    try:
        HEADER_CACHE_PATH.write_text(json.dumps(cache))
    except Exception:
        pass

# ---------------- Header parsing & binary reading ----------------
def parse_header(path):
    text = None
    for enc in ('utf-8','latin-1','cp1252'):
        try:
            with open(path, 'r', encoding=enc) as f:
                text = f.read()
            break
        except Exception:
            continue
    if text is None:
        raise IOError("Cannot read header " + str(path))
    lines = text.splitlines()
    header = {}; filedesc_list=[]; in_fd=False; cur={}
    for raw in lines:
        line = raw.strip()
        if not line: continue
        if line.startswith("FileDescBegin"):
            in_fd=True; cur={}; continue
        if line.startswith("FileDescEnd"):
            in_fd=False; filedesc_list.append(cur); cur={}; continue
        if ':' in line:
            k,v = line.split(':',1); k=k.strip(); v=v.strip()
            if in_fd: cur[k]=v
            else: header[k]=v
        else:
            m=re.match(r'(?P<k>[^:]+):\s*(?P<v>.+)', raw)
            if m:
                k=m.group('k').strip(); v=m.group('v').strip()
                if in_fd: cur[k]=v
                else: header[k]=v
    for fd in filedesc_list:
        for key in ('Scale','Offset'):
            if key in fd:
                try: fd[key]=float(fd[key])
                except:
                    try: fd[key]=float(str(fd[key]).split()[0])
                    except: pass
    for k in list(header.keys()):
        v = header[k]
        try: header[k]=float(v)
        except: header[k]=v
    return header, filedesc_list

def find_dtype_for_file(path, expected_pixels):
    candidate = [np.int16, np.uint16, np.int32, np.uint32, np.int64, np.float32, np.float64]
    filesize = Path(path).stat().st_size
    for dt in candidate:
        s = np.dtype(dt).itemsize
        if filesize % s != 0: continue
        cnt = filesize // s
        if cnt == expected_pixels:
            return np.fromfile(str(path), dtype=dt).astype(np.float64), dt
    for dt in (np.float32, np.int32, np.int16, np.uint16):
        try:
            data = np.fromfile(str(path), dtype=dt)
            if data.size >= expected_pixels:
                return data[:expected_pixels].astype(np.float64), dt
        except Exception:
            pass
    data = np.fromfile(str(path), dtype=np.uint8)
    if data.size >= expected_pixels:
        return data[:expected_pixels].astype(np.float64), np.uint8
    raise IOError("Cannot read binary file " + str(path))

def read_channel_file(path, xpix, ypix, scale=1.0, offset=0.0):
    expected = int(xpix) * int(ypix)
    arr, dt = find_dtype_for_file(path, expected)
    arr = arr.reshape((int(ypix), int(xpix)))
    return arr * float(scale) + float(offset)

def normalize_unit_and_data(data2d, unit_str):
    if not unit_str:
        return unit_str, data2d
    u = unit_str.strip().lower()
    if u in ('a','amp','amps','ampere','amperes') or re.fullmatch(r'a(\b.*)?', u):
        return 'pA', data2d * 1e12
    if 'pa' in u and 'p' in u:
        return 'pA', data2d
    if u in ('v','volt','volts','voltage') or re.fullmatch(r'v(\b.*)?', u):
        return 'mV', data2d * 1e3
    if 'mv' in u:
        return 'mV', data2d
    return unit_str, data2d

# ---------------- Virtual copy filters ----------------
def flatten_remove_median(img, axis='both'):
    """Subtract row/column medians from image."""
    arr = np.asarray(img, dtype=float)
    out = arr.copy()
    if axis in ('both', 'row', 0):
        med = np.nanmedian(out, axis=1, keepdims=True)
        out = out - med
    if axis in ('both', 'col', 1):
        med = np.nanmedian(out, axis=0, keepdims=True)
        out = out - med
    return out

def subtract_best_fit_plane(img):
    """Subtract best fit plane ax + by + c."""
    arr = np.asarray(img, dtype=float)
    h, w = arr.shape
    y, x = np.mgrid[:h, :w]
    A = np.c_[x.ravel(), y.ravel(), np.ones_like(x).ravel()]
    C, _, _, _ = np.linalg.lstsq(A, arr.ravel(), rcond=None)
    plane = (C[0]*x + C[1]*y + C[2])
    return arr - plane

def subtract_2nd_order_plane(img):
    """Subtract quadratic plane ax^2 + by^2 + cxy + dx + ey + f."""
    arr = np.asarray(img, dtype=float)
    h, w = arr.shape
    y, x = np.mgrid[:h, :w]
    A = np.c_[x.ravel()**2, y.ravel()**2, x.ravel()*y.ravel(), x.ravel(), y.ravel(), np.ones_like(x).ravel()]
    C, _, _, _ = np.linalg.lstsq(A, arr.ravel(), rcond=None)
    plane = (C[0]*x**2 + C[1]*y**2 + C[2]*x*y + C[3]*x + C[4]*y + C[5])
    return arr - plane

try:
    from scipy.ndimage import gaussian_filter as _scipy_gaussian
    _GAUSS_BACKEND = 'scipy'
except Exception:
    try:
        import cv2 as _cv2
        _GAUSS_BACKEND = 'cv2'
    except Exception:
        _GAUSS_BACKEND = None

def gaussian_filter_image(img, sigma):
    """Gaussian blur using scipy/cv2 fallback."""
    arr = np.asarray(img, dtype=float)
    if _GAUSS_BACKEND == 'scipy':
        return _scipy_gaussian(arr, sigma=sigma)
    if _GAUSS_BACKEND == 'cv2':
        k = int(max(3, (round(sigma*6) // 2) * 2 + 1))
        return _cv2.GaussianBlur(arr, (k, k), sigma)
    raise RuntimeError("Gaussian filter requires scipy or OpenCV.")

def highpass_filter(img, sigma):
    """High-pass filter = img - low-pass."""
    arr = np.asarray(img, dtype=float)
    lp = gaussian_filter_image(arr, sigma)
    return arr - lp

FILTER_DEFINITIONS = {
    'flatten': {'label': 'Flatten (row/col median)', 'needs_gaussian': False},
    'tilt': {'label': 'Tilt correction (plane)', 'needs_gaussian': False},
    'plane2': {'label': 'Global plane fit (2nd order)', 'needs_gaussian': False},
    'highpass': {'label': 'High-pass (Gaussian)', 'needs_gaussian': True, 'default_sigma': 2.0},
    'lowpass': {'label': 'Low-pass (Gaussian)', 'needs_gaussian': True, 'default_sigma': 2.0},
}

def _gaussian_available():
    return _GAUSS_BACKEND is not None
def _filter_signature(spec):
    if not spec:
        return tuple()
    sig = []
    for step in spec.get('steps', []):
        params = tuple(sorted((step.get('params') or {}).items()))
        sig.append((step.get('key'), params))
    return tuple(sig)

# ---------------- Spectroscopy helpers ----------------
SPECTRO_TIME_FORMATS = [
    '%Y-%m-%d %H:%M:%S',
    '%Y-%m-%d %H:%M',
    '%d/%m/%Y %H:%M:%S',
    '%d/%m/%Y %H:%M',
    '%d-%m-%Y %H:%M:%S',
    '%m/%d/%Y %H:%M:%S',
    '%m/%d/%Y %H:%M',
    '%Y-%m-%d',
    '%d/%m/%Y',
    '%m/%d/%Y',
]

def _parse_datetime_from_tokens(date_tok, time_tok):
    candidates = []
    if date_tok and time_tok:
        candidates.append(f"{date_tok} {time_tok}")
    if date_tok:
        candidates.append(date_tok)
    if time_tok:
        candidates.append(time_tok)
    for cand in candidates:
        for fmt in SPECTRO_TIME_FORMATS:
            try:
                return datetime.strptime(cand.strip(), fmt)
            except Exception:
                continue
    return None

def read_dat_spectroscopy(path):
    """
    Parse an Omicron-style spectroscopy .dat/.txt file.
    Returns dict with keys: path, x, y, time (datetime), V (np.ndarray), channels (dict)
    """
    path = Path(path)
    text = path.read_text(encoding='utf-8', errors='ignore')
    header_lines = []
    data_header = None
    data_lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith(';') or line.startswith('#'):
            header_lines.append(line.lstrip(';#').strip())
            continue
        if data_header is None:
            data_header = line
            continue
        data_lines.append(line)
    if data_header is None or not data_lines:
        raise ValueError(f"No numeric data found in {path}")
    col_names = [c for c in re.split(r'\s+', data_header.strip()) if c]
    if not col_names:
        raise ValueError(f"Cannot parse column names from {path}")
    data_str = "\n".join(data_lines)
    try:
        arr = np.loadtxt(io.StringIO(data_str))
    except Exception:
        raise
    arr = np.atleast_2d(arr)
    if arr.shape[1] != len(col_names):
        # attempt to pad/truncate if mismatch
        min_cols = min(arr.shape[1], len(col_names))
        arr = arr[:, :min_cols]
        col_names = col_names[:min_cols]
    # metadata from header
    x = y = None
    date_tok = time_tok = None
    for h in header_lines:
        lower = h.lower()
        if 'x/y-pos' in lower or 'x/y pos' in lower:
            nums = re.findall(r'[-+]?\d+(?:\.\d+)?', h)
            if len(nums) >= 2:
                x = float(nums[0]); y = float(nums[1])
        elif 'xpos' in lower and 'ypos' not in lower and x is None:
            try: x = float(h.split(':',1)[1].strip())
            except Exception: pass
        elif 'ypos' in lower and 'xpos' not in lower and y is None:
            try: y = float(h.split(':',1)[1].strip())
            except Exception: pass
        if 'date' in lower and date_tok is None and ':' in h:
            date_tok = h.split(':',1)[1].strip()
        if any(tok in lower for tok in ('time', 'timestamp')) and time_tok is None and ':' in h:
            time_tok = h.split(':',1)[1].strip()
    dt = _parse_datetime_from_tokens(date_tok, time_tok)
    if dt is None:
        dt = datetime.fromtimestamp(path.stat().st_mtime)
    # data extraction
    bias_idx = None
    channels = {}
    for idx, name in enumerate(col_names):
        lname = name.lower()
        if lname in ('bias','v','voltage'):
            bias_idx = idx
            continue
    if bias_idx is None:
        bias_idx = 2 if arr.shape[1] > 2 else 0
    V = np.asarray(arr[:, bias_idx], dtype=float)
    for idx, name in enumerate(col_names):
        if idx == bias_idx:
            continue
        channels[name] = np.asarray(arr[:, idx], dtype=float)
    # ensure df alias
    if 'df' not in channels and arr.shape[1] > 5:
        channels['df'] = np.asarray(arr[:, min(5, arr.shape[1]-1)], dtype=float)
    return {
        'path': path,
        'x': x,
        'y': y,
        'time': dt,
        'V': V,
        'channels': channels,
    }

def find_last_image_for_spec(spec_time, image_list):
    """
    Given a datetime spec_time and iterable of {'path': Path, 'time': datetime},
    return the entry representing the latest image acquired before the spectroscopy.
    If none exist earlier, return the closest in absolute time difference.
    """
    if spec_time is None or not image_list:
        return None
    earlier = [img for img in image_list if img.get('time') and img['time'] < spec_time]
    if earlier:
        earlier.sort(key=lambda i: i['time'], reverse=True)
        return earlier[0]
    # fallback to nearest by absolute delta
    def _delta(img):
        if not img.get('time'):
            return float('inf')
        return abs((img['time'] - spec_time).total_seconds())
    candidates = [img for img in image_list if img.get('time')]
    if not candidates:
        return None
    candidates.sort(key=_delta)
    return candidates[0]

def fit_parabola_bias(bias, values):
    """
    Fit df(V) data to df = a (V - b)^2 + c.
    Returns dict with fit params, errors, rmse, fitted curve callable.
    """
    bias = np.asarray(bias, dtype=float)
    values = np.asarray(values, dtype=float)
    mask = np.isfinite(bias) & np.isfinite(values)
    if mask.sum() < 3:
        raise ValueError("Need at least 3 valid points for fit")
    x = bias[mask]
    y = values[mask]
    x0 = float(np.mean(x))
    z = x - x0
    X = np.vstack((z**2, z, np.ones_like(z))).T
    coeffs, _, _, _ = np.linalg.lstsq(X, y, rcond=None)
    A, Bz, Cz = coeffs
    if abs(A) < 1e-12:
        raise ValueError("Parabola fit degenerates (a≈0)")
    y_fit = X @ coeffs
    residuals = y - y_fit
    rmse = float(np.sqrt(np.mean(residuals**2)))
    XtX = X.T @ X
    try:
        cov_ABC = np.linalg.inv(XtX)
    except np.linalg.LinAlgError:
        raise ValueError("Cannot invert design matrix for fit")
    dof = max(1, len(x) - 3)
    sigma2 = float(np.sum(residuals**2) / dof)
    cov_ABC = cov_ABC * sigma2
    cov_ABC = cov_ABC * sigma2
    T = np.array([
        [1.0, 0.0, 0.0],
        [-2.0*x0, 1.0, 0.0],
        [x0*x0, -x0, 1.0],
    ])
    coeffs_prime = T @ np.array([A, Bz, Cz])
    cov_prime = T @ cov_ABC @ T.T
    A, B, C = coeffs_prime
    varA = cov_prime[0,0]; varB = cov_prime[1,1]; varC = cov_prime[2,2]
    covAB = cov_prime[0,1]; covAC = cov_prime[0,2]; covBC = cov_prime[1,2]
    a = A
    b = -B / (2.0 * A)
    db_dA = B / (2.0 * (A**2))
    db_dB = -1.0 / (2.0 * A)
    varb = (db_dA**2)*varA + (db_dB**2)*varB + 2*db_dA*db_dB*covAB
    b_err = float(np.sqrt(varb)) if varb > 0 else 0.0
    c = C - a * (b**2)
    dc_dA = -(b**2) - 2*a*b*db_dA
    dc_dB = -2*a*b*db_dB
    dc_dC = 1.0
    varc = (dc_dA**2)*varA + (dc_dB**2)*varB + (dc_dC**2)*varC + \
           2*dc_dA*dc_dB*covAB + 2*dc_dA*dc_dC*covAC + 2*dc_dB*dc_dC*covBC
    c_err = float(np.sqrt(varc)) if varc > 0 else 0.0
    a_err = float(np.sqrt(varA)) if varA > 0 else 0.0
    def evaluate(v):
        v = np.asarray(v, dtype=float)
        return a * (v - b)**2 + c
    return {
        'a': float(a),
        'b': float(b),
        'c': float(c),
        'a_err': float(a_err),
        'b_err': float(b_err),
        'c_err': float(c_err),
        'rmse': rmse,
        'func': evaluate
    }

# ---------------- thumbnail helper: with badge drawing ----------------
def array_to_qimage(arr, cmap_name='viridis', vmin=None, vmax=None, gamma=1.0):
    """
    Convert numpy array -> QImage using cached 8-bit LUTs per colormap to reduce CPU time.
    """
    arr = np.asarray(arr, dtype=np.float32)
    if arr.size == 0:
        return QtGui.QImage()
    try:
        if vmin is None or vmax is None:
            vmin = float(np.nanpercentile(arr, 1.0))
            vmax = float(np.nanpercentile(arr, 99.0))
    except Exception:
        vmin = float(np.nanmin(arr)); vmax = float(np.nanmax(arr))
    if vmin == vmax:
        vmin = float(np.nanmin(arr)); vmax = float(np.nanmax(arr))
    denom = (vmax - vmin) or 1.0
    norm = np.clip((arr - vmin) / denom, 0.0, 1.0)
    if gamma != 1.0:
        norm = norm ** (1.0 / gamma)
    indices = (norm * 255.0).astype(np.uint8)
    lut_key = ('lut', cmap_name)
    lut = _LUT_CACHE.get(lut_key)
    if lut is None:
        try:
            cmap = colormaps.get_cmap(cmap_name)
        except Exception:
            cmap = colormaps.get_cmap('viridis')
        lut = (cmap(np.linspace(0.0, 1.0, 256)) * 255).astype(np.uint8)
        _LUT_CACHE[lut_key] = lut
    rgba8 = lut[indices]
    h, w = rgba8.shape[:2]
    try:
        img = QtGui.QImage(rgba8.data, w, h, rgba8.strides[0], QtGui.QImage.Format_RGBA8888)
        return img.copy()
    except Exception:
        buf = bytes(rgba8.ravel().tolist())
        return QtGui.QImage(buf, w, h, QtGui.QImage.Format_RGBA8888).copy()


# cache for generated icons to avoid regenerating
_CMAP_ICON_CACHE = {}

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


# ---------------- MultiPreview canvas (unchanged) ----------------
class MultiPreviewCanvas(FigureCanvas):
    def __init__(self, parent=None, figsize=(6,6)):
        # enable constrained_layout so Matplotlib reserves space for tick labels
        # and colorbars when the canvas is resized by the UI
        self.fig = Figure(figsize=figsize, constrained_layout=True)
        super().__init__(self.fig)
        if parent is not None:
            self.setParent(parent)
        self.views = []
        self._ax_view_map = {}
        self._copy_feedback_handler = None
        self._drag_candidate = None  # (view, QPoint start, QImage cache)
        # profile (interactive line) state
        self.profile_enabled = False
        self.profile_pts = None  # (x0, y0, x1, y1) in data coords of main ax
        self._profile_line = None
        self._profile_p0 = None
        self._profile_p1 = None
        self._cids = []
        self._base_click_cid = self.mpl_connect('button_press_event', self._on_base_click)
        self._dragging = None  # 'p0' or 'p1'
        self.main_ax = None
        self.profile_callback = None  # callable(x_px, vals, length_nm)

    def set_views(self, views):
        self.views = views[:]
        self._redraw()

    def clear_views(self):
        self.views = []
        self._redraw()

    def _redraw(self):
        self.fig.clf()
        self._ax_view_map = {}
        n = len(self.views)
        if n == 0:
            # nothing to show
            self.draw_idle()
            return

        # make the thumbnail grid responsive to the current canvas pixel aspect:
        # choose number of columns based on canvas width/height so thumbnails
        # adapt when the preview area is resized.
        try:
            w_px, h_px = self.get_width_height()
            aspect = float(max(1, w_px)) / float(max(1, h_px))
        except Exception:
            aspect = 1.0
        cols = max(1, int(math.ceil(math.sqrt(n * aspect))))
        rows = int(math.ceil(n / cols))
        for i, v in enumerate(self.views):
            ax = self.fig.add_subplot(rows, cols, i+1)
            self._ax_view_map[ax] = v
            if i == 0:
                self.main_ax = ax
            arr = np.asarray(v['arr'])
            extent = v.get('extent', None)
            cmap = v.get('cmap', 'viridis')
            if extent is None:
                im = ax.imshow(arr, origin='upper', interpolation='nearest', cmap=cmap)
            else:
                im = ax.imshow(arr, extent=extent, origin='upper', interpolation='nearest', aspect='equal', cmap=cmap)
            unit = v.get('unit', '')
            if unit:
                cbar = self.fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(unit)
            title = v.get('title', '')
            ax.set_title(title, fontsize=9)
            ax.tick_params(labelsize=8)
        # constrained_layout=True set on Figure should handle spacing.
        # fall back to tight_layout if constrained fails (defensive).
        try:
            # prefer constrained_layout, but calling tight_layout as fallback
            # can help older matplotlib versions. Both are safe inside try/except.
            self.fig.tight_layout()
        except Exception:
            try:
                self.fig.set_constrained_layout(True)
            except Exception:
                pass
        # if profile mode is enabled, (re)create artists on main ax
        if self.profile_enabled:
            self._ensure_profile_artists()
        # use draw_idle() to avoid blocking UI thread while Qt is resizing
        self.draw_idle()

    # ---------- Interactive profile helpers ----------
    def set_profile_callback(self, cb):
        self.profile_callback = cb

    def set_copy_feedback_handler(self, handler):
        self._copy_feedback_handler = handler

    def enable_profile(self, enable:bool):
        if enable == self.profile_enabled:
            return
        self.profile_enabled = enable
        if enable:
            self._connect_profile_events()
            self._ensure_profile_artists()
        else:
            self._disconnect_profile_events()
            self._clear_profile_artists()
            self.profile_pts = None
        self.draw_idle()

    def _connect_profile_events(self):
        if self._cids:
            return
        self._cids = [
            self.mpl_connect('button_press_event', self._on_press),
            self.mpl_connect('button_release_event', self._on_release),
            self.mpl_connect('motion_notify_event', self._on_motion),
        ]

    def _disconnect_profile_events(self):
        for cid in self._cids:
            try: self.mpl_disconnect(cid)
            except Exception: pass
        self._cids = []

    def _ensure_profile_artists(self):
        if self.main_ax is None:
            return
        if self._profile_line is None:
            # initialize points centered if not set
            if self.profile_pts is None:
                try:
                    v0 = self.views[0]
                    arr = np.asarray(v0['arr'])
                    h, w = arr.shape
                    if v0.get('extent') is None:
                        x0 = w*0.25; y0 = h*0.5; x1 = w*0.75; y1 = h*0.5
                    else:
                        xmin, xmax, ymin, ymax = v0['extent'][0], v0['extent'][1], v0['extent'][2], v0['extent'][3]
                        # our code sets extent as [0, XRange, YRange, 0], so choose a centered horizontal line
                        x0 = xmin + 0.25*(xmax - xmin); x1 = xmin + 0.75*(xmax - xmin)
                        y0 = ymax + 0.5*(ymin - ymax); y1 = y0
                except Exception:
                    x0 = 0.25; x1 = 0.75; y0 = y1 = 0.5
                self.profile_pts = (x0, y0, x1, y1)
            x0, y0, x1, y1 = self.profile_pts
            self._profile_line, = self.main_ax.plot([x0,x1],[y0,y1], color='yellow', lw=2, alpha=0.9, zorder=9)
            self._profile_p0, = self.main_ax.plot([x0],[y0], marker='o', color='yellow', ms=7, mec='black', mew=1.0, zorder=10)
            self._profile_p1, = self.main_ax.plot([x1],[y1], marker='o', color='yellow', ms=7, mec='black', mew=1.0, zorder=10)

    def _clear_profile_artists(self):
        for art in (self._profile_line, self._profile_p0, self._profile_p1):
            try:
                if art is not None:
                    art.remove()
            except Exception:
                pass
        self._profile_line = self._profile_p0 = self._profile_p1 = None
        self.draw_idle()

    def _update_profile_artists(self):
        if self._profile_line is None or self._profile_p0 is None or self._profile_p1 is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        self._profile_line.set_data([x0,x1],[y0,y1])
        self._profile_p0.set_data([x0],[y0])
        self._profile_p1.set_data([x1],[y1])
        self.draw_idle()
        self._emit_profile()

    def _pt_distance_pixels(self, x, y, xp, yp):
        try:
            p_scr = self.main_ax.transData.transform((x, y))
            q_scr = self.main_ax.transData.transform((xp, yp))
            dx = p_scr[0] - q_scr[0]; dy = p_scr[1] - q_scr[1]
            return (dx*dx + dy*dy) ** 0.5
        except Exception:
            return float('inf')

    def _on_press(self, event):
        if not self.profile_enabled or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        if event.button != 1:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        if self.profile_pts is None:
            self.profile_pts = (x, y, x, y)
            self._ensure_profile_artists()
            self._dragging = 'p1'
            self._update_profile_artists()
            return
        x0, y0, x1, y1 = self.profile_pts
        d0 = self._pt_distance_pixels(x, y, x0, y0)
        d1 = self._pt_distance_pixels(x, y, x1, y1)
        thresh = 10.0  # pixels
        if d0 <= thresh or d0 <= d1:
            if d0 <= thresh:
                self._dragging = 'p0'
                return
        if d1 <= thresh:
            self._dragging = 'p1'
            return
        # else: start a new line from here
        self.profile_pts = (x, y, x, y)
        self._dragging = 'p1'
        self._update_profile_artists()

    def _on_motion(self, event):
        if not self.profile_enabled or self._dragging is None or event.inaxes is None or event.inaxes is not self.main_ax:
            return
        x, y = event.xdata, event.ydata
        if x is None or y is None:
            return
        x0, y0, x1, y1 = self.profile_pts
        if self._dragging == 'p0':
            self.profile_pts = (x, y, x1, y1)
        elif self._dragging == 'p1':
            self.profile_pts = (x0, y0, x, y)
        self._update_profile_artists()

    def _on_release(self, event):
        if not self.profile_enabled:
            return
        self._dragging = None

    def _emit_profile(self):
        if not callable(self.profile_callback):
            return
        try:
            v0 = self.views[0]
            arr = np.asarray(v0['arr'], dtype=float)
            h, w = arr.shape
            extent = v0.get('extent', None)
            x0, y0, x1, y1 = self.profile_pts
            # map data coords to pixel indices
            if extent is None:
                c0 = x0; r0 = y0; c1 = x1; r1 = y1
                length_nm = None
            else:
                xmin, xmax = extent[0], extent[1]
                ymin, ymax = extent[2], extent[3]
                # our extent is [0, XRange, YRange, 0] so use ranges directly
                xr = (xmax - xmin) if (xmax is not None and xmin is not None) else 1.0
                yr = (ymin - ymax) if (ymin is not None and ymax is not None) else 1.0  # note inverted
                c0 = (x0 - xmin) / (xr + 1e-12) * (w - 1)
                c1 = (x1 - xmin) / (xr + 1e-12) * (w - 1)
                # y increases downward in array index
                # since extent top=ymax (0) and bottom=ymin (YRange), map linearly
                r0 = (y0 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                r1 = (y1 - ymax) / (ymin - ymax + 1e-12) * (h - 1)
                # physical length in nm using data coords directly
                try:
                    dx_nm = (x1 - x0); dy_nm = (y1 - y0)
                    length_nm = float((dx_nm*dx_nm + dy_nm*dy_nm) ** 0.5)
                except Exception:
                    length_nm = None
            # sample along the line using bilinear interpolation
            n = int(max(2, round(((c1 - c0)**2 + (r1 - r0)**2) ** 0.5) + 1))
            t = np.linspace(0.0, 1.0, n)
            cc = c0 + (c1 - c0) * t
            rr = r0 + (r1 - r0) * t
            rr = np.clip(rr, 0, h - 1)
            cc = np.clip(cc, 0, w - 1)
            i0 = np.floor(rr).astype(int)
            j0 = np.floor(cc).astype(int)
            i1 = np.clip(i0 + 1, 0, h - 1)
            j1 = np.clip(j0 + 1, 0, w - 1)
            wy = rr - i0
            wx = cc - j0
            vals = (
                (1 - wy) * (1 - wx) * arr[i0, j0] +
                wy * (1 - wx) * arr[i1, j0] +
                (1 - wy) * wx * arr[i0, j1] +
                wy * wx * arr[i1, j1]
            )
            x_px = np.linspace(0.0, float(n - 1), n)
            unit = v0.get('unit', None)
            self.profile_callback(x_px, vals, length_nm, unit)
        except Exception:
            pass

    def _on_base_click(self, event):
        if event is None or event.inaxes is None:
            return
        ax = event.inaxes
        view = self._ax_view_map.get(ax)
        if event.button == 3:
            self._show_context_menu(event, view)
            return
        if event.button != 1:
            return
        if getattr(event, 'dblclick', False):
            if view:
                self._copy_view_to_clipboard(view)
            return
        if view and getattr(event, 'guiEvent', None) is not None:
            pos = event.guiEvent.globalPos()
            self._drag_candidate = {'view': view, 'start': QtCore.QPoint(pos), 'image': None}

    def _copy_view_to_clipboard(self, view):
        try:
            qimg = self._view_to_qimage(view)
            QtWidgets.QApplication.clipboard().setImage(qimg)
            if callable(self._copy_feedback_handler):
                self._copy_feedback_handler(view)
        except Exception:
            pass

    def _view_to_qimage(self, view):
        arr = np.asarray(view.get('arr'))
        cmap = view.get('cmap', 'viridis')
        return array_to_qimage(arr, cmap_name=cmap)

    def _show_context_menu(self, event, view):
        if view is None or getattr(event, 'guiEvent', None) is None:
            return
        menu = QtWidgets.QMenu(self)
        copy_act = menu.addAction("Copy image")
        save_act = menu.addAction("Save image as...")
        chosen = menu.exec_(event.guiEvent.globalPos())
        if chosen == copy_act:
            self._copy_view_to_clipboard(view)
        elif chosen == save_act:
            self._save_view_to_file(view)

    def _save_view_to_file(self, view):
        try:
            title = view.get('title') or 'view'
            default = f"{title}.png"
            path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save view", default, "PNG Files (*.png)")
            if not path:
                return
            qimg = self._view_to_qimage(view)
            qimg.save(path, "PNG")
        except Exception:
            QtWidgets.QMessageBox.warning(self, "Save view", "Unable to save image.")

    def _start_drag(self, view, qimg=None):
        try:
            if qimg is None:
                qimg = self._view_to_qimage(view)
            pix = QtGui.QPixmap.fromImage(qimg)
            drag = QtGui.QDrag(self)
            mime = QtCore.QMimeData()
            mime.setImageData(qimg)
            drag.setMimeData(mime)
            drag.setPixmap(pix.scaled(128, 128, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
            drag.exec_(QtCore.Qt.CopyAction)
        except Exception:
            pass

    def mouseMoveEvent(self, event):
        if self._drag_candidate:
            start = self._drag_candidate.get('start')
            if start is not None:
                if (event.globalPos() - start).manhattanLength() >= 10:
                    view = self._drag_candidate.get('view')
                    qimg = self._drag_candidate.get('image')
                    if qimg is None and view is not None:
                        qimg = self._view_to_qimage(view)
                        self._drag_candidate['image'] = qimg
                    if view is not None and qimg is not None:
                        self._start_drag(view, qimg)
                    self._drag_candidate = None
                    super().mouseMoveEvent(event)
                    return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_candidate = None
        super().mouseReleaseEvent(event)

class ProfileDialog(QtWidgets.QDialog):
    """Dialog showing the sampled profile and basic stats."""
    def __init__(self, x, vals, length_nm=None, parent=None, unit=None):
        super().__init__(parent)
        self.setWindowTitle('Profile measurement')
        self.resize(600, 320)
        self._unit = unit
        v = QtWidgets.QVBoxLayout()
        # matplotlib canvas for plot
        fig = Figure(figsize=(6,3))
        self.canvas = FigureCanvas(fig)
        self.ax = fig.add_subplot(111)
        (self._line,) = self.ax.plot(x, vals)
        self.ax.set_xlabel('Distance (px)')
        self.ax.set_ylabel(f"Value ({unit})" if unit else 'Value')
        self.ax_top = self.ax.twiny()
        self.ax_top.set_xlabel('Distance (nm)')
        self._set_top_axis(length_nm, len(x))
        v.addWidget(self.canvas)
        # stats area
        self.stats = QtWidgets.QLabel(self._fmt_length(length_nm))
        v.addWidget(self.stats)
        btn = QtWidgets.QPushButton('Close')
        btn.clicked.connect(self.accept)
        v.addWidget(btn, alignment=QtCore.Qt.AlignRight)
        self.setLayout(v)

    def _fmt_length(self, length_nm):
        return f"Length: {length_nm:.3f} nm" if length_nm is not None else "Length: N/A"

    def _set_top_axis(self, length_nm, n_pts):
        try:
            self.ax_top.set_xlim(self.ax.get_xlim())
            if length_nm is None or n_pts <= 1:
                self.ax_top.set_xticks([])
                self.ax_top.set_xticklabels([])
            else:
                ticks = self.ax.get_xticks()
                scale = float(length_nm) / float(n_pts - 1)
                self.ax_top.set_xticks(ticks)
                self.ax_top.set_xticklabels([f"{t*scale:.1f}" for t in ticks])
        except Exception:
            pass

    def update_data(self, x, vals, length_nm=None):
        try:
            self._line.set_data(x, vals)
            self.ax.relim(); self.ax.autoscale_view()
            self._set_top_axis(length_nm, len(x))
            self.stats.setText(self._fmt_length(length_nm))
            self.canvas.draw_idle()
        except Exception:
            pass

class SpectroscopyPopup(QtWidgets.QDialog):
    """Popup window showing spectroscopy curves for a given file."""
    def __init__(self, spec, parent=None):
        super().__init__(parent)
        self.spec = spec
        self.setWindowTitle(f"Spectroscopy: {Path(spec['path']).name}")
        self.resize(720, 520)
        layout = QtWidgets.QVBoxLayout()
        meta_txt = f"File: {Path(spec['path']).name}\nPosition: {spec.get('x','?')}/{spec.get('y','?')} nm\nTime: {spec.get('time')}"
        self.meta_label = QtWidgets.QLabel(meta_txt)
        self.meta_label.setWordWrap(True)
        layout.addWidget(self.meta_label)

        selector_layout = QtWidgets.QHBoxLayout()
        selector_layout.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        selector_layout.addWidget(self.channel_combo, 1)
        self.fit_btn = QtWidgets.QPushButton("Fit parabola")
        selector_layout.addWidget(self.fit_btn)
        layout.addLayout(selector_layout)

        self.fig = Figure(figsize=(6,4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        layout.addWidget(self.canvas, 1)
        self.fit_result_label = QtWidgets.QLabel("")
        self.fit_result_label.setWordWrap(True)
        layout.addWidget(self.fit_result_label)
        self.setLayout(layout)

        self.V = np.asarray(spec.get('V', []), dtype=float)
        self.channels = {name: np.asarray(vals, dtype=float) for name, vals in (spec.get('channels', {}) or {}).items()}
        for name in self.channels.keys():
            self.channel_combo.addItem(name)
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        self.fit_btn.clicked.connect(self._on_fit_clicked)
        self._last_fit_result = None
        if self.channel_combo.count():
            self.channel_combo.setCurrentIndex(0)
            self._plot_selected_channel()
        else:
            self.ax.text(0.5, 0.5, "No channels", ha='center', va='center', transform=self.ax.transAxes)
            self.canvas.draw()
        self._update_fit_button()

    def _on_channel_changed(self, name):
        self._last_fit_result = None
        self.fit_result_label.setText("")
        self._plot_selected_channel()
        self._update_fit_button()

    def _plot_selected_channel(self):
        self.ax.clear()
        name = self.channel_combo.currentText()
        if not name or name not in self.channels or not self.V.size:
            self.canvas.draw_idle()
            return
        self.ax.plot(self.V, self.channels[name], color='#c94cfa', lw=1.5, label='Data')
        self.ax.set_xlabel("Bias (V)")
        self.ax.set_ylabel(name)
        self.ax.grid(True, alpha=0.2)
        if self._last_fit_result and self._last_fit_result.get('channel') == name:
            self._draw_fit_overlay(self._last_fit_result)
        else:
            handles, labels = self.ax.get_legend_handles_labels()
            if handles:
                self.ax.legend()
        self.canvas.draw_idle()

    def _draw_fit_overlay(self, res):
        if not self.V.size:
            return
        x_dense = np.linspace(np.nanmin(self.V), np.nanmax(self.V), 400)
        y_dense = res['func'](x_dense)
        self.ax.plot(x_dense, y_dense, '--', color='#ff8c00', lw=1.5, label='Fit')
        b = res['b']; c = res['c']; b_err = res.get('b_err', 0.0)
        self.ax.errorbar([b], [c], xerr=[b_err], fmt='o', color='#004c99', ecolor='#004c99', capsize=4, label='LCPD')
        self.ax.legend()
        text = (
            f"a = {res['a']:.4g} +/- {res['a_err']:.2g}\n"
            f"b (LCPD) = {res['b']*1000:.2f} +/- {res['b_err']*1000:.2f} mV\n"
            f"c = {res['c']:.4g} +/- {res['c_err']:.2g} Hz\n"
            f"RMSE = {res['rmse']:.4g}"
        )
        self.fit_result_label.setText(text)

    def _on_fit_clicked(self):
        name = self.channel_combo.currentText()
        if not name or name not in self.channels:
            return
        try:
            res = fit_parabola_bias(self.V, self.channels[name])
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Fit failed", str(e))
            return
        res['channel'] = name
        self._last_fit_result = res
        self._plot_selected_channel()

    def _update_fit_button(self):
        enable = bool(self.channel_combo.count() and self.V.size)
        self.fit_btn.setEnabled(enable)

class CustomFilterDialog(QtWidgets.QDialog):
    """Dialog to assemble custom filter pipelines."""
    def __init__(self, parent=None, base_image=None, apply_step_func=None):
        super().__init__(parent)
        self.setWindowTitle("Custom filter pipeline")
        self.resize(460, 480)
        self.base_image = base_image
        self.apply_step = apply_step_func
        self._pipeline = []
        layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.filter_combo = QtWidgets.QComboBox()
        for key, info in FILTER_DEFINITIONS.items():
            self.filter_combo.addItem(info['label'], key)
        form.addRow("Filter", self.filter_combo)
        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(["both","row","col"])
        form.addRow("Axis", self.axis_combo)
        self.sigma_spin = QtWidgets.QDoubleSpinBox()
        self.sigma_spin.setRange(0.1, 50.0); self.sigma_spin.setSingleStep(0.1); self.sigma_spin.setValue(2.0)
        form.addRow("Sigma", self.sigma_spin)
        layout.addLayout(form)
        btn_row = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("Add step")
        remove_btn = QtWidgets.QPushButton("Remove selected")
        btn_row.addWidget(add_btn); btn_row.addWidget(remove_btn)
        layout.addLayout(btn_row)
        self.pipeline_list = QtWidgets.QListWidget()
        layout.addWidget(self.pipeline_list, 1)
        self.preview_cb = QtWidgets.QCheckBox("Preview on current image")
        layout.addWidget(self.preview_cb)
        self.preview_label = QtWidgets.QLabel("Preview unavailable")
        self.preview_label.setFixedHeight(160)
        self.preview_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(self.preview_label)
        name_row = QtWidgets.QHBoxLayout()
        name_row.addWidget(QtWidgets.QLabel("Name prefix:"))
        self.name_edit = QtWidgets.QLineEdit("Custom")
        name_row.addWidget(self.name_edit)
        layout.addLayout(name_row)
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btn_box)
        add_btn.clicked.connect(self._on_add_step)
        remove_btn.clicked.connect(self._on_remove_step)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        self.preview_cb.toggled.connect(self._update_preview)

    def _current_step(self):
        key = self.filter_combo.currentData()
        params = {}
        if key == 'flatten':
            params['axis'] = self.axis_combo.currentText()
        if key in ('highpass','lowpass'):
            params['sigma'] = float(self.sigma_spin.value())
        return {'key': key, 'params': params}

    def _on_add_step(self):
        step = self._current_step()
        label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
        self._pipeline.append(step)
        self.pipeline_list.addItem(f"{len(self._pipeline)}. {label}")
        self._update_preview()

    def _on_remove_step(self):
        row = self.pipeline_list.currentRow()
        if row >= 0:
            self.pipeline_list.takeItem(row)
            del self._pipeline[row]
            self.pipeline_list.clear()
            for idx, step in enumerate(self._pipeline, 1):
                label = FILTER_DEFINITIONS.get(step['key'], {}).get('label', step['key'])
                self.pipeline_list.addItem(f"{idx}. {label}")
            self._update_preview()

    def _update_preview(self):
        if not self.preview_cb.isChecked() or self.base_image is None or not self.apply_step:
            self.preview_label.setText("Preview unavailable")
            self.preview_label.setPixmap(QtGui.QPixmap())
            return
        arr = np.asarray(self.base_image, dtype=float)
        for step in self._pipeline:
            arr = self.apply_step(arr, step)
        qimg = array_to_qimage(arr)
        pix = QtGui.QPixmap.fromImage(qimg).scaled(self.preview_label.width(), self.preview_label.height(),
                                                   QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.preview_label.setPixmap(pix)
        self.preview_label.setText("")

    def pipeline_steps(self):
        return list(self._pipeline)

    def pipeline_label(self):
        return self.name_edit.text().strip() or "Custom"

class BatchExportSignals(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int, str)
    finished = QtCore.pyqtSignal(list, list, bool)

class BatchExportWorker(QtCore.QRunnable):
    def __init__(self, parent, paths, config, out_dir):
        super().__init__()
        self.parent = parent
        self.paths = [str(p) for p in paths]
        self.config = config
        self.out_dir = Path(out_dir)
        self.signals = BatchExportSignals()
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        saved = []
        errors = []
        total = len(self.paths)
        for idx, path in enumerate(self.paths, 1):
            if self._cancelled:
                break
            try:
                result = self.parent.render_and_save_file_using_config(Path(path), self.config, self.out_dir)
                saved.extend(result)
            except Exception as e:
                errors.append(f"{Path(path).name}: {e}")
            self.signals.progress.emit(idx, total, path)
        self.signals.finished.emit(saved, errors, self._cancelled)


class _SpectroFitWorker(QtCore.QObject):
    finished = QtCore.pyqtSignal(list, list)

    def __init__(self, specs, channel):
        super().__init__()
        self.specs = list(specs)
        self.channel = channel

    def run(self):
        results = []
        logs = []
        for spec in self.specs:
            name = Path(spec['path']).name
            V = np.asarray(spec.get('V', []), dtype=float)
            channels = spec.get('channels') or {}
            data = channels.get(self.channel)
            if data is None or not V.size:
                logs.append(f"{name}: channel '{self.channel}' unavailable")
                continue
            try:
                res = fit_parabola_bias(V, data)
                res['spec'] = spec
                results.append(res)
                logs.append(f"{name}: fit ok (RMSE {res['rmse']:.3g})")
            except Exception as e:
                logs.append(f"{name}: {e}")
        self.finished.emit(results, logs)


# ------------------------------------------
# NEW: background thumbnail worker + signals
# ------------------------------------------
class ThumbSignals(QtCore.QObject):
    thumb_ready = QtCore.pyqtSignal(object)  # emits (data_key, cmap, QImage)


class ThumbWorker(QtCore.QRunnable):
    def __init__(self, data_key, thumb_arr, cmap_name):
        super().__init__()
        self.data_key = data_key
        self.thumb_arr = np.asarray(thumb_arr, dtype=np.float32)
        self.cmap_name = cmap_name

    def run(self):
        try:
            qimg = array_to_qimage(self.thumb_arr, cmap_name=self.cmap_name)
            thumb_signals_global.thumb_ready.emit((self.data_key, self.cmap_name, qimg))
        except Exception:
            pass


thumb_signals_global = ThumbSignals()


class SpectroscopyCompareDialog(QtWidgets.QDialog):
    """Modern comparison UI for spectroscopy overlays and fitting."""
    def __init__(self, specs, parent=None):
        super().__init__(parent)
        self.specs = list(specs)
        self._line_map = {}
        self._legend_map = {}
        self._fit_results = {}
        self._fit_thread = None
        self._fit_worker = None
        self._popup_refs = []
        self.setWindowTitle("Spectroscopy comparison")
        self.resize(1250, 640)
        self._build_ui()
        self._populate_list()
        self._populate_channels()
        self._update_plot()

    def _build_ui(self):
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_layout = QtWidgets.QVBoxLayout()
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        # Left panel: filter + list
        left = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left)
        left_layout.setContentsMargins(4,4,4,4)
        self.filter_edit = QtWidgets.QLineEdit()
        self.filter_edit.setPlaceholderText("Filter spectra...")
        self.filter_edit.textChanged.connect(self._apply_filter)
        self.spec_list = QtWidgets.QListWidget()
        self.spec_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.spec_list.itemChanged.connect(self._on_item_check_changed)
        self.spec_list.itemSelectionChanged.connect(self._on_list_selection_changed)
        self.spec_list.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.spec_list.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.spec_list.customContextMenuRequested.connect(self._on_list_context_menu)
        left_layout.addWidget(self.filter_edit)
        left_layout.addWidget(self.spec_list, 1)
        splitter.addWidget(left)
        splitter.setStretchFactor(0, 0)

        # Center panel: plot + status
        center = QtWidgets.QWidget()
        center_layout = QtWidgets.QVBoxLayout(center)
        center_layout.setContentsMargins(4,4,4,4)
        self.fig = Figure(figsize=(5,4))
        self.canvas = FigureCanvas(self.fig)
        self.ax = self.fig.add_subplot(111)
        self.ax.grid(True, alpha=0.2)
        center_layout.addWidget(self.canvas, 1)
        self.status_label = QtWidgets.QLabel("0 selected / 0 total")
        center_layout.addWidget(self.status_label)
        splitter.addWidget(center)
        splitter.setStretchFactor(1, 2)
        self.canvas.mpl_connect('pick_event', self._on_legend_pick)

        # Right panel: controls + results
        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(6,6,6,6)
        channel_row = QtWidgets.QHBoxLayout()
        channel_row.addWidget(QtWidgets.QLabel("Channel:"))
        self.channel_combo = QtWidgets.QComboBox()
        self.channel_combo.currentTextChanged.connect(self._on_channel_changed)
        channel_row.addWidget(self.channel_combo, 1)
        right_layout.addLayout(channel_row)

        btn_row = QtWidgets.QHBoxLayout()
        self.fit_selected_btn = QtWidgets.QPushButton("Fit selected (F)")
        self.fit_all_btn = QtWidgets.QPushButton("Fit all")
        self.export_btn = QtWidgets.QPushButton("Export CSV")
        btn_row.addWidget(self.fit_selected_btn)
        btn_row.addWidget(self.fit_all_btn)
        btn_row.addWidget(self.export_btn)
        right_layout.addLayout(btn_row)
        self.fit_selected_btn.clicked.connect(self._fit_selected)
        self.fit_all_btn.clicked.connect(self._fit_all)
        self.export_btn.clicked.connect(self._export_csv)

        QtWidgets.QShortcut(QtGui.QKeySequence("F"), self, activated=self._fit_selected)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+E"), self, activated=self._export_csv)

        self.options_toggle = QtWidgets.QToolButton()
        self.options_toggle.setText("Fit options")
        self.options_toggle.setCheckable(True)
        self.options_toggle.setChecked(False)
        self.options_toggle.setArrowType(QtCore.Qt.RightArrow)
        self.options_toggle.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.options_toggle.toggled.connect(self._on_options_toggled)
        right_layout.addWidget(self.options_toggle)

        self.options_body = QtWidgets.QWidget()
        form = QtWidgets.QFormLayout(self.options_body)
        self.degree_spin = QtWidgets.QSpinBox()
        self.degree_spin.setRange(2, 2)
        self.degree_spin.setValue(2)
        self.degree_spin.setEnabled(False)
        form.addRow("Degree", self.degree_spin)
        self.mask_min = QtWidgets.QDoubleSpinBox(); self.mask_min.setRange(-1e6, 1e6); self.mask_min.setSuffix(" V")
        self.mask_max = QtWidgets.QDoubleSpinBox(); self.mask_max.setRange(-1e6, 1e6); self.mask_max.setSuffix(" V")
        form.addRow("Mask min", self.mask_min)
        form.addRow("Mask max", self.mask_max)
        self.options_body.setVisible(False)
        right_layout.addWidget(self.options_body)

        self.results_table = QtWidgets.QTableWidget(0, 10)
        self.results_table.setHorizontalHeaderLabels(["File","X (nm)","Y (nm)","a","±a","b (mV)","±b","c (Hz)","±c","RMSE"])
        self.results_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.results_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.results_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.results_table.itemSelectionChanged.connect(self._on_table_selection)
        self.results_table.itemDoubleClicked.connect(self._on_table_double_clicked)
        self.results_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._on_table_context_menu)
        right_layout.addWidget(self.results_table, 1)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(100)
        right_layout.addWidget(self.log)

        splitter.addWidget(right)
        splitter.setStretchFactor(2, 1)

    def _populate_list(self):
        self.spec_list.blockSignals(True)
        self.spec_list.clear()
        self._item_map = {}
        for spec in self.specs:
            item = QtWidgets.QListWidgetItem(Path(spec['path']).name)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsSelectable)
            item.setCheckState(QtCore.Qt.Checked)
            item.setData(QtCore.Qt.UserRole, spec)
            self.spec_list.addItem(item)
            self._item_map[str(Path(spec['path']))] = item
        self.spec_list.blockSignals(False)

    def _populate_channels(self):
        channels = sorted({name for spec in self.specs for name in (spec.get('channels') or {}).keys()})
        self.channel_combo.blockSignals(True)
        self.channel_combo.clear()
        for name in channels:
            self.channel_combo.addItem(name)
        if channels:
            self.channel_combo.setCurrentText('df' if 'df' in channels else channels[0])
        self.channel_combo.blockSignals(False)

    def _apply_filter(self, text):
        text = text.lower()
        for i in range(self.spec_list.count()):
            item = self.spec_list.item(i)
            item.setHidden(text not in item.text().lower())
        self._update_status()

    def _checked_items(self):
        return [self.spec_list.item(i) for i in range(self.spec_list.count())
                if self.spec_list.item(i).checkState() == QtCore.Qt.Checked and not self.spec_list.item(i).isHidden()]

    def _selected_items(self):
        return [item for item in self._checked_items() if item.isSelected()]

    def _on_channel_changed(self):
        self._fit_results = {}
        self._populate_results_table()
        self._update_plot()

    def _on_item_check_changed(self):
        self._update_plot()

    def _on_list_selection_changed(self):
        self._update_plot()

    def _update_plot(self):
        channel = self.channel_combo.currentText()
        self.ax.clear()
        self.ax.grid(True, alpha=0.2)
        self._line_map.clear()
        self._legend_map.clear()
        selected_ids = {str(Path(item.data(QtCore.Qt.UserRole)['path'])) for item in self._selected_items()}
        colors = itertools.cycle(matplotlib.cm.get_cmap('tab10').colors)
        plotted = 0
        for item in self._checked_items():
            spec = item.data(QtCore.Qt.UserRole)
            spec_id = str(Path(spec['path']))
            channels = spec.get('channels') or {}
            data = channels.get(channel)
            V = np.asarray(spec.get('V', []), dtype=float)
            if data is None or not V.size:
                continue
            color = next(colors)
            highlight = spec_id in selected_ids or not selected_ids
            line, = self.ax.plot(V, data, color=color, lw=2.4 if highlight else 1.2,
                                 alpha=1.0 if highlight else 0.4, label=Path(spec['path']).name)
            self._line_map[spec_id] = line
            plotted += 1
            if spec_id in self._fit_results:
                self._draw_fit_for_spec(spec_id, color)
        if plotted == 0:
            self.ax.text(0.5,0.5,"No data for selected items", ha='center', va='center', transform=self.ax.transAxes)
        else:
            legend = self.ax.legend(loc='best', fontsize=8)
            if legend:
                legend.set_draggable(True)
                for leg_line, text in zip(legend.get_lines(), legend.get_texts()):
                    leg_line.set_picker(True)
                    spec_id = self._spec_id_by_name(text.get_text())
                    if spec_id:
                        self._legend_map[leg_line] = spec_id
        self.ax.set_xlabel("Bias (V)")
        self.ax.set_ylabel(channel)
        self.canvas.draw_idle()
        self._update_status(plotted)

    def _draw_fit_for_spec(self, spec_id, color):
        res = self._fit_results.get(spec_id)
        if not res:
            return
        spec = res.get('spec')
        V = np.asarray(spec.get('V', []), dtype=float)
        if not V.size:
            return
        x_dense = np.linspace(np.nanmin(V), np.nanmax(V), 400)
        self.ax.plot(x_dense, res['func'](x_dense), '--', color=color, lw=1.2)
        b = res['b']; c = res['c']; be = res.get('b_err', 0.0)
        self.ax.errorbar([b], [c], xerr=[be], fmt='o', color=color, ecolor=color, capsize=3)

    def _spec_id_by_name(self, name):
        for spec in self.specs:
            if Path(spec['path']).name == name:
                return str(Path(spec['path']))
        return None

    def _on_legend_pick(self, event):
        spec_id = self._legend_map.get(event.artist)
        if not spec_id:
            return
        line = self._line_map.get(spec_id)
        if not line:
            return
        visible = not line.get_visible()
        line.set_visible(visible)
        event.artist.set_alpha(1.0 if visible else 0.2)
        self.canvas.draw_idle()

    def _update_status(self, plotted=None):
        total = sum(1 for i in range(self.spec_list.count()) if not self.spec_list.item(i).isHidden())
        checked = len(self._checked_items())
        text = f"{checked} selected / {total} total"
        if plotted is not None:
            text += f" — showing {plotted}"
        self.status_label.setText(text)

    def _show_popup_for_spec(self, spec):
        dlg = SpectroscopyPopup(spec, parent=self)
        dlg.show()
        self._popup_refs.append(dlg)

    def _on_item_double_clicked(self, item):
        self._show_popup_for_spec(item.data(QtCore.Qt.UserRole))

    def _on_list_context_menu(self, pos):
        item = self.spec_list.itemAt(pos)
        if not item:
            return
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        if menu.exec_(self.spec_list.mapToGlobal(pos)) == act:
            self._show_popup_for_spec(item.data(QtCore.Qt.UserRole))

    def _on_table_context_menu(self, pos):
        row = self.results_table.indexAt(pos).row()
        if row < 0:
            return
        spec_id = self.results_table.item(row,0).data(QtCore.Qt.UserRole)
        menu = QtWidgets.QMenu(self)
        act = menu.addAction("Open popup")
        if menu.exec_(self.results_table.mapToGlobal(pos)) == act:
            spec = self._spec_by_id(spec_id)
            if spec:
                self._show_popup_for_spec(spec)

    def _on_table_double_clicked(self, item):
        spec_id = self.results_table.item(item.row(),0).data(QtCore.Qt.UserRole)
        spec = self._spec_by_id(spec_id)
        if spec:
            self._show_popup_for_spec(spec)

    def _on_table_selection(self):
        row = self.results_table.currentRow()
        if row < 0:
            return
        spec_id = self.results_table.item(row,0).data(QtCore.Qt.UserRole)
        item = self._item_map.get(spec_id)
        if item:
            self.spec_list.setCurrentItem(item, QtCore.QItemSelectionModel.SelectCurrent)
            self._update_plot()

    def _spec_by_id(self, spec_id):
        for spec in self.specs:
            if str(Path(spec['path'])) == spec_id:
                return spec
        return None

    def _fit_selected(self):
        items = self._selected_items() or self._checked_items()
        self._start_fit([item.data(QtCore.Qt.UserRole) for item in items])

    def _fit_all(self):
        self._start_fit([item.data(QtCore.Qt.UserRole) for item in self._checked_items()])

    def _start_fit(self, specs):
        if not specs or self._fit_thread:
            if not specs:
                self._log("Nothing to fit.")
            return
        channel = self.channel_combo.currentText()
        self._set_busy(True, f"Fitting {len(specs)} spectra...")
        self._fit_worker = _SpectroFitWorker(specs, channel)
        self._fit_thread = QtCore.QThread(self)
        self._fit_worker.moveToThread(self._fit_thread)
        self._fit_thread.started.connect(self._fit_worker.run)
        self._fit_worker.finished.connect(self._on_fit_finished)
        self._fit_worker.finished.connect(self._fit_thread.quit)
        self._fit_thread.finished.connect(self._cleanup_fit_thread)
        self._fit_thread.start()

    def _cleanup_fit_thread(self):
        self._fit_thread.deleteLater()
        self._fit_thread = None
        self._fit_worker = None
        self._set_busy(False, "Fit ready.")

    def _on_fit_finished(self, results, logs):
        for msg in logs:
            self._log(msg)
        for res in results:
            spec = res.get('spec')
            if spec:
                self._fit_results[str(Path(spec['path']))] = res
        self._populate_results_table()
        self._update_plot()

    def _populate_results_table(self):
        rows = []
        for spec_id, res in self._fit_results.items():
            spec = res.get('spec')
            if not spec:
                continue
            xs = spec.get('x')
            ys = spec.get('y')
            rows.append((spec_id, Path(spec['path']).name,
                         "n/a" if xs is None else f"{xs:.1f}",
                         "n/a" if ys is None else f"{ys:.1f}",
                         f"{res['a']:.4g}", f"{res['a_err']:.2g}",
                         f"{res['b']:.2f}", f"{res['b_err']:.2f}",
                         f"{res['c']:.4g}", f"{res['c_err']:.2g}",
                         f"{res['rmse']:.4g}"))
        self.results_table.setRowCount(len(rows))
        for r, data in enumerate(rows):
            spec_id, name, xval, yval, a, ae, b, be, c, ce, rmse = data
            values = [name, xval, yval, a, ae, b, be, c, ce, rmse]
            for col, val in enumerate(values):
                item = QtWidgets.QTableWidgetItem(val)
                if col == 0:
                    item.setData(QtCore.Qt.UserRole, spec_id)
                self.results_table.setItem(r, col, item)

    def _export_csv(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export CSV", "spectroscopy_fit.csv", "CSV Files (*.csv)")
        if not path:
            return
        headers = ["File","X (nm)","Y (nm)","a","da","b (mV)","db","c (Hz)","dc","RMSE"]
        with open(path, 'w', newline='') as f:
            f.write(",".join(headers) + "\n")
            for row in range(self.results_table.rowCount()):
                vals = [self.results_table.item(row, col).text() if self.results_table.item(row, col) else ""
                        for col in range(self.results_table.columnCount())]
                f.write(",".join(vals) + "\n")
        self._log(f"Exported to {path}")

    def _set_busy(self, busy, message):
        self.fit_selected_btn.setEnabled(not busy)
        self.fit_all_btn.setEnabled(not busy)
        self.export_btn.setEnabled(not busy)
        if busy:
            self.status_label.setText(message)

    def _on_options_toggled(self, checked):
        self.options_toggle.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.options_body.setVisible(checked)

    def _log(self, text):
        self.log.appendPlainText(text)

# ---------------- Detection helpers ----------------

def _detect_dtype_for_file(path, expected_pixels):
    """
    Choose a dtype candidate for a binary channel file by filesize heuristics.
    Returns a numpy dtype (e.g. np.int16) or None on failure.
    This DOES NOT read the whole file; it only inspects the filesize.
    """
    candidate = [np.int16, np.uint16, np.int32, np.uint32, np.int64, np.float32, np.float64, np.uint8]
    filesize = Path(path).stat().st_size
    for dt in candidate:
        s = np.dtype(dt).itemsize
        # accept if filesize is at least expected_pixels * itemsize (some files may have padding)
        if filesize >= expected_pixels * s and (filesize % s == 0 or (filesize // s) >= expected_pixels):
            return dt
    # fallback: try float32
    return np.float32

def _sample_topo_std_from_file(header, fd, txt_path, n_samples=32):
    """
    Sample a few random pixels from the binary file indicated by fd (a FileDesc dict).
    Returns (std_in_nm, mode_val_nm) or raises on error.
    Uses np.memmap so it doesn't load the entire image into RAM.
    Assumes xPixel and yPixel are in header.
    """
    fname = fd.get('FileName')
    if not fname:
        raise IOError("No FileName in FileDesc")
    binpath = Path(txt_path).parent / fname
    xpix = int(header.get('xPixel', 128))
    ypix = int(header.get('yPixel', xpix))
    expected = int(xpix) * int(ypix)
    dt = _detect_dtype_for_file(binpath, expected)
    if dt is None:
        raise IOError("Cannot detect dtype for %s" % str(binpath))
    # memmap the file as 1D of expected length (mode='r'); memmap avoids loading all into RAM
    try:
        mm = np.memmap(str(binpath), dtype=dt, mode='r')
    except Exception as e:
        raise
    # ensure we have at least expected elements
    if mm.size < expected:
        raise IOError("Binary size smaller than expected pixels for %s" % str(binpath))
    # shape view (do not copy)
    view1d = mm[:expected]
    total = view1d.size
    n = min(n_samples, total)
    if n <= 0:
        raise ValueError("No pixels available for sampling")
    idx = np.random.randint(0, total, size=n)
    vals = np.asarray(view1d[idx], dtype=np.float64)
    # apply scale + offset if present (FileDesc has Scale/Offset keys)
    scale = float(fd.get('Scale', 1.0) or 1.0)
    offset = float(fd.get('Offset', 0.0) or 0.0)
    vals = vals * scale + offset
    # physical unit handling: try to decide if value is already in nm; if PhysUnit mentions 'nm' assume nm
    phys = (fd.get('PhysUnit','') or '').lower()
    if 'nm' in phys:
        vals_nm = vals
    else:
        # no robust conversion known -> assume values are in nm already (preserve legacy behavior)
        vals_nm = vals
    std_nm = float(np.nanstd(vals_nm))
    # approximate mode from small sample (for abs_z estimation): use median of samples as cheap proxy
    mode_est = float(np.median(vals_nm))
    return std_nm, mode_est
def header_indicates_constant(header):
    """Return 'CH' or 'CC' or None based on header textual indicators."""
    if not header: return None
    # combine keys & values into lowercase string for searching
    combined = " ".join([f"{k}:{str(v)}" for k,v in header.items()]).lower()
    # look for phrases that indicate constant-height or constant-current
    if any(kw in combined for kw in ('constant-height', 'constant height', 'constantheight', 'constheight', 'mode: constant', 'scanmode: constant', 'operationmode: constant')):
        return 'CH'
    if any(kw in combined for kw in ('constant-current', 'constant current', 'constantcurrent', 'feedback: current', 'mode: current', 'scanmode: current')):
        return 'CC'
    # some vendors write "constant" but not long form; look for 'constant' plus 'height'/'current' nearby
    if 'constant' in combined:
        if 'height' in combined: return 'CH'
        if 'current' in combined: return 'CC'
    return None


def _find_topography_channel(fds):
    """
    Return index of the TRUE topographic channel or None if not found.
    Priority:
      1. Caption contains 'topo' (case insensitive)
      2. FileName contains 'topo'
      3. Caption contains 'height' but avoid sensor/feedback/setpoint channels
      4. Otherwise return None
    """
    if not fds:
        return None
    # step 1: caption has "topo"
    for i, fd in enumerate(fds):
        cap = (fd.get("Caption","") or "").lower()
        if "topo" in cap:
            return i
    # step 2: filename has "topo"
    for i, fd in enumerate(fds):
        fn = (fd.get("FileName","") or "").lower()
        if "topo" in fn:
            return i
    # step 3: caption has "height" but avoid sensor/feedback channels
    for i, fd in enumerate(fds):
        cap = (fd.get("Caption","") or "").lower()
        if "height" in cap and "sensor" not in cap and "feedback" not in cap and "setpoint" not in cap:
            return i
    return None

def filedesc_indicates_current_or_topo(fd):
    """Return 'current' or 'topo' or None based on FileDesc keys/Caption/FileName."""
    s = " ".join([str(fd.get(k,'')) for k in ('Caption','FileName','PhysUnit')]).lower()
    if any(tok in s for tok in ('current','i ', 'it_', 'lia', 'amp','pa','a ')):
        return 'current'
    if any(tok in s for tok in ('topo','height','z ')):
        return 'topo'
    return None

# ---------------- Main widget ----------------
class SXMGridViewer(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SXM Grid Viewer")
        self.resize(1250, 840)

        self.config = load_config()
        self.last_dir = Path(self.config.get("last_dir", str(Path.cwd())))
        self.last_channel_index = int(self.config.get("last_channel_index", 0))
        self.thumb_cmap = self.config.get("thumbnail_cmap", "viridis")
        self.preview_cmap = self.config.get("preview_cmap", "viridis")
        self.spec_folder_path = Path(self.config.get("spectra_folder", str(self.last_dir)))
        self.show_spectra = bool(self.config.get("show_spectra", True))
        self.tags = self.config.get("tags", {})  # persistent tags: {path: {"tag":"constant-height","abs_z_pm":int,...}}

        self.files = []
        self.headers = {}
        self.thumb_cache = {}
        self._thumb_data_cache = {}
        self._topo_stats_cache = {}
        self._channel_data_cache = OrderedDict()
        self._channel_cache_lock = threading.Lock()
        self._filtered_channel_cache = OrderedDict()
        self._filtered_cache_lock = threading.Lock()
        self.per_file_channel_cmap = {}
        self.last_preview = None
        self.spectros = []
        self.spectros_by_image = defaultdict(list)
        self._spectro_cache = {}
        self.image_time_index = {}
        self._spectro_popups = []
        self._multi_spectro_popups = []
        self._multi_spec_selection = []
        self._multi_spec_selection_keys = set()
        self.thumb_multi_select = set()
        self._batch_export_progress = None
        self._batch_export_worker = None
        self.virtual_copies = {}
        self.virtual_copy_order = []
        self.thumbnail_filters = {}
        self.header_cache = load_header_cache()
        self._header_cache_dirty = False
        # Deprecated: previously stored concrete arrays for extra views
        # self.added_views kept for backward compatibility but not used for rendering
        self.added_views = []
        # New: store extra view specifications to rebuild per selected file
        # Each spec: { 'caption': str, 'index': int, 'cmap': str }
        self.extra_view_specs = []
        # Thumbnail helpers: mapping from file path -> container widget for selection styling
        self.thumb_widgets = {}
        # NEW: background thumbnail worker integration
        try:
            self.thumb_pool = QtCore.QThreadPool.globalInstance()
        except Exception:
            self.thumb_pool = None
        try:
            thumb_signals_global.thumb_ready.connect(self._on_thumb_ready)
        except Exception:
            pass
        self.selected_file_for_thumbs = None

        # fonts
        base_font = QtGui.QFont("Segoe UI", 11)
        bold_font = QtGui.QFont("Segoe UI", 11, QtGui.QFont.Bold)
        meta_font = QtGui.QFont("Consolas", 13)
        try:
            app = QtWidgets.QApplication.instance()
            if app is not None:
                app.setFont(base_font)
        except Exception:
            pass

        # UI: left controls + meta + inspector; right thumbs + preview
        left_v = QtWidgets.QVBoxLayout(); left_v.setSpacing(8)
        path_h = QtWidgets.QHBoxLayout()
        self.path_le = QtWidgets.QLineEdit(str(self.last_dir))
        self.open_btn = QtWidgets.QPushButton("Open folder")
        path_h.addWidget(self.path_le); path_h.addWidget(self.open_btn); left_v.addLayout(path_h)

        spec_path_h = QtWidgets.QHBoxLayout()
        spec_path_h.addWidget(QtWidgets.QLabel("Spectra folder:"))
        self.spec_folder_le = QtWidgets.QLineEdit(str(self.spec_folder_path))
        self.spec_folder_le.setPlaceholderText("Defaults to SXM folder")
        self.spec_folder_btn = QtWidgets.QPushButton("Browse")
        spec_path_h.addWidget(self.spec_folder_le, 1)
        spec_path_h.addWidget(self.spec_folder_btn)
        left_v.addLayout(spec_path_h)

        controls_h = QtWidgets.QHBoxLayout()
        self.channel_label = QtWidgets.QLabel("Channel:")
        self.channel_label.setFont(bold_font)
        self.channel_dropdown = QtWidgets.QComboBox(); self.channel_dropdown.setMinimumWidth(380)
        self.thumb_cmap_combo = QtWidgets.QComboBox(); self.preview_cmap_combo = QtWidgets.QComboBox()
        
        # populate colormap combos with all available matplotlib colormaps and icons
        try:
            cmap_list = sorted(colormaps.keys())
        except Exception:
            cmap_list = ['viridis','plasma','inferno','magma','cividis','gray','hot','coolwarm','turbo']
        for m in cmap_list:
            try:
                icon = _colormap_icon(m, width=96, height=14)
            except Exception:
                icon = QIcon()
            self.thumb_cmap_combo.addItem(icon, m)
            self.preview_cmap_combo.addItem(icon, m)

        self.thumb_cmap_combo.setCurrentText(self.thumb_cmap); self.preview_cmap_combo.setCurrentText(self.preview_cmap)
        controls_h.addWidget(self.channel_label); controls_h.addWidget(self.channel_dropdown)
        controls_h.addWidget(QtWidgets.QLabel("Thumb cmap:")); controls_h.addWidget(self.thumb_cmap_combo)
        controls_h.addWidget(QtWidgets.QLabel("Preview cmap:")); controls_h.addWidget(self.preview_cmap_combo)
        # Dark mode toggle
        self.dark_mode = bool(self.config.get('dark_mode', False))
        self.dark_mode_cb = QtWidgets.QCheckBox('Dark mode')
        self.dark_mode_cb.setChecked(self.dark_mode)
        controls_h.addWidget(self.dark_mode_cb)
        left_v.addLayout(controls_h)

        self.meta_box = QtWidgets.QTextEdit(); self.meta_box.setReadOnly(True); self.meta_box.setFont(meta_font); self.meta_box.setMinimumWidth(380)
        left_v.addWidget(self.meta_box, 1)

        tag_h = QtWidgets.QHBoxLayout()
        self.tag_ch_btn = QtWidgets.QPushButton("Tag as CH")
        self.tag_cc_btn = QtWidgets.QPushButton("Tag as CC")
        self.untag_btn = QtWidgets.QPushButton("Untag")
        
        # Purge config button
        self.purge_config_btn = QtWidgets.QPushButton('Purge config')
        tag_h.addWidget(self.purge_config_btn)
        tag_h.addWidget(self.tag_ch_btn); tag_h.addWidget(self.tag_cc_btn); tag_h.addWidget(self.untag_btn)
        left_v.addLayout(tag_h)

        # NOTE:
        # Removed the "File channels (selected file)" inspector (list + cmap + "Show channel" button).
        # That UI duplicated functionality already provided via the thumbnails and the "Add channel view"
        # dialog. We rely on thumbnails + Add dialog going forward, so we keep the left panel slimmer.

        left_w = QtWidgets.QWidget(); left_w.setLayout(left_v)

        # Right panel with splitter for thumbnails/preview
        title_lbl = QtWidgets.QLabel("Selected channel"); title_lbl.setFont(bold_font)
        self.scroll = QtWidgets.QScrollArea(); self.thumb_container = QtWidgets.QWidget(); self.thumb_layout = QtWidgets.QGridLayout(); self.thumb_layout.setSpacing(14)
        self.thumb_container.setLayout(self.thumb_layout); self.scroll.setWidgetResizable(True); self.scroll.setWidget(self.thumb_container)
        thumbs_panel = QtWidgets.QWidget()
        thumbs_panel_layout = QtWidgets.QVBoxLayout(); thumbs_panel_layout.setContentsMargins(0,0,0,0)
        thumbs_panel_layout.addWidget(title_lbl)
        thumbs_panel_layout.addWidget(self.scroll, 1)
        thumbs_toolbar = QtWidgets.QHBoxLayout()
        thumbs_toolbar.addWidget(QtWidgets.QLabel('Sort:'))
        self.thumb_sort_combo = QtWidgets.QComboBox()
        self.thumb_sort_combo.addItems(['Name (A-Z)', 'Date (new-old)', 'Date (old-new)', 'Tag (CH-CC-U)'])
        thumbs_toolbar.addWidget(self.thumb_sort_combo)
        thumbs_toolbar.addSpacing(8)
        thumbs_toolbar.addWidget(QtWidgets.QLabel('Filter:'))
        self.thumb_filter_combo = QtWidgets.QComboBox()
        self.thumb_filter_combo.addItems(['Name (A-Z)', 'Date (new-old)', 'Date (old-new)', 'Tag (CH-CC-U)'])
        thumbs_toolbar.addWidget(self.thumb_filter_combo)
        thumbs_panel_layout.addLayout(thumbs_toolbar)
        # restore sort/filter from config if present
        try:
            sort_label = self.config.get('thumb_sort', 'Name (A-Z)')
            if sort_label in [self.thumb_sort_combo.itemText(i) for i in range(self.thumb_sort_combo.count())]:
                self.thumb_sort_combo.setCurrentText(sort_label)
            filt_label = self.config.get('thumb_filter', 'All')
            if filt_label in [self.thumb_filter_combo.itemText(i) for i in range(self.thumb_filter_combo.count())]:
                self.thumb_filter_combo.setCurrentText(filt_label)
        except Exception:
            pass
        thumbs_panel.setLayout(thumbs_panel_layout)

        preview_panel = QtWidgets.QWidget()
        preview_panel_layout = QtWidgets.QVBoxLayout(); preview_panel_layout.setContentsMargins(0,0,0,0)
        self.preview_canvas = MultiPreviewCanvas(self, figsize=(6,5))
        self.preview_canvas.set_copy_feedback_handler(self._on_view_copied)
        preview_panel_layout.addWidget(self.preview_canvas, 1)
        extra_h = QtWidgets.QHBoxLayout()
        self.add_view_btn = QtWidgets.QPushButton("Add channel view"); self.clear_views_btn = QtWidgets.QPushButton("Clear extra views")
        self.export_pngs_btn = QtWidgets.QPushButton("Export PNGs")
        extra_h.addWidget(self.add_view_btn); extra_h.addWidget(self.clear_views_btn); extra_h.addWidget(self.export_pngs_btn)
        self.measure_profile_btn = QtWidgets.QPushButton('Measure profile')
        extra_h.addWidget(self.measure_profile_btn)
        self.show_spectra_cb = QtWidgets.QCheckBox("Show spectroscopies")
        self.show_spectra_cb.setChecked(self.show_spectra)
        extra_h.addWidget(self.show_spectra_cb)
        self.clear_spec_selection_btn = QtWidgets.QPushButton("Clear spec selection")
        extra_h.addWidget(self.clear_spec_selection_btn)
        self.spec_selection_label = QtWidgets.QLabel("Spectra selected: 0")
        font_small = QtGui.QFont("Segoe UI", 9)
        self.spec_selection_label.setFont(font_small)
        extra_h.addWidget(self.spec_selection_label)
        self.export_selected_btn = QtWidgets.QPushButton("Export selected (same view)")
        extra_h.addWidget(self.export_selected_btn)
        preview_panel_layout.addLayout(extra_h)
        preview_panel.setLayout(preview_panel_layout)

        right_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        right_splitter.addWidget(thumbs_panel)
        right_splitter.addWidget(preview_panel)
        right_splitter.setStretchFactor(0, 3)
        right_splitter.setStretchFactor(1, 2)
        right_w = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(); right_layout.setContentsMargins(0,0,0,0)
        right_layout.addWidget(right_splitter, 1)
        right_w.setLayout(right_layout)

        main_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        main_splitter.addWidget(left_w)
        main_splitter.addWidget(right_w)
        # prefer to give the right pane more space by default
        main_splitter.setStretchFactor(0, 1)
        main_splitter.setStretchFactor(1, 3)

        # Prevent panes from collapsing to zero width when the user drags the splitter.
        # This avoids the left inspector disappearing when the user expands the thumbnails.
        try:
            main_splitter.setCollapsible(0, False)
            main_splitter.setCollapsible(1, False)
        except Exception:
            # older PyQt versions may not support setCollapsible; ignore safely
            pass

        # Ensure the left widget cannot shrink below a useful width
        try:
            left_w.setMinimumWidth(360)
        except Exception:
            pass

        # Set reasonable initial sizes (left, right). Adjust these numbers to taste.
        try:
            main_splitter.setSizes([360, 1080])
        except Exception:
            pass
        main_layout = QtWidgets.QHBoxLayout()
        main_layout.addWidget(main_splitter)
        self.setLayout(main_layout)

        # signals
        self.open_btn.clicked.connect(self.open_folder_dialog)
        self.path_le.returnPressed.connect(self.open_folder_by_path)
        self.spec_folder_btn.clicked.connect(self.on_spec_folder_browse)
        self.spec_folder_le.returnPressed.connect(self.on_spec_folder_entered)
        self.channel_dropdown.currentIndexChanged.connect(self.on_channel_dropdown_changed)
        self.thumb_cmap_combo.currentIndexChanged.connect(self.on_thumb_cmap_changed)
        self.preview_cmap_combo.currentIndexChanged.connect(self.on_preview_cmap_changed)
        self.thumb_sort_combo.currentIndexChanged.connect(self.on_thumb_sort_changed)
        self.thumb_filter_combo.currentIndexChanged.connect(self.on_thumb_filter_changed)
        # no size slider callback
        # inspector widgets removed -> no connections required here
        self.add_view_btn.clicked.connect(self.on_add_view)
        self.clear_views_btn.clicked.connect(self.on_clear_views)
        self.export_pngs_btn.clicked.connect(self.on_export_pngs)
        self.measure_profile_btn.clicked.connect(self._on_start_profile)
        self.show_spectra_cb.toggled.connect(self.on_show_spectra_toggled)
        self.clear_spec_selection_btn.clicked.connect(self.on_clear_spec_selection)
        self.export_selected_btn.clicked.connect(self.on_export_selected_same_view)
        self.tag_ch_btn.clicked.connect(lambda: self.on_manual_tag('constant-height'))
        self.tag_cc_btn.clicked.connect(lambda: self.on_manual_tag('constant-current'))
        self.untag_btn.clicked.connect(lambda: self.on_manual_tag(None))
        self.dark_mode_cb.toggled.connect(self.on_dark_mode_toggled)

        # autoload
        if self.last_dir.exists():
            QtCore.QTimer.singleShot(50, lambda: self.load_folder(self.last_dir))
        try:
            self.purge_config_btn.clicked.connect(self._on_purge_config)
        except Exception:
            pass
        # apply initial dark mode palette
        try:
            self._apply_dark_mode(self.dark_mode)
        except Exception:
            pass

    def _apply_dark_mode(self, enabled: bool):
        app = QtWidgets.QApplication.instance()
        if app is None:
            return
        if enabled:
            app.setStyle('Fusion')
            palette = QtGui.QPalette()
            palette.setColor(QtGui.QPalette.Window, QtGui.QColor(53,53,53))
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Base, QtGui.QColor(35,35,35))
            palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(53,53,53))
            palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Button, QtGui.QColor(53,53,53))
            palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
            palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42,130,218))
            palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
            app.setPalette(palette)
        else:
            app.setPalette(app.style().standardPalette())

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
        self.config['dark_mode'] = self.dark_mode; save_config(self.config)
        self._apply_dark_mode(self.dark_mode)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- folder load & auto-detection ----------
    def open_folder_dialog(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Select data folder", str(self.last_dir))
        if d:
            self.load_folder(Path(d))

    def open_folder_by_path(self):
        p = Path(self.path_le.text().strip())
        if p.exists() and p.is_dir():
            self.load_folder(p)

    def load_folder(self, folder:Path):
        folder = Path(folder)
        self.last_dir = folder
        self.path_le.setText(str(folder))
        # persist last dir early
        self.config['last_dir'] = str(folder)
        save_config(self.config)

        txts = sorted(folder.glob("*.txt"))
        self.files = txts
        self.headers.clear()
        self._invalidate_thumbnail_cache()
        self._invalidate_channel_cache()
        self.thumb_multi_select = set()
        cache_hits = 0
        cache_miss = 0
        for t in txts:
            cached = self._get_cached_header(t)
            if cached:
                hdr, fds = cached
                cache_hits += 1
            else:
                try:
                    hdr, fds = parse_header(t)
                    cache_miss += 1
                    self._store_header_cache(t, hdr, fds)
                except Exception:
                    continue
            self.headers[str(t)] = (hdr, fds)
        if cache_miss:
            self._save_header_cache()
        if not self.headers:
            self.meta_box.setPlainText("No valid .txt headers found")
            self.clear_thumbs(); return
        self._build_image_timestamp_index()

        # build channel dropdown from first header
        first_key = next(iter(self.headers))
        _, first_fds = self.headers[first_key]
        labels = []
        for idx, fd in enumerate(first_fds):
            cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            labels.append(f"{idx}: {cap}")
        max_channels = max(len(v[1]) for v in self.headers.values())
        if max_channels > len(labels):
            for idx in range(len(labels), max_channels):
                labels.append(f"{idx}: chan{idx}")

        self.channel_dropdown.blockSignals(True)
        self.channel_dropdown.clear()
        for lab in labels:
            self.channel_dropdown.addItem(lab)
            self.channel_dropdown.setItemData(self.channel_dropdown.count()-1, lab, QtCore.Qt.ToolTipRole)
        self.channel_dropdown.setMinimumWidth(380)
        if 0 <= self.last_channel_index < self.channel_dropdown.count():
            self.channel_dropdown.setCurrentIndex(self.last_channel_index)
        else:
            self.last_channel_index = 0; self.channel_dropdown.setCurrentIndex(0)
        self.channel_dropdown.blockSignals(False)

        # set cmaps
        try: self.thumb_cmap_combo.setCurrentText(self.thumb_cmap)
        except: pass
        try: self.preview_cmap_combo.setCurrentText(self.preview_cmap)
        except: pass
        # set icon sizes for cmap combos
        try:
            self.thumb_cmap_combo.setIconSize(QtCore.QSize(96, 14))
            self.preview_cmap_combo.setIconSize(QtCore.QSize(96, 14))
        except Exception:
            pass

        # auto-detect tags for files not already tagged
        self._auto_detect_tags_for_folder()

        # load spectroscopy markers referencing this folder
        self._reload_spectros(refresh=False)

        QtCore.QTimer.singleShot(0, lambda: self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex()))

    def _auto_detect_tags_for_folder(self):
        """Auto-detect CH/CC for files that are not already present in self.tags.
        This version runs sampling-only detection (fast) and executes only once per-folder per-session.
        """
        # per-session guard: run for this folder only once
        if not hasattr(self, '_auto_detect_done_folders'):
            self._auto_detect_done_folders = set()
        cur_folder = str(self.last_dir)
        if cur_folder in self._auto_detect_done_folders:
            return
        self._auto_detect_done_folders.add(cur_folder)

        # iterate files in order and analyze headers, and fall back to topo-statistic if needed
        for p in self.files:
            key = str(p)
            if key in self.tags:
                continue  # don't overwrite manual tags
            hdr, fds = self.headers.get(key, (None, None))
            # header-driven detection first
            tag = header_indicates_constant(hdr)
            if tag == 'CH':
                self.tags[key] = {'tag': 'constant-height', 'abs_z_pm': None}
                continue
            elif tag == 'CC':
                self.tags[key] = {'tag': 'constant-current'}
                continue

            # fallback: try to find a topo-like channel and sample a few pixels
            topo_idx = None
            if fds:
                topo_idx = _find_topography_channel(fds)
                if topo_idx is None and len(fds) > 0:
                    topo_idx = 0

                if topo_idx is not None:
                    fd = fds[topo_idx]
                    stats = None
                    cache_key = (key, fd.get('FileName'))
                    if not hasattr(self, '_topo_stats_cache'):
                        self._topo_stats_cache = {}
                    if cache_key in self._topo_stats_cache:
                        stats = self._topo_stats_cache[cache_key]
                    else:
                        try:
                            stats = _sample_topo_std_from_file(hdr, fd, Path(key), n_samples=32)
                            self._topo_stats_cache[cache_key] = stats
                        except Exception:
                            stats = None
                    if stats is None:
                        continue
                    std_nm, mode_est_nm = stats
                    if std_nm <= CH_STD_THRESHOLD_NM:
                        abs_pm = int(round(mode_est_nm * 1000.0)) if mode_est_nm is not None else None
                        self.tags[key] = {'tag': 'constant-height', 'abs_z_pm': abs_pm}
                        continue
                    # values varied -> treat channel as constant-current reference
                    self.tags[key] = {'tag': 'constant-current'}
                    continue


        # persist tags after the initial auto pass
        self.config['tags'] = self.tags
        save_config(self.config)

    # ---------- thumbnails population with badge overlay ----------
    def clear_thumbs(self):
        while self.thumb_layout.count():
            item = self.thumb_layout.takeAt(0); w = item.widget()
            if w: w.setParent(None)
        self.thumb_widgets = {}

    def populate_thumbnails_for_channel(self, channel_idx:int):
        self.clear_thumbs()
        max_cols = 4; row = 0; col = 0
        thumb_w = 160
        thumb_h = 120
        cmap_name = self.thumb_cmap_combo.currentText() or self.thumb_cmap
        self.meta_box.setPlainText(f"Building thumbnails for channel {channel_idx} ...")
        files_iter = list(self.files)

        # apply filter (by tag)
        filt = (self.thumb_filter_combo.currentText() if hasattr(self, 'thumb_filter_combo') else 'All')
        if filt != 'All':
            def include(path_str):
                tag = (self.tags.get(path_str, {}) or {}).get('tag', None)
                if filt == 'CH only':
                    return tag == 'constant-height'
                if filt == 'CC only':
                    return tag == 'constant-current'
                if filt == 'Untagged':
                    return tag is None
                return True
            files_iter = [t for t in files_iter if include(str(t))]

        # apply sorting
        sort_mode = (self.thumb_sort_combo.currentText() if hasattr(self, 'thumb_sort_combo') else 'Name (Aâ†’Z)')
        if sort_mode.startswith('Name'):
            files_iter.sort(key=lambda p: Path(p).name.lower())
        elif 'Date (new' in sort_mode or 'Date (old' in sort_mode:
            rev = ('new' in sort_mode)
            def sort_key_date(p):
                hdr = self.headers.get(str(p), (None, None))[0]
                dt = self._parse_header_datetime(hdr)
                return dt
            files_iter.sort(key=sort_key_date, reverse=rev)
        elif sort_mode.startswith('Tag'):
            order = {'constant-height': 0, 'constant-current': 1, None: 2}
            files_iter.sort(key=lambda p: (order.get((self.tags.get(str(p), {}) or {}).get('tag', None), 2), Path(p).name.lower()))

        for i, t in enumerate(files_iter):
            key = str(t)
            if key not in self.headers: continue
            header, fds = self.headers[key]
            marker_defs = []
            if channel_idx < 0 or channel_idx >= len(fds):
                pix = QtGui.QPixmap(thumb_w, thumb_h); pix.fill(QtGui.QColor('black'))
            else:
                fd = fds[channel_idx]; fname = fd.get("FileName")
                try:
                    data_key, thumb_arr = self._get_thumbnail_array(key, channel_idx, header, fd, thumb_w, thumb_h)
                    pix_cache_key = (tuple(data_key), cmap_name)
                    base_pix = self.thumb_cache.get(pix_cache_key)
                    if base_pix is None:
                        # --------------- DISK CACHE CHECK ---------------
                        try:
                            THUMB_DISK_CACHE_DIR.mkdir(parents=True, exist_ok=True)
                            hsh = hashlib.sha256(repr((data_key, cmap_name)).encode()).hexdigest()
                            cache_file = THUMB_DISK_CACHE_DIR / f"{hsh}.png"
                            if cache_file.exists():
                                base_pix = QtGui.QPixmap(str(cache_file)).scaled(
                                    thumb_w, thumb_h, QtCore.Qt.KeepAspectRatio, QtCore.Qt.FastTransformation)
                                self.thumb_cache[pix_cache_key] = base_pix
                        except Exception:
                            pass

                    if base_pix is None:
                        # --------------- PLACEHOLDER + ASYNC WORKER ---------------
                        base_pix = QtGui.QPixmap(thumb_w, thumb_h)
                        base_pix.fill(QtGui.QColor(40, 40, 40))
                        try:
                            worker = ThumbWorker(data_key, thumb_arr, cmap_name)
                            if self.thumb_pool is not None:
                                self.thumb_pool.start(worker)
                            else:
                                raise RuntimeError("no thread pool")
                        except Exception:
                            # fallback synchronous generation
                            try:
                                qimg = array_to_qimage(thumb_arr, cmap_name=cmap_name)
                                base_pix = QtGui.QPixmap.fromImage(qimg).scaled(
                                    thumb_w, thumb_h, QtCore.Qt.KeepAspectRatio, QtCore.Qt.FastTransformation)
                                self.thumb_cache[pix_cache_key] = base_pix
                            except Exception:
                                pass
                    pix = base_pix.copy()
                except Exception:
                    pix = QtGui.QPixmap(thumb_w, thumb_h)
                    pix.fill(QtGui.QColor('black'))
                # add badge/border if tagged
                taginfo = self.tags.get(key, {})
                if taginfo:
                    tag = taginfo.get('tag', None)
                    painter = QtGui.QPainter(pix)
                    pen = QtGui.QPen()
                    pen.setWidth(4)
                    if tag == 'constant-height':
                        pen.setColor(QtGui.QColor(0, 180, 0))  # green border
                        painter.setPen(pen)
                        painter.drawRect(2,2,pix.width()-5,pix.height()-5)
                        # small label
                        font = QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold)
                        painter.setFont(font)
                        painter.setPen(QtGui.QColor(255,255,255))
                        painter.drawText(6, 18, "CH")
                    elif tag == 'constant-current':
                        pen.setColor(QtGui.QColor(30, 100, 200))  # blue
                        painter.setPen(pen)
                        painter.drawRect(2,2,pix.width()-5,pix.height()-5)
                        font = QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold)
                        painter.setFont(font)
                        painter.setPen(QtGui.QColor(255,255,255))
                        painter.drawText(6, 18, "CC")
                    painter.end()
                if key in self.thumbnail_filters:
                    painter = QtGui.QPainter(pix)
                    painter.setBrush(QtGui.QColor(160, 16, 239, 220))
                    painter.setPen(QtGui.QPen(QtGui.QColor('black')))
                    painter.drawEllipse(pix.width()-24, 6, 18, 18)
                    painter.setPen(QtGui.QColor('white'))
                    painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
                    painter.drawText(QtCore.QRect(pix.width()-24, 6, 18, 18), QtCore.Qt.AlignCenter, "F")
                    painter.end()
                # overlay spectroscopies if available
                try:
                    xpix = int(header.get('xPixel', 128)); ypix = int(header.get('yPixel', xpix))
                    marker_defs = self._draw_spectro_markers_on_pixmap(pix, header, key, xpix, ypix)
                except Exception:
                    marker_defs = []

            lbl = QtWidgets.QLabel(); lbl.setObjectName("thumb_image_label")
            lbl.setPixmap(pix); lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setProperty("file_path", str(t)); lbl.setProperty("channel_index", int(channel_idx))
            lbl.setProperty("spec_markers", marker_defs)
            lbl.setMouseTracking(True)
            lbl.mousePressEvent = self._make_thumb_click_handler(lbl)
            lbl.mouseMoveEvent = self._make_thumb_move_handler(lbl)
            lbl.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            lbl.customContextMenuRequested.connect(lambda pos, lb=lbl: self._on_thumb_context_menu(lb, pos))
            vbox = QtWidgets.QVBoxLayout(); w = QtWidgets.QWidget()
            vbox.addWidget(lbl)
            cap = QtWidgets.QLabel(Path(t).name); cap.setAlignment(QtCore.Qt.AlignCenter); cap.setMaximumHeight(18)
            cap.setFont(QtGui.QFont("Segoe UI", 9)); vbox.addWidget(cap); w.setLayout(vbox)
            self.thumb_layout.addWidget(w, row, col)
            # remember widget for selection highlighting
            self.thumb_widgets[str(t)] = w
            try:
                if str(t) in getattr(self, 'thumb_multi_select', set()):
                    w.setStyleSheet("border: 3px solid #a010ef;")
                elif str(t) == str(getattr(self, 'selected_file_for_thumbs', None)):
                    w.setStyleSheet("border: 3px solid #5f10a3;")
                else:
                    w.setStyleSheet("")
            except Exception:
                pass
            col += 1
            if col >= max_cols:
                col = 0; row += 1
        self.meta_box.setPlainText(f"Thumbnails built for channel {channel_idx}  (thumb cmap: {cmap_name})")
    def _thumbnail_filter_signature(self, file_key):
        spec = self.thumbnail_filters.get(str(file_key))
        return _filter_signature(spec)

    def _downsample_for_thumbnail(self, arr, thumb_w, thumb_h):
        arr = np.asarray(arr, dtype=float)
        if arr.size == 0:
            return arr
        h, w = arr.shape
        if h > thumb_h or w > thumb_w:
            ys = np.linspace(0, h - 1, thumb_h).astype(int)
            xs = np.linspace(0, w - 1, thumb_w).astype(int)
            return arr[np.ix_(ys, xs)]
        return arr

    def _on_thumb_ready(self, payload):
        """
        Slot invoked when a ThumbWorker finishes converting data -> QImage.
        payload: (data_key, cmap_name, QImage)
        """
        try:
            data_key, cmap_name, qimg = payload
            if not isinstance(data_key, (tuple, list)) or len(data_key) < 6:
                return
            thumb_w, thumb_h = data_key[-2], data_key[-1]
            pix = QtGui.QPixmap.fromImage(qimg).scaled(
                thumb_w, thumb_h, QtCore.Qt.KeepAspectRatio, QtCore.Qt.FastTransformation)
            cache_key = (tuple(data_key), cmap_name)
            self.thumb_cache[cache_key] = pix
            # persist to disk for later runs (best-effort)
            try:
                THUMB_DISK_CACHE_DIR.mkdir(parents=True, exist_ok=True)
                hsh = hashlib.sha256(repr(cache_key).encode()).hexdigest()
                cache_file = THUMB_DISK_CACHE_DIR / f"{hsh}.png"
                pix.toImage().save(str(cache_file), "PNG")
            except Exception:
                pass
            file_key = str(data_key[0])
            widget = self.thumb_widgets.get(file_key)
            if widget is None:
                return
            lbl = widget.findChild(QtWidgets.QLabel, "thumb_image_label")
            if lbl is None:
                return
            # ensure we are still displaying same channel
            current_idx = lbl.property("channel_index")
            if current_idx is None or int(current_idx) != int(data_key[1]):
                return
            lbl.setPixmap(pix)
            lbl.update()
        except Exception:
            pass

    def _get_thumbnail_array(self, file_key, channel_idx, header, fd, thumb_w, thumb_h):
        filter_sig = self._thumbnail_filter_signature(file_key)
        fname = fd.get("FileName")
        if not fname:
            raise ValueError("Missing FileName for channel")
        bin_path = Path(file_key).parent / fname
        try:
            bin_mtime = bin_path.stat().st_mtime
        except Exception:
            bin_mtime = 0.0
        data_key = (file_key, channel_idx, bin_mtime, filter_sig, thumb_w, thumb_h)
        cached = self._thumb_data_cache.get(data_key)
        if cached is not None:
            return data_key, cached
        _, arr_conv = self._get_filtered_channel_array(file_key, channel_idx, header, fd)
        thumb_arr = self._downsample_for_thumbnail(arr_conv, thumb_w, thumb_h)
        self._thumb_data_cache[data_key] = thumb_arr
        return data_key, thumb_arr

    def _invalidate_thumbnail_cache(self, paths=None):
        if not paths:
            self._thumb_data_cache.clear()
            self.thumb_cache.clear()
            return
        path_set = {str(Path(p)) for p in paths}
        data_keys = [k for k in self._thumb_data_cache.keys() if k[0] in path_set]
        for k in data_keys:
            self._thumb_data_cache.pop(k, None)
        pix_keys = [k for k in self.thumb_cache.keys() if k[0][0] in path_set]
        for k in pix_keys:
            self.thumb_cache.pop(k, None)

    def _channel_cache_key(self, file_key, channel_idx, fd):
        fname = fd.get('FileName')
        if not fname:
            raise ValueError("Missing FileName for channel")
        bin_path = Path(file_key).parent / fname
        try:
            mtime = bin_path.stat().st_mtime
        except Exception:
            mtime = 0.0
        return (str(bin_path), int(channel_idx), mtime)

    def _get_channel_array(self, file_key, channel_idx, header, fd):
        key = self._channel_cache_key(file_key, channel_idx, fd)
        cache = self._channel_data_cache
        with self._channel_cache_lock:
            arr = cache.get(key)
            if arr is not None:
                cache.move_to_end(key)
                return arr
        xpix = int(header.get('xPixel', 128))
        ypix = int(header.get('yPixel', xpix))
        bin_path = Path(key[0])
        arr = read_channel_file(bin_path, xpix, ypix,
                                scale=fd.get('Scale', 1.0), offset=fd.get('Offset', 0.0))
        with self._channel_cache_lock:
            cache[key] = arr
            while len(cache) > CHANNEL_DATA_CACHE_LIMIT:
                cache.popitem(last=False)
        return arr

    def _get_filtered_channel_array(self, file_key, channel_idx, header, fd):
        file_key = str(file_key)
        channel_key = self._channel_cache_key(file_key, channel_idx, fd)
        arr = self._get_channel_array(file_key, channel_idx, header, fd)
        unit = fd.get('PhysUnit','')
        unit_final, arr_conv = normalize_unit_and_data(arr, unit)
        spec = self.thumbnail_filters.get(file_key)
        sig = _filter_signature(spec)
        cache_key = (channel_key, unit_final, sig)
        with self._filtered_cache_lock:
            cached = self._filtered_channel_cache.get(cache_key)
            if cached is not None:
                self._filtered_channel_cache.move_to_end(cache_key)
                return unit_final, cached
        result = np.asarray(arr_conv, dtype=float)
        if sig:
            result = self._apply_filter_pipeline(result, spec.get('steps', []))
        with self._filtered_cache_lock:
            self._filtered_channel_cache[cache_key] = result
            while len(self._filtered_channel_cache) > FILTERED_CACHE_LIMIT:
                self._filtered_channel_cache.popitem(last=False)
        return unit_final, result

    def _invalidate_channel_cache(self, paths=None):
        with self._channel_cache_lock:
            if not paths:
                self._channel_data_cache.clear()
                with self._filtered_cache_lock:
                    self._filtered_channel_cache.clear()
                return
            parent_dirs = {str(Path(p).parent) for p in paths}
            to_remove = [k for k in self._channel_data_cache.keys() if str(Path(k[0]).parent) in parent_dirs]
            for k in to_remove:
                self._channel_data_cache.pop(k, None)
        self._invalidate_filtered_cache(paths)

    def _invalidate_filtered_cache(self, paths=None):
        with self._filtered_cache_lock:
            if not paths:
                self._filtered_channel_cache.clear()
                return
            parent_dirs = {str(Path(p).parent) for p in paths}
            to_remove = [k for k in self._filtered_channel_cache.keys()
                        if str(Path(k[0][0]).parent) in parent_dirs]
            for k in to_remove:
                self._filtered_channel_cache.pop(k, None)

    def on_thumb_sort_changed(self, idx):
        try:
            self.config['thumb_sort'] = self.thumb_sort_combo.currentText(); save_config(self.config)
        except Exception:
            pass
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def on_thumb_filter_changed(self, idx):
        try:
            self.config['thumb_filter'] = self.thumb_filter_combo.currentText(); save_config(self.config)
        except Exception:
            pass
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    # removed size change handler

    def _parse_header_datetime(self, header):
        """Return a sortable key (float timestamp) parsed from header Date/Time if possible; otherwise 0.0.
        Accepts common formats, falls back to 0.0 on failure."""
        try:
            date = str(header.get('Date', '') or '').strip()
            time = str(header.get('Time', '') or '').strip()
            if not date and not time:
                return 0.0
            candidates = []
            if date and time:
                candidates.append(f"{date} {time}")
            if date:
                candidates.append(date)
            fmts = [
                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M:%S', '%d/%m/%Y %H:%M:%S',
                '%d-%m-%Y %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'
            ]
            for s in candidates:
                for fmt in fmts:
                    try:
                        dt = datetime.strptime(s, fmt)
                        return dt.timestamp()
                    except Exception:
                        continue
            return 0.0
        except Exception:
            return 0.0

    def _header_datetime_dt(self, header, path):
        try:
            ts = float(self._parse_header_datetime(header or {}))
            if ts <= 0:
                ts = Path(path).stat().st_mtime
            return datetime.fromtimestamp(ts)
        except Exception:
            return datetime.fromtimestamp(Path(path).stat().st_mtime)

    def _build_image_timestamp_index(self):
        self.image_time_index = {}
        self.image_meta = []
        for p in self.files:
            header, _ = self.headers.get(str(p), (None, None))
            if header is None:
                continue
            dt = self._header_datetime_dt(header, p)
            self.image_time_index[str(p)] = dt
            self.image_meta.append({'path': Path(p), 'time': dt})

    def _build_metadata_html(self, header_path:Path, header:dict, fd:dict, channel_idx:int, unit_final:str, arr_conv:np.ndarray) -> str:
        """Return HTML for the metadata pane with clearer styling and sections."""
        def esc(s):
            try:
                return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            except Exception:
                return ''
        dark = bool(getattr(self, 'dark_mode', False))
        text_color = '#e0e0e0' if dark else '#222'
        label_color = '#a0a0a0' if dark else '#555'
        filename = header_path.name
        date = header.get('Date', '')
        time = header.get('Time', '')
        bias = header.get('Bias', None); bias_unit = header.get('BiasPhysUnit', '')
        setp = header.get('SetPoint', None); setp_unit = header.get('SetPointPhysUnit', '')
        user = header.get('UserName', '')
        cap = fd.get('Caption','')
        phys_orig = fd.get('PhysUnit','')
        scale = fd.get('Scale','')
        offset = fd.get('Offset','')
        # stats
        try:
            flat = np.asarray(arr_conv).ravel()
            vmin = np.nanmin(flat); vmax = np.nanmax(flat); vmed = np.nanmedian(flat)
            stats = f"min={vmin:.6g} | max={vmax:.6g} | median={vmed:.6g}"
        except Exception:
            stats = "min/max/median: N/A"
        # tags
        taginfo = self.tags.get(str(header_path), {})
        tag_label = taginfo.get('tag', None)
        tag_chip = ''
        if tag_label == 'constant-height':
            chip_color = '#2e7d32'; chip_text = 'CH'
        elif tag_label == 'constant-current':
            chip_color = '#1565c0'; chip_text = 'CC'
        else:
            chip_color = None; chip_text = ''
        if chip_color:
            tag_chip = f"<span style='background:{chip_color};color:#fff;border-radius:10px;padding:2px 8px;font-weight:600'>" \
                       f"{chip_text}</span> <span style='color:#555'>({esc(tag_label)})</span>"
        # abs z + dzs
        ch_lines = ''
        if tag_label == 'constant-height':
            abs_pm = taginfo.get('abs_z_pm', None)
            if abs_pm is not None:
                abs_nm = abs_pm/1000.0
                ch_lines += f"<div>Const-height (abs z): <b>{abs_nm:.3f} nm</b></div>"
            dz_prev_nonch, prevname = self._dz_vs_last_before_ch(header_path)
            if dz_prev_nonch is not None:
                ch_lines += f"<div>dz vs prev non-CH (<i>{esc(prevname)}</i>): <b>{dz_prev_nonch:+.0f} pm</b> ({dz_prev_nonch/1000.0:+.3f} nm)</div>"
            dz_prev_ch, prevch_name = self._dz_vs_previous_ch(header_path)
            if dz_prev_ch is not None:
                ch_lines += f"<div>dz vs prev CH (<i>{esc(prevch_name)}</i>): <b>{dz_prev_ch:+.0f} pm</b> ({dz_prev_ch/1000.0:+.3f} nm)</div>"

        # control params
        params = {}
        def collect_params(d):
            for k,v in (d or {}).items():
                kl = str(k).lower()
                if any(tok in kl for tok in ('ki','kp','pll','ampl','amplitude','amp','setpoint','natural','natfreq','freq','f0','kpl','kipl','lockin')):
                    try:
                        params[k] = float(v)
                    except Exception:
                        params[k] = v
        collect_params(header); collect_params(fd)
        params_rows = ''.join([f"<tr><td>{esc(k)}</td><td style='text-align:right'>{esc(v)}</td></tr>" for k,v in params.items()])

        spec_section = ''
        spec_entries = self.spectros_by_image.get(str(header_path), [])
        if self.show_spectra and spec_entries:
            rows = []
            for idx, spec in enumerate(spec_entries[:6], 1):
                name = Path(spec['path']).name
                xs = spec.get('x')
                ys = spec.get('y')
                pos_txt = f"{xs:.1f}/{ys:.1f} nm" if xs is not None and ys is not None else "n/a"
                rows.append(f"<tr><td>S{idx}</td><td>{esc(name)}</td><td style='text-align:right'>{esc(pos_txt)}</td></tr>")
            if len(spec_entries) > 6:
                rows.append(f"<tr><td colspan='3' style='text-align:center;color:{label_color}'>+ {len(spec_entries)-6} more…</td></tr>")
            spec_section = f"""
            <div style='height:6px'></div>
            <div style='font-weight:600; color:{label_color}; margin-bottom:2px'>Spectroscopies ({len(spec_entries)})</div>
            <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
              {''.join(rows)}
            </table>
            """

        html = f"""
        <div style='font-family:Segoe UI, Arial; font-size:14px; color:{text_color}'>
          <div style='font-weight:600; font-size:16px; margin-bottom:4px'>{esc(filename)} {tag_chip}</div>
          <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
            <tr><td style='color:{label_color}'>Date</td><td style='text-align:right'>{esc(date) or '&nbsp;'}</td></tr>
            <tr><td style='color:{label_color}'>Time</td><td style='text-align:right'>{esc(time) or '&nbsp;'}</td></tr>
            <tr><td style='color:{label_color}'>Bias</td><td style='text-align:right'>{'' if bias is None else esc(bias)} {esc(bias_unit)}</td></tr>
            <tr><td style='color:{label_color}'>SetPoint</td><td style='text-align:right'>{'' if setp is None else esc(setp)} {esc(setp_unit)}</td></tr>
            <tr><td style='color:{label_color}'>User</td><td style='text-align:right'>{esc(user)}</td></tr>
          </table>
          <div style='height:6px'></div>
          {spec_section}
          <div style='height:6px'></div>
          <div style='font-weight:600; color:%s; margin-bottom:2px'>Channel</div>
          <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
            <tr><td style='color:{label_color}'>Index</td><td style='text-align:right'>{channel_idx}</td></tr>
            <tr><td style='color:{label_color}'>Caption</td><td style='text-align:right'>{esc(cap)}</td></tr>
            <tr><td style='color:{label_color}'>Unit (orig)</td><td style='text-align:right'>{esc(phys_orig)}</td></tr>
            <tr><td style='color:{label_color}'>Shown unit</td><td style='text-align:right'><b>{esc(unit_final)}</b></td></tr>
            <tr><td style='color:{label_color}'>Scale</td><td style='text-align:right'>{esc(scale)}</td></tr>
            <tr><td style='color:{label_color}'>Offset</td><td style='text-align:right'>{esc(offset)}</td></tr>
            <tr><td style='color:{label_color}'>Stats</td><td style='text-align:right'>{esc(stats)}</td></tr>
          </table>
          <div style='height:6px'></div>
          {ch_lines}
          {('<div style=\'height:6px\'></div><div style=\'font-weight:600; color:#333; margin-bottom:2px\'>Control params</div>' if params_rows else '')}
          {('<table style=\'width:100%; border-collapse:collapse\' cellspacing=\'0\' cellpadding=\'2\'>' + params_rows + '</table>') if params_rows else ''}
        </div>
        """
        return html

    def _refresh_thumb_selection_styles(self):
        sel = str(getattr(self, 'selected_file_for_thumbs', '') or '')
        multi = getattr(self, 'thumb_multi_select', set())
        for fp, w in list(getattr(self, 'thumb_widgets', {}).items()):
            try:
                if str(fp) in multi:
                    w.setStyleSheet("border: 3px solid #a010ef;")
                elif str(fp) == sel and sel:
                    w.setStyleSheet("border: 3px solid #5f10a3;")
                else:
                    w.setStyleSheet("")
            except Exception:
                continue

    def _make_thumb_click_handler(self, label_widget):
        def handler(event):
            if event.button() != QtCore.Qt.LeftButton:
                return
            if self._handle_spec_marker_click(label_widget, event):
                return
            fp = label_widget.property("file_path")
            ch_idx = int(label_widget.property("channel_index"))
            mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
            if mods & QtCore.Qt.ShiftModifier:
                self._toggle_thumb_multi_selection(fp)
                return
            if mods & QtCore.Qt.ControlModifier:
                self._toggle_thumb_multi_selection(fp)
                return
            self._clear_thumb_multi_selection(update_styles=False)
            self.on_thumbnail_clicked(fp, ch_idx)
        return handler

    def _make_thumb_move_handler(self, label_widget):
        def handler(event):
            if not self._handle_spec_hover(label_widget, event):
                QtWidgets.QLabel.mouseMoveEvent(label_widget, event)
        return handler

    # ---------- thumbnail clicked -> preview + inspector populate ----------
    def on_thumbnail_clicked(self, header_path_str, channel_idx):
        """
        Thumbnail clicked -> preview.
        We no longer populate a persistent per-file inspector list (UI removed).
        Instead we:
          - show the clicked channel in the main preview,
          - update the thumb selection highlight,
          - record the current file header path and channel index so dialogs like
            "Add channel view" can reuse them via current_inspector_* attributes.
        """
        # show preview as before
        self.show_file_channel(header_path_str, channel_idx)
        # highlight selection in thumbnails
        try:
            self.selected_file_for_thumbs = str(header_path_str)
            self._refresh_thumb_selection_styles()
        except Exception:
            pass
        # record the header and channel idx for dialogs that expect them
        key = str(header_path_str)
        self.current_inspector_header = key
        self.current_inspector_channel = int(channel_idx)

    # NOTE: removed on_file_channel_selected and on_file_channel_show_clicked
    # These functions supported the removed per-file inspector UI. The same "show channel"
    # functionality is available via the thumbnail UI and the "Add channel view" dialog.

    # ---------- preview + metadata ---------- 
    def show_file_channel(self, header_path_str, channel_idx:int, use_local_cmap=False):
        self.last_preview = (str(header_path_str), int(channel_idx))
        header_path = Path(header_path_str)
        # track selected file for thumbnail highlighting
        try:
            self.selected_file_for_thumbs = str(header_path)
            self._refresh_thumb_selection_styles()
        except Exception:
            pass
        file_key = str(header_path)
        header, fds = self.headers.get(file_key, (None,None))
        if header is None or channel_idx < 0 or channel_idx >= len(fds): return
        fd = fds[channel_idx]; fname = fd.get("FileName")
        try:
            xpix = int(header.get('xPixel', 128)); ypix = int(header.get('yPixel', xpix))
            unit_final, arr_conv = self._get_filtered_channel_array(file_key, channel_idx, header, fd)
        except Exception as e:
            self.meta_box.setPlainText("Error reading channel: %s" % str(e)); return

        XScanRange = float(header.get('XScanRange', 0.0)) if header.get('XScanRange') else None
        YScanRange = float(header.get('YScanRange', 0.0)) if header.get('YScanRange') else None
        extent = [0.0, float(XScanRange), float(YScanRange), 0.0] if (XScanRange and YScanRange) else None

        cmap_to_use = self.preview_cmap_combo.currentText() or self.preview_cmap
        if use_local_cmap:
            cmap_to_use = self.per_file_channel_cmap.get((file_key, channel_idx), cmap_to_use)

        # build views (main + dynamic extras based on current file)
        views = []
        main = {'arr': arr_conv, 'extent': extent, 'cmap': cmap_to_use, 'unit': unit_final, 'title': f"{Path(header_path).name} {fd.get('Caption','')}"}
        views.append(main)

        # Rebuild extra views for the currently selected file using stored specifications
        for spec in getattr(self, 'extra_view_specs', []):
            try:
                # Find matching channel in this file (by caption first, then by index)
                idx2 = self._find_channel_index_for_spec(fds, spec)
                if idx2 is None:
                    continue
                fd2 = fds[idx2]
                unit2_final, arr2_conv = self._get_filtered_channel_array(file_key, idx2, header, fd2)
                cmap2 = self._resolve_extra_spec_cmap(spec, file_key)
                title2 = f"{Path(header_path).name} {fd2.get('Caption','')}"
                views.append({'arr': arr2_conv, 'extent': extent, 'cmap': cmap2, 'unit': unit2_final, 'title': title2})
            except Exception:
                # Skip extra view if anything fails for this file
                continue

        self.preview_canvas.set_views(views)

        # Styled HTML metadata
        try:
            html = self._build_metadata_html(header_path, header, fd, channel_idx, unit_final, arr_conv)
            self.meta_box.setHtml(html)
        except Exception:
            self.meta_box.setPlainText(f"File: {header_path.name}")

    def get_current_detail_config(self):
        """Return JSON-friendly configuration describing current detail view state."""
        cfg = {'channels': [], 'cmaps': {}, 'vmin_vmax': {}, 'figure_size': list(self.preview_canvas.fig.get_size_inches())}
        main_desc = None
        if self.last_preview:
            file_key = str(self.last_preview[0])
            header, fds = self.headers.get(file_key, (None, None))
            if header and fds:
                idx = int(self.last_preview[1])
                if 0 <= idx < len(fds):
                    cap = fds[idx].get('Caption', fds[idx].get('FileName', f"chan{idx}"))
                    key = f"idx_{idx}_{cap}"
                    main_desc = {'type': 'index', 'index': idx, 'caption': cap, 'key': key}
                    cfg['channels'].append(main_desc)
                    cmap = self.per_file_channel_cmap.get((file_key, idx), self.preview_cmap_combo.currentText() or self.preview_cmap)
                    cfg['cmaps'][key] = cmap
                    cfg['vmin_vmax'][key] = None
        # include extra views
        for spec in getattr(self, 'extra_view_specs', []):
            key = f"spec_{spec.get('caption','')}#{spec.get('index',-1)}"
            desc = {'type': 'spec', 'spec': spec.copy(), 'key': key}
            cfg['channels'].append(desc)
            cfg['cmaps'][key] = spec.get('cmap', self.preview_cmap_combo.currentText() or self.preview_cmap)
            cfg['vmin_vmax'][key] = None
        return cfg

    def _apply_filters_to_array(self, file_path, arr):
        spec = self.thumbnail_filters.get(str(file_path))
        if not spec:
            return arr
        return self._apply_filter_pipeline(arr, spec.get('steps', []))

    def _apply_filter_pipeline(self, arr, steps):
        result = np.asarray(arr, dtype=float)
        for step in steps:
            result = self._run_filter_step(result, step)
        return result

    def _run_filter_step(self, arr, step):
        key = step.get('key')
        params = step.get('params', {})
        try:
            if key == 'flatten':
                axis = params.get('axis', 'both')
                return flatten_remove_median(arr, axis=axis)
            if key == 'tilt':
                return subtract_best_fit_plane(arr)
            if key == 'plane2':
                return subtract_2nd_order_plane(arr)
            if key == 'lowpass':
                sigma = params.get('sigma', 2.0)
                return gaussian_filter_image(arr, sigma)
            if key == 'highpass':
                sigma = params.get('sigma', 2.0)
                return highpass_filter(arr, sigma)
        except Exception:
            pass
        return arr

    # ---------- dz helpers ----------
    def _dz_vs_previous_ch(self, header_path:Path):
        """Return dz pm and previous CH filename (most recent earlier file that is CH)."""
        key = str(header_path)
        info = self.tags.get(key, {})
        cur_abs = info.get('abs_z_pm', None)
        if cur_abs is None: return None, None
        try: idx = self.files.index(header_path)
        except ValueError:
            idx = None
            for i,p in enumerate(self.files):
                if str(p) == str(header_path): idx = i; break
        if idx is None: return None, None
        for j in range(idx-1, -1, -1):
            keyj = str(self.files[j]); infoj = self.tags.get(keyj, {})
            if infoj.get('tag') == 'constant-height' and infoj.get('abs_z_pm') is not None:
                return (cur_abs - infoj.get('abs_z_pm')), Path(keyj).name
        return None, None

    def _dz_vs_last_before_ch(self, header_path:Path):
        """Return dz pm vs last previous file that is not CH (e.g., last topo or CC before starting CH)."""
        key = str(header_path)
        info = self.tags.get(key, {})
        cur_abs = info.get('abs_z_pm', None)
        if cur_abs is None: return None, None
        try: idx = self.files.index(header_path)
        except ValueError:
            idx = None
            for i,p in enumerate(self.files):
                if str(p) == str(header_path): idx = i; break
        if idx is None: return None, None
        # search backwards for first previous file that is NOT CH
        for j in range(idx-1, -1, -1):
            keyj = str(self.files[j]); infoj = self.tags.get(keyj, {})
            if infoj.get('tag') != 'constant-height' and infoj.get('abs_z_pm') is not None:
                return (cur_abs - infoj.get('abs_z_pm')), Path(keyj).name
        return None, None

    # ---------- Add / Clear extra views ----------
    def on_add_view(self):
        if not hasattr(self, 'current_inspector_header') or self.current_inspector_header is None:
            QtWidgets.QMessageBox.information(self, "No file selected", "Please select a thumbnail first.")
            return
        hdr_path = Path(self.current_inspector_header); header, fds = self.headers.get(str(hdr_path), (None, None))
        if header is None: return
        dlg = QtWidgets.QDialog(self); dlg.setWindowTitle("Add channel view")
        v = QtWidgets.QVBoxLayout()
        listw = QtWidgets.QListWidget()
        for idx, fd in enumerate(fds):
            cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            it = QtWidgets.QListWidgetItem(f"{idx}: {cap}"); it.setData(QtCore.Qt.UserRole, idx); listw.addItem(it)
        v.addWidget(listw)
        hm = QtWidgets.QHBoxLayout()
        hm.addWidget(QtWidgets.QLabel("Cmap:"))
        cmapcombo = QtWidgets.QComboBox()
        # Populate cmap list with icons, falling back to a fixed list if colormaps is unavailable
        try:
            cmap_names = sorted(colormaps.keys())
        except Exception:
            cmap_names = ['viridis','plasma','inferno','magma','cividis','gray','hot','coolwarm','turbo']
        for name in cmap_names:
            try:
                icon = _colormap_icon(name, width=96, height=14)
            except Exception:
                icon = QIcon()
            cmapcombo.addItem(icon, name)
        if 'viridis' in cmap_names:
            try:
                cmapcombo.setCurrentText('viridis')
            except Exception:
                pass
        hm.addWidget(cmapcombo)
        v.addLayout(hm)
        btn_h = QtWidgets.QHBoxLayout(); add_btn = QtWidgets.QPushButton("Add"); cancel_btn = QtWidgets.QPushButton("Cancel")
        btn_h.addWidget(add_btn); btn_h.addWidget(cancel_btn); v.addLayout(btn_h)
        dlg.setLayout(v)
        add_btn.clicked.connect(dlg.accept); cancel_btn.clicked.connect(dlg.reject)
        if dlg.exec_() != QtWidgets.QDialog.Accepted: return
        sel = listw.currentItem()
        if not sel: QtWidgets.QMessageBox.information(self, "Choose channel", "Please select a channel to add."); return
        idx = sel.data(QtCore.Qt.UserRole); cmap = cmapcombo.currentText()
        # Record spec by caption and index; rebuild dynamically for selected file
        fd = fds[idx]
        cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
        key = str(hdr_path)
        spec = self._ensure_extra_spec_entry(cap, idx, cmap)
        self._set_extra_spec_override(spec, key, cmap)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _get_cached_header(self, path):
        """Return cached (header, fds) tuple if file is unchanged."""
        entry = self.header_cache.get(str(path))
        if not entry:
            return None
        try:
            mtime = Path(path).stat().st_mtime
        except Exception:
            return None
        if abs(entry.get('mtime', 0.0) - mtime) > 1e-6:
            return None
        header = entry.get('header')
        fds = entry.get('fds')
        if header is None or fds is None:
            return None
        return header, fds

    def _store_header_cache(self, path, header, fds):
        """Store parsed header info for future sessions."""
        try:
            mtime = Path(path).stat().st_mtime
        except Exception:
            return
        self.header_cache[str(path)] = {
            'mtime': mtime,
            'header': header,
            'fds': fds,
        }
        self._header_cache_dirty = True

    def _save_header_cache(self):
        if getattr(self, '_header_cache_dirty', False):
            save_header_cache(self.header_cache)
            self._header_cache_dirty = False

    def on_clear_views(self):
        self.added_views = []
        self.extra_view_specs = []
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- helpers for extra view mapping ----------
    def _find_existing_extra_spec(self, caption, idx):
        """Return the stored spec entry for a given caption/index combo if it exists."""
        cap_norm = (caption or '').strip().lower()
        try:
            idx = int(idx)
        except Exception:
            idx = -1
        for spec in getattr(self, 'extra_view_specs', []):
            spec_cap = (spec.get('caption') or '').strip().lower()
            try:
                spec_idx = int(spec.get('index', -1))
            except Exception:
                spec_idx = -1
            if cap_norm and spec_cap and cap_norm == spec_cap:
                return spec
            if (not cap_norm) and idx != -1 and idx == spec_idx:
                return spec
        return None

    def _ensure_extra_spec_entry(self, caption, idx, cmap):
        """Fetch an existing spec entry or create a new one."""
        spec = self._find_existing_extra_spec(caption, idx)
        if spec is None:
            spec = {'caption': caption, 'index': int(idx), 'cmap': str(cmap), 'cmap_overrides': {}}
            self.extra_view_specs.append(spec)
        else:
            spec.setdefault('cmap_overrides', {})
            if 'cmap' not in spec or not spec['cmap']:
                spec['cmap'] = str(cmap)
        return spec

    def _resolve_extra_spec_cmap(self, spec, file_key):
        """Choose the best cmap for a spec, honoring per-file overrides when available."""
        if not spec:
            return self.preview_cmap_combo.currentText() or self.preview_cmap
        overrides = spec.get('cmap_overrides') or {}
        if file_key in overrides:
            return overrides[file_key]
        return spec.get('cmap', self.preview_cmap_combo.currentText() or self.preview_cmap)

    def _set_extra_spec_override(self, spec, file_key, cmap):
        """Store the cmap override for a spec/file pair."""
        if spec is None:
            return
        od = spec.setdefault('cmap_overrides', {})
        od[file_key] = str(cmap)

    def _find_channel_index_for_spec(self, fds, spec):
        """Given the list of file descriptors for a file and a spec dict
        {'caption': str, 'index': int, ...}, return the best matching channel index.
        Prefers exact caption match (case-insensitive), then substring match, then stored index.
        Returns None if no suitable channel is found.
        """
        if not fds:
            return None
        target_cap = (spec.get('caption') or '').strip().lower()
        if target_cap:
            # exact caption match
            for i, fd in enumerate(fds):
                cap_i = (fd.get('Caption','') or '').strip().lower()
                if cap_i == target_cap:
                    return i
            # substring caption match
            for i, fd in enumerate(fds):
                cap_i = (fd.get('Caption','') or '').strip().lower()
                if target_cap in cap_i and cap_i:
                    return i
            # try FileName match if caption didn't work
            for i, fd in enumerate(fds):
                fn_i = (fd.get('FileName','') or '').strip().lower()
                if fn_i == target_cap or (target_cap and target_cap in fn_i):
                    return i
        # fallback to stored index
        try:
            idx = int(spec.get('index', -1))
        except Exception:
            idx = -1
        if 0 <= idx < len(fds):
            return idx
        return None

    # ---------- Export PNGs ----------
    def _sanitize_filename_component(self, s: str) -> str:
        try:
            s = str(s)
        except Exception:
            s = ""
        # Replace invalid Windows filename chars and compress spaces
        s = re.sub(r'[<>:"/\\|?*]+', '_', s)
        s = s.strip().replace(' ', '_')
        s = re.sub(r'_+', '_', s)
        return s or "unnamed"

    def on_export_pngs(self):
        # Export high-quality PNGs for the currently selected file's visible channels (main + extras)
        if not self.last_preview:
            QtWidgets.QMessageBox.information(self, "No selection", "Select a file/channel first.")
            return
        header_path_str, channel_idx = self.last_preview
        header_path = Path(header_path_str)
        file_key = str(header_path)
        header, fds = self.headers.get(file_key, (None, None))
        if header is None:
            QtWidgets.QMessageBox.information(self, "No header", "Cannot find header for current selection.")
            return

        # Choose output directory
        default_dir = str(getattr(self, 'last_dir', header_path.parent))
        out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select export folder", default_dir)
        if not out_dir:
            return

        # Common geometry/extent
        try:
            xpix = int(header.get('xPixel', 128)); ypix = int(header.get('yPixel', xpix))
        except Exception:
            xpix = 128; ypix = 128
        XScanRange = float(header.get('XScanRange', 0.0)) if header.get('XScanRange') else None
        YScanRange = float(header.get('YScanRange', 0.0)) if header.get('YScanRange') else None
        extent = [0.0, float(XScanRange), float(YScanRange), 0.0] if (XScanRange and YScanRange) else None

        # Build list of (arr, cmap, unit, caption, idx)
        exports = []

        # Main channel
        try:
            fd = fds[channel_idx]
            unit_final, arr_conv = self._get_filtered_channel_array(file_key, channel_idx, header, fd)
            cap = fd.get('Caption', fd.get('FileName', f"chan{channel_idx}"))
            cmap_main = self.per_file_channel_cmap.get((file_key, int(channel_idx)), self.preview_cmap_combo.currentText() or self.preview_cmap)
            exports.append({'arr': arr_conv, 'cmap': cmap_main, 'unit': unit_final, 'caption': cap, 'idx': int(channel_idx)})
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Export", f"Failed main channel: {e}")

        # Extra channels (based on spec mapping for current file)
        for spec in getattr(self, 'extra_view_specs', []):
            try:
                idx2 = self._find_channel_index_for_spec(fds, spec)
                if idx2 is None:
                    continue
                fd2 = fds[idx2]
                unit2_final, arr2_conv = self._get_filtered_channel_array(file_key, idx2, header, fd2)
                cap2 = fd2.get('Caption', fd2.get('FileName', f"chan{idx2}"))
                cmap2 = self._resolve_extra_spec_cmap(spec, file_key)
                exports.append({'arr': arr2_conv, 'cmap': cmap2, 'unit': unit2_final, 'caption': cap2, 'idx': int(idx2)})
            except Exception:
                continue

        if not exports:
            QtWidgets.QMessageBox.information(self, "Export", "No channels to export.")
            return

        # Metadata for naming
        date = self._sanitize_filename_component(header.get('Date', ''))
        file_base = self._sanitize_filename_component(header_path.stem)
        try:
            folder_name = self._sanitize_filename_component(str(header_path.parent.name))
        except Exception:
            folder_name = ''

        # Save each channel as a separate high-DPI PNG
        from matplotlib.figure import Figure
        for item in exports:
            try:
                fig = Figure(figsize=(6, 5), dpi=300)
                ax = fig.add_subplot(1,1,1)
                arr = np.asarray(item['arr'])
                cmapname = item.get('cmap', 'viridis')
                if extent is None:
                    im = ax.imshow(arr, origin='upper', interpolation='nearest', cmap=cmapname)
                else:
                    im = ax.imshow(arr, extent=extent, origin='upper', interpolation='nearest', aspect='equal', cmap=cmapname)
                unit = item.get('unit') or ''
                if unit:
                    cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                    cbar.set_label(unit)
                # include channel caption + acquisition date + folder so the PNG carries context
                display_date = header.get('Date', '') if header is not None else ''
                try:
                    folder_disp = header_path.parent.name
                except Exception:
                    folder_disp = ''
                chan_name = item.get('caption') or f"chan{item.get('idx',0)}"
                title_parts = [p for p in (chan_name, display_date, folder_disp) if p]
                ax.set_title("   ".join(title_parts))
                try:
                    fig.tight_layout()
                except Exception:
                    pass

                chan_name = self._sanitize_filename_component(item.get('caption') or f"chan{item.get('idx',0)}")
                parts = [p for p in (chan_name, file_base, folder_name, date) if p]
                fname = "__".join(parts) + ".png"
                out_path = str(Path(out_dir) / fname)
                fig.savefig(out_path, dpi=300, bbox_inches='tight')
            except Exception as e:
                # keep going for other channels
                print('Export failed for a channel:', e)

        QtWidgets.QMessageBox.information(self, "Export", f"Exported {len(exports)} PNG(s) to\n{out_dir}")

    def render_and_save_file_using_config(self, header_path, config, out_dir):
        """
        Render the given file using the supplied config (as returned by get_current_detail_config)
        and save a multi-panel PNG. Returns a list with the saved file path.
        """
        header_path = Path(header_path)
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        header, fds = self.headers.get(str(header_path), (None, None))
        if header is None or fds is None:
            header, fds = parse_header(header_path)
            self.headers[str(header_path)] = (header, fds)
        try:
            xpix = int(header.get('xPixel', 128))
            ypix = int(header.get('yPixel', xpix))
        except Exception:
            xpix = 128; ypix = 128
        XScanRange = float(header.get('XScanRange', 0.0)) if header.get('XScanRange') else None
        YScanRange = float(header.get('YScanRange', 0.0)) if header.get('YScanRange') else None
        extent = [0.0, float(XScanRange), float(YScanRange), 0.0] if (XScanRange and YScanRange) else None
        render_items = []
        for desc in config.get('channels', []):
            key = desc.get('key') or f"idx_{desc.get('index')}"
            idx = None
            if desc.get('type') == 'index':
                idx = int(desc.get('index', -1))
            elif desc.get('type') == 'spec':
                idx = self._find_channel_index_for_spec(fds, desc.get('spec'))
            if idx is None or idx < 0 or idx >= len(fds):
                continue
            fd = fds[idx]
            fname = fd.get('FileName')
            try:
                unit_final, arr_conv = self._get_filtered_channel_array(str(header_path), idx, header, fd)
            except Exception:
                continue
            label = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
            cmap = config.get('cmaps', {}).get(key, self.preview_cmap_combo.currentText() or self.preview_cmap)
            v_range = config.get('vmin_vmax', {}).get(key)
            vmin = vmax = None
            if isinstance(v_range, (list, tuple)) and len(v_range) == 2:
                vmin, vmax = v_range
            render_items.append({'arr': arr_conv, 'extent': extent, 'unit': unit_final, 'label': label, 'cmap': cmap, 'vmin': vmin, 'vmax': vmax})
        if not render_items:
            raise ValueError("No matching channels for export.")
        fig_size = config.get('figure_size', (6, 5))
        if not isinstance(fig_size, (list, tuple)) or len(fig_size) != 2:
            fig_size = (6, 5)
        fig_w, fig_h = fig_size
        fig = Figure(figsize=(fig_w, fig_h), dpi=300)
        total = len(render_items)
        cols = int(math.ceil(math.sqrt(total)))
        rows = int(math.ceil(total / cols))
        for i, item in enumerate(render_items, 1):
            ax = fig.add_subplot(rows, cols, i)
            im = ax.imshow(item['arr'], extent=item['extent'], origin='upper', interpolation='nearest',
                           aspect='equal' if item['extent'] else 'auto', cmap=item['cmap'],
                           vmin=item['vmin'], vmax=item['vmax'])
            ax.set_title(item['label'], fontsize=9)
            ax.tick_params(labelsize=8)
            if item['unit']:
                cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(item['unit'])
        try:
            fig.tight_layout()
        except Exception:
            pass
        base = self._sanitize_filename_component(header_path.stem)
        chlist = "_".join([self._sanitize_filename_component(it['label']) for it in render_items])
        fname = f"{base}__channels_{chlist}.png"
        out_path = out_dir / fname
        counter = 1
        while out_path.exists():
            out_path = out_dir / f"{base}__channels_{chlist}_{counter}.png"
            counter += 1
        fig.savefig(out_path, dpi=300, bbox_inches='tight')
        return [str(out_path)]

    # ---------- Profile measurement (interactive line) ----------
    def _on_start_profile(self):
        # toggle interactive line profile mode
        views = getattr(self.preview_canvas, 'views', [])
        if not views:
            QtWidgets.QMessageBox.information(self, "Measure profile", "No image to measure. Load a channel first.")
            return
        active = getattr(self.preview_canvas, 'profile_enabled', False)
        if not active:
            # enter profile mode
            self.preview_canvas.set_profile_callback(self._on_profile_updated)
            self.preview_canvas.enable_profile(True)
            try: self.measure_profile_btn.setText('Exit profile')
            except Exception: pass
            self.meta_box.setPlainText("Profile mode: drag the yellow endpoints on the main image. Close to exit.")
            self._profile_dialog = None
        else:
            # exit profile mode
            self.preview_canvas.enable_profile(False)
            try: self.measure_profile_btn.setText('Measure profile')
            except Exception: pass
            try:
                if hasattr(self, '_profile_dialog') and self._profile_dialog is not None:
                    self._profile_dialog.close()
            except Exception:
                pass

    def _on_profile_updated(self, x_px, vals, length_nm, unit):
        # create or update a persistent profile dialog
        try:
            if not hasattr(self, '_profile_dialog') or self._profile_dialog is None:
                self._profile_dialog = ProfileDialog(x_px, vals, length_nm=length_nm, parent=self, unit=unit)
                self._profile_dialog.show()
            else:
                self._profile_dialog.update_data(x_px, vals, length_nm=length_nm)
        except Exception:
            pass

    def _on_view_copied(self, view):
        title = view.get('title') or 'View'
        msg = f"Copied '{title}' to clipboard"
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), msg)
        except Exception:
            pass

    # ---------- manual tagging (still available) ----------
    def on_manual_tag(self, tag):
        if self.last_preview is None:
            QtWidgets.QMessageBox.information(self, "No file selected", "Please select a thumbnail first."); return
        header_path_str, ch_idx = self.last_preview; header_path = Path(header_path_str); key = str(header_path)
        if tag is None:
            if key in self.tags: del self.tags[key]
        else:
            info = {'tag': tag}
            if tag == 'constant-height':
                try:
                    hdr, fds = self.headers.get(key)
                    topo_idx = _find_topography_channel(fds)
                    if topo_idx is None:
                        topo_idx = ch_idx
                    fd = fds[topo_idx]
                    arr = self._get_channel_array(key, topo_idx, hdr, fd)
                    phys = (fd.get('PhysUnit','') or '').lower()
                    arr_nm = arr
                    hist, edges = np.histogram(arr_nm.ravel(), bins=200)
                    imax = int(np.argmax(hist))
                    mode_val = 0.5*(edges[imax] + edges[imax+1])
                    abs_pm = int(round(mode_val * 1000.0))
                    info['abs_z_pm'] = abs_pm
                except Exception:
                    info['abs_z_pm'] = None
            self.tags[key] = info
        self.config['tags'] = self.tags; save_config(self.config)
        # refresh thumbnails & preview (so badges/metadata update)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- Spectroscopy helpers ----------
    def on_spec_folder_browse(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select spectroscopy folder", str(self.spec_folder_path))
        if folder:
            self.spec_folder_le.setText(folder)
            self._set_spec_folder(Path(folder))

    def on_spec_folder_entered(self):
        text = self.spec_folder_le.text().strip()
        if not text:
            return
        self._set_spec_folder(Path(text))

    def _set_spec_folder(self, path:Path):
        try:
            self.spec_folder_path = Path(path)
            self.config['spectra_folder'] = str(self.spec_folder_path)
            save_config(self.config)
        except Exception:
            pass
        self._reload_spectros(refresh=True)

    def _reload_spectros(self, refresh=True):
        try:
            folder = getattr(self, 'spec_folder_path', None) or self.last_dir
            folder = Path(folder)
        except Exception:
            folder = self.last_dir
        if not self.show_spectra:
            self.spectros = []
            self.spectros_by_image = defaultdict(list)
            self._clear_multi_spec_selection()
            return
        self.spectros = self._scan_spectros(folder)
        self._assign_spectros_to_images()
        self._clear_multi_spec_selection()
        if refresh:
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _scan_spectros(self, folder:Path):
        specs = []
        if not folder or not Path(folder).exists():
            return specs
        patterns = ("*.dat","*.DAT","*.txt","*.TXT")
        cache = self._spectro_cache
        seen_keys = set()
        for pat in patterns:
            for f in sorted(folder.glob(pat)):
                p = Path(f)
                if p.is_dir():
                    continue
                key = str(p)
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                try:
                    mtime = p.stat().st_mtime
                except Exception:
                    mtime = 0.0
                cached = cache.get(key)
                if cached and abs(cached.get('mtime', 0.0) - mtime) <= 1e-6:
                    spec = cached.get('data')
                else:
                    try:
                        spec = read_dat_spectroscopy(p)
                    except Exception:
                        continue
                    cache[key] = {'mtime': mtime, 'data': spec}
                if spec is not None:
                    specs.append(spec)
        stale = [k for k in list(cache.keys()) if k not in seen_keys]
        for k in stale:
            cache.pop(k, None)
        specs.sort(key=lambda s: s.get('time') or datetime.min)
        return specs

    def _assign_spectros_to_images(self):
        self.spectros_by_image = defaultdict(list)
        images = getattr(self, 'image_meta', [])
        if not images:
            return
        for spec in self.spectros:
            match = find_last_image_for_spec(spec.get('time'), images)
            if not match:
                continue
            image_key = str(match['path'])
            spec['image_key'] = image_key
            self.spectros_by_image[image_key].append(spec)
        for k in list(self.spectros_by_image.keys()):
            self.spectros_by_image[k].sort(key=lambda s: s.get('time') or datetime.min)

    def _map_spec_to_pixels(self, spec, header, xpix, ypix):
        try:
            x = float(spec.get('x'))
            y = float(spec.get('y'))
        except Exception:
            return None
        if x is None or y is None:
            return None
        try:
            x_center = float(header.get('xCenter', 0.0))
            y_center = float(header.get('yCenter', 0.0))
            x_range = float(header.get('XScanRange', header.get('ScanRange', 0.0)) or 0.0)
            y_range = float(header.get('YScanRange', header.get('ScanRange', 0.0)) or 0.0)
        except Exception:
            return None
        if x_range <= 0 or y_range <= 0:
            return None
        xmin = x_center - x_range/2.0
        ymin = y_center - y_range/2.0
        xmax = x_center + x_range/2.0
        ymax = y_center + y_range/2.0
        if xmax == xmin or ymax == ymin:
            return None
        frac_x = (x - xmin) / (xmax - xmin)
        frac_y = (ymax - y) / (ymax - ymin)
        if not (0.0 <= frac_x <= 1.0 and 0.0 <= frac_y <= 1.0):
            return None
        cols = max(1, int(xpix) - 1)
        rows = max(1, int(ypix) - 1)
        col = frac_x * cols
        row = frac_y * rows
        return col, row

    def _draw_spectro_markers_on_pixmap(self, pixmap, header, file_key, xpix, ypix):
        if not self.show_spectra:
            return []
        specs = self.spectros_by_image.get(file_key, [])
        if not specs:
            return []
        markers = []
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
        w_scale = pixmap.width() / max(1, xpix - 1)
        h_scale = pixmap.height() / max(1, ypix - 1)
        for idx, spec in enumerate(specs, 1):
            coords = self._map_spec_to_pixels(spec, header, xpix, ypix)
            if coords is None:
                continue
            col, row = coords
            x = col * w_scale
            y = row * h_scale
            radius = 9
            center = QtCore.QPointF(x, y)
            painter.setBrush(QtGui.QColor(70, 150, 220, 220))
            pen = QtGui.QPen(QtGui.QColor(20, 20, 20))
            pen.setWidth(1)
            painter.setPen(pen)
            painter.drawEllipse(center, radius, radius)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255)))
            painter.setFont(QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold))
            painter.drawText(QtCore.QRectF(x-radius, y-radius, radius*2, radius*2), QtCore.Qt.AlignCenter, str(idx))
            markers.append({'rect': QtCore.QRectF(x-radius, y-radius, radius*2, radius*2), 'spec': spec, 'label': idx})
        painter.end()
        return markers

    def _label_pos_to_pix_coords(self, label_widget, pos):
        pix = label_widget.pixmap()
        if pix is None:
            return None
        offset_x = (label_widget.width() - pix.width()) / 2.0
        offset_y = (label_widget.height() - pix.height()) / 2.0
        x = pos.x() - offset_x
        y = pos.y() - offset_y
        if x < 0 or y < 0 or x > pix.width() or y > pix.height():
            return None
        return x, y

    def _handle_spec_marker_click(self, label_widget, event):
        if getattr(event, 'button', None) and event.button() != QtCore.Qt.LeftButton:
            return False
        if not self.show_spectra:
            return False
        markers = label_widget.property("spec_markers") or []
        if not markers:
            return False
        coords = self._label_pos_to_pix_coords(label_widget, event.pos())
        if coords is None:
            return False
        x, y = coords
        for info in markers:
            rect = info.get('rect')
            if rect and rect.contains(x, y):
                mods = event.modifiers() if event is not None else QtCore.Qt.NoModifier
                if mods & QtCore.Qt.ShiftModifier:
                    self._toggle_multi_spec_selection(info.get('spec'))
                else:
                    self._clear_multi_spec_selection()
                    self._open_spectroscopy_popup(info.get('spec'))
                return True
        return False

    def _handle_spec_hover(self, label_widget, event):
        if not self.show_spectra:
            QtWidgets.QToolTip.hideText()
            return False
        markers = label_widget.property("spec_markers") or []
        if not markers:
            QtWidgets.QToolTip.hideText()
            return False
        coords = self._label_pos_to_pix_coords(label_widget, event.pos())
        if coords is None:
            QtWidgets.QToolTip.hideText()
            return False
        x, y = coords
        for info in markers:
            rect = info.get('rect')
            if rect and rect.contains(x, y):
                tooltip = Path(info['spec']['path']).name
                QtWidgets.QToolTip.showText(label_widget.mapToGlobal(event.pos()), tooltip)
                return True
        QtWidgets.QToolTip.hideText()
        return False

    def _open_spectroscopy_popup(self, spec):
        if not spec:
            return
        try:
            dlg = SpectroscopyPopup(spec, parent=self)
            dlg.show()
            self._spectro_popups.append(dlg)
            dlg.finished.connect(lambda _: self._spectro_popups.remove(dlg) if dlg in self._spectro_popups else None)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Spectroscopy", str(e))

    def _on_thumb_context_menu(self, label_widget, pos):
        fp = str(label_widget.property("file_path"))
        targets = list(self.thumb_multi_select) if self.thumb_multi_select and fp in self.thumb_multi_select else [fp]
        menu = QtWidgets.QMenu(self)
        sub = menu.addMenu("Apply filter")
        for key, info in FILTER_DEFINITIONS.items():
            act = QtWidgets.QAction(info['label'], menu)
            if info.get('needs_gaussian') and not _gaussian_available():
                act.setEnabled(False)
                act.setToolTip("Requires scipy or OpenCV.")
            act.triggered.connect(lambda _, k=key, paths=list(targets): self._apply_filter_to_paths(paths, k))
            sub.addAction(act)
        custom_act = QtWidgets.QAction("Custom pipeline...", menu)
        custom_act.triggered.connect(lambda _, paths=list(targets), focus=fp: self._open_custom_filter_dialog(paths, focus))
        sub.addAction(custom_act)
        clear_one = QtWidgets.QAction("Clear filter", menu)
        clear_one.triggered.connect(lambda _, paths=[fp]: self._clear_filter_for_paths(paths))
        menu.addAction(clear_one)
        if len(targets) > 1:
            clear_sel = QtWidgets.QAction("Clear filter (selected)", menu)
            clear_sel.triggered.connect(lambda _, paths=list(targets): self._clear_filter_for_paths(paths))
            menu.addAction(clear_sel)
        menu.exec_(label_widget.mapToGlobal(pos))

    def _apply_filter_to_paths(self, paths, filter_key=None, pipeline=None, label=None):
        if not paths:
            return
        if len(paths) > 12:
            ret = QtWidgets.QMessageBox.question(self, "Filters", f"Apply filter to {len(paths)} images? This may use significant memory.",
                                                 QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, QtWidgets.QMessageBox.No)
            if ret != QtWidgets.QMessageBox.Yes:
                return
        if filter_key and FILTER_DEFINITIONS.get(filter_key, {}).get('needs_gaussian') and not _gaussian_available():
            QtWidgets.QMessageBox.warning(self, "Filters", "Gaussian filters require scipy or OpenCV.")
            return
        if pipeline is None:
            params = {}
            if filter_key in ('highpass', 'lowpass'):
                params['sigma'] = FILTER_DEFINITIONS.get(filter_key, {}).get('default_sigma', 2.0)
            step = {'key': filter_key, 'params': params}
            spec_steps = [step]
            spec_label = FILTER_DEFINITIONS.get(filter_key, {}).get('label', filter_key)
        else:
            spec_steps = pipeline
            spec_label = label or 'Custom'
        path_keys = {str(Path(p)) for p in paths}
        for key in path_keys:
            steps_copy = [dict(step) for step in spec_steps]
            self.thumbnail_filters[key] = {'steps': steps_copy, 'label': spec_label}
        self._invalidate_thumbnail_cache(path_keys)
        self._invalidate_filtered_cache(path_keys)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview and str(self.last_preview[0]) in path_keys:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _clear_filter_for_paths(self, paths):
        changed = False
        path_keys = {str(Path(p)) for p in paths}
        for key in path_keys:
            if self.thumbnail_filters.pop(key, None) is not None:
                changed = True
        if changed:
            self._invalidate_thumbnail_cache(path_keys)
            self._invalidate_filtered_cache(path_keys)
            self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
            if self.last_preview and str(self.last_preview[0]) in path_keys:
                self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def _open_custom_filter_dialog(self, paths, focus_path):
        base_arr = None
        try:
            focus_key = str(focus_path)
            header, fds = self.headers.get(focus_key, (None, None))
            if header and fds:
                idx = None
                if self.last_preview and str(self.last_preview[0]) == focus_key:
                    idx = int(self.last_preview[1])
                if idx is None:
                    idx = 0
                if 0 <= idx < len(fds):
                    fd = fds[idx]
                    arr = self._get_channel_array(focus_key, idx, header, fd)
                    base_arr = normalize_unit_and_data(arr, fd.get('PhysUnit',''))[1]
        except Exception:
            base_arr = None
        dlg = CustomFilterDialog(self, base_arr, self._run_filter_step)
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            pipeline = dlg.pipeline_steps()
            if pipeline:
                self._apply_filter_to_paths(paths, pipeline=pipeline, label=dlg.pipeline_label())

    def _toggle_thumb_multi_selection(self, file_path):
        path = str(file_path)
        if not hasattr(self, 'thumb_multi_select'):
            self.thumb_multi_select = set()
        if path in self.thumb_multi_select:
            self.thumb_multi_select.remove(path)
        else:
            self.thumb_multi_select.add(path)
        self._refresh_thumb_selection_styles()

    def _clear_thumb_multi_selection(self, update_styles=True):
        self.thumb_multi_select = set()
        if update_styles:
            self._refresh_thumb_selection_styles()

    def _toggle_multi_spec_selection(self, spec):
        if not spec:
            return
        key = str(Path(spec['path']))
        if key in self._multi_spec_selection_keys:
            self._multi_spec_selection = [s for s in self._multi_spec_selection if str(Path(s['path'])) != key]
            self._multi_spec_selection_keys.remove(key)
        else:
            self._multi_spec_selection.append(spec)
            self._multi_spec_selection_keys.add(key)
        self._update_spec_selection_label()
        if len(self._multi_spec_selection) >= 2:
            self._open_multi_spectroscopy_popup()

    def _update_spec_selection_label(self):
        count = len(self._multi_spec_selection)
        if hasattr(self, 'spec_selection_label'):
            self.spec_selection_label.setText(f"Spectra selected: {count}")

    def _clear_multi_spec_selection(self):
        self._multi_spec_selection = []
        self._multi_spec_selection_keys = set()
        for dlg in list(self._multi_spectro_popups):
            try:
                dlg.close()
            except Exception:
                pass
        self._multi_spectro_popups = []
        self._update_spec_selection_label()

    def on_clear_spec_selection(self):
        self._clear_multi_spec_selection()

    def _open_multi_spectroscopy_popup(self):
        specs = list(self._multi_spec_selection)
        if len(specs) < 2:
            return
        # close previous multi popups
        for dlg in list(self._multi_spectro_popups):
            try:
                dlg.close()
            except Exception:
                pass
        self._multi_spectro_popups = []
        dlg = SpectroscopyCompareDialog(specs, parent=self)
        dlg.show()
        self._multi_spectro_popups.append(dlg)
        dlg.finished.connect(lambda _: self._multi_spectro_popups.remove(dlg) if dlg in self._multi_spectro_popups else None)

    def on_spec_coord_mode_changed(self, idx):
        try:
            self.spec_coord_mode = self.spec_coord_combo.currentText()
        except Exception:
            self.spec_coord_mode = 'Auto'
        self.config['spec_coord_mode'] = self.spec_coord_mode; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_spec_invert_changed(self, checked: bool):
        self.spec_invert_y = bool(checked)
        self.config['spectro_invert_y'] = self.spec_invert_y; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_dark_mode_toggled(self, checked: bool):
        self.dark_mode = bool(checked)
        self.config['dark_mode'] = self.dark_mode; save_config(self.config)
        self._apply_dark_mode(self.dark_mode)
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    # ---------- control callbacks ----------
    def on_channel_dropdown_changed(self, idx):
        self.last_channel_index = int(idx); self.config['last_channel_index'] = self.last_channel_index; save_config(self.config)
        self.populate_thumbnails_for_channel(idx)

    def on_thumb_cmap_changed(self, idx):
        self.thumb_cmap = self.thumb_cmap_combo.currentText(); self.config['thumbnail_cmap'] = self.thumb_cmap; save_config(self.config)
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())

    def on_preview_cmap_changed(self, idx):
        self.preview_cmap = self.preview_cmap_combo.currentText(); self.config['preview_cmap'] = self.preview_cmap; save_config(self.config)
        if self.last_preview: self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_show_spectra_toggled(self, checked):
        self.show_spectra = bool(checked)
        self.config['show_spectra'] = self.show_spectra; save_config(self.config)
        if self.show_spectra:
            self._reload_spectros(refresh=False)
        else:
            self.spectros = []
            self.spectros_by_image = defaultdict(list)
            self._clear_multi_spec_selection()
        self.populate_thumbnails_for_channel(self.channel_dropdown.currentIndex())
        if self.last_preview:
            self.show_file_channel(self.last_preview[0], self.last_preview[1])

    def on_export_selected_same_view(self):
        targets = list(getattr(self, 'thumb_multi_select', set()))
        if not targets:
            if getattr(self, 'selected_file_for_thumbs', None):
                targets = [self.selected_file_for_thumbs]
            elif self.last_preview:
                targets = [self.last_preview[0]]
        if not targets:
            QtWidgets.QMessageBox.information(self, "Export", "No thumbnails selected.")
            return
        config = self.get_current_detail_config()
        if not config.get('channels'):
            QtWidgets.QMessageBox.information(self, "Export", "No channels configured to export.")
            return
        out_dir = QtWidgets.QFileDialog.getExistingDirectory(self, "Select export folder", str(self.last_dir))
        if not out_dir:
            return
        worker = BatchExportWorker(self, targets, config, out_dir)
        worker.signals.progress.connect(self._on_batch_export_progress)
        worker.signals.finished.connect(self._on_batch_export_finished)
        self._batch_export_worker = worker
        progress = QtWidgets.QProgressDialog("Exporting...", "Cancel", 0, len(targets), self)
        progress.setWindowTitle("Batch export")
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.canceled.connect(worker.cancel)
        progress.show()
        self._batch_export_progress = progress
        QtCore.QThreadPool.globalInstance().start(worker)

    def _on_batch_export_progress(self, current, total, path):
        dlg = getattr(self, '_batch_export_progress', None)
        if dlg is None:
            return
        dlg.setMaximum(total)
        dlg.setValue(current)
        dlg.setLabelText(f"Exporting {Path(path).name} ({current}/{total})")

    def _on_batch_export_finished(self, saved_paths, errors, cancelled):
        dlg = getattr(self, '_batch_export_progress', None)
        if dlg is not None:
            dlg.close()
            self._batch_export_progress = None
        self._batch_export_worker = None
        msg_lines = [f"Saved {len(saved_paths)} file(s)."]
        if saved_paths:
            preview_paths = "\n".join(saved_paths[:5])
            msg_lines.append(preview_paths + ("\n..." if len(saved_paths) > 5 else ""))
        if cancelled:
            msg_lines.append("Operation cancelled.")
        if errors:
            msg_lines.append("Errors:\n" + "\n".join(errors[:10]))
        QtWidgets.QMessageBox.information(self, "Batch export", "\n".join(msg_lines))

    def _on_purge_config(self):
        """Purge stored configuration data (tags, last_dir, cmaps) and clear runtime caches."""
        try:
            # backup current config
            try:
                if CONFIG_PATH.exists():
                    CONFIG_PATH.with_suffix('.bak').write_text(CONFIG_PATH.read_text())
            except Exception:
                pass
            # clear in-memory
            self.tags = {}
            self._invalidate_thumbnail_cache()
            self._invalidate_channel_cache()
            self.per_file_channel_cmap.clear()
            # clear config file on disk
            try:
                if CONFIG_PATH.exists():
                    CONFIG_PATH.unlink()
            except Exception:
                pass
            # reset defaults
            self.config = {}
            self.last_dir = Path.cwd()
            self.last_channel_index = 0
            QtWidgets.QMessageBox.information(self, 'Purge config', 'Configuration and tags purged. Please reopen your folder.')
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, 'Purge failed', str(e))

# ---------- run ----------
def main():
    app = QtWidgets.QApplication(sys.argv)
    try: app.setFont(QtGui.QFont("Segoe UI", 11))
    except Exception: pass
    w = SXMGridViewer(); w.show(); sys.exit(app.exec_())

if __name__ == "__main__":
    main()
