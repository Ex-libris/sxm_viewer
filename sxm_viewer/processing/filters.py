"""Image filtering helpers used throughout the viewer.

Provides: flatten (row/col median), tilt (plane subtraction),
2nd-order plane, line-by-line flatten (per-line median/mean/poly),
Gaussian low/high-pass, Laplacian edge detection,
logarithmic dynamic-range compression, histogram equalisation,
and CLAHE (Contrast Limited Adaptive Histogram Equalization).
"""
from __future__ import annotations

import numpy as np


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

def _fit_plane_coeffs(arr, x, y, degree=1):
    """Least-squares fit a plane (degree=1) or quadratic surface (degree=2)
    to ``arr`` over pixel grid ``x``/``y``. Returns ``(coeffs, surface,
    residual_std)`` - shared by subtract_best_fit_plane/subtract_2nd_order_plane
    and by the auto-enhance diagnosis, which needs the residual to decide
    whether a tilt/plane correction is warranted at all.
    """
    if degree == 1:
        A = np.c_[x.ravel(), y.ravel(), np.ones_like(x).ravel()]
    else:
        A = np.c_[x.ravel()**2, y.ravel()**2, x.ravel()*y.ravel(), x.ravel(), y.ravel(), np.ones_like(x).ravel()]
    C, _, _, _ = np.linalg.lstsq(A, arr.ravel(), rcond=None)
    if degree == 1:
        surface = C[0]*x + C[1]*y + C[2]
    else:
        surface = C[0]*x**2 + C[1]*y**2 + C[2]*x*y + C[3]*x + C[4]*y + C[5]
    residual_std = float(np.nanstd(arr - surface))
    return C, surface, residual_std


def subtract_best_fit_plane(img):
    """Subtract best fit plane ax + by + c."""
    arr = np.asarray(img, dtype=float)
    h, w = arr.shape
    y, x = np.mgrid[:h, :w]
    _, plane, _ = _fit_plane_coeffs(arr, x, y, degree=1)
    return arr - plane

def subtract_2nd_order_plane(img):
    """Subtract quadratic plane ax^2 + by^2 + cxy + dx + ey + f."""
    arr = np.asarray(img, dtype=float)
    h, w = arr.shape
    y, x = np.mgrid[:h, :w]
    _, plane, _ = _fit_plane_coeffs(arr, x, y, degree=2)
    return arr - plane


def line_flatten_image(img, axis='row', method='median'):
    """
    Line-by-line background removal for SPM stripe artifacts.

    Subtracts a per-line statistic or polynomial fit to remove the
    DC offset, tilt, or curvature that varies from one scan line to
    the next.  This is the standard fix for 1/f-noise "sawtooth"
    artifacts in STM/AFM images.

    Parameters
    ----------
    img : ndarray
        Input 2-D image.
    axis : str
        'row' — process horizontal scan lines (default for STM).
        'col' — process vertical columns.
        'both' — process rows then columns.
    method : str
        'median' — subtract per-line median (robust, default).
        'mean'   — subtract per-line arithmetic mean.
        'poly1'  — subtract per-line 1st-order polynomial fit.
        'poly2'  — subtract per-line 2nd-order polynomial fit.

    Returns
    -------
    ndarray
        Flattened image with the same shape as input.
    """
    arr = np.asarray(img, dtype=float).copy()
    methods = ('median', 'mean', 'poly1', 'poly2')

    def _process_one_axis(data, ax):
        out = data.copy()
        n = data.shape[ax]
        x = np.arange(data.shape[1 - ax], dtype=float)
        for i in range(n):
            line = (out[i, :] if ax == 0 else out[:, i])
            valid = line[~np.isnan(line)]
            if len(valid) == 0:
                continue
            m = str(method).lower()
            if m == 'median':
                offset = np.nanmedian(line)
            elif m == 'mean':
                offset = np.nanmean(line)
            elif m in ('poly1', 'poly2'):
                mask = ~np.isnan(line)
                if mask.sum() < (2 if m == 'poly1' else 3):
                    offset = np.nanmedian(line)
                else:
                    deg = 1 if m == 'poly1' else 2
                    coeffs = np.polyfit(x[mask], line[mask], deg)
                    offset = np.polyval(coeffs, x)
                    offset[~mask] = 0
            else:
                raise ValueError(f"Unknown method: {m!r}, expected one of {methods}")
            if ax == 0:
                out[i, :] = line - offset
            else:
                out[:, i] = line - offset
        return out

    axis = str(axis).lower()
    if axis in ('row', 'both'):
        arr = _process_one_axis(arr, 0)
    if axis in ('col', 'both'):
        arr = _process_one_axis(arr, 1)
    return arr


def _robust_std(values):
    """MAD-based robust standard deviation estimate (resistant to the
    outliers/glitches this module's detectors are themselves looking for -
    a plain np.std would be skewed by the very artifacts being measured)."""
    values = np.asarray(values, dtype=float)
    finite = values[np.isfinite(values)]
    if finite.size < 2:
        return 0.0
    mad = np.nanmedian(np.abs(finite - np.nanmedian(finite)))
    return float(1.4826 * mad)


def _row_glitch_scores(row_stat):
    """Per-row "how far from a smooth trend" score: each interior row's stat
    compared to the straight-line average of its immediate neighbors. This
    cancels out any linear gradient (tilt) exactly, so it stays near zero for
    smooth topography and only spikes at an isolated glitched row - unlike
    comparing against a small local neighborhood's own median/MAD, which is
    itself too noisy an estimate (too few points) to reliably tell a real
    glitch apart from ordinary row-to-row noise.
    """
    h = row_stat.shape[0]
    deviation = np.zeros(h, dtype=float)
    if h < 3:
        return deviation
    deviation[1:-1] = row_stat[1:-1] - 0.5 * (row_stat[:-2] + row_stat[2:])
    return deviation


def _percentile_ratio_outlier_mask(values, ratio=25.0, ref_pct=95.0, min_samples=10):
    """Flag entries whose magnitude exceeds ``ratio`` times the ``ref_pct``
    percentile of |values| - i.e. it must clearly exceed even the roughest
    slice of otherwise-ordinary data, not just be a few standard deviations
    out. Real SPM topography's fine-scale roughness is heavy-tailed (its
    residuals are nowhere near Gaussian), so a MAD/z-score threshold that
    works on smooth synthetic test data ends up flagging a large fraction of
    perfectly ordinary texture on real scans - this scales with whatever the
    data's own "normal roughness" ceiling actually is instead of assuming a
    fixed statistical shape.

    ``ref_pct`` deliberately stops short of the very top of the distribution
    (95th, not 99th): a single genuine glitch also nudges its immediate
    neighbors' scores (see _row_glitch_scores), so a reference percentile
    too close to the max can end up measuring the glitch's own ripple
    instead of the data's ordinary spread, masking the very thing being
    looked for. ``ratio=25`` was picked empirically to sit an order of
    magnitude above real scan data's own worst-case ratio at this
    reference point while remaining far below a genuine synthetic
    glitch/spike's ratio - see the auto-enhance test suite.
    """
    values = np.asarray(values, dtype=float)
    finite = values[np.isfinite(values)]
    if finite.size < min_samples:
        return np.zeros(values.shape, dtype=bool)
    ref = np.percentile(np.abs(finite), float(ref_pct))
    if ref <= 0:
        return np.zeros(values.shape, dtype=bool)
    return np.abs(values) > (float(ratio) * ref)


def repair_bad_lines(img, ratio=25.0):
    """
    Detect and repair single glitched scan lines (e.g. a tip glitch or scan
    error affecting one row) by interpolating from the nearest unaffected
    rows above/below.

    A row is flagged when it deviates from the straight-line trend of its
    immediate neighbors (see _row_glitch_scores) by more than ``ratio`` times
    the 99th-percentile deviation across the whole image (see
    _percentile_ratio_outlier_mask) - distinct from line_flatten_image, which
    removes a per-line offset/tilt from every row rather than detecting and
    replacing specific bad ones.

    Parameters
    ----------
    img : ndarray
        Input 2-D image.
    ratio : float
        How many times the 99th-percentile row deviation a row's own
        deviation must exceed to be treated as glitched. Default 4.0.

    Returns
    -------
    ndarray
        Image with flagged rows replaced by interpolation from the nearest
        good rows; unchanged if no rows are flagged.
    """
    arr = np.asarray(img, dtype=float)
    out = arr.copy()
    h = arr.shape[0]
    if h < 6:
        return out
    row_stat = np.nanmedian(arr, axis=1)
    deviation = _row_glitch_scores(row_stat)
    bad = _percentile_ratio_outlier_mask(deviation[1:-1], ratio=ratio)
    bad = np.concatenate(([False], bad, [False]))  # edge rows have no well-defined score
    if not bad.any():
        return out
    good_idx = np.where(~bad)[0]
    if good_idx.size == 0:
        return out
    for i in np.where(bad)[0]:
        above = good_idx[good_idx < i]
        below = good_idx[good_idx > i]
        if above.size and below.size:
            a, b = above[-1], below[0]
            w = (i - a) / float(b - a)
            out[i, :] = (1 - w) * out[a, :] + w * out[b, :]
        elif above.size:
            out[i, :] = out[above[-1], :]
        elif below.size:
            out[i, :] = out[below[0], :]
    return out


try:
    from scipy.ndimage import median_filter as _scipy_median_filter
except Exception:
    _scipy_median_filter = None


def remove_spikes(img, ratio=25.0, window=3):
    """
    Replace isolated outlier pixels (e.g. tip-contamination spikes or dead
    pixels) with the local median, leaving the rest of the image untouched -
    unlike a Gaussian blur, only pixels that are extreme outliers relative to
    their immediate neighborhood are modified.

    Requires scipy. Falls back to returning the input unchanged when scipy is
    not available.

    Parameters
    ----------
    img : ndarray
        Input 2-D image.
    ratio : float
        How many times the 99th-percentile local-median residual a pixel's
        own residual must exceed to be treated as a spike (see
        _percentile_ratio_outlier_mask) - not a fixed z-score, since real
        SPM topography's fine-scale roughness is heavy-tailed enough that a
        Gaussian-style threshold flags a large fraction of ordinary texture
        on real scans, even though it looks fine on smooth synthetic data.
        Default 4.0.
    window : int
        Size of the local median window (odd, >= 3). Default 3.

    Returns
    -------
    ndarray
        Image with flagged pixels replaced by the local median; unchanged if
        no pixels are flagged or scipy is unavailable.
    """
    arr = np.asarray(img, dtype=float)
    if _scipy_median_filter is None:
        return arr.copy()
    w = int(window)
    if w < 3:
        w = 3
    if w % 2 == 0:
        w += 1
    med = _scipy_median_filter(arr, size=w)
    residual = arr - med
    out = arr.copy()
    spike_mask = _percentile_ratio_outlier_mask(residual, ratio=ratio, min_samples=50)
    out[spike_mask] = med[spike_mask]
    return out


try:
    from scipy.ndimage import gaussian_filter as _scipy_gaussian
    _GAUSS_BACKEND = 'scipy'
except Exception:
    try:
        import cv2 as _cv2
        _GAUSS_BACKEND = 'cv2'
    except Exception:
        _GAUSS_BACKEND = None

try:
    from skimage.exposure import equalize_adapthist as _sk_equalize_adapthist
    _SKIMAGE_AVAILABLE = True
except Exception:
    _SKIMAGE_AVAILABLE = False

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

def laplacian_filter_image(img, sigma=0.0, neighbors=8, absolute=True):
    """
    Discrete Laplacian edge response with optional Gaussian pre-smoothing.

    Parameters
    ----------
    sigma : float
        Gaussian sigma applied before Laplacian. If smoothing backends are not
        available, the raw image is used.
    neighbors : int
        4 or 8-neighbor Laplacian stencil.
    absolute : bool
        Return absolute edge magnitude when True, signed response when False.
    """
    arr = np.asarray(img, dtype=float)
    work = arr
    try:
        sigma = float(sigma)
    except Exception:
        sigma = 0.0
    if sigma > 0.0:
        try:
            work = gaussian_filter_image(arr, sigma)
        except Exception:
            work = arr

    pad = np.pad(work, 1, mode='edge')
    c = pad[1:-1, 1:-1]
    up = pad[:-2, 1:-1]
    down = pad[2:, 1:-1]
    left = pad[1:-1, :-2]
    right = pad[1:-1, 2:]
    try:
        neigh = int(neighbors)
    except Exception:
        neigh = 8
    if neigh == 4:
        out = up + down + left + right - (4.0 * c)
    else:
        ul = pad[:-2, :-2]
        ur = pad[:-2, 2:]
        dl = pad[2:, :-2]
        dr = pad[2:, 2:]
        out = up + down + left + right + ul + ur + dl + dr - (8.0 * c)
    if bool(absolute):
        out = np.abs(out)
    return out


def log_filter_image(img, epsilon=1e-3):
    """
    Apply log10 transform to compress dynamic range.

    Shifts data to positive (subtracts the minimum), then applies
    log10 with an epsilon floor relative to the data range to avoid
    log(0).  The ``epsilon`` parameter is interpreted as a **fraction**
    of ``max - min``, with an absolute floor of 1e-12.

    Parameters
    ----------
    img : ndarray
        Input image.
    epsilon : float
        Fraction of the data range added as offset before log10.
        Default 1e-3 (0.1% of range) gives meaningful contrast for
        SPM signals that span several orders of magnitude.
        Use smaller values (e.g. 1e-5) for stronger compression.

    Returns
    -------
    ndarray
        Log10-transformed image.
    """
    arr = np.asarray(img, dtype=float)
    arr_min = np.nanmin(arr)
    arr = arr - arr_min
    data_range = np.nanmax(arr) - np.nanmin(arr)
    effective_eps = max(1e-12, float(epsilon) * float(data_range))
    return np.log10(arr + effective_eps)


def clahe_filter_image(img, clip_limit=0.03, tile_size=8):
    """
    Apply Contrast Limited Adaptive Histogram Equalization (CLAHE).

    Enhances local contrast by equalizing the histogram in small tiles
    and blending the results using bilinear interpolation.  The
    ``clip_limit`` controls how much contrast amplification is allowed
    (lower values reduce noise amplification).  ``tile_size`` sets the
    size of each local tile in pixels.

    Requires scikit-image.  Falls back to returning the input image
    unchanged when the library is not installed.

    Parameters
    ----------
    img : ndarray
        Input floating-point image.
    clip_limit : float
        Clipping limit for contrast limiting (0.001 - 0.5).
        Default 0.03.
    tile_size : int
        Size of the local tile grid.  Common values: 4, 8, 16, 32.
        Default 8.

    Returns
    -------
    ndarray
        CLAHE-equalized image with same shape and value range as input.
    """
    arr = np.asarray(img, dtype=float)
    if not _SKIMAGE_AVAILABLE:
        raise RuntimeError(
            "CLAHE filter requires scikit-image (pip install scikit-image)"
        )
    # Normalise to [0,1] for skimage
    valid = arr[~np.isnan(arr)]
    if len(valid) < 2:
        return arr
    lo, hi = float(np.nanmin(arr)), float(np.nanmax(arr))
    rng = hi - lo
    if rng <= 0:
        return arr
    norm = (arr - lo) / rng
    # Apply CLAHE
    clip_limit = max(0.001, min(0.5, float(clip_limit)))
    tile_size = int(tile_size)
    if tile_size < 2:
        tile_size = 8
    equalized = _sk_equalize_adapthist(
        norm, kernel_size=int(tile_size), clip_limit=clip_limit, nbins=256
    )
    return equalized * rng + lo


def histogram_equalize_image(img):
    """
    Apply histogram equalization for maximum contrast.

    Maps pixel intensities so the output histogram is approximately
    flat (uniform), spreading all intensity levels evenly across
    the available range.

    Parameters
    ----------
    img : ndarray
        Input image.

    Returns
    -------
    ndarray
        Equalized image in [0, 1] range.
    """
    arr = np.asarray(img, dtype=float)
    valid = arr[~np.isnan(arr)]
    if len(valid) < 2:
        return arr
    sorted_valid = np.sort(valid)
    out = np.interp(arr, sorted_valid, np.linspace(0, 1, len(sorted_valid)))
    return out


FILTER_DEFINITIONS = {
    'flatten': {'label': 'Flatten (row/col median)', 'needs_gaussian': False},
    'tilt': {'label': 'Tilt correction (plane)', 'needs_gaussian': False},
    'plane2': {'label': 'Global plane fit (2nd order)', 'needs_gaussian': False},
    'line_flatten': {
        'label': 'Line-by-line flatten',
        'needs_gaussian': False,
        'default_axis': 'row',
        'default_method': 'median',
    },
    'highpass': {'label': 'High-pass (Gaussian)', 'needs_gaussian': True, 'default_sigma': 2.0},
    'lowpass': {'label': 'Low-pass (Gaussian)', 'needs_gaussian': True, 'default_sigma': 2.0},
    'laplacian': {
        'label': 'Laplacian (edge response)',
        'needs_gaussian': False,
        'default_sigma': 0.6,
        'default_neighbors': 8,
        'default_absolute': True,
    },
    'log': {
        'label': 'Logarithmic (dynamic range)',
        'needs_gaussian': False,
        'default_epsilon': 1e-3,
    },
    'histeq': {
        'label': 'Histogram equalization',
        'needs_gaussian': False,
    },
    'clahe': {
        'label': 'CLAHE (adaptive equalization)',
        'needs_gaussian': False,
        'default_clip_limit': 0.03,
        'default_tile_size': 8,
    },
    'line_repair': {
        'label': 'Repair glitched scan lines',
        'needs_gaussian': False,
        'default_ratio': 25.0,
    },
    'spike_removal': {
        'label': 'Remove isolated spikes',
        'needs_gaussian': False,
        'default_ratio': 25.0,
        'default_window': 3,
    },
}

def _gaussian_available():
    """Return True when a Gaussian filtering backend (scipy or OpenCV) is available."""
    return _GAUSS_BACKEND is not None

def _filter_signature(spec):
    """Return a hashable signature for a filter pipeline spec."""
    if not spec:
        return tuple()
    sig = []
    for step in spec.get('steps', []):
        params = tuple(sorted((step.get('params') or {}).items()))
        sig.append((step.get('key'), params))
    return tuple(sig)


__all__ = [
    "flatten_remove_median",
    "subtract_best_fit_plane",
    "subtract_2nd_order_plane",
    "line_flatten_image",
    "gaussian_filter_image",
    "highpass_filter",
    "laplacian_filter_image",
    "log_filter_image",
    "histogram_equalize_image",
    "clahe_filter_image",
    "repair_bad_lines",
    "remove_spikes",
    "FILTER_DEFINITIONS",
    "_gaussian_available",
    "_filter_signature",
    "_fit_plane_coeffs",
    "_robust_std",
    "_row_glitch_scores",
    "_percentile_ratio_outlier_mask",
]



