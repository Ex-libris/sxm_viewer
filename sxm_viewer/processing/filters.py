"""Image filtering helpers used throughout the viewer."""
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
    "gaussian_filter_image",
    "highpass_filter",
    "FILTER_DEFINITIONS",
    "_gaussian_available",
    "_filter_signature",
]
