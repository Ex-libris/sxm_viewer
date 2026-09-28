"""Pure measurement-math helpers for the profile dialog's measurement band.

Kept separate from ``profile_dialog.py`` so the numbers shown in the readout
strip - the ones a user might actually cite - can be unit-tested without a
QApplication. Deliberately depends on nothing but numpy and
``profile_units`` (also numpy-only) - this dialog has to run on offline
measurement PCs with a locked-down Python install, so nothing here should
ever become a reason to add a dependency.
"""
from __future__ import annotations

from typing import Optional, Tuple

import numpy as np

from .profile_units import si_scale


def fmt_tol(value: float, tol: float, unit: str, digits: int = 3) -> str:
    """Format ``value`` with an explicit ``+/- tol``, sharing one SI prefix.

    e.g. ``fmt_tol(139.3e-9, 0.23e-9, "m", 3)`` -> ``"139 +/- 0.23 nm"``.
    ``tol`` is scaled by the same factor as ``value`` so the pair reads as
    one physically consistent measurement instead of two different prefixes
    (the old dialog's flat ``.3f`` formatting was the root cause of showing
    three decimal digits on ~10 nm sampling - precision the data couldn't
    support).
    """
    scale, prefix = si_scale(value)
    value_str = ("%%.%dg" % digits) % (value * scale)
    tol_str = "%.2g" % (abs(tol) * scale)
    return f"{value_str} ± {tol_str} {prefix}{unit}"


def find_peaks(y, frac: float = 0.10) -> np.ndarray:
    """Local maxima above ``frac`` of the full range. No scipy dependency."""
    y = np.asarray(y, dtype=float)
    if y.size < 3:
        return np.array([], dtype=int)
    mask = (y[1:-1] > y[:-2]) & (y[1:-1] >= y[2:])
    idx = np.where(mask)[0] + 1
    span = np.ptp(y)
    thr = y.min() + frac * (span if span else 1.0)
    return idx[y[idx] > thr]


def band_statistics(x, y, lo: float, hi: float) -> dict:
    """Stats for the sub-range of ``(x, y)`` inside ``[lo, hi]``.

    Returns a dict with ``n`` (point count in band), and - only when
    ``n >= 2`` - ``mean``, ``rms``, ``slope`` (per unit of ``x``), and
    ``y_left``/``y_right`` (interpolated y at the band edges). Missing keys
    when ``n < 2`` signal "not enough data," not zero.
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    lo, hi = (lo, hi) if lo <= hi else (hi, lo)
    y_left = float(np.interp(lo, x, y)) if x.size else float('nan')
    y_right = float(np.interp(hi, x, y)) if x.size else float('nan')
    mask = (x >= lo) & (x <= hi)
    n = int(mask.sum())
    result = {"n": n, "y_left": y_left, "y_right": y_right}
    if n >= 2:
        seg_x, seg_y = x[mask], y[mask]
        result["mean"] = float(seg_y.mean())
        result["rms"] = float(seg_y.std())
        try:
            result["slope"] = float(np.polyfit(seg_x, seg_y, 1)[0])
        except Exception:
            result["slope"] = float('nan')
    return result


def peaks_in_range(x, peak_indices, lo: float, hi: float) -> int:
    """Count of ``peak_indices`` (into ``x``) whose x value falls in ``[lo, hi]``."""
    x = np.asarray(x, dtype=float)
    if len(peak_indices) == 0:
        return 0
    lo, hi = (lo, hi) if lo <= hi else (hi, lo)
    xs = x[np.asarray(peak_indices, dtype=int)]
    return int(((xs >= lo) & (xs <= hi)).sum())


def snap_edges(edges: Tuple[float, float], peak_x, tol_px: float, px_size: float) -> Tuple[float, float]:
    """Snap each of ``edges`` to the nearest value in ``peak_x`` within ``tol_px`` pixels.

    An edge with no peak inside the tolerance is returned unchanged. Pure
    function - callers are responsible for guarding against feeding the
    result back into a signal handler that would re-trigger this (see
    ``ProfileDialog._snap_region``'s ``_region_snapping`` flag).
    """
    peak_x = np.asarray(peak_x, dtype=float)
    if peak_x.size == 0 or px_size <= 0:
        return tuple(edges)
    tol = tol_px * px_size
    out = []
    for v in edges:
        j = int(np.argmin(np.abs(peak_x - v)))
        out.append(float(peak_x[j]) if abs(peak_x[j] - v) < tol else float(v))
    lo, hi = out
    return (lo, hi) if lo <= hi else (hi, lo)


__all__ = ["fmt_tol", "find_peaks", "band_statistics", "peaks_in_range", "snap_edges"]
