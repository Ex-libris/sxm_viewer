"""Diagnosis logic for the "Let the robot" auto-enhance feature.

Given a raw 2-D scan array, decides which filter-pipeline steps (in the same
``{"key": ..., "params": {...}}`` shape ``FilterController`` already consumes)
would improve it, plus a plain-language summary of what was found - without
applying anything itself or touching Qt/GUI state, so it can be unit-tested
in isolation the same way ``FilterController``'s pure methods are.

Incomplete-scan cropping is deliberately NOT handled here: detecting the
valid scan region (``detect_valid_scan_region``) lives in the GUI layer
(``gui/thumbnail_render.py``), and applying a crop needs the canvas's own
extent-recomputation helpers - both are orchestration concerns, so the
caller (the "Let the robot" click handler) runs that step itself, before
calling ``diagnose()`` on the (possibly already-cropped) array.
"""
from __future__ import annotations

import numpy as np

from .filters import (
    _fit_plane_coeffs,
    _row_glitch_scores,
    _percentile_ratio_outlier_mask,
    subtract_best_fit_plane,
    subtract_2nd_order_plane,
    repair_bad_lines,
    remove_spikes,
)

try:
    from scipy.ndimage import median_filter as _median_filter
except Exception:
    _median_filter = None

try:
    from scipy.ndimage import gaussian_filter as _gaussian_filter
except Exception:
    _gaussian_filter = None


def diagnose_tilt(arr):
    """Decide between no correction / tilt / plane2 by comparing residual
    reduction from 1st- vs 2nd-order plane fits against the raw data's own
    spread. Returns (step_or_None, summary)."""
    arr = np.asarray(arr, dtype=float)
    h, w = arr.shape
    if h < 3 or w < 3:
        return None, ""
    y, x = np.mgrid[:h, :w]
    raw_std = float(np.nanstd(arr))
    if raw_std <= 0:
        return None, ""
    _, _, tilt_std = _fit_plane_coeffs(arr, x, y, degree=1)
    _, _, plane_std = _fit_plane_coeffs(arr, x, y, degree=2)
    # A 1st-order fit barely reducing the spread means the image is already flat.
    if tilt_std >= raw_std * 0.98:
        return None, ""
    # 2nd order meaningfully better than 1st -> real curvature, not just tilt.
    if plane_std < tilt_std * 0.85:
        return {"key": "plane2", "params": {}}, "corrected curvature (2nd-order plane fit)"
    return {"key": "tilt", "params": {}}, "corrected tilt"


def diagnose_bad_lines(arr, ratio=25.0):
    """Decide whether any row looks like a single glitched scan line, using
    the same detection logic as repair_bad_lines (kept in sync deliberately -
    this only counts rows, it doesn't repair them itself)."""
    arr = np.asarray(arr, dtype=float)
    h = arr.shape[0]
    if h < 6:
        return None, ""
    row_stat = np.nanmedian(arr, axis=1)
    deviation = _row_glitch_scores(row_stat)
    bad = _percentile_ratio_outlier_mask(deviation[1:-1], ratio=ratio)
    bad = np.concatenate(([False], bad, [False]))
    bad_count = int(np.sum(bad))
    if bad_count == 0:
        return None, ""
    step = {"key": "line_repair", "params": {"ratio": ratio}}
    plural = "s" if bad_count != 1 else ""
    return step, f"repaired {bad_count} glitched scan line{plural}"


def diagnose_spikes(arr, ratio=25.0, window=3):
    """Decide whether isolated outlier pixels (tip-contamination spikes,
    dead pixels) are present."""
    if _median_filter is None:
        return None, ""
    arr = np.asarray(arr, dtype=float)
    if arr.shape[0] < window or arr.shape[1] < window:
        return None, ""
    med = _median_filter(arr, size=window)
    residual = arr - med
    spike_mask = _percentile_ratio_outlier_mask(residual, ratio=ratio, min_samples=50)
    spike_count = int(np.sum(spike_mask))
    if spike_count == 0:
        return None, ""
    step = {"key": "spike_removal", "params": {"ratio": ratio, "window": window}}
    plural = "s" if spike_count != 1 else ""
    return step, f"removed {spike_count} isolated spike{plural}"


def diagnose_noise(arr, hf_fraction_threshold=0.18, sigma=1.0):
    """Decide whether high-frequency noise dominates enough of the total
    variation to warrant a light Gaussian denoise, via the fraction of the
    image's total spread that a light blur removes."""
    if _gaussian_filter is None:
        return None, ""
    arr = np.asarray(arr, dtype=float)
    h, w = arr.shape
    if h < 5 or w < 5:
        return None, ""
    total_std = float(np.nanstd(arr))
    if total_std <= 0:
        return None, ""
    smoothed = _gaussian_filter(arr, sigma=sigma)
    hf_std = float(np.nanstd(arr - smoothed))
    fraction = hf_std / total_std
    if fraction < float(hf_fraction_threshold):
        return None, ""
    step = {"key": "lowpass", "params": {"sigma": sigma}}
    return step, "applied a light denoise"


def diagnose(arr):
    """Diagnose a (possibly already-cropped) scan array.

    Returns (steps, summaries):
        steps - ordered list of {"key": ..., "params": {...}} dicts, ready
            for FilterController._set_filter_pipeline_on_canvas.
        summaries - matching plain-language descriptions, one per step, for
            the toast/Activity Log.

    Order matters, and each check runs on the *cumulative* result of the
    ones before it (tilt/plane correction first, then line repair, then
    spike removal, then noise reduction last) - not independently on the
    original array. Otherwise a whole-row glitch would get flagged by both
    diagnose_bad_lines and diagnose_spikes (every pixel in an offset row
    also looks like a spike in isolation), double-counting the same
    underlying problem and producing a misleading summary, when repairing
    the line first would have already resolved it before spike-detection
    ever saw it - exactly the same sequential semantics
    FilterController._apply_filter_pipeline already uses when it actually
    applies these steps for real.
    """
    working = np.asarray(arr, dtype=float).copy()
    steps = []
    summaries = []

    tilt_step, tilt_summary = diagnose_tilt(working)
    if tilt_step is not None:
        steps.append(tilt_step)
        summaries.append(tilt_summary)
        working = subtract_2nd_order_plane(working) if tilt_step["key"] == "plane2" else subtract_best_fit_plane(working)

    lines_step, lines_summary = diagnose_bad_lines(working)
    if lines_step is not None:
        steps.append(lines_step)
        summaries.append(lines_summary)
        working = repair_bad_lines(working, **lines_step["params"])

    spikes_step, spikes_summary = diagnose_spikes(working)
    if spikes_step is not None:
        steps.append(spikes_step)
        summaries.append(spikes_summary)
        working = remove_spikes(working, **spikes_step["params"])

    noise_step, noise_summary = diagnose_noise(working)
    if noise_step is not None:
        steps.append(noise_step)
        summaries.append(noise_summary)

    return steps, summaries


__all__ = [
    "diagnose",
    "diagnose_tilt",
    "diagnose_bad_lines",
    "diagnose_spikes",
    "diagnose_noise",
]
