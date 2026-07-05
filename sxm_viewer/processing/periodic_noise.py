"""FFT-based periodic-noise detection and removal.

A raster scan turns time into space: a temporal noise source (turbo pump
vibration, mains hum, floor vibration) shows up in the image as a spatial
frequency determined by the scan's own timing - fast noise (much quicker
than a single line) aliases into ripple *within* a line (fast/x axis), slow
noise (comparable to the line rate) aliases into row-to-row banding
(slow/y axis). In the FFT this can show up as a narrow isolated peak
(coherent, phase-stable noise) or as a whole band/streak spanning most of
one axis (noise whose phase drifts line-to-line, or classic scan-line
banding) - removal has to handle both shapes, not just small dots, or a
strong streak survives essentially untouched no matter which point you
notch out.

The hard part isn't finding sharp FFT features - it's telling noise apart
from a genuine periodic surface feature (atomic lattice, moire pattern),
which also produces sharp peaks. This module only detects and classifies
candidates with a plain-language tip; removal only ever touches regions a
human explicitly selected or drew (see gui/dialogs/periodic_noise.py) -
nothing here is meant to run unattended.
"""
from __future__ import annotations

import numpy as np

from .filters import _fit_plane_coeffs

try:
    from scipy.ndimage import gaussian_filter as _gaussian_filter
    from scipy.ndimage import label as _label
except Exception:  # pragma: no cover - optional dependency
    _gaussian_filter = None
    _label = None


def _detrend(arr):
    arr = np.asarray(arr, dtype=float)
    h, w = arr.shape
    y, x = np.mgrid[:h, :w]
    _, plane, _ = _fit_plane_coeffs(arr, x, y, degree=1)
    return arr - plane, plane


def fft_magnitude(arr):
    """Return (freqs_x, freqs_y, magnitude) for the detrended array's FFT.

    freqs_x/freqs_y are in cycles/pixel, ranging over [-0.5, 0.5) (Nyquist),
    matching np.fft.fftfreq's convention. magnitude is fftshift'd so index
    (0, 0) of the returned array corresponds to the lowest frequencies and
    the center corresponds to zero frequency.

    The detrended array is normalized to unit std before the FFT - real SPM
    data is in meters (topography height differences of order 1e-8 to
    1e-11), and detect_candidate_regions' background-ratio test relies on
    log1p(magnitude) providing genuine logarithmic compression, which only
    happens when magnitude is order-1 or larger. At the real, tiny physical
    scale log1p(x) ~= x (no compression at all), silently making every
    threshold calibrated against synthetic O(1)-amplitude test data useless.
    Normalizing first makes detection scale-invariant.

    Windowed with a 2D Hann taper before the FFT - a real scan's edges
    almost never tile periodically, and an un-windowed FFT of a
    non-periodic signal leaks that edge discontinuity into the spectrum.
    Windowing only affects detection here - apply_mask_filter deliberately
    does its own unwindowed FFT/inverse-FFT, since windowing would need to
    be undone for an accurate reconstruction.
    """
    detrended, _ = _detrend(arr)
    std = float(np.nanstd(detrended))
    if std > 0:
        detrended = detrended / std
    h, w = detrended.shape
    window = np.outer(np.hanning(h), np.hanning(w))
    windowed = detrended * window
    spectrum = np.fft.fftshift(np.fft.fft2(windowed))
    magnitude = np.abs(spectrum)
    freqs_x = np.fft.fftshift(np.fft.fftfreq(w))
    freqs_y = np.fft.fftshift(np.fft.fftfreq(h))
    return freqs_x, freqs_y, magnitude


def _fold_to_nyquist(cycles_per_sample):
    """Fold an arbitrarily large cycles-per-sample count to the aliased
    frequency actually observable at 1 sample/pixel (or /row), in [0, 0.5]."""
    frac = cycles_per_sample - np.floor(cycles_per_sample)
    if frac > 0.5:
        frac = 1.0 - frac
    return frac


def expected_peak_frequencies(line_time_s, pixel_dwell_s, candidate_hz=(50, 60, 500, 1000, 1500), max_harmonic=3):
    """For each candidate temporal frequency and its low harmonics, compute
    where it would alias to as fast-axis ripple (fx, fy=0) and as slow-axis
    banding (fx=0, fy) - a real source could show up via either mechanism
    depending on its relative coupling to the fast vs. slow scan motion.

    Returns a list of {"hz": float, "axis": "fast"|"slow", "fx": float, "fy": float}.
    """
    expected = []
    for base_hz in candidate_hz:
        for n in range(1, max_harmonic + 1):
            hz = base_hz * n
            fx = _fold_to_nyquist(hz * pixel_dwell_s)
            expected.append({"hz": hz, "axis": "fast", "fx": fx, "fy": 0.0})
            fy = _fold_to_nyquist(hz * line_time_s)
            expected.append({"hz": hz, "axis": "slow", "fx": 0.0, "fy": fy})
    return expected


_DC_EXCLUDE_RADIUS = 3  # bins around zero frequency - broad low-freq content, not a narrow feature
_MATCH_TOLERANCE = 0.015  # cycles/pixel - how close a region must be to an expected position
_LATTICE_RADIUS_TOLERANCE = 0.03  # cycles/pixel - radius bucket width for symmetry grouping
_STREAK_SPAN_FRACTION = 0.3  # a blob spanning this much of an axis while narrow in the other is a "streak"
_STREAK_NARROW_FRACTION = 0.06


def detect_candidate_regions(magnitude, freqs_x, freqs_y, expected=None, min_ratio=4.0, max_regions=12):
    """Find contiguous regions of the spectrum that stand out against the
    local background, classify each against `expected` positions and
    against other detected regions (for lattice-like symmetry), and return a
    plain-language tip per region.

    A region can be point-like (a small blob, typical of coherent single-
    frequency noise) or streak-like (a blob spanning much of one axis while
    staying narrow in the other - typical of scan-line banding or noise
    whose phase isn't stable line-to-line). Streaks are widened to the full
    visible extent of that axis: thresholding can fragment a visually
    continuous streak into several disconnected blobs wherever it dips
    below the noise floor along its length, and a partial-width notch would
    leave most of a real streak's energy untouched.

    Returns a list of dicts: {"fx_min", "fx_max", "fy_min", "fy_max",
    "fx", "fy" (representative point), "strength", "tip_category", "tip",
    "hz", "is_streak"}. tip_category is one of "noise", "caution",
    "unclear" - callers/UI decide how to color/sort these, not this module.
    """
    if _gaussian_filter is None or _label is None:
        return []
    magnitude = np.asarray(magnitude, dtype=float)
    h, w = magnitude.shape
    cy, cx = h // 2, w // 2
    log_mag = np.log1p(magnitude)

    background = _gaussian_filter(log_mag, sigma=max(h, w) / 12.0)
    ratio = log_mag - background

    # Exclude the DC/low-frequency blob - not the narrow features this is
    # meant to find, and the plane/tilt fit already removed the gradient
    # component that would otherwise dominate here.
    yy, xx = np.mgrid[:h, :w]
    dc_mask = (yy - cy) ** 2 + (xx - cx) ** 2 <= _DC_EXCLUDE_RADIUS ** 2
    ratio_masked = np.where(dc_mask, -np.inf, ratio)

    above = ratio_masked >= min_ratio
    blobs = []
    if above.any():
        labeled, num_labels = _label(above, structure=np.ones((3, 3)))
        for label_id in range(1, num_labels + 1):
            blob = labeled == label_id
            rows, cols = np.where(blob)
            row_min, row_max = int(rows.min()), int(rows.max())
            col_min, col_max = int(cols.min()), int(cols.max())
            blob_ratio = np.where(blob, ratio_masked, -np.inf)
            prow, pcol = np.unravel_index(np.argmax(blob_ratio), blob_ratio.shape)
            row_span = row_max - row_min + 1
            col_span = col_max - col_min + 1
            is_streak = (
                (row_span >= _STREAK_SPAN_FRACTION * h and col_span <= _STREAK_NARROW_FRACTION * w)
                or (col_span >= _STREAK_SPAN_FRACTION * w and row_span <= _STREAK_NARROW_FRACTION * h)
            )
            if is_streak:
                if row_span >= col_span:
                    row_min, row_max = 0, h - 1
                else:
                    col_min, col_max = 0, w - 1
            blobs.append({
                "row_min": row_min, "row_max": row_max, "col_min": col_min, "col_max": col_max,
                "prow": int(prow), "pcol": int(pcol),
                "strength": float(ratio_masked[prow, pcol]),
                "is_streak": is_streak,
            })

    # A streak whose brightness varies along its length (phase-drifting
    # noise, or a coherent tone competing with genuinely strong topographic
    # content near the streak) can dip below min_ratio often enough that
    # per-pixel thresholding fragments it into small, scattered blobs that
    # never trigger the span-based is_streak check above. Directly testing
    # whether a whole column/row's fraction of moderately-elevated bins is
    # unusually high (a distributed signal, not a single sharp point)
    # catches this even when no individual bin crosses the full threshold.
    col_fraction = np.mean(ratio_masked >= min_ratio * 0.5, axis=0)
    row_fraction = np.mean(ratio_masked >= min_ratio * 0.5, axis=1)
    col_strength = np.max(ratio_masked, axis=0)
    row_strength = np.max(ratio_masked, axis=1)
    used_cols = set()
    used_rows = set()
    for col in np.argsort(col_fraction)[::-1]:
        col = int(col)
        if col_fraction[col] < 0.25 or col in used_cols or not np.isfinite(col_strength[col]):
            continue
        # Widen from this seed column to its full width: adjacent columns
        # whose fraction is still a substantial share of the seed's are
        # part of the same streak, not separate features - a single-bin
        # notch would leave most of a visually-wider bright band untouched.
        col_min = col_max = col
        threshold = col_fraction[col] * 0.5
        while col_min > 0 and col_fraction[col_min - 1] >= threshold:
            col_min -= 1
        while col_max < w - 1 and col_fraction[col_max + 1] >= threshold:
            col_max += 1
        used_cols.update(range(col_min, col_max + 1))
        prow = int(np.argmax(ratio_masked[:, col]))
        blobs.append({
            "row_min": 0, "row_max": h - 1, "col_min": col_min, "col_max": col_max,
            "prow": prow, "pcol": col,
            "strength": float(col_strength[col]),
            "is_streak": True,
        })
    for row in np.argsort(row_fraction)[::-1]:
        row = int(row)
        if row_fraction[row] < 0.25 or row in used_rows or not np.isfinite(row_strength[row]):
            continue
        row_min = row_max = row
        threshold = row_fraction[row] * 0.5
        while row_min > 0 and row_fraction[row_min - 1] >= threshold:
            row_min -= 1
        while row_max < h - 1 and row_fraction[row_max + 1] >= threshold:
            row_max += 1
        used_rows.update(range(row_min, row_max + 1))
        pcol = int(np.argmax(ratio_masked[row, :]))
        blobs.append({
            "row_min": row_min, "row_max": row_max, "col_min": 0, "col_max": w - 1,
            "prow": row, "pcol": pcol,
            "strength": float(row_strength[row]),
            "is_streak": True,
        })

    if not blobs:
        return []
    # Streaks first, then by strength: a streak's representative peak is
    # found via the same argmax as a point-blob at that same location, so
    # they can end up with an identical (prow, pcol) and therefore an
    # identical mirror pair-key. Sorting by strength alone lets whichever
    # was merely first in the list claim that pair-key - if that happened
    # to be the narrow point-blob, the far more informative wide streak got
    # silently discarded as a "duplicate" of a tiny fraction of itself.
    # Streaks must win that conflict, since the point is already contained
    # inside the streak's bounding box.
    blobs.sort(key=lambda b: (b["is_streak"], b["strength"]), reverse=True)

    # A real image's FFT is conjugate-symmetric, so every region has a
    # mirror at (-fx, -fy) with identical content - keep only one of each
    # pair, always the canonical (representative point with non-negative
    # fx, or fx==0 and non-negative fy) member.
    kept = []
    seen_pairs = set()
    for blob in blobs:
        prow, pcol = blob["prow"], blob["pcol"]
        mirror_prow, mirror_pcol = (h - prow) % h, (w - pcol) % w
        pair_key = frozenset({(prow, pcol), (mirror_prow, mirror_pcol)})
        if pair_key in seen_pairs:
            continue
        seen_pairs.add(pair_key)
        fx_p, fy_p = float(freqs_x[pcol]), float(freqs_y[prow])
        fx_min_o = float(freqs_x[blob["col_min"]])
        fx_max_o = float(freqs_x[blob["col_max"]])
        fy_min_o = float(freqs_y[blob["row_min"]])
        fy_max_o = float(freqs_y[blob["row_max"]])
        if fx_p < 0 or (fx_p == 0 and fy_p < 0):
            # Reflect the whole bounding box through the origin in float
            # frequency space (a simple negate-and-swap) rather than
            # reflecting integer indices - index reflection for an n-length
            # fftshift'd axis is (n - i) % n, which wraps at i=0 and breaks
            # for a bounding box that already spans the full axis (as a
            # streak's does), while negating the frequency values has no
            # such edge case.
            fx_min, fx_max = -fx_max_o, -fx_min_o
            fy_min, fy_max = -fy_max_o, -fy_min_o
            fx_rep, fy_rep = -fx_p, -fy_p
        else:
            fx_min, fx_max = fx_min_o, fx_max_o
            fy_min, fy_max = fy_min_o, fy_max_o
            fx_rep, fy_rep = fx_p, fy_p
        kept.append({
            "fx_min": fx_min, "fx_max": fx_max,
            "fy_min": fy_min, "fy_max": fy_max,
            "fx": fx_rep, "fy": fy_rep,
            "strength": blob["strength"], "is_streak": blob["is_streak"],
            "shape": "rect", "mode": "remove",
        })
        if len(kept) >= max_regions:
            break

    for region in kept:
        region["radius"] = float(np.hypot(region["fx"], region["fy"]))

    timing_available = expected is not None
    for region in kept:
        radius = region["radius"]
        candidates = []
        if expected:
            for exp in expected:
                dist = np.hypot(region["fx"] - exp["fx"], region["fy"] - exp["fy"])
                if dist <= _MATCH_TOLERANCE:
                    candidates.append((exp, dist))
        match = None
        if candidates:
            # Harmonics of different base frequencies can alias to the same
            # spatial position - genuinely ambiguous from the image alone,
            # so among equally-close matches, prefer the lowest Hz: a
            # fundamental is a more plausible real-world source than an
            # obscure harmonic of something else.
            exp, dist = min(candidates, key=lambda ed: (round(ed[1], 6), ed[0]["hz"]))
            match = exp
        if match is not None:
            region["hz"] = match["hz"]
            region["tip_category"] = "noise"
            region["tip"] = (
                f"Matches ~{match['hz']:g} Hz (pump/mains harmonic) - likely external noise, safe to remove."
            )
        elif region["is_streak"]:
            region["hz"] = None
            region["tip_category"] = "noise"
            region["tip"] = (
                "Spans most of one axis at a fixed frequency on the other - the signature of "
                "scan-line banding or phase-drifting noise, not real surface structure. Likely safe to remove."
            )
        else:
            siblings = [
                r for r in kept
                if r is not region and abs(r["radius"] - radius) <= _LATTICE_RADIUS_TOLERANCE
            ]
            if len(siblings) >= 2:
                region["hz"] = None
                region["tip_category"] = "caution"
                region["tip"] = (
                    "Positioned like real periodic surface structure (e.g. atomic lattice or "
                    "moire pattern) - review carefully before removing."
                )
            else:
                region["hz"] = None
                region["tip_category"] = "unclear"
                region["tip"] = (
                    "Origin unclear - no timing match, no lattice-like symmetry. Inspect the image before removing."
                )
        if not timing_available and region["tip_category"] != "noise":
            region["tip"] = (
                "Scan timing not available for this file, so pump/mains-frequency matching "
                "couldn't be checked. " + region["tip"]
            )

    kept.sort(key=lambda r: r["strength"], reverse=True)
    return kept


def _region_bounds(region):
    """Accept either a rich region dict (from detect_candidate_regions) or a
    plain (shape, mode, fx_min, fx_max, fy_min, fy_max) tuple (the hashable
    form stored in a filter step's params, since FILTER_DEFINITIONS steps
    must round-trip through _filter_signature, which requires every param
    value to be hashable - a list of dicts isn't). Returns
    (shape, mode, fx_min, fx_max, fy_min, fy_max)."""
    if isinstance(region, dict):
        return (
            str(region.get("shape", "rect")), str(region.get("mode", "remove")),
            float(region["fx_min"]), float(region["fx_max"]),
            float(region["fy_min"]), float(region["fy_max"]),
        )
    shape, mode, fx_min, fx_max, fy_min, fy_max = region
    return str(shape), str(mode), float(fx_min), float(fx_max), float(fy_min), float(fy_max)


def _outside_fraction(fxx, fyy, shape, fx_min, fx_max, fy_min, fy_max, taper):
    """0 inside the shape, ramping smoothly to 1 over `taper` (cycles/pixel)
    outside its boundary, clipped to [0, 1]."""
    if shape == "ellipse":
        cx, cy = (fx_min + fx_max) / 2.0, (fy_min + fy_max) / 2.0
        rx = max((fx_max - fx_min) / 2.0, 1e-6)
        ry = max((fy_max - fy_min) / 2.0, 1e-6)
        # Normalized radial distance (1.0 = exactly on the boundary), then
        # converted back to a cycles/pixel-equivalent distance so the same
        # `taper` value means the same thing for both shapes.
        r_norm = np.sqrt(((fxx - cx) / rx) ** 2 + ((fyy - cy) / ry) ** 2)
        dist_outside = (r_norm - 1.0) * ((rx + ry) / 2.0)
    else:
        dist_outside = np.maximum(
            np.maximum(fx_min - fxx, fxx - fx_max),
            np.maximum(fy_min - fyy, fyy - fy_max),
        )
    return np.clip(dist_outside / taper, 0.0, 1.0)


def apply_mask_filter(arr, regions, taper=0.01):
    """Attenuate the given frequency regions (in cycles/pixel, from
    detect_candidate_regions or manually drawn) and their conjugate
    mirrors, with a smooth cosine taper at each edge (not a hard cutoff, to
    avoid ringing), then invert back to real space. A region can be a
    rectangle or an ellipse, and can span an entire axis (a streak/band),
    unlike a single circular notch, so a strong scan-line-banding artifact
    actually gets removed rather than leaving most of its energy untouched.

    Each region has a mode: "remove" (attenuate) or "protect" (force the
    result back toward unattenuated, even where a "remove" region overlaps
    it). Protect regions are applied after all remove regions, letting a
    user draw one broad, sloppy removal region and then carve out a
    smaller protected zone (e.g. the DC-adjacent core, where genuine
    low-frequency topography and noise close to zero frequency can't be
    separated - removing one always costs some of the other) rather than
    needing to get every removal region's boundary surgically precise.

    taper is in cycles/pixel - how far outside each region's edge the
    attenuation smoothly ramps back up to (or, for protect, away from) 1.0.
    """
    if not regions:
        return np.asarray(arr, dtype=float).copy()
    detrended, plane = _detrend(arr)
    h, w = detrended.shape
    spectrum = np.fft.fftshift(np.fft.fft2(detrended))
    freqs_x = np.fft.fftshift(np.fft.fftfreq(w))
    freqs_y = np.fft.fftshift(np.fft.fftfreq(h))
    fxx, fyy = np.meshgrid(freqs_x, freqs_y)
    taper = max(float(taper), 1e-6)

    atten = np.ones((h, w), dtype=float)
    for region in regions:
        shape, mode, fx_min, fx_max, fy_min, fy_max = _region_bounds(region)
        if mode != "remove":
            continue
        for x_lo, x_hi, y_lo, y_hi in (
            (fx_min, fx_max, fy_min, fy_max),
            (-fx_max, -fx_min, -fy_max, -fy_min),
        ):
            outside_frac = _outside_fraction(fxx, fyy, shape, x_lo, x_hi, y_lo, y_hi, taper)
            gain = 0.5 - 0.5 * np.cos(np.pi * outside_frac)
            atten *= np.where(outside_frac >= 1.0, 1.0, 0.02 + 0.98 * gain)

    for region in regions:
        shape, mode, fx_min, fx_max, fy_min, fy_max = _region_bounds(region)
        if mode != "protect":
            continue
        for x_lo, x_hi, y_lo, y_hi in (
            (fx_min, fx_max, fy_min, fy_max),
            (-fx_max, -fx_min, -fy_max, -fy_min),
        ):
            outside_frac = _outside_fraction(fxx, fyy, shape, x_lo, x_hi, y_lo, y_hi, taper)
            # 1.0 at the protected region's center, smoothly fading to 0
            # (no effect) beyond its taper zone - blending atten toward 1.0
            # rather than hard-setting it, so the transition stays smooth.
            protect_strength = 1.0 - outside_frac
            atten = atten * (1.0 - protect_strength) + protect_strength

    filtered_spectrum = spectrum * atten
    result = np.fft.ifft2(np.fft.ifftshift(filtered_spectrum)).real
    return result + plane


__all__ = [
    "fft_magnitude",
    "expected_peak_frequencies",
    "detect_candidate_regions",
    "apply_mask_filter",
]
