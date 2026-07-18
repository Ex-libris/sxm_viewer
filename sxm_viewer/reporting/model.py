"""Derive the report structure from a collected folder payload.

Input payload schema (plain data, built by ``gui/controllers/report.py``):

``payload["images"]`` - list of dicts per loaded image::

    key (str), name (str), time (datetime|None),
    extent ([x0, x1, y1, y0] nm, same convention as _header_extent),
    angle (deg), xpix, ypix,
    bias_v (float|None), z_level_nm (float|None, median of the raw topo
    channel - for a constant-height scan this is the absolute tip height),
    z_spread_nm (float|None, p98-p2 of the raw topo channel: ~0 means a
    genuinely flat constant-height plane),
    topo (2D np.ndarray|None), topo_unit (str),
    meta (dict of short display strings)

``payload["specs"]`` - list of dicts per spectroscopy point::

    name, path, x, y (nm|None), time (datetime|None),
    image_key (str|None), marker_frac ((fx, fy) in [0,1]|None),
    V (np.ndarray|None), curve (np.ndarray|None), curve_channel (str),
    off_frame_direction, off_frame_distance_nm,
    assignment_reason_label, assignment_confidence,
    is_matrix (bool), matrix_dataset (str|None),
    grid_row, grid_col, grid_rows, grid_cols, matrix_index

``payload["grids"]`` - list of dicts per matrix dataset::

    dataset (str), name (str), image_key (str|None),
    spec_indices (indices into payload["specs"])

Everything here is pure numpy/scipy - no Qt (see package docstring).
"""
from __future__ import annotations

import math
from datetime import datetime

import numpy as np

from .channels import classify_channel

try:
    from scipy.cluster.vq import kmeans2
except Exception:  # pragma: no cover - scipy is a hard app dependency
    kmeans2 = None

try:
    from ..data.spectroscopy import fit_parabola_bias
except Exception:  # pragma: no cover
    fit_parabola_bias = None


# ---------------------------------------------------------------------------
# K-Means helpers
# ---------------------------------------------------------------------------

def _zscore(features):
    """Column-wise z-score; constant columns become 0 instead of NaN."""
    feats = np.asarray(features, dtype=float)
    mean = np.nanmean(feats, axis=0)
    std = np.nanstd(feats, axis=0)
    std[std == 0] = 1.0
    return (feats - mean) / std


def _kmeans_labels(features, k, seed=1234):
    """Labels from scipy kmeans2, or None when clustering is unusable."""
    if kmeans2 is None or k < 1:
        return None
    feats = np.asarray(features, dtype=float)
    if feats.ndim != 2 or feats.shape[0] < k:
        return None
    if k == 1:
        return np.zeros(feats.shape[0], dtype=int)
    try:
        _, labels = kmeans2(feats, k, minit="++", seed=seed, missing="raise")
    except Exception:
        return None
    return np.asarray(labels, dtype=int)


def _inertia(features, labels):
    feats = np.asarray(features, dtype=float)
    total = 0.0
    for lab in np.unique(labels):
        members = feats[labels == lab]
        if members.size:
            total += float(((members - members.mean(axis=0)) ** 2).sum())
    return total


def choose_k_elbow(features, k_max, improvement_threshold=0.2, seed=1234):
    """Scan k=1..k_max and keep adding clusters while each extra cluster
    still removes at least ``improvement_threshold`` of the remaining
    within-cluster variance (a plain elbow rule). Returns (k, labels)."""
    feats = np.asarray(features, dtype=float)
    n = feats.shape[0] if feats.ndim == 2 else 0
    if n == 0:
        return 0, None
    best_labels = np.zeros(n, dtype=int)
    if n == 1 or k_max <= 1:
        return 1, best_labels
    best_k = 1
    total_inertia = _inertia(feats, best_labels)
    prev_inertia = total_inertia
    for k in range(2, min(k_max, n) + 1):
        if total_inertia <= 0 or prev_inertia <= 0:
            break
        if prev_inertia / total_inertia < 0.05:
            # Current clustering already explains >95% of the variance -
            # further splits only subdivide genuinely tight groups (the
            # relative-improvement test below stays spuriously high once
            # the residual is tiny).
            break
        labels = _kmeans_labels(feats, k, seed=seed)
        if labels is None or len(np.unique(labels)) < k:
            break
        inertia = _inertia(feats, labels)
        improvement = (prev_inertia - inertia) / prev_inertia
        if improvement < improvement_threshold:
            break
        best_k, best_labels, prev_inertia = k, labels, inertia
    return best_k, best_labels


# ---------------------------------------------------------------------------
# Sample regions (space+time K-Means over images): scans acquired in the
# same sample area during the same time window group into one "Region".
# ---------------------------------------------------------------------------

def _image_features(image):
    """(center_x_nm, center_y_nm, epoch_seconds) or None when unusable."""
    extent = image.get("extent")
    if not extent or len(extent) != 4:
        return None
    x0, x1, y1, y0 = [float(v) for v in extent]
    if not all(math.isfinite(v) for v in (x0, x1, y0, y1)):
        return None
    cx = 0.5 * (x0 + x1)
    cy = 0.5 * (y0 + y1)
    t = image.get("time")
    epoch = t.timestamp() if isinstance(t, datetime) else None
    return cx, cy, epoch


def build_session_regions(images, k_max=8):
    """Cluster images into sample "Regions" by (center x, center y, time) -
    i.e. scans taken in the same area of the sample around the same time.

    Returns (regions, image_region) where regions is a list of dicts
    ``{"label", "image_keys", "t0", "t1"}`` ordered by mean acquisition
    time and image_region maps every image key to its region index.
    Images without usable features land in the region of the closest-in
    -time clustered image (or region 0 as a last resort).
    """
    usable = []
    for img in images:
        feats = _image_features(img)
        if feats is not None:
            usable.append((img, feats))
    image_region = {}
    if not usable:
        if not images:
            return [], {}
        region = {"label": "Region 1", "image_keys": [str(i["key"]) for i in images],
                  "t0": None, "t1": None}
        return [region], {str(i["key"]): 0 for i in images}

    have_time = [f[2] is not None for _, f in usable]
    median_epoch = float(np.median([f[2] for _, f in usable if f[2] is not None])) if any(have_time) else 0.0
    matrix = [[f[0], f[1], f[2] if f[2] is not None else median_epoch] for _, f in usable]
    feats = _zscore(matrix)
    _, labels = choose_k_elbow(feats, k_max)
    if labels is None:
        labels = np.zeros(len(usable), dtype=int)

    groups = {}
    for (img, _), lab in zip(usable, labels):
        groups.setdefault(int(lab), []).append(img)

    def _mean_epoch(imgs):
        times = [i["time"].timestamp() for i in imgs if isinstance(i.get("time"), datetime)]
        return float(np.mean(times)) if times else float("inf")

    ordered = sorted(groups.values(), key=_mean_epoch)
    regions = []
    for idx, imgs in enumerate(ordered):
        imgs.sort(key=lambda i: i.get("time") or datetime.min)
        times = [i["time"] for i in imgs if isinstance(i.get("time"), datetime)]
        regions.append({
            "label": f"Region {idx + 1}",
            "image_keys": [str(i["key"]) for i in imgs],
            "t0": min(times) if times else None,
            "t1": max(times) if times else None,
        })
        for img in imgs:
            image_region[str(img["key"])] = idx

    # Park images that couldn't be clustered with their nearest-in-time region.
    clustered_times = [(i.get("time"), image_region[str(i["key"])])
                       for imgs in ordered for i in imgs if isinstance(i.get("time"), datetime)]
    for img in images:
        key = str(img["key"])
        if key in image_region:
            continue
        t = img.get("time")
        region_idx = 0
        if isinstance(t, datetime) and clustered_times:
            region_idx = min(clustered_times, key=lambda ct: abs((ct[0] - t).total_seconds()))[1]
        image_region[key] = region_idx
        regions[region_idx]["image_keys"].append(key)
    return regions, image_region


# ---------------------------------------------------------------------------
# Image sequences (data-sets): the same frame re-scanned while stepping one
# acquisition parameter - e.g. constant-height current/df images taken at a
# series of tip-sample distances (z_rel), or scans of the same spot at a
# series of biases. Detected by footprint identity (center + size + angle
# within tolerance), then classified by which parameter actually varies.
# ---------------------------------------------------------------------------

_SEQ_CENTER_FRAC = 0.25   # allowed center drift, as a fraction of frame size
_SEQ_SIZE_FRAC = 0.10     # allowed relative width/height mismatch
_SEQ_ANGLE_TOL_DEG = 2.0
_SEQ_BIAS_TOL_V = 2e-3    # bias span below this counts as "constant"
_SEQ_Z_TOL_NM = 5e-3      # 5 pm; z-level span below this is just noise
_SEQ_Z_FLAT_NM = 0.02     # topo spread above this is real topography, so the
                          # median z is surface, not a controlled tip height


def _sequence_geometry(image):
    """(cx, cy, w, h, angle) of an image's frame in nm, or None."""
    extent = image.get("extent")
    if not extent or len(extent) != 4:
        return None
    x0, x1, y1, y0 = [float(v) for v in extent]
    if not all(math.isfinite(v) for v in (x0, x1, y0, y1)):
        return None
    w, h = abs(x1 - x0), abs(y1 - y0)
    if w <= 0 or h <= 0:
        return None
    return (0.5 * (x0 + x1), 0.5 * (y0 + y1), w, h,
            float(image.get("angle") or 0.0))


def _same_footprint(g1, g2):
    cx1, cy1, w1, h1, a1 = g1
    cx2, cy2, w2, h2, a2 = g2
    if abs(w1 - w2) > _SEQ_SIZE_FRAC * max(w1, w2):
        return False
    if abs(h1 - h2) > _SEQ_SIZE_FRAC * max(h1, h2):
        return False
    da = abs(a1 - a2) % 360.0
    if min(da, 360.0 - da) > _SEQ_ANGLE_TOL_DEG:
        return False
    mean_span = 0.25 * (w1 + h1 + w2 + h2)
    return math.hypot(cx1 - cx2, cy1 - cy2) <= _SEQ_CENTER_FRAC * mean_span


def _finite_or_none(value):
    try:
        num = float(value)
    except (TypeError, ValueError):
        return None
    return num if math.isfinite(num) else None


def build_image_sequences(images, image_region=None):
    """Group images that re-scan the same frame into parameter series.

    Returns a list of dicts ordered by first-acquisition time::

        label (str), kind ("height"|"bias"|"bias+height"|"repeat"),
        kind_label (str), region (int|None),
        bias_span_v (float), z_span_nm (float),
        t0, t1 (datetime|None),
        members: [{key, name, time, tag, bias_v, z_level_nm, z_rel_pm}]

    ``z_rel_pm`` is the tip height relative to the sequence's first image
    (image 0 at z_rel = 0, positive = retracted), only set for genuinely
    flat (constant-height) frames. Footprint-matched pairs with nothing
    varying are dropped as plain re-scans; groups of >= 3 are kept even
    then (a time series at one spot is itself worth reporting).
    """
    usable = []
    for img in images:
        geom = _sequence_geometry(img)
        if geom is not None:
            usable.append((img, geom))
    # Time-ordered so "image 0" of each sequence is the earliest one.
    usable.sort(key=lambda ig: (ig[0].get("time") is None,
                                ig[0].get("time") or datetime.min,
                                str(ig[0].get("name") or "")))
    groups = []
    for img, geom in usable:
        for grp in groups:
            if _same_footprint(geom, grp["geom"]):
                grp["images"].append(img)
                break
        else:
            groups.append({"geom": geom, "images": [img]})

    sequences = []
    for grp in groups:
        members = grp["images"]
        if len(members) < 2:
            continue
        biases = [_finite_or_none(i.get("bias_v")) for i in members]
        # A median-z value is only a *controlled* tip height when the frame
        # is genuinely flat (constant-height); on real topography it tracks
        # the surface and would fake a height series from drift alone.
        zs = []
        for img in members:
            z = _finite_or_none(img.get("z_level_nm"))
            spread = _finite_or_none(img.get("z_spread_nm"))
            zs.append(z if (z is not None and spread is not None
                            and spread <= _SEQ_Z_FLAT_NM) else None)
        bias_vals = [b for b in biases if b is not None]
        z_vals = [z for z in zs if z is not None]
        bias_span = (max(bias_vals) - min(bias_vals)) if len(bias_vals) >= 2 else 0.0
        z_span = (max(z_vals) - min(z_vals)) if len(z_vals) >= 2 else 0.0
        bias_varies = bias_span > _SEQ_BIAS_TOL_V
        z_varies = z_span > _SEQ_Z_TOL_NM
        if bias_varies and z_varies:
            kind, kind_label = "bias+height", "Bias + height series"
        elif bias_varies:
            kind, kind_label = "bias", "Bias series"
        elif z_varies:
            kind, kind_label = "height", "Height (tip-sample distance) series"
        else:
            kind, kind_label = "repeat", "Repeated frame (time series)"
        if kind == "repeat" and len(members) < 3:
            continue  # a single re-scan of a frame is not a data-set
        z_ref = next((z for z in zs if z is not None), None)
        entries = []
        for img, bias, z in zip(members, biases, zs):
            entries.append({
                "key": str(img["key"]),
                "name": img.get("name") or "",
                "time": img.get("time"),
                "tag": img.get("tag"),
                "bias_v": bias,
                "z_level_nm": z,
                "z_rel_pm": ((z - z_ref) * 1e3)
                            if (z is not None and z_ref is not None) else None,
            })
        times = [e["time"] for e in entries if isinstance(e["time"], datetime)]
        region = None
        if image_region:
            votes = [image_region.get(e["key"]) for e in entries
                     if image_region.get(e["key"]) is not None]
            if votes:
                region = max(set(votes), key=votes.count)
        sequences.append({
            "kind": kind,
            "kind_label": kind_label,
            "region": region,
            "bias_span_v": float(bias_span),
            "z_span_nm": float(z_span),
            "t0": min(times) if times else None,
            "t1": max(times) if times else None,
            "members": entries,
        })
    sequences.sort(key=lambda s: (s["t0"] is None, s["t0"] or datetime.min))
    for idx, seq in enumerate(sequences):
        seq["label"] = f"Sequence {idx + 1}"
    return sequences


# ---------------------------------------------------------------------------
# Grid geometry (ports of MatrixSpectroViewer helpers)
# ---------------------------------------------------------------------------
# These deliberately mirror _grid_dims/_spec_grid_row_col/_grid_xy_coords/
# _grid_local_pitch/_grid_local_orientation in
# gui/dialogs/spectroscopy_dialogs.py (which are methods on a QDialog, so
# they can't be imported here without pulling Qt in). If the conventions
# there change - especially the local-frame flip rules - these must change
# with them, or report grid maps will mirror relative to the Grid Map
# Explorer's relative view.

def grid_dims(specs):
    if not specs:
        return 0, 0, True
    rows = max((s.get("grid_rows") or 0) for s in specs)
    cols = max((s.get("grid_cols") or 0) for s in specs)
    zero_based = True
    if rows and cols:
        matrix_indices = [s.get("matrix_index") for s in specs if s.get("matrix_index") is not None]
        if matrix_indices:
            min_idx, max_idx = min(matrix_indices), max(matrix_indices)
            if min_idx >= 1 and max_idx == rows * cols:
                zero_based = False
    return rows, cols, zero_based


def spec_grid_row_col(spec, rows, cols, zero_based):
    row = spec.get("grid_row")
    col = spec.get("grid_col")
    if row is not None and col is not None:
        try:
            return int(row), int(col)
        except Exception:
            pass
    if not cols:
        return None
    idx = spec.get("matrix_index")
    if idx is None:
        return None
    try:
        idx_val = int(idx)
    except Exception:
        return None
    if not zero_based:
        idx_val -= 1
    return idx_val // cols, idx_val % cols


def grid_xy_coords(specs, rows, cols, zero_based):
    if not specs or not rows or not cols:
        return None, None
    X = np.full((rows, cols), np.nan, dtype=float)
    Y = np.full((rows, cols), np.nan, dtype=float)
    for spec in specs:
        rc = spec_grid_row_col(spec, rows, cols, zero_based)
        if rc is None or not (0 <= rc[0] < rows and 0 <= rc[1] < cols):
            continue
        x, y = spec.get("x"), spec.get("y")
        if x is None or y is None:
            continue
        X[rc[0], rc[1]] = float(x)
        Y[rc[0], rc[1]] = float(y)
    return X, Y


def grid_local_pitch(X, Y, rows, cols):
    dx = dy = 1.0
    if cols > 1:
        step = np.hypot(np.diff(X, axis=1), np.diff(Y, axis=1))
        finite = step[np.isfinite(step)]
        if finite.size:
            dx = float(np.nanmedian(finite)) or 1.0
    if rows > 1:
        step = np.hypot(np.diff(X, axis=0), np.diff(Y, axis=0))
        finite = step[np.isfinite(step)]
        if finite.size:
            dy = float(np.nanmedian(finite)) or 1.0
    return dx, dy


def grid_local_orientation(X, Y, rows, cols, angle_deg=0.0):
    """Flatten-orientation flips, decided in the ANCHOR IMAGE's raster
    frame (rotate coords by the anchor scan angle first) so the local view
    matches the thumbnail/reference image - same convention as
    _grid_local_orientation in spectroscopy_dialogs.py; angle 0 reduces to
    the plain north-up test."""
    row_flip = False
    col_flip = False
    theta = math.radians(float(angle_deg or 0.0))
    if theta:
        cos_t, sin_t = math.cos(theta), math.sin(theta)
        U = X * cos_t - Y * sin_t
        V = X * sin_t + Y * cos_t
    else:
        U, V = X, Y
    if rows > 1:
        v_first = np.nanmean(V[0, :])
        v_last = np.nanmean(V[-1, :])
        if np.isfinite(v_first) and np.isfinite(v_last):
            row_flip = bool(v_last > v_first)
    if cols > 1:
        u_first = np.nanmean(U[:, 0])
        u_last = np.nanmean(U[:, -1])
        if np.isfinite(u_first) and np.isfinite(u_last):
            col_flip = bool(u_last < u_first)
    return row_flip, col_flip


# ---------------------------------------------------------------------------
# Per-grid model: slice map + curve-shape K-Means
# ---------------------------------------------------------------------------

def _resample_curves(specs):
    """Common-axis curve matrix. Returns (V_ref, matrix, spec_indices_used)."""
    ref_v = None
    for spec in specs:
        v, c = spec.get("V"), spec.get("curve")
        if v is None or c is None:
            continue
        v = np.asarray(v, dtype=float)
        if v.size >= 2 and (ref_v is None or v.size > ref_v.size):
            ref_v = v
    if ref_v is None:
        return None, None, []
    order = np.argsort(ref_v)
    ref_v = ref_v[order]
    rows, used = [], []
    for i, spec in enumerate(specs):
        v, c = spec.get("V"), spec.get("curve")
        if v is None or c is None:
            continue
        v = np.asarray(v, dtype=float)
        c = np.asarray(c, dtype=float)
        if v.size != c.size or v.size < 2:
            continue
        o = np.argsort(v)
        resampled = np.interp(ref_v, v[o], c[o])
        if not np.all(np.isfinite(resampled)):
            continue
        rows.append(resampled)
        used.append(i)
    if not rows:
        return None, None, []
    return ref_v, np.vstack(rows), used


def build_grid_model(grid, specs, k_max=4, anchor_angle=0.0):
    """Slice map, curve-cluster map and representative curves for one
    matrix dataset. Arrays come back already display-flipped for the local
    frame (imshow with origin='upper', extent (0, cols*dx, rows*dy, 0)),
    matching the Grid Map Explorer's "Relative axes" convention (oriented
    into the anchor image's raster frame via ``anchor_angle``)."""
    rows, cols, zero_based = grid_dims(specs)
    if not rows or not cols:
        return None
    X, Y = grid_xy_coords(specs, rows, cols, zero_based)
    if X is None:
        return None
    dx, dy = grid_local_pitch(X, Y, rows, cols)
    row_flip, col_flip = grid_local_orientation(X, Y, rows, cols, angle_deg=anchor_angle)

    ref_v, curve_matrix, used = _resample_curves(specs)

    first = specs[0] if specs else {}
    x_label = str(first.get("x_label") or "Bias (V)")
    x_is_bias = x_label.lower().startswith("bias")

    slice_map = np.full((rows, cols), np.nan, dtype=float)
    cluster_map = np.full((rows, cols), np.nan, dtype=float)
    cluster_curves = []
    bias_value = None
    if ref_v is not None and curve_matrix is not None and len(used):
        # Slice at the smallest |bias| available (user heuristic: "the slice
        # value at some arbitrary small bias"); for non-bias sweeps just use
        # the middle of the axis.
        slice_idx = int(np.argmin(np.abs(ref_v))) if x_is_bias else ref_v.size // 2
        bias_value = float(ref_v[slice_idx])
        for local_i, spec_i in enumerate(used):
            rc = spec_grid_row_col(specs[spec_i], rows, cols, zero_based)
            if rc is None or not (0 <= rc[0] < rows and 0 <= rc[1] < cols):
                continue
            slice_map[rc[0], rc[1]] = curve_matrix[local_i, slice_idx]

        # K-Means on normalized curve shapes. Smooth first: point-to-point
        # noise inflates within-cluster variance uniformly across any split,
        # which drowns the elbow test's relative-improvement signal and made
        # a real grid with obvious molecule/substrate contrast come out k=1.
        smoothed = curve_matrix
        if curve_matrix.shape[1] >= 15:
            kernel = np.ones(5) / 5.0
            smoothed = np.apply_along_axis(
                lambda row: np.convolve(row, kernel, mode="same"), 1, curve_matrix)
        std = smoothed.std(axis=1, keepdims=True)
        std[std == 0] = 1.0
        normed = (smoothed - smoothed.mean(axis=1, keepdims=True)) / std
        k, labels = choose_k_elbow(normed, k_max, improvement_threshold=0.10)
        if labels is None:
            labels = np.zeros(len(used), dtype=int)
            k = 1
        for local_i, spec_i in enumerate(used):
            rc = spec_grid_row_col(specs[spec_i], rows, cols, zero_based)
            if rc is None or not (0 <= rc[0] < rows and 0 <= rc[1] < cols):
                continue
            cluster_map[rc[0], rc[1]] = float(labels[local_i])
        for lab in range(int(k)):
            member_mask = labels == lab
            members = curve_matrix[member_mask]
            if not members.shape[0]:
                continue
            # A real member spectrum (the one nearest the cluster centroid
            # in normalized-shape space), NOT an average - averaging grid
            # spectra washes out the physics (user heuristic #6).
            centroid = normed[member_mask].mean(axis=0)
            rep_local = int(np.argmin(((normed[member_mask] - centroid) ** 2).sum(axis=1)))
            rep_row = int(np.flatnonzero(member_mask)[rep_local])
            cluster_curves.append({
                "label": lab,
                "count": int(members.shape[0]),
                "V": ref_v,
                "curve": curve_matrix[rep_row],
            })

    kpfm = _grid_kpfm_maps(specs, rows, cols, zero_based) if x_is_bias else None

    def _display(arr):
        out = arr
        if row_flip:
            out = np.flipud(out)
        if col_flip:
            out = np.fliplr(out)
        return out

    if kpfm is not None:
        kpfm = {key: (_display(val) if isinstance(val, np.ndarray) else val)
                for key, val in kpfm.items()}

    return {
        "dataset": grid.get("dataset"),
        "name": grid.get("name") or str(grid.get("dataset") or ""),
        "image_key": grid.get("image_key"),
        "rows": rows,
        "cols": cols,
        "n_points": len(specs),
        "local_dx": dx,
        "local_dy": dy,
        "row_flip": row_flip,
        "col_flip": col_flip,
        "X": X,
        "Y": Y,
        "slice_map": _display(slice_map),
        "cluster_map": _display(cluster_map),
        "cluster_curves": cluster_curves,
        "bias_value": bias_value,
        "curve_channel": first.get("curve_channel") or "",
        "x_label": x_label,
        "kpfm": kpfm,
        "time": first.get("time"),
    }


def _grid_kpfm_maps(specs, rows, cols, zero_based, min_ok_fraction=0.6):
    """Per-pixel parabola (KPFM) fit maps for a bias-swept grid whose
    frequency-shift channel actually varies with bias: LCPD (vertex bias,
    -b/2a) and the c coefficient, each with 1-sigma errors. Returns None
    when the grid has no varying frequency-shift channel or too few pixels
    fit cleanly."""
    if fit_parabola_bias is None:
        return None
    lcpd = np.full((rows, cols), np.nan, dtype=float)
    lcpd_err = np.full((rows, cols), np.nan, dtype=float)
    c_map = np.full((rows, cols), np.nan, dtype=float)
    c_err = np.full((rows, cols), np.nan, dtype=float)
    n_total = 0
    n_ok = 0
    channel_used = None
    v_min = v_max = None
    for spec in specs:
        rc = spec_grid_row_col(spec, rows, cols, zero_based)
        if rc is None or not (0 <= rc[0] < rows and 0 <= rc[1] < cols):
            continue
        v = spec.get("V")
        curves = spec.get("curves") or {}
        df = None
        for name, values in curves.items():
            if classify_channel(name) == "freq":
                df = values
                channel_used = channel_used or name
                break
        if v is None or df is None:
            continue
        v = np.asarray(v, dtype=float)
        df = np.asarray(df, dtype=float)
        if v.size != df.size or v.size < 5:
            continue
        n_total += 1
        try:
            fit = fit_parabola_bias(v, df)
        except Exception:
            continue
        a, b = fit["a"], fit["b"]
        if not (np.isfinite(a) and np.isfinite(b)) or a == 0:
            continue
        vertex = -b / (2.0 * a)
        if v_min is None:
            v_min, v_max = float(np.nanmin(v)), float(np.nanmax(v))
        span = (v_max - v_min) or 1.0
        # A vertex far outside the sweep means the "parabola" is really a
        # line/noise at this pixel - drop it rather than distorting the map.
        if not (v_min - span <= vertex <= v_max + span):
            continue
        # 1-sigma error via propagation from the (uncorrelated) a/b errors.
        vertex_err = abs(vertex) * math.sqrt(
            (fit["a_err"] / a) ** 2 + ((fit["b_err"] / b) ** 2 if b else 0.0))
        r, ccol = rc
        lcpd[r, ccol] = vertex
        lcpd_err[r, ccol] = vertex_err
        c_map[r, ccol] = fit["c"]
        c_err[r, ccol] = fit["c_err"]
        n_ok += 1
    if not n_total or (n_ok / float(n_total)) < min_ok_fraction:
        return None
    return {
        "lcpd": lcpd,
        "lcpd_err": lcpd_err,
        "c": c_map,
        "c_err": c_err,
        "n_ok": n_ok,
        "n_total": n_total,
        "channel": channel_used or "",
    }


# ---------------------------------------------------------------------------
# Flagged items ("needs attention" appendix)
# ---------------------------------------------------------------------------

def build_flagged(specs):
    off_frame, low_confidence, unassigned = [], [], []
    for spec in specs:
        if not spec.get("image_key"):
            unassigned.append(spec)
            continue
        if spec.get("off_frame_direction"):
            off_frame.append(spec)
        conf = str(spec.get("assignment_confidence") or "").lower()
        if conf == "low":
            low_confidence.append(spec)
    # One appendix row per matrix dataset, not one per grid point.
    def _dedupe(items):
        seen, out = set(), []
        for spec in items:
            key = spec.get("matrix_dataset") if spec.get("is_matrix") else id(spec)
            if key in seen:
                continue
            seen.add(key)
            out.append(spec)
        return out
    return {
        "off_frame": _dedupe(off_frame),
        "low_confidence": _dedupe(low_confidence),
        "unassigned": _dedupe(unassigned),
    }


# ---------------------------------------------------------------------------
# Top level
# ---------------------------------------------------------------------------

def build_report_model(payload):
    images = list(payload.get("images") or [])
    specs = list(payload.get("specs") or [])
    grids = list(payload.get("grids") or [])

    regions, image_region = build_session_regions(images)
    sequences = build_image_sequences(images, image_region)

    angle_by_key = {str(i["key"]): float(i.get("angle") or 0.0) for i in images}
    grid_models = []
    for grid in grids:
        g_specs = [specs[i] for i in grid.get("spec_indices") or [] if 0 <= i < len(specs)]
        anchor_angle = angle_by_key.get(str(grid.get("image_key") or ""), 0.0)
        model = build_grid_model(grid, g_specs, anchor_angle=anchor_angle)
        if model is not None:
            grid_models.append(model)

    times = [i["time"] for i in images if isinstance(i.get("time"), datetime)]
    times += [s["time"] for s in specs if isinstance(s.get("time"), datetime)]
    single_specs = [s for s in specs if not s.get("is_matrix")]

    specs_by_image = {}
    for spec in single_specs:
        key = spec.get("image_key")
        if key:
            specs_by_image.setdefault(str(key), []).append(spec)

    return {
        "payload": payload,
        "regions": regions,
        "image_region": image_region,
        "sequences": sequences,
        "specs_by_image": specs_by_image,
        "grids": grid_models,
        "flagged": build_flagged(specs),
        "summary": {
            "folder": payload.get("folder") or "",
            "n_images": len(images),
            "n_sequences": len(sequences),
            "n_specs": len(single_specs),
            "n_grid_points": sum(1 for s in specs if s.get("is_matrix")),
            "n_grids": len(grid_models),
            "t0": min(times) if times else None,
            "t1": max(times) if times else None,
        },
    }
