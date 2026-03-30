"""Interactive A/B image comparison for preview and popup canvases.

The controller keeps lightweight frozen copies of the selected views so a
comparison stays valid even if the source preview is changed or the source
popup is closed. The dialog focuses on rigid alignment controls that are
practical for microscopy troubleshooting: manual rotation/translation plus
auto-translate / auto-rigid helpers.
"""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

import numpy as np
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from ..._shared import QtCore, QtGui, QtWidgets

try:  # pragma: no cover - optional dependency
    from scipy import ndimage

    _HAS_SCIPY = True
except Exception:  # pragma: no cover - optional dependency
    ndimage = None
    _HAS_SCIPY = False

try:  # pragma: no cover - optional dependency
    from skimage.registration import phase_cross_correlation

    _HAS_SKIMAGE = True
except Exception:  # pragma: no cover - optional dependency
    phase_cross_correlation = None
    _HAS_SKIMAGE = False

try:  # pragma: no cover - optional dependency
    from skimage import exposure as _sk_exposure

    _HAS_SKIMAGE_EXPOSURE = True
except Exception:  # pragma: no cover - optional dependency
    _sk_exposure = None
    _HAS_SKIMAGE_EXPOSURE = False


def _safe_label(snapshot):
    if not snapshot:
        return "Empty"
    return str(snapshot.get("label") or "Unnamed view")


def _finite_fill(arr):
    data = np.asarray(arr, dtype=float)
    finite = np.isfinite(data)
    if not finite.any():
        return np.zeros_like(data, dtype=float), finite, 0.0
    fill = float(np.nanmedian(data[finite]))
    out = np.array(data, copy=True)
    out[~finite] = fill
    return out, finite, fill


def _resize_nearest(arr, shape):
    data = np.asarray(arr, dtype=float)
    if tuple(data.shape[:2]) == tuple(shape):
        return np.array(data, copy=True)
    rows = np.clip(np.rint(np.linspace(0, data.shape[0] - 1, shape[0])).astype(int), 0, data.shape[0] - 1)
    cols = np.clip(np.rint(np.linspace(0, data.shape[1] - 1, shape[1])).astype(int), 0, data.shape[1] - 1)
    return np.array(data[rows][:, cols], copy=True)


def _resize_to_shape(arr, shape):
    data = np.asarray(arr, dtype=float)
    if tuple(data.shape[:2]) == tuple(shape):
        return np.array(data, copy=True)
    if _HAS_SCIPY and ndimage is not None:
        zoom = (float(shape[0]) / float(max(1, data.shape[0])), float(shape[1]) / float(max(1, data.shape[1])))
        try:
            return np.array(ndimage.zoom(data, zoom=zoom, order=1, mode="nearest"), copy=True)
        except Exception:
            pass
    return _resize_nearest(data, shape)


def _coord_to_fractional_index(coord, start, end, size):
    """Map axis-space coordinates onto fractional pixel indices for resampling."""
    if size <= 1 or abs(float(end) - float(start)) <= 1e-12:
        return np.zeros_like(coord, dtype=float)
    coord_arr = np.asarray(coord, dtype=float)
    if float(end) > float(start):
        t_val = (coord_arr - float(start)) / float(end - start)
    else:
        t_val = (coord_arr - float(end)) / float(start - end)
    return t_val * float(size - 1)


def _axis_coords_for_size(start, end, size):
    """Return the axis coordinate of each pixel index using the canvas' extent convention."""
    if size <= 0:
        return np.empty((0,), dtype=float)
    if size == 1:
        return np.array([float(end if float(end) <= float(start) else start)], dtype=float)
    frac = np.linspace(0.0, 1.0, int(size), dtype=float)
    if float(end) > float(start):
        return float(start) + (float(end) - float(start)) * frac
    return float(end) + (float(start) - float(end)) * frac


def _resample_to_reference_grid(arr, source_extent, ref_shape, ref_extent):
    """Resample an image onto the reference snapshot grid using the physical extents when available."""
    data = np.asarray(arr, dtype=float)
    if tuple(data.shape[:2]) == tuple(ref_shape) and (
        source_extent is None or ref_extent is None or tuple(source_extent) == tuple(ref_extent)
    ):
        return np.array(data, copy=True)
    if source_extent is None or ref_extent is None or not _HAS_SCIPY or ndimage is None:
        return _resize_to_shape(data, ref_shape)
    try:
        sx0, sx1, sy1, sy0 = (float(v) for v in source_extent)
        rx0, rx1, ry1, ry0 = (float(v) for v in ref_extent)
    except Exception:
        return _resize_to_shape(data, ref_shape)
    data_filled, finite, fill = _finite_fill(data)
    ref_h = max(1, int(ref_shape[0]))
    ref_w = max(1, int(ref_shape[1]))
    world_x = _axis_coords_for_size(rx0, rx1, ref_w)
    world_y = _axis_coords_for_size(ry1, ry0, ref_h)
    grid_x, grid_y = np.meshgrid(world_x, world_y)
    src_cols = _coord_to_fractional_index(grid_x, sx0, sx1, data.shape[1])
    src_rows = _coord_to_fractional_index(grid_y, sy1, sy0, data.shape[0])
    coords = np.vstack([src_rows.ravel(), src_cols.ravel()])
    sampled = ndimage.map_coordinates(
        data_filled,
        coords,
        order=1,
        mode="constant",
        cval=fill,
    ).reshape((ref_h, ref_w))
    sampled_mask = ndimage.map_coordinates(
        finite.astype(float),
        coords,
        order=0,
        mode="constant",
        cval=0.0,
    ).reshape((ref_h, ref_w)) > 0.5
    sampled = np.array(sampled, copy=True)
    sampled[~sampled_mask] = np.nan
    return sampled


def _window(shape):
    if len(shape) != 2:
        return 1.0
    wy = np.hanning(max(2, int(shape[0])))
    wx = np.hanning(max(2, int(shape[1])))
    return np.outer(wy, wx)


def _registration_image(arr):
    data, finite, _ = _finite_fill(arr)
    if finite.any():
        center = float(np.mean(data[finite]))
        scale = float(np.std(data[finite]))
    else:
        center = 0.0
        scale = 1.0
    scale = max(scale, 1e-6)
    normalized = (data - center) / scale
    return normalized * _window(normalized.shape)


def _phase_correlation_shift(reference, moving):
    ref = _registration_image(reference)
    mov = _registration_image(moving)
    if _HAS_SKIMAGE and phase_cross_correlation is not None:
        try:
            shift, error, _ = phase_cross_correlation(ref, mov, upsample_factor=10)
            dy, dx = shift
            score = 1.0 - float(error)
            return float(dx), float(dy), score
        except Exception:
            pass
    fft_ref = np.fft.fft2(ref)
    fft_mov = np.fft.fft2(mov)
    cross_power = fft_ref * np.conj(fft_mov)
    denom = np.abs(cross_power)
    denom = np.where(denom < 1e-12, 1.0, denom)
    corr = np.fft.ifft2(cross_power / denom)
    corr_abs = np.abs(corr)
    peak = np.unravel_index(int(np.argmax(corr_abs)), corr_abs.shape)
    shifts = []
    for idx, size in zip(peak, corr_abs.shape):
        shift = float(idx)
        if shift > (size / 2.0):
            shift -= float(size)
        shifts.append(shift)
    dy, dx = shifts
    return float(dx), float(dy), float(corr_abs[peak])


def _translate_integer(arr, shift_x, shift_y):
    data = np.asarray(arr, dtype=float)
    out = np.full(data.shape, np.nan, dtype=float)
    mask = np.zeros(data.shape, dtype=bool)
    dx = int(round(float(shift_x)))
    dy = int(round(float(shift_y)))
    src_y0 = max(0, -dy)
    src_y1 = min(data.shape[0], data.shape[0] - dy) if dy >= 0 else data.shape[0]
    src_x0 = max(0, -dx)
    src_x1 = min(data.shape[1], data.shape[1] - dx) if dx >= 0 else data.shape[1]
    dst_y0 = max(0, dy)
    dst_x0 = max(0, dx)
    height = max(0, src_y1 - src_y0)
    width = max(0, src_x1 - src_x0)
    if height <= 0 or width <= 0:
        return out, mask
    out[dst_y0 : dst_y0 + height, dst_x0 : dst_x0 + width] = data[src_y0:src_y1, src_x0:src_x1]
    mask[dst_y0 : dst_y0 + height, dst_x0 : dst_x0 + width] = True
    return out, mask


def _transform_with_mask(arr, rotation_deg, shift_x, shift_y):
    data, finite, fill = _finite_fill(arr)
    if _HAS_SCIPY and ndimage is not None:
        try:
            rotation_xy = _rotation_matrix(rotation_deg)
            rotation_rc = np.array(
                [
                    [rotation_xy[1, 1], rotation_xy[1, 0]],
                    [rotation_xy[0, 1], rotation_xy[0, 0]],
                ],
                dtype=float,
            )
            inverse_rc = rotation_rc.T
            center_rc = np.array(
                [
                    (max(1, int(data.shape[0])) - 1.0) / 2.0,
                    (max(1, int(data.shape[1])) - 1.0) / 2.0,
                ],
                dtype=float,
            )
            shift_rc = np.array([float(shift_y), float(shift_x)], dtype=float)
            offset = center_rc - inverse_rc @ (center_rc + shift_rc)
            transformed = ndimage.affine_transform(
                data,
                inverse_rc,
                offset=offset,
                order=1,
                mode="constant",
                cval=fill,
            )
            transformed_mask = ndimage.affine_transform(
                finite.astype(float),
                inverse_rc,
                offset=offset,
                order=0,
                mode="constant",
                cval=0.0,
            )
            valid = np.asarray(transformed_mask) > 0.5
            transformed = np.array(transformed, copy=True)
            transformed[~valid] = np.nan
            return transformed, valid
        except Exception:
            pass
    if abs(float(rotation_deg)) > 1e-9:
        raise RuntimeError("Rotation alignment requires scipy.")
    return _translate_integer(data, shift_x, shift_y)


def _match_intensity(reference, moving, mode):
    mode = str(mode or "None").strip().lower()
    valid = np.isfinite(reference) & np.isfinite(moving)
    if not valid.any():
        return np.array(moving, copy=True), 0.0
    ref = np.asarray(reference, dtype=float)[valid]
    mov = np.asarray(moving, dtype=float)[valid]
    if mode == "median match":
        offset = float(np.median(ref) - np.median(mov))
    elif mode == "mean match":
        offset = float(np.mean(ref) - np.mean(mov))
    else:
        offset = 0.0
    out = np.array(moving, copy=True)
    finite = np.isfinite(out)
    out[finite] += offset
    return out, offset


def _alignment_metrics(reference, moving):
    valid = np.isfinite(reference) & np.isfinite(moving)
    overlap = int(np.count_nonzero(valid))
    if overlap <= 1:
        return {
            "overlap": 0,
            "coverage": 0.0,
            "corr": float("nan"),
            "rmse": float("nan"),
            "mean_delta": float("nan"),
        }
    ref = np.asarray(reference, dtype=float)[valid]
    mov = np.asarray(moving, dtype=float)[valid]
    ref_center = ref - np.mean(ref)
    mov_center = mov - np.mean(mov)
    denom = float(np.linalg.norm(ref_center) * np.linalg.norm(mov_center))
    corr = float(np.dot(ref_center, mov_center) / denom) if denom > 1e-12 else float("nan")
    diff = ref - mov
    rmse = float(np.sqrt(np.mean(diff**2)))
    return {
        "overlap": overlap,
        "coverage": float(overlap) / float(reference.size or 1),
        "corr": corr,
        "rmse": rmse,
        "mean_delta": float(np.mean(diff)),
    }


def _alignment_score(reference, moving):
    """Collapse compare metrics into a single score for local rigid-refinement searches."""
    metrics = _alignment_metrics(reference, moving)
    corr = metrics.get("corr")
    coverage = float(metrics.get("coverage", 0.0) or 0.0)
    rmse = metrics.get("rmse")
    if corr is None or not np.isfinite(corr):
        corr = -1.0
    if rmse is None or not np.isfinite(rmse):
        rmse = 1e9
    score = float(corr) + (0.25 * coverage) - (0.05 * float(rmse))
    return score, metrics


def _image_center_xy(shape):
    """Return the pixel-space center used by scipy-style in-place rotations."""
    height = int(shape[0]) if shape else 0
    width = int(shape[1]) if shape else 0
    return np.array([(max(1, width) - 1.0) / 2.0, (max(1, height) - 1.0) / 2.0], dtype=float)


def _rotation_matrix(rotation_deg):
    """Return the point transform that matches scipy.ndimage.rotate(..., axes=(1, 0))."""
    radians = float(np.deg2rad(float(rotation_deg)))
    cos_a = float(np.cos(radians))
    sin_a = float(np.sin(radians))
    return np.array([[cos_a, -sin_a], [sin_a, cos_a]], dtype=float)


def _transform_points(points, shape, rotation_deg=0.0, shift_x=0.0, shift_y=0.0):
    """Apply the compare dialog transform to point pairs in image pixel coordinates."""
    pts = np.asarray(points, dtype=float)
    if pts.size == 0:
        return np.empty((0, 2), dtype=float)
    center = _image_center_xy(shape)
    rotated = (pts - center) @ _rotation_matrix(rotation_deg).T + center
    return rotated + np.array([float(shift_x), float(shift_y)], dtype=float)


def _inverse_transform_points(points, shape, rotation_deg=0.0, shift_x=0.0, shift_y=0.0):
    """Map displayed aligned-B landmark clicks back onto the original B pixel grid."""
    pts = np.asarray(points, dtype=float)
    if pts.size == 0:
        return np.empty((0, 2), dtype=float)
    center = _image_center_xy(shape)
    unshifted = pts - np.array([float(shift_x), float(shift_y)], dtype=float)
    return (unshifted - center) @ _rotation_matrix(rotation_deg) + center


def _fit_translation_from_points(points_a, points_b):
    """Estimate a pure translation from corresponding landmark pairs."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    if pts_a.shape != pts_b.shape or pts_a.ndim != 2 or pts_a.shape[1] != 2 or pts_a.shape[0] < 1:
        raise ValueError("Translation fit needs at least one complete A/B landmark pair.")
    delta = pts_a - pts_b
    shift = np.mean(delta, axis=0)
    residual = delta - shift
    rmse = float(np.sqrt(np.mean(np.sum(residual**2, axis=1)))) if residual.size else 0.0
    return 0.0, float(shift[0]), float(shift[1]), rmse


def _fit_shift_for_rotation(points_a, points_b, moving_shape, rotation_deg):
    """Estimate the post-rotation shift in the same parameterization used by the compare dialog."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    rotated = _transform_points(pts_b, moving_shape, rotation_deg=rotation_deg, shift_x=0.0, shift_y=0.0)
    delta = pts_a - rotated
    shift = np.mean(delta, axis=0)
    transformed = rotated + shift
    residual = pts_a - transformed
    rmse = float(np.sqrt(np.mean(np.sum(residual**2, axis=1)))) if residual.size else 0.0
    return float(shift[0]), float(shift[1]), rmse


def _wrap_angle_deg(angle_deg):
    """Normalize an angle to the [-180, 180) range used by the compare spin box."""
    wrapped = (float(angle_deg) + 180.0) % 360.0 - 180.0
    return 180.0 if wrapped == -180.0 else wrapped


def _pairwise_angle_candidates(points_a, points_b):
    """Derive rigid-rotation guesses from the directions of matched landmark pairs."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    if pts_a.shape != pts_b.shape or pts_a.ndim != 2 or pts_a.shape[0] < 2:
        return []
    candidates = []
    seen = set()
    for i in range(int(pts_a.shape[0]) - 1):
        for j in range(i + 1, int(pts_a.shape[0])):
            vec_a = pts_a[j] - pts_a[i]
            vec_b = pts_b[j] - pts_b[i]
            if np.linalg.norm(vec_a) < 1e-6 or np.linalg.norm(vec_b) < 1e-6:
                continue
            angle_a = float(np.degrees(np.arctan2(vec_a[1], vec_a[0])))
            angle_b = float(np.degrees(np.arctan2(vec_b[1], vec_b[0])))
            candidate = _wrap_angle_deg(angle_a - angle_b)
            key = round(candidate, 6)
            if key in seen:
                continue
            seen.add(key)
            candidates.append(candidate)
    return candidates


def _best_rigid_angle_candidate(points_a, points_b, moving_shape, seed_angle):
    """Pick the lowest-residual coarse angle before local rigid refinement."""
    candidates = [_wrap_angle_deg(seed_angle)]
    candidates.extend(_pairwise_angle_candidates(points_a, points_b))
    # A full coarse sweep keeps the fit robust when the initial seed lands in the
    # wrong basin, which can happen with sparse or symmetric landmark choices.
    candidates.extend(np.linspace(-180.0, 178.0, 180))
    best = None
    seen = set()
    for angle in candidates:
        trial_angle = _wrap_angle_deg(angle)
        key = round(trial_angle, 6)
        if key in seen:
            continue
        seen.add(key)
        shift_x, shift_y, rmse = _fit_shift_for_rotation(points_a, points_b, moving_shape, trial_angle)
        if best is None or rmse < best[3]:
            best = (trial_angle, shift_x, shift_y, rmse)
    return best


def _refine_rigid_angle(points_a, points_b, moving_shape, seed_angle):
    """Refine the rigid-fit angle against the dialog's actual landmark residual model."""
    best = _best_rigid_angle_candidate(points_a, points_b, moving_shape, seed_angle)
    for step, radius in ((1.0, 12.0), (0.2, 2.0), (0.05, 0.4)):
        center = float(seed_angle) if best is None else float(best[0])
        count = max(3, int(round((2.0 * radius) / step)) + 1)
        angles = np.linspace(center - radius, center + radius, count)
        current_best = best
        for angle in angles:
            trial_angle = _wrap_angle_deg(angle)
            shift_x, shift_y, rmse = _fit_shift_for_rotation(points_a, points_b, moving_shape, trial_angle)
            if current_best is None or rmse < current_best[3]:
                current_best = (trial_angle, shift_x, shift_y, rmse)
        best = current_best
    return best


def _fit_rigid_from_points(points_a, points_b, moving_shape):
    """Estimate rotation plus translation from corresponding landmark pairs."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    if pts_a.shape != pts_b.shape or pts_a.ndim != 2 or pts_a.shape[1] != 2 or pts_a.shape[0] < 2:
        raise ValueError("Rigid fit needs at least two complete A/B landmark pairs.")
    centroid_a = np.mean(pts_a, axis=0)
    centroid_b = np.mean(pts_b, axis=0)
    aa = pts_a - centroid_a
    bb = pts_b - centroid_b
    covariance = bb.T @ aa
    u_mat, _singular, vt_mat = np.linalg.svd(covariance)
    rotation = vt_mat.T @ u_mat.T
    if np.linalg.det(rotation) < 0.0:
        vt_mat[-1, :] *= -1.0
        rotation = vt_mat.T @ u_mat.T
    angle = _wrap_angle_deg(np.degrees(np.arctan2(rotation[1, 0], rotation[0, 0])))
    refined = _refine_rigid_angle(pts_a, pts_b, moving_shape, angle)
    if refined is None:
        shift_x, shift_y, rmse = _fit_shift_for_rotation(pts_a, pts_b, moving_shape, angle)
        return angle, shift_x, shift_y, rmse
    return refined


def _is_rigid_mode(mode_text):
    """Return True when the selected compare mode is the rotation+translation variant."""
    return str(mode_text or "").strip().lower().startswith("rigid")


def _finite_minmax(arr, fallback=(0.0, 1.0)):
    """Return a finite plotting range, widening zero-span inputs just enough for imshow."""
    finite = np.asarray(arr, dtype=float)
    finite = finite[np.isfinite(finite)]
    if finite.size == 0:
        return float(fallback[0]), float(fallback[1])
    low = float(np.min(finite))
    high = float(np.max(finite))
    if abs(high - low) <= 1e-12:
        pad = max(abs(low) * 0.01, 1e-6)
        return low - pad, high + pad
    return low, high


def _equalize_hist_display(arr):
    """Histogram-equalize finite pixels for display while preserving NaN-masked regions."""
    data = np.asarray(arr, dtype=float)
    out = np.full(data.shape, np.nan, dtype=float)
    finite = np.isfinite(data)
    if not finite.any():
        return out
    values = np.asarray(data[finite], dtype=float)
    if _HAS_SKIMAGE_EXPOSURE and _sk_exposure is not None:
        try:
            out[finite] = np.asarray(_sk_exposure.equalize_hist(values), dtype=float)
            return out
        except Exception:
            pass
    order = np.argsort(values, kind="mergesort")
    ranks = np.empty(order.shape[0], dtype=float)
    if order.size == 1:
        ranks[order] = 0.5
    else:
        ranks[order] = np.linspace(0.0, 1.0, order.size, dtype=float)
    out[finite] = ranks
    return out


def _apply_display_stretch(arr, mode_text):
    """Apply display-only contrast stretching without modifying the scientific data arrays."""
    mode = str(mode_text or "linear").strip().lower()
    if mode.startswith("hist"):
        return _equalize_hist_display(arr)
    return np.array(arr, copy=True, dtype=float)


def _sanitize_filename_token(text):
    """Collapse arbitrary labels into filesystem-safe filename fragments."""
    raw = str(text or "").strip()
    if not raw:
        return "compare"
    cleaned = "".join(ch if ch.isalnum() or ch in {"-", "_", "."} else "_" for ch in raw)
    cleaned = cleaned.strip("._")
    return cleaned or "compare"


def _axis_to_display_pixel(point_x, point_y, extent, shape, origin="upper"):
    """Map rendered axis coordinates onto pixel coordinates for an imshow extent/origin pair."""
    height = max(1, int(shape[0]))
    width = max(1, int(shape[1]))
    cols = max(width - 1, 1)
    rows = max(height - 1, 1)
    xmin, xmax, ymin, ymax = (float(v) for v in extent)
    span_x = float(xmax - xmin)
    span_y = float(ymax - ymin)
    col = 0.0 if abs(span_x) <= 1e-12 else ((float(point_x) - xmin) / span_x) * float(cols)
    if abs(span_y) <= 1e-12:
        row_use = 0.0
    elif str(origin or "upper").lower() == "upper":
        row_use = ((ymax - float(point_y)) / span_y) * float(rows)
    else:
        row_use = ((float(point_y) - ymin) / span_y) * float(rows)
    row = float(rows) - row_use if str(origin or "upper").lower() == "lower" and rows > 0 else row_use
    return np.array(
        [
            float(np.clip(col, 0.0, max(0.0, float(width - 1)))),
            float(np.clip(row, 0.0, max(0.0, float(height - 1)))),
        ],
        dtype=float,
    )


def _display_pixel_to_axis(point, extent, shape, origin="upper"):
    """Map stored image pixel coordinates back onto displayed axis coordinates."""
    height = max(1, int(shape[0]))
    width = max(1, int(shape[1]))
    cols = max(width - 1, 1)
    rows = max(height - 1, 1)
    xmin, xmax, ymin, ymax = (float(v) for v in extent)
    col = float(point[0])
    row = float(point[1])
    lower_origin = str(origin or "upper").lower() == "lower"
    row_use = float(rows) - row if lower_origin and rows > 0 else row
    x_axis = xmin if cols == 0 else xmin + (col / float(cols)) * (xmax - xmin)
    if rows == 0:
        y_axis = ymax if not lower_origin else ymin
    elif not lower_origin:
        y_axis = ymax - (row_use / float(rows)) * (ymax - ymin)
    else:
        y_axis = ymin + (row_use / float(rows)) * (ymax - ymin)
    return float(x_axis), float(y_axis)


def _sample_line_profile(arr, start_px, end_px, *, order=1):
    """Sample a line profile from a 2-D array with scipy-style interpolation when available."""
    data = np.asarray(arr, dtype=float)
    delta = np.asarray(end_px, dtype=float) - np.asarray(start_px, dtype=float)
    steps = max(2, int(np.ceil(np.hypot(delta[0], delta[1]) * 2.0)) + 1)
    cols = np.linspace(float(start_px[0]), float(end_px[0]), steps, dtype=float)
    rows = np.linspace(float(start_px[1]), float(end_px[1]), steps, dtype=float)
    coords = np.vstack([rows, cols])
    filled, finite, fill = _finite_fill(data)
    if _HAS_SCIPY and ndimage is not None:
        samples = ndimage.map_coordinates(filled, coords, order=order, mode="constant", cval=fill)
        valid = ndimage.map_coordinates(finite.astype(float), coords, order=0, mode="constant", cval=0.0) > 0.5
    else:
        row_idx = np.clip(np.rint(rows).astype(int), 0, max(0, data.shape[0] - 1))
        col_idx = np.clip(np.rint(cols).astype(int), 0, max(0, data.shape[1] - 1))
        samples = filled[row_idx, col_idx]
        valid = finite[row_idx, col_idx]
    samples = np.asarray(samples, dtype=float)
    samples[~valid] = np.nan
    return samples, cols, rows


def _sample_image_value(arr, point_px):
    """Return an interpolated value at a single pixel-space coordinate."""
    values, _cols, _rows = _sample_line_profile(arr, point_px, point_px, order=1)
    return float(values[0]) if values.size else float("nan")


class ImageCompareController:
    """Manage Compare A/B slots and the compare popup lifecycle."""

    def __init__(self, viewer):
        self.viewer = viewer
        self._slot_a = None
        self._slot_b = None
        self._dialog = None

    def menu_state(self):
        """Return lightweight state so canvas menus can enable/label compare actions."""
        return {
            "has_a": self._slot_a is not None,
            "has_b": self._slot_b is not None,
            "label_a": _safe_label(self._slot_a),
            "label_b": _safe_label(self._slot_b),
        }

    def handle_menu_action(self, action, view, canvas=None):
        """Dispatch compare actions coming from a preview canvas context menu."""
        action = str(action or "").strip().lower()
        try:
            if action == "set_a":
                self.set_slot("a", view, canvas=canvas)
            elif action == "set_b":
                self.set_slot("b", view, canvas=canvas)
            elif action == "compare_with_a":
                if self._slot_a is None:
                    self._show_missing_slot("A")
                    return
                self.set_slot("b", view, canvas=canvas, announce=False)
                self.open_compare_dialog()
            elif action == "compare_with_b":
                if self._slot_b is None:
                    self._show_missing_slot("B")
                    return
                self.set_slot("a", view, canvas=canvas, announce=False)
                self.open_compare_dialog()
            elif action == "open_compare":
                self.open_compare_dialog()
            elif action == "swap_compare":
                self.swap_slots()
            elif action == "clear_compare":
                self.clear_slots()
        except ValueError as exc:
            QtWidgets.QMessageBox.information(self.viewer, "Image comparison", str(exc))

    def set_slot(self, slot_name, view, canvas=None, announce=True):
        """Freeze the current view into Compare A or B."""
        snapshot = self._freeze_view(view, canvas=canvas)
        if slot_name == "a":
            self._slot_a = snapshot
        else:
            self._slot_b = snapshot
        if announce:
            self._show_feedback(f"Compare {slot_name.upper()}: {snapshot['label']}")
        self._sync_dialog()

    def swap_slots(self):
        """Swap A and B and refresh an open dialog."""
        if self._slot_a is None and self._slot_b is None:
            return
        self._slot_a, self._slot_b = self._slot_b, self._slot_a
        self._show_feedback("Swapped compare A and B")
        self._sync_dialog()

    def clear_slots(self):
        """Clear transient compare selections without closing the dialog."""
        self._slot_a = None
        self._slot_b = None
        self._show_feedback("Cleared compare A/B selection")
        self._sync_dialog()

    def open_compare_dialog(self):
        """Show or refresh the comparison popup when both slots are populated."""
        if self._slot_a is None or self._slot_b is None:
            self._show_missing_slot("A and B")
            return None
        if self._dialog is not None:
            try:
                if self._dialog.isVisible():
                    self._dialog.set_snapshots(self._slot_a, self._slot_b)
                    self._dialog.raise_()
                    self._dialog.activateWindow()
                    return self._dialog
            except Exception:
                self._dialog = None

        dlg = ImageCompareDialog(self, self.viewer, self._slot_a, self._slot_b)
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.finished.connect(lambda _=None: self._on_dialog_finished(dlg))
        self._dialog = dlg
        self._register_popup(dlg)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        return dlg

    def _sync_dialog(self):
        if self._dialog is None:
            return
        try:
            if self._slot_a is None or self._slot_b is None:
                self._dialog.set_slots_pending(self._slot_a, self._slot_b)
            else:
                self._dialog.set_snapshots(self._slot_a, self._slot_b)
        except Exception:
            pass

    def _freeze_view(self, view, canvas=None):
        """Create a self-contained compare snapshot from a live view dictionary."""
        if not isinstance(view, dict):
            raise ValueError("Comparison requires a valid image view.")
        arr = view.get("arr")
        if arr is None:
            raise ValueError("Selected view does not contain image data.")
        data = np.asarray(arr, dtype=float)
        if data.ndim < 2 or data.size == 0:
            raise ValueError("Selected view is empty.")
        rel_axes = False
        if canvas is not None and hasattr(canvas, "_use_relative_axes"):
            try:
                rel_axes = bool(canvas._use_relative_axes(view))
            except Exception:
                rel_axes = False
        elif view.get("relative_axes"):
            rel_axes = True
        if rel_axes:
            data = np.flipud(data)
        raw_extent = view.get("extent_raw")
        if raw_extent is None:
            raw_extent = view.get("extent")
        extent = raw_extent
        if canvas is not None and hasattr(canvas, "_display_extent_for_view"):
            try:
                extent = canvas._display_extent_for_view(view, raw_extent)
            except Exception:
                extent = raw_extent
        if extent is not None:
            try:
                extent = tuple(float(v) for v in extent)
            except Exception:
                extent = None
        title = str(view.get("title") or "").strip()
        meta = dict(view.get("meta") or {})
        path_text = str(view.get("path") or meta.get("path") or meta.get("file_path") or "").strip()
        axis_unit = str(view.get("axis_unit") or meta.get("axis_unit") or ("px" if extent is None else "nm") or "px")
        color_unit = str(view.get("unit") or view.get("colorbar_label") or "").strip()
        channel_idx = view.get("channel_idx")
        if channel_idx is None:
            channel_idx = meta.get("channel_index")
        filename = Path(path_text).name if path_text else "Image"
        label_parts = [filename]
        if title and title != filename:
            label_parts.append(title)
        label = " | ".join(part for part in label_parts if part)
        if not label:
            label = title or "Image"
        return {
            "arr": np.array(data, copy=True),
            "extent": extent,
            "extent_raw": extent,
            "axis_unit": axis_unit,
            "unit": color_unit,
            "cmap": view.get("cmap", "viridis"),
            "clim": tuple(view.get("clim")) if view.get("clim") else None,
            "path": path_text,
            "channel_idx": channel_idx,
            "title": title or filename,
            "label": label,
        }

    def _show_feedback(self, text):
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self.viewer, self.viewer.rect(), 2200)
        except Exception:
            pass

    def _show_missing_slot(self, slot_text):
        QtWidgets.QMessageBox.information(
            self.viewer,
            "Image comparison",
            f"Set compare {slot_text} first from a preview or popup image.",
        )

    def _register_popup(self, dlg):
        refs = getattr(self.viewer, "_popup_refs", None)
        if refs is not None and dlg not in refs:
            refs.append(dlg)
        controller = getattr(self.viewer, "quick_crop_controller", None)
        if controller is not None:
            try:
                controller.update_popup_actions()
            except Exception:
                pass

    def _on_dialog_finished(self, dlg):
        if getattr(self.viewer, "_popup_refs", None) and dlg in self.viewer._popup_refs:
            self.viewer._popup_refs.remove(dlg)
        controller = getattr(self.viewer, "quick_crop_controller", None)
        if controller is not None:
            try:
                controller.update_popup_actions()
            except Exception:
                pass
        if self._dialog is dlg:
            self._dialog = None


class _LegacyImageCompareDialog(QtWidgets.QDialog):
    """Popup that aligns B onto A and renders the comparison views."""

    def __init__(self, controller, viewer, snapshot_a, snapshot_b):
        super().__init__(viewer)
        self.controller = controller
        self.viewer = viewer
        self._snapshot_a = None
        self._snapshot_b = None
        self._base_a = None
        self._base_b = None
        self._compare_axes = {}
        self._landmark_pairs = []
        self._landmark_overlay_artists = []
        self._last_landmark_rmse = None
        self._update_timer = QtCore.QTimer(self)
        self._update_timer.setSingleShot(True)
        self._update_timer.setInterval(40)
        self._update_timer.timeout.connect(self._rebuild_views)
        self.setWindowTitle("Compare A/B")
        self.setMinimumSize(860, 720)
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
            | QtCore.Qt.WindowSystemMenuHint
        )

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_row = QtWidgets.QHBoxLayout()
        header_row.setSpacing(8)
        header_row.addWidget(QtWidgets.QLabel("A:", self))
        self.label_a = QtWidgets.QLabel("", self)
        self.label_a.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        header_row.addWidget(self.label_a, 1)
        header_row.addWidget(QtWidgets.QLabel("B:", self))
        self.label_b = QtWidgets.QLabel("", self)
        self.label_b.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        header_row.addWidget(self.label_b, 1)
        layout.addLayout(header_row)

        controls = QtWidgets.QHBoxLayout()
        controls.setSpacing(8)
        controls.addWidget(QtWidgets.QLabel("Auto mode:", self))
        self.auto_mode_combo = QtWidgets.QComboBox(self)
        self.auto_mode_combo.addItems(["Translate only", "Rigid (rotate + shift)"])
        self.auto_mode_combo.setToolTip(
            "Translate only keeps the current rotation fixed. Rigid solves rotation plus translation."
        )
        controls.addWidget(self.auto_mode_combo)
        controls.addWidget(QtWidgets.QLabel("Intensity:", self))
        self.intensity_combo = QtWidgets.QComboBox(self)
        self.intensity_combo.addItems(["None", "Median match", "Mean match"])
        self.intensity_combo.setToolTip("Match the baseline of B to A before computing difference maps.")
        controls.addWidget(self.intensity_combo)
        controls.addWidget(QtWidgets.QLabel("Rotation:", self))
        self.rotation_spin = QtWidgets.QDoubleSpinBox(self)
        self.rotation_spin.setRange(-180.0, 180.0)
        self.rotation_spin.setDecimals(2)
        self.rotation_spin.setSingleStep(1.0)
        self.rotation_spin.setSuffix(" deg")
        self.rotation_spin.setKeyboardTracking(False)
        self.rotation_spin.setToolTip("Manual rotation applied to B before comparison.")
        controls.addWidget(self.rotation_spin)
        controls.addWidget(QtWidgets.QLabel("Shift X:", self))
        self.shift_x_spin = QtWidgets.QDoubleSpinBox(self)
        self.shift_x_spin.setRange(-4096.0, 4096.0)
        self.shift_x_spin.setDecimals(2)
        self.shift_x_spin.setSingleStep(0.5)
        self.shift_x_spin.setSuffix(" px")
        self.shift_x_spin.setKeyboardTracking(False)
        self.shift_x_spin.setToolTip("Horizontal shift on the resampled A grid.")
        controls.addWidget(self.shift_x_spin)
        controls.addWidget(QtWidgets.QLabel("Shift Y:", self))
        self.shift_y_spin = QtWidgets.QDoubleSpinBox(self)
        self.shift_y_spin.setRange(-4096.0, 4096.0)
        self.shift_y_spin.setDecimals(2)
        self.shift_y_spin.setSingleStep(0.5)
        self.shift_y_spin.setSuffix(" px")
        self.shift_y_spin.setKeyboardTracking(False)
        self.shift_y_spin.setToolTip("Vertical shift on the resampled A grid.")
        controls.addWidget(self.shift_y_spin)

        self.auto_btn = QtWidgets.QPushButton("Auto fit", self)
        self.auto_btn.clicked.connect(self._auto_fit)
        controls.addWidget(self.auto_btn)
        self.reset_btn = QtWidgets.QPushButton("Reset", self)
        self.reset_btn.clicked.connect(self._reset_transform)
        controls.addWidget(self.reset_btn)
        self.swap_btn = QtWidgets.QPushButton("Swap A/B", self)
        self.swap_btn.clicked.connect(self._swap_slots)
        controls.addWidget(self.swap_btn)
        controls.addStretch(1)
        layout.addLayout(controls)

        landmark_row = QtWidgets.QHBoxLayout()
        landmark_row.setSpacing(8)
        self.pick_landmarks_btn = QtWidgets.QPushButton("Pick landmarks", self)
        self.pick_landmarks_btn.setCheckable(True)
        self.pick_landmarks_btn.setToolTip(
            "Alternate clicks between the A and B panels to define matching reference points."
        )
        self.pick_landmarks_btn.toggled.connect(self._on_pick_landmarks_toggled)
        landmark_row.addWidget(self.pick_landmarks_btn)
        landmark_row.addWidget(QtWidgets.QLabel("Point fit:", self))
        self.landmark_mode_combo = QtWidgets.QComboBox(self)
        self.landmark_mode_combo.addItems(["Rigid (rotate + shift)", "Translate only"])
        self.landmark_mode_combo.setToolTip(
            "Rigid uses two or more pairs to solve rotation plus shift. Translate only keeps rotation fixed."
        )
        self.landmark_mode_combo.currentIndexChanged.connect(lambda *_args: self._update_landmark_status())
        landmark_row.addWidget(self.landmark_mode_combo)
        self.fit_landmarks_btn = QtWidgets.QPushButton("Fit from points", self)
        self.fit_landmarks_btn.clicked.connect(self._fit_from_landmarks)
        landmark_row.addWidget(self.fit_landmarks_btn)
        self.undo_landmarks_btn = QtWidgets.QPushButton("Undo point", self)
        self.undo_landmarks_btn.clicked.connect(self._undo_landmark)
        landmark_row.addWidget(self.undo_landmarks_btn)
        self.clear_landmarks_btn = QtWidgets.QPushButton("Clear points", self)
        self.clear_landmarks_btn.clicked.connect(self._clear_landmarks)
        landmark_row.addWidget(self.clear_landmarks_btn)
        self.landmark_status = QtWidgets.QLabel("", self)
        self.landmark_status.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        landmark_row.addWidget(self.landmark_status, 1)
        layout.addLayout(landmark_row)

        self.metrics_label = QtWidgets.QLabel("", self)
        self.metrics_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        layout.addWidget(self.metrics_label)

        self.canvas = MultiPreviewCanvas(self, figsize=(8.5, 7.0))
        try:
            self.canvas.set_render_suspended(True)
        except Exception:
            pass
        self.canvas.set_compact_size_hints(True)
        self.canvas.set_window_arrange_callback(viewer.on_arrange_popouts)
        self.canvas.set_window_minimize_callback(viewer.on_minimize_popouts)
        self.canvas.set_window_restore_callback(viewer.on_restore_popouts)
        self.canvas.set_window_close_callback(viewer.on_close_popouts)
        self.canvas.show_molecules = False
        self.canvas._show_acquisition_overlay = False
        self.canvas._show_shortcut_hint = False
        self.canvas._show_profile_overlays = False
        self.canvas._show_angle_overlays = False
        self.canvas.scale_bar_enabled = False
        try:
            self.canvas.set_measurement_shortcuts_enabled(False)
        except Exception:
            self.canvas._measurement_shortcuts_enabled = False
        try:
            self.canvas.set_plot_font_family_callback(lambda fam: viewer.set_plot_font_family(fam))
            self.canvas.set_plot_font_family(getattr(viewer, "_plot_font_family", "sans-serif"))
        except Exception:
            pass
        self.canvas._detail_dark = bool(getattr(viewer, "detail_dark_view", False))
        self.canvas._detail_grid = bool(getattr(viewer, "detail_grid_view", False))
        self._landmark_click_cid = self.canvas.mpl_connect("button_press_event", self._on_canvas_click)
        layout.addWidget(self.canvas, 1)

        self.intensity_combo.currentIndexChanged.connect(lambda *_args: self._schedule_update())
        self.rotation_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.shift_x_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.shift_y_spin.valueChanged.connect(lambda *_args: self._schedule_update())

        self.set_snapshots(snapshot_a, snapshot_b)
        self._update_landmark_status()
        try:
            self.canvas.set_render_suspended(False)
        except Exception:
            pass
        self.resize(1020, 820)

    def set_slots_pending(self, snapshot_a, snapshot_b):
        """Keep the dialog open but show that one of the slots is missing."""
        self._snapshot_a = snapshot_a
        self._snapshot_b = snapshot_b
        self._base_a = None
        self._base_b = None
        self._compare_axes = {}
        self._clear_landmark_overlays()
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self.label_a.setText(_safe_label(snapshot_a))
        self.label_b.setText(_safe_label(snapshot_b))
        self.metrics_label.setText("Both compare slots must be populated to render the comparison.")
        self._set_transform_controls(0.0, 0.0, 0.0)
        self.canvas.clear_views()
        self._update_landmark_status()

    def set_snapshots(self, snapshot_a, snapshot_b):
        """Replace A/B sources and rebuild the comparison views."""
        self._snapshot_a = snapshot_a
        self._snapshot_b = snapshot_b
        self.label_a.setText(_safe_label(snapshot_a))
        self.label_b.setText(_safe_label(snapshot_b))
        if snapshot_a is None or snapshot_b is None:
            self.set_slots_pending(snapshot_a, snapshot_b)
            return
        self._base_a = np.array(snapshot_a.get("arr"), copy=True)
        self._base_b = _resample_to_reference_grid(
            snapshot_b.get("arr"),
            snapshot_b.get("extent"),
            self._base_a.shape,
            snapshot_a.get("extent"),
        )
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._clear_landmark_overlays()
        span = float(max(self._base_a.shape[:2]))
        for spin in (self.shift_x_spin, self.shift_y_spin):
            spin.blockSignals(True)
            spin.setRange(-span, span)
            spin.blockSignals(False)
        self._set_transform_controls(0.0, 0.0, 0.0)
        self.setWindowTitle(f"Compare A/B - {_safe_label(snapshot_a)} vs {_safe_label(snapshot_b)}")
        self._update_landmark_status()
        self._schedule_update(immediate=True)

    def _schedule_update(self, immediate=False):
        if immediate:
            self._update_timer.stop()
            self._rebuild_views()
            return
        self._update_timer.start()

    def _set_transform_controls(self, rotation_deg, shift_x, shift_y):
        """Update the transform widgets atomically so new slots do not inherit stale alignment."""
        self.rotation_spin.blockSignals(True)
        self.shift_x_spin.blockSignals(True)
        self.shift_y_spin.blockSignals(True)
        self.rotation_spin.setValue(float(rotation_deg))
        self.shift_x_spin.setValue(float(shift_x))
        self.shift_y_spin.setValue(float(shift_y))
        self.rotation_spin.blockSignals(False)
        self.shift_x_spin.blockSignals(False)
        self.shift_y_spin.blockSignals(False)

    def _reset_transform(self):
        """Reset manual alignment without clearing the current landmark pairs."""
        self._set_transform_controls(0.0, 0.0, 0.0)
        self._schedule_update(immediate=True)

    def _swap_slots(self):
        """Delegate slot swapping to the controller so menus and popup state stay in sync."""
        self.controller.swap_slots()

    def _auto_fit(self):
        """Estimate a transform from the image content instead of user-supplied landmarks."""
        if self._base_a is None or self._base_b is None:
            return
        mode = str(self.auto_mode_combo.currentText() or "Translate only")
        try:
            if _is_rigid_mode(mode):
                rotation, shift_x, shift_y = self._estimate_rigid_transform(self._base_a, self._base_b)
                self.rotation_spin.setValue(rotation)
            else:
                rotation = float(self.rotation_spin.value())
                shift_x, shift_y = self._estimate_translation(self._base_a, self._base_b, rotation)
            self.shift_x_spin.setValue(shift_x)
            self.shift_y_spin.setValue(shift_y)
            self._schedule_update(immediate=True)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to estimate alignment.\n{exc}")

    def _estimate_translation(self, reference, moving, rotation_deg):
        rotated, _ = _transform_with_mask(moving, rotation_deg, 0.0, 0.0)
        shift_x, shift_y, _ = _phase_correlation_shift(reference, rotated)
        return float(shift_x), float(shift_y)

    def _estimate_rigid_transform(self, reference, moving):
        if not _HAS_SCIPY or ndimage is None:
            raise RuntimeError("Rigid auto-alignment requires scipy.")
        ref_ds, mov_ds, scale = self._downsample_pair(reference, moving)
        coarse_angles = np.arange(-180.0, 180.0, 15.0, dtype=float)
        best = None
        for angles in (coarse_angles, None):
            if angles is None and best is not None:
                angles = np.arange(best[0] - 15.0, best[0] + 15.01, 3.0, dtype=float)
            for angle in angles:
                try:
                    shift_x, shift_y = self._estimate_translation(ref_ds, mov_ds, float(angle))
                    aligned, _ = _transform_with_mask(mov_ds, float(angle), shift_x, shift_y)
                    metrics = _alignment_metrics(ref_ds, aligned)
                    score = metrics.get("corr")
                    if score is None or not np.isfinite(score):
                        continue
                    if best is None or score > best[1]:
                        best = (float(angle), float(score), float(shift_x), float(shift_y))
                except Exception:
                    continue
        if best is None:
            raise RuntimeError("No valid rigid alignment candidate was found.")
        angle, _score, shift_x, shift_y = best
        if abs(scale - 1.0) > 1e-9:
            shift_x /= scale
            shift_y /= scale
        return float(angle), float(shift_x), float(shift_y)

    def _refine_rigid_transform_from_seed(self, reference, moving, seed_rotation, seed_shift_x, seed_shift_y):
        """Refine a landmark-seeded rigid transform against the image content on the compare grid."""
        if not _HAS_SCIPY or ndimage is None:
            return float(seed_rotation), float(seed_shift_x), float(seed_shift_y)
        ref_ds, mov_ds, scale = self._downsample_pair(reference, moving)
        seed_rotation = float(seed_rotation)
        seed_shift_x = float(seed_shift_x)
        seed_shift_y = float(seed_shift_y)
        scaled_seed_x = seed_shift_x * scale
        scaled_seed_y = seed_shift_y * scale
        best = None
        try:
            initial_aligned, _ = _transform_with_mask(mov_ds, seed_rotation, scaled_seed_x, scaled_seed_y)
            initial_score, initial_metrics = _alignment_score(ref_ds, initial_aligned)
            best = (
                _wrap_angle_deg(seed_rotation),
                float(initial_score),
                float(scaled_seed_x),
                float(scaled_seed_y),
                initial_metrics,
            )
        except Exception:
            best = None
        for radius, step in ((10.0, 1.0), (2.0, 0.25), (0.4, 0.05)):
            center = float(seed_rotation) if best is None else float(best[0])
            count = max(3, int(round((2.0 * radius) / step)) + 1)
            angles = np.linspace(center - radius, center + radius, count)
            current_best = best
            for angle in angles:
                trial_angle = _wrap_angle_deg(angle)
                try:
                    shift_x, shift_y = self._estimate_translation(ref_ds, mov_ds, trial_angle)
                    aligned, _ = _transform_with_mask(mov_ds, trial_angle, shift_x, shift_y)
                    score, metrics = _alignment_score(ref_ds, aligned)
                except Exception:
                    continue
                if current_best is None or score > current_best[1]:
                    current_best = (
                        float(trial_angle),
                        float(score),
                        float(shift_x),
                        float(shift_y),
                        metrics,
                    )
            best = current_best
        if best is None:
            return float(seed_rotation), float(seed_shift_x), float(seed_shift_y)
        angle, _score, shift_x_ds, shift_y_ds, _metrics = best
        if abs(scale - 1.0) > 1e-9:
            shift_x_ds /= scale
            shift_y_ds /= scale
        return float(angle), float(shift_x_ds), float(shift_y_ds)

    def _downsample_pair(self, reference, moving, max_size=256):
        ref = np.asarray(reference, dtype=float)
        mov = np.asarray(moving, dtype=float)
        longest = float(max(ref.shape[:2]))
        if longest <= float(max_size):
            return ref, mov, 1.0
        scale = float(max_size) / longest
        new_shape = (
            max(48, int(round(ref.shape[0] * scale))),
            max(48, int(round(ref.shape[1] * scale))),
        )
        return _resize_to_shape(ref, new_shape), _resize_to_shape(mov, new_shape), float(new_shape[1]) / float(ref.shape[1])

    def _rebuild_views(self):
        if self._snapshot_a is None or self._snapshot_b is None or self._base_a is None or self._base_b is None:
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        try:
            aligned_b, _mask = _transform_with_mask(self._base_b, rotation, shift_x, shift_y)
        except Exception as exc:
            self.metrics_label.setText(str(exc))
            return
        aligned_b, offset = _match_intensity(self._base_a, aligned_b, self.intensity_combo.currentText())
        diff = np.array(self._base_a, copy=True) - np.array(aligned_b, copy=True)
        abs_diff = np.abs(diff)
        metrics = _alignment_metrics(self._base_a, aligned_b)
        self.metrics_label.setText(
            "Overlap: {coverage:.1%} | Corr: {corr:.4f} | RMSE: {rmse:.4g} | Mean(A-B): {mean_delta:.4g} | "
            "Offset(B): {offset:.4g} | rot={rotation:.2f} deg, dx={shift_x:.2f} px, dy={shift_y:.2f} px".format(
                coverage=metrics.get("coverage", 0.0),
                corr=metrics.get("corr", float("nan")),
                rmse=metrics.get("rmse", float("nan")),
                mean_delta=metrics.get("mean_delta", float("nan")),
                offset=offset,
                rotation=rotation,
                shift_x=shift_x,
                shift_y=shift_y,
            )
        )
        views = self._build_compare_views(aligned_b, diff, abs_diff, offset)
        self.canvas.set_views(views, preserve_profiles=False)
        self._refresh_compare_axes()
        self._draw_landmark_overlays()

    def _build_compare_views(self, aligned_b, diff, abs_diff, offset):
        extent = self._snapshot_a.get("extent")
        unit = str(self._snapshot_a.get("unit") or "").strip()
        axis_unit = self._snapshot_a.get("axis_unit")
        view_a = self._view_payload(
            self._base_a,
            title=f"A: {_safe_label(self._snapshot_a)}",
            cmap=self._snapshot_a.get("cmap", "viridis"),
            clim=self._snapshot_a.get("clim"),
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=unit,
        )
        clim_b = self._snapshot_b.get("clim")
        if clim_b is not None:
            try:
                clim_b = (float(clim_b[0]) + float(offset), float(clim_b[1]) + float(offset))
            except Exception:
                clim_b = None
        view_b = self._view_payload(
            aligned_b,
            title=f"B aligned: {_safe_label(self._snapshot_b)}",
            cmap=self._snapshot_b.get("cmap", "viridis"),
            clim=clim_b,
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=unit,
        )
        finite_diff = diff[np.isfinite(diff)]
        if finite_diff.size:
            diff_span = float(np.nanpercentile(np.abs(finite_diff), 98))
        else:
            diff_span = 1.0
        diff_span = max(diff_span, 1e-9)
        finite_abs = abs_diff[np.isfinite(abs_diff)]
        if finite_abs.size:
            abs_span = float(np.nanpercentile(finite_abs, 98))
        else:
            abs_span = 1.0
        abs_span = max(abs_span, 1e-9)
        view_diff = self._view_payload(
            diff,
            title="A - B",
            cmap="coolwarm",
            clim=(-diff_span, diff_span),
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=f"diff {unit}".strip(),
        )
        view_abs = self._view_payload(
            abs_diff,
            title="|A - B|",
            cmap="inferno",
            clim=(0.0, abs_span),
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=f"abs diff {unit}".strip(),
        )
        return [view_a, view_b, view_diff, view_abs]

    @staticmethod
    def _view_payload(arr, *, title, cmap, clim, extent, axis_unit, colorbar_label):
        payload = {
            "arr": np.array(arr, copy=True),
            "title": title,
            "cmap": cmap,
            "axis_unit": axis_unit,
            "colorbar_label": colorbar_label,
            "unit": colorbar_label,
        }
        if clim is not None:
            payload["clim"] = clim
        if extent is not None:
            payload["extent"] = extent
            payload["extent_raw"] = extent
        return payload

    def _on_pick_landmarks_toggled(self, _checked):
        """Update the instruction label when interactive landmark picking is toggled."""
        if self.pick_landmarks_btn.isChecked() and not self._landmark_pairs:
            try:
                self._set_transform_controls(0.0, 0.0, 0.0)
                self._schedule_update(immediate=True)
            except Exception:
                pass
        self._update_landmark_status()
        if self.pick_landmarks_btn.isChecked():
            self._show_landmark_hint("Click a reference point in A, then the matching point in B.")

    def _fit_from_landmarks(self):
        """Solve the exact landmark-defined alignment from the currently completed point pairs."""
        if self._base_a is None or self._base_b is None:
            return
        points_a, points_b = self._complete_landmark_arrays()
        mode = str(self.landmark_mode_combo.currentText() or "Rigid (rotate + shift)")
        rigid_hint = None
        try:
            if _is_rigid_mode(mode):
                rotation, shift_x, shift_y, rmse = _fit_rigid_from_points(points_a, points_b, self._base_b.shape)
            else:
                rotation, shift_x, shift_y, rmse = _fit_translation_from_points(points_a, points_b)
                if int(points_a.shape[0]) >= 2:
                    try:
                        rigid_rotation, rigid_dx, rigid_dy, rigid_rmse = _fit_rigid_from_points(
                            points_a,
                            points_b,
                            self._base_b.shape,
                        )
                        if np.isfinite(rigid_rmse) and rigid_rmse + 1e-6 < (0.75 * max(rmse, 1e-6)):
                            rigid_hint = (
                                "Translate-only fit keeps rotation fixed. "
                                f"Rigid fit would reduce point RMSE to {rigid_rmse:.2f} px "
                                f"(rot={rigid_rotation:.2f} deg, dx={rigid_dx:.2f} px, dy={rigid_dy:.2f} px)."
                            )
                    except Exception:
                        rigid_hint = None
            self._set_transform_controls(rotation, shift_x, shift_y)
            self._last_landmark_rmse = float(rmse)
            self._update_landmark_status()
            self._schedule_update(immediate=True)
            if rigid_hint:
                self._show_landmark_hint(rigid_hint)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to fit the selected landmarks.\n{exc}")

    def _undo_landmark(self):
        """Remove the last picked point so a mis-click can be corrected quickly."""
        if not self._landmark_pairs:
            return
        last = self._landmark_pairs[-1]
        if last.get("b") is not None:
            last["b"] = None
        else:
            self._landmark_pairs.pop()
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _clear_landmarks(self):
        """Drop all manually picked points while leaving the current transform untouched."""
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._clear_landmark_overlays()
        self._update_landmark_status()

    def _update_landmark_status(self):
        """Refresh the compact status text and enablement for landmark-related controls."""
        complete_pairs = sum(1 for pair in self._landmark_pairs if pair.get("a") is not None and pair.get("b") is not None)
        total_pairs = len(self._landmark_pairs)
        expected = self._expected_landmark_target()
        next_index = total_pairs + 1 if expected == "a" else total_pairs
        mode_text = str(self.landmark_mode_combo.currentText() or "Rigid (rotate + shift)")
        required_pairs = 2 if _is_rigid_mode(mode_text) else 1
        if self._base_a is None or self._base_b is None:
            text = "Landmarks unavailable until both compare slots are set."
        elif self.pick_landmarks_btn.isChecked():
            text = f"{complete_pairs} pair(s) stored. Next click: {expected.upper()}{max(1, next_index)}."
        elif complete_pairs:
            text = f"{complete_pairs} pair(s) stored."
        else:
            text = "Toggle 'Pick landmarks' to add matching A/B reference points."
        if self._base_a is not None and complete_pairs < required_pairs:
            fit_kind = "rigid" if _is_rigid_mode(mode_text) else "translate-only"
            text += f" {required_pairs} complete pair(s) required for {fit_kind} fit."
        if self._last_landmark_rmse is not None and np.isfinite(self._last_landmark_rmse):
            text += f" Fit RMSE: {self._last_landmark_rmse:.2f} px."
        self.landmark_status.setText(text)
        self.fit_landmarks_btn.setEnabled(
            complete_pairs >= required_pairs and self._base_a is not None and self._base_b is not None
        )
        self.undo_landmarks_btn.setEnabled(bool(self._landmark_pairs))
        self.clear_landmarks_btn.setEnabled(bool(self._landmark_pairs))

    def _expected_landmark_target(self):
        """Return which panel should receive the next click in the alternating A/B workflow."""
        if not self._landmark_pairs or self._landmark_pairs[-1].get("b") is not None:
            return "a"
        return "b"

    def _complete_landmark_arrays(self):
        """Return completed landmark pairs as two Nx2 float arrays."""
        points_a = []
        points_b = []
        for pair in self._landmark_pairs:
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is None or point_b is None:
                continue
            points_a.append(point_a)
            points_b.append(point_b)
        if not points_a:
            raise ValueError("Pick at least one complete A/B landmark pair first.")
        return np.asarray(points_a, dtype=float), np.asarray(points_b, dtype=float)

    def _refresh_compare_axes(self):
        """Cache the first two compare axes so clicks and overlays only target A and B."""
        axes = list(getattr(self.canvas, "_ax_view_map", {}).keys())
        self._compare_axes = {
            "a": axes[0] if len(axes) > 0 else None,
            "b": axes[1] if len(axes) > 1 else None,
        }

    def _axis_role(self, ax):
        """Map a matplotlib axis back to the A or B compare panel."""
        if not self._compare_axes:
            self._refresh_compare_axes()
        for role, target_ax in self._compare_axes.items():
            if ax is target_ax:
                return role
        return None

    def _on_canvas_click(self, event):
        """Capture landmark clicks on the A/B panels without touching the diff panels."""
        if not self.pick_landmarks_btn.isChecked() or self._base_a is None or self._base_b is None:
            return
        if getattr(event, "button", None) != 1 or event.xdata is None or event.ydata is None:
            return
        role = self._axis_role(getattr(event, "inaxes", None))
        if role not in {"a", "b"}:
            return
        expected = self._expected_landmark_target()
        if role != expected:
            self._show_landmark_hint(f"Click the next reference point in {expected.upper()} first.")
            return
        try:
            point = self._event_to_landmark_point(role, event)
        except ValueError as exc:
            self._show_landmark_hint(str(exc))
            return
        if role == "a":
            self._landmark_pairs.append({"a": point, "b": None})
        else:
            self._landmark_pairs[-1]["b"] = point
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _event_to_landmark_point(self, role, event):
        """Convert a canvas click into stored A-grid or raw-B pixel coordinates."""
        ax = getattr(event, "inaxes", None)
        view = getattr(self.canvas, "_ax_view_map", {}).get(ax)
        if view is None:
            raise ValueError("The clicked panel does not contain a comparison image.")
        arr = np.asarray(view.get("arr")) if view.get("arr") is not None else None
        if arr is None or arr.ndim < 2 or arr.size == 0:
            raise ValueError("The clicked comparison image is empty.")
        point = self._axis_to_display_pixel(ax, view, event.xdata, event.ydata)
        if role == "a":
            if not self._point_in_bounds(point, self._base_a.shape):
                raise ValueError("A landmark click must stay inside the visible image.")
            return tuple(point)
        base_point = _inverse_transform_points(
            [point],
            self._base_b.shape,
            rotation_deg=float(self.rotation_spin.value()),
            shift_x=float(self.shift_x_spin.value()),
            shift_y=float(self.shift_y_spin.value()),
        )[0]
        if not self._point_in_bounds(base_point, self._base_b.shape):
            raise ValueError("The clicked B point maps outside the source image. Reset or refine the transform first.")
        return tuple(base_point)

    def _point_in_bounds(self, point, shape):
        """Return True when a pixel-space landmark lies inside the target image."""
        if point is None or shape is None or len(shape) < 2:
            return False
        x_val = float(point[0])
        y_val = float(point[1])
        height = int(shape[0])
        width = int(shape[1])
        return 0.0 <= x_val <= float(max(0, width - 1)) and 0.0 <= y_val <= float(max(0, height - 1))


def _icd_on_profile_mode_toggled(self, checked):
    """Switch into line-profile drawing mode without interfering with landmark picking."""
    if checked and self.pick_landmarks_btn.isChecked():
        self.pick_landmarks_btn.blockSignals(True)
        self.pick_landmarks_btn.setChecked(False)
        self.pick_landmarks_btn.blockSignals(False)
        self._update_landmark_status()
    cursor = QtCore.Qt.CrossCursor if checked else QtCore.Qt.ArrowCursor
    self.canvas.setCursor(QtGui.QCursor(cursor))


def _icd_clear_profile_overlay(self):
    """Remove the current ROI line overlay from all image panels."""
    for artist in self._profile_overlay_artists:
        try:
            artist.remove()
        except Exception:
            pass
    self._profile_overlay_artists = []


def _icd_draw_profile_overlay(self):
    """Show the active line ROI across all four image panels simultaneously."""
    self._clear_profile_overlay()
    if self._profile_line is None or not self._image_axes:
        self.canvas.draw_idle()
        return
    start = np.asarray(self._profile_line["start"], dtype=float)
    end = np.asarray(self._profile_line["end"], dtype=float)
    for ax in self._image_axes.values():
        line, = ax.plot(
            [start[0], end[0]],
            [start[1], end[1]],
            color="#ffd24d",
            linewidth=1.5,
            zorder=28,
        )
        start_marker = ax.scatter([start[0]], [start[1]], s=20, c="#ffd24d", edgecolors="black", zorder=29)
        end_marker = ax.scatter([end[0]], [end[1]], s=20, c="#ffd24d", edgecolors="black", zorder=29)
        self._profile_overlay_artists.extend([line, start_marker, end_marker])
    self.canvas.draw_idle()


def _icd_set_profile_line(self, start_axis, end_axis):
    """Store a shared axis-space profile ROI and refresh its overlays and line plot."""
    self._profile_line = {
        "start": (float(start_axis[0]), float(start_axis[1])),
        "end": (float(end_axis[0]), float(end_axis[1])),
    }
    self._draw_profile_overlay()
    self._plot_profile_panel()
    self.canvas.draw_idle()


def _icd_ensure_crosshair_artists(self):
    """Create hidden crosshair artists lazily so hover updates only move line objects."""
    if self._crosshair_artists or not self._image_axes:
        return
    for role, ax in self._image_axes.items():
        vline = ax.axvline(0.0, color="white", linewidth=0.8, alpha=0.8, visible=False, zorder=24)
        hline = ax.axhline(0.0, color="white", linewidth=0.8, alpha=0.8, visible=False, zorder=24)
        self._crosshair_artists[role] = (vline, hline)


def _icd_update_crosshair(self, x_val, y_val, *, redraw=True):
    """Link a hover crosshair across all image panels and update the status readout."""
    if not self._render_state or not self._image_axes:
        return
    point = self._axis_to_pixel(x_val, y_val)
    if not self._point_in_bounds(point, self._render_state.get("shape")):
        self._clear_crosshair(redraw=redraw)
        return
    self._ensure_crosshair_artists()
    for vline, hline in self._crosshair_artists.values():
        vline.set_xdata([float(x_val), float(x_val)])
        hline.set_ydata([float(y_val), float(y_val)])
        vline.set_visible(True)
        hline.set_visible(True)
    self._hover_axis_point = (float(x_val), float(y_val))
    self._update_hover_label(point, float(x_val), float(y_val))
    if redraw:
        self.canvas.draw_idle()


def _icd_clear_crosshair(self, *, redraw=True):
    """Hide the linked crosshair and clear the hover readout."""
    for vline, hline in self._crosshair_artists.values():
        vline.set_visible(False)
        hline.set_visible(False)
    self._hover_axis_point = None
    self.hover_label.setText("")
    if redraw:
        self.canvas.draw_idle()


def _icd_update_hover_label(self, point_px, x_val, y_val):
    """Format the current hover location and interpolated A/B/diff values for the status line."""
    axis_unit = str(self._render_state.get("axis_unit") or "nm")
    value_unit = str(self._render_state.get("value_unit") or "nm")
    a_val = _sample_image_value(self._render_state.get("a_raw"), point_px)
    b_val = _sample_image_value(self._render_state.get("b_raw"), point_px)
    diff_val = _sample_image_value(self._render_state.get("diff"), point_px)

    def _fmt(value):
        return "--" if value is None or not np.isfinite(value) else f"{float(value):.4g}"

    self.hover_label.setText(
        f"x: {x_val:.4g} {axis_unit} | y: {y_val:.4g} {axis_unit} | "
        f"A: {_fmt(a_val)} {value_unit} | B: {_fmt(b_val)} {value_unit} | A-B: {_fmt(diff_val)} {value_unit}"
    )


def _icd_on_canvas_press(self, event):
    """Handle landmark picking and profile-line starts on the compare figure."""
    if getattr(event, "button", None) != 1 or event.xdata is None or event.ydata is None:
        return
    role = self._axis_role(getattr(event, "inaxes", None))
    if role not in {"a", "b", "diff", "abs"}:
        return
    if self.profile_mode_btn.isChecked():
        self._profile_drag_active = True
        self._set_profile_line((event.xdata, event.ydata), (event.xdata, event.ydata))
        return
    if not self.pick_landmarks_btn.isChecked() or role not in {"a", "b"}:
        return
    expected = self._expected_landmark_target()
    if role != expected:
        self._show_landmark_hint(f"Click the next reference point in {expected.upper()} first.")
        return
    try:
        point = self._event_to_landmark_point(role, event.xdata, event.ydata)
    except ValueError as exc:
        self._show_landmark_hint(str(exc))
        return
    if role == "a":
        self._landmark_pairs.append({"a": point, "b": None})
    else:
        self._landmark_pairs[-1]["b"] = point
    self._last_landmark_rmse = None
    self._update_landmark_status()
    self._draw_landmark_overlays()


def _icd_on_canvas_motion(self, event):
    """Update the linked crosshair and live profile while the user moves across image panels."""
    role = self._axis_role(getattr(event, "inaxes", None))
    if self._profile_drag_active:
        if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
            self._set_profile_line(self._profile_line["start"], (event.xdata, event.ydata))
            self._update_crosshair(event.xdata, event.ydata, redraw=False)
            self.canvas.draw_idle()
        return
    if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
        self._update_crosshair(event.xdata, event.ydata)
    else:
        self._clear_crosshair()


def _icd_on_canvas_release(self, event):
    """Finish a profile drag without disturbing normal hover crosshair updates."""
    if not self._profile_drag_active:
        return
    self._profile_drag_active = False
    role = self._axis_role(getattr(event, "inaxes", None))
    if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
        self._set_profile_line(self._profile_line["start"], (event.xdata, event.ydata))
        self._update_crosshair(event.xdata, event.ydata)
    else:
        self.canvas.draw_idle()


def _icd_on_canvas_leave(self, _event):
    """Clear the linked crosshair when the pointer leaves the compare figure."""
    if not self._profile_drag_active:
        self._clear_crosshair()


def _icd_show_landmark_hint(self, text):
    """Show lightweight point-picking feedback without interrupting the dialog workflow."""
    try:
        QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self, self.rect(), 2200)
    except Exception:
        pass


def _icd_default_export_stem(self):
    """Build a deterministic export stem from the current compare filenames plus a timestamp."""
    source_a = Path(str((self._snapshot_a or {}).get("path") or _safe_label(self._snapshot_a)))
    source_b = Path(str((self._snapshot_b or {}).get("path") or _safe_label(self._snapshot_b)))
    stem_a = _sanitize_filename_token(source_a.stem or source_a.name or "A")
    stem_b = _sanitize_filename_token(source_b.stem or source_b.name or "B")
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{stem_a}__vs__{stem_b}_{timestamp}"


def _icd_export_comparison(self):
    """Save a 300 DPI four-panel PNG and a CSV of the current A-B difference image."""
    state = dict(self._render_state or {})
    if not state:
        QtWidgets.QMessageBox.information(self, "Export", "No comparison data is available to export.")
        return
    default_dir = None
    path_a = str((self._snapshot_a or {}).get("path") or "").strip()
    if path_a:
        try:
            default_dir = Path(path_a).resolve().parent
        except Exception:
            default_dir = None
    if default_dir is None:
        viewer_dir = getattr(self.viewer, "last_dir", None)
        if viewer_dir:
            try:
                default_dir = Path(str(viewer_dir))
            except Exception:
                default_dir = None
    if default_dir is None:
        default_dir = Path.cwd()
    stem = self._default_export_stem()
    default_path = default_dir / f"{stem}.png"
    png_path, _selected = QtWidgets.QFileDialog.getSaveFileName(
        self,
        "Export comparison figure",
        str(default_path),
        "PNG Files (*.png)",
    )
    if not png_path:
        return
    png_path = Path(png_path)
    if png_path.suffix.lower() != ".png":
        png_path = png_path.with_suffix(".png")
    csv_path = png_path.with_name(f"{png_path.stem}_AminusB.csv")
    try:
        self._export_png(png_path)
        self._export_diff_csv(csv_path)
        try:
            self.viewer.last_dir = str(png_path.parent)
        except Exception:
            pass
        QtWidgets.QMessageBox.information(
            self,
            "Export",
            f"Saved\n{png_path}\nand\n{csv_path}",
        )
    except Exception as exc:
        QtWidgets.QMessageBox.warning(self, "Export", f"Unable to export the comparison.\n{exc}")


def _icd_export_png(self, path):
    """Create a standalone 2x2 export figure with current display ranges and colorbars."""
    state = dict(self._render_state or {})
    fig = Figure(figsize=(9.0, 8.2))
    axes = fig.subplots(2, 2)
    panels = [
        (axes[0][0], state["a_display"], state["title_a"], state["topo_cmap"], state["a_clim"], state["display_unit"]),
        (axes[0][1], state["b_display"], state["title_b"], state["topo_cmap"], state["b_clim"], state["display_unit"]),
        (axes[1][0], state["diff"], "A - B", "RdBu_r", state["diff_clim"], f"A-B ({state['value_unit']})"),
        (axes[1][1], state["abs_diff"], "|A - B|", "magma", state["abs_clim"], f"|A-B| ({state['value_unit']})"),
    ]
    for ax, arr, title, cmap, clim, cbar_label in panels:
        image = ax.imshow(
            np.asarray(arr, dtype=float),
            cmap=str(cmap),
            vmin=float(clim[0]),
            vmax=float(clim[1]),
            origin=self.IMAGE_ORIGIN,
            extent=state["extent"],
            interpolation="nearest",
            aspect="equal",
        )
        ax.set_title(str(title), fontsize=10)
        cbar = fig.colorbar(image, ax=ax, fraction=0.046, pad=0.03)
        cbar.set_label(str(cbar_label))
        ax.tick_params(labelsize=8)
    fig.tight_layout()
    fig.savefig(str(path), dpi=300, bbox_inches="tight")


def _icd_export_diff_csv(self, path):
    """Save the current A-B image as a CSV table with axis coordinates expressed in the current unit."""
    state = dict(self._render_state or {})
    diff = np.asarray(state.get("diff"), dtype=float)
    extent = state.get("extent")
    shape = state.get("shape")
    axis_unit = str(state.get("axis_unit") or "nm")
    if diff.ndim != 2 or extent is None or not shape:
        raise ValueError("The current difference image is unavailable.")
    x_coords = _axis_coords_for_size(extent[0], extent[1], shape[1])
    y_coords = _axis_coords_for_size(extent[2], extent[3], shape[0])
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow([f"y_{axis_unit}/x_{axis_unit}", *[f"{val:.9g}" for val in x_coords]])
        for y_val, row in zip(y_coords, diff):
            writer.writerow([f"{float(y_val):.9g}", *[f"{float(val):.9g}" if np.isfinite(val) else "" for val in row]])


_LegacyImageCompareDialog._on_profile_mode_toggled = _icd_on_profile_mode_toggled
_LegacyImageCompareDialog._clear_profile_overlay = _icd_clear_profile_overlay
_LegacyImageCompareDialog._draw_profile_overlay = _icd_draw_profile_overlay
_LegacyImageCompareDialog._set_profile_line = _icd_set_profile_line
_LegacyImageCompareDialog._ensure_crosshair_artists = _icd_ensure_crosshair_artists
_LegacyImageCompareDialog._update_crosshair = _icd_update_crosshair
_LegacyImageCompareDialog._clear_crosshair = _icd_clear_crosshair
_LegacyImageCompareDialog._update_hover_label = _icd_update_hover_label
_LegacyImageCompareDialog._on_canvas_press = _icd_on_canvas_press
_LegacyImageCompareDialog._on_canvas_motion = _icd_on_canvas_motion
_LegacyImageCompareDialog._on_canvas_release = _icd_on_canvas_release
_LegacyImageCompareDialog._on_canvas_leave = _icd_on_canvas_leave
_LegacyImageCompareDialog._show_landmark_hint = _icd_show_landmark_hint
_LegacyImageCompareDialog._default_export_stem = _icd_default_export_stem
_LegacyImageCompareDialog._export_comparison = _icd_export_comparison
_LegacyImageCompareDialog._export_png = _icd_export_png
_LegacyImageCompareDialog._export_diff_csv = _icd_export_diff_csv

if False:

    def _on_profile_mode_toggled(self, checked):
        """Switch into line-profile drawing mode without interfering with landmark picking."""
        if checked and self.pick_landmarks_btn.isChecked():
            self.pick_landmarks_btn.blockSignals(True)
            self.pick_landmarks_btn.setChecked(False)
            self.pick_landmarks_btn.blockSignals(False)
            self._update_landmark_status()
        cursor = QtCore.Qt.CrossCursor if checked else QtCore.Qt.ArrowCursor
        self.canvas.setCursor(QtGui.QCursor(cursor))

    def _clear_profile_overlay(self):
        """Remove the current ROI line overlay from all image panels."""
        for artist in self._profile_overlay_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._profile_overlay_artists = []

    def _draw_profile_overlay(self):
        """Show the active line ROI across all four image panels simultaneously."""
        self._clear_profile_overlay()
        if self._profile_line is None or not self._image_axes:
            self.canvas.draw_idle()
            return
        start = np.asarray(self._profile_line["start"], dtype=float)
        end = np.asarray(self._profile_line["end"], dtype=float)
        for ax in self._image_axes.values():
            line, = ax.plot(
                [start[0], end[0]],
                [start[1], end[1]],
                color="#ffd24d",
                linewidth=1.5,
                zorder=28,
            )
            start_marker = ax.scatter([start[0]], [start[1]], s=20, c="#ffd24d", edgecolors="black", zorder=29)
            end_marker = ax.scatter([end[0]], [end[1]], s=20, c="#ffd24d", edgecolors="black", zorder=29)
            self._profile_overlay_artists.extend([line, start_marker, end_marker])
        self.canvas.draw_idle()

    def _set_profile_line(self, start_axis, end_axis):
        """Store a shared axis-space profile ROI and refresh its overlays and line plot."""
        self._profile_line = {
            "start": (float(start_axis[0]), float(start_axis[1])),
            "end": (float(end_axis[0]), float(end_axis[1])),
        }
        self._draw_profile_overlay()
        self._plot_profile_panel()
        self.canvas.draw_idle()

    def _ensure_crosshair_artists(self):
        """Create hidden crosshair artists lazily so hover updates only move line objects."""
        if self._crosshair_artists or not self._image_axes:
            return
        for role, ax in self._image_axes.items():
            vline = ax.axvline(0.0, color="white", linewidth=0.8, alpha=0.8, visible=False, zorder=24)
            hline = ax.axhline(0.0, color="white", linewidth=0.8, alpha=0.8, visible=False, zorder=24)
            self._crosshair_artists[role] = (vline, hline)

    def _update_crosshair(self, x_val, y_val, *, redraw=True):
        """Link a hover crosshair across all image panels and update the status readout."""
        if not self._render_state or not self._image_axes:
            return
        point = self._axis_to_pixel(x_val, y_val)
        if not self._point_in_bounds(point, self._render_state.get("shape")):
            self._clear_crosshair(redraw=redraw)
            return
        self._ensure_crosshair_artists()
        for vline, hline in self._crosshair_artists.values():
            vline.set_xdata([float(x_val), float(x_val)])
            hline.set_ydata([float(y_val), float(y_val)])
            vline.set_visible(True)
            hline.set_visible(True)
        self._hover_axis_point = (float(x_val), float(y_val))
        self._update_hover_label(point, float(x_val), float(y_val))
        if redraw:
            self.canvas.draw_idle()

    def _clear_crosshair(self, *, redraw=True):
        """Hide the linked crosshair and clear the hover readout."""
        for vline, hline in self._crosshair_artists.values():
            vline.set_visible(False)
            hline.set_visible(False)
        self._hover_axis_point = None
        self.hover_label.setText("")
        if redraw:
            self.canvas.draw_idle()

    def _update_hover_label(self, point_px, x_val, y_val):
        """Format the current hover location and interpolated A/B/diff values for the status line."""
        axis_unit = str(self._render_state.get("axis_unit") or "nm")
        value_unit = str(self._render_state.get("value_unit") or "nm")
        a_val = _sample_image_value(self._render_state.get("a_raw"), point_px)
        b_val = _sample_image_value(self._render_state.get("b_raw"), point_px)
        diff_val = _sample_image_value(self._render_state.get("diff"), point_px)

        def _fmt(value):
            return "--" if value is None or not np.isfinite(value) else f"{float(value):.4g}"

        self.hover_label.setText(
            f"x: {x_val:.4g} {axis_unit} | y: {y_val:.4g} {axis_unit} | "
            f"A: {_fmt(a_val)} {value_unit} | B: {_fmt(b_val)} {value_unit} | A-B: {_fmt(diff_val)} {value_unit}"
        )

    def _on_canvas_press(self, event):
        """Handle landmark picking and profile-line starts on the compare figure."""
        if getattr(event, "button", None) != 1 or event.xdata is None or event.ydata is None:
            return
        role = self._axis_role(getattr(event, "inaxes", None))
        if role not in {"a", "b", "diff", "abs"}:
            return
        if self.profile_mode_btn.isChecked():
            self._profile_drag_active = True
            self._set_profile_line((event.xdata, event.ydata), (event.xdata, event.ydata))
            return
        if not self.pick_landmarks_btn.isChecked() or role not in {"a", "b"}:
            return
        expected = self._expected_landmark_target()
        if role != expected:
            self._show_landmark_hint(f"Click the next reference point in {expected.upper()} first.")
            return
        try:
            point = self._event_to_landmark_point(role, event.xdata, event.ydata)
        except ValueError as exc:
            self._show_landmark_hint(str(exc))
            return
        if role == "a":
            self._landmark_pairs.append({"a": point, "b": None})
        else:
            self._landmark_pairs[-1]["b"] = point
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _on_canvas_motion(self, event):
        """Update the linked crosshair and live profile while the user moves across image panels."""
        role = self._axis_role(getattr(event, "inaxes", None))
        if self._profile_drag_active:
            if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
                self._set_profile_line(self._profile_line["start"], (event.xdata, event.ydata))
                self._update_crosshair(event.xdata, event.ydata, redraw=False)
                self.canvas.draw_idle()
            return
        if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
            self._update_crosshair(event.xdata, event.ydata)
        else:
            self._clear_crosshair()

    def _on_canvas_release(self, event):
        """Finish a profile drag without disturbing normal hover crosshair updates."""
        if not self._profile_drag_active:
            return
        self._profile_drag_active = False
        role = self._axis_role(getattr(event, "inaxes", None))
        if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
            self._set_profile_line(self._profile_line["start"], (event.xdata, event.ydata))
            self._update_crosshair(event.xdata, event.ydata)
        else:
            self.canvas.draw_idle()

    def _on_canvas_leave(self, _event):
        """Clear the linked crosshair when the pointer leaves the compare figure."""
        if not self._profile_drag_active:
            self._clear_crosshair()

    def _show_landmark_hint(self, text):
        """Show lightweight point-picking feedback without interrupting the dialog workflow."""
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self, self.rect(), 2200)
        except Exception:
            pass

    def _default_export_stem(self):
        """Build a deterministic export stem from the current compare filenames plus a timestamp."""
        source_a = Path(str((self._snapshot_a or {}).get("path") or _safe_label(self._snapshot_a)))
        source_b = Path(str((self._snapshot_b or {}).get("path") or _safe_label(self._snapshot_b)))
        stem_a = _sanitize_filename_token(source_a.stem or source_a.name or "A")
        stem_b = _sanitize_filename_token(source_b.stem or source_b.name or "B")
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        return f"{stem_a}__vs__{stem_b}_{timestamp}"

    def _export_comparison(self):
        """Save a 300 DPI four-panel PNG and a CSV of the current A-B difference image."""
        state = dict(self._render_state or {})
        if not state:
            QtWidgets.QMessageBox.information(self, "Export", "No comparison data is available to export.")
            return
        default_dir = None
        path_a = str((self._snapshot_a or {}).get("path") or "").strip()
        if path_a:
            try:
                default_dir = Path(path_a).resolve().parent
            except Exception:
                default_dir = None
        if default_dir is None:
            viewer_dir = getattr(self.viewer, "last_dir", None)
            if viewer_dir:
                try:
                    default_dir = Path(str(viewer_dir))
                except Exception:
                    default_dir = None
        if default_dir is None:
            default_dir = Path.cwd()
        stem = self._default_export_stem()
        default_path = default_dir / f"{stem}.png"
        png_path, _selected = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export comparison figure",
            str(default_path),
            "PNG Files (*.png)",
        )
        if not png_path:
            return
        png_path = Path(png_path)
        if png_path.suffix.lower() != ".png":
            png_path = png_path.with_suffix(".png")
        csv_path = png_path.with_name(f"{png_path.stem}_AminusB.csv")
        try:
            self._export_png(png_path)
            self._export_diff_csv(csv_path)
            try:
                self.viewer.last_dir = str(png_path.parent)
            except Exception:
                pass
            QtWidgets.QMessageBox.information(
                self,
                "Export",
                f"Saved\n{png_path}\nand\n{csv_path}",
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Export", f"Unable to export the comparison.\n{exc}")

    def _export_png(self, path):
        """Create a standalone 2x2 export figure with current display ranges and colorbars."""
        state = dict(self._render_state or {})
        fig = Figure(figsize=(9.0, 8.2))
        axes = fig.subplots(2, 2)
        panels = [
            (axes[0][0], state["a_display"], state["title_a"], state["topo_cmap"], state["a_clim"], state["display_unit"]),
            (axes[0][1], state["b_display"], state["title_b"], state["topo_cmap"], state["b_clim"], state["display_unit"]),
            (axes[1][0], state["diff"], "A - B", "RdBu_r", state["diff_clim"], f"A-B ({state['value_unit']})"),
            (axes[1][1], state["abs_diff"], "|A - B|", "magma", state["abs_clim"], f"|A-B| ({state['value_unit']})"),
        ]
        for ax, arr, title, cmap, clim, cbar_label in panels:
            image = ax.imshow(
                np.asarray(arr, dtype=float),
                cmap=str(cmap),
                vmin=float(clim[0]),
                vmax=float(clim[1]),
                origin=self.IMAGE_ORIGIN,
                extent=state["extent"],
                interpolation="nearest",
                aspect="equal",
            )
            ax.set_title(str(title), fontsize=10)
            cbar = fig.colorbar(image, ax=ax, fraction=0.046, pad=0.03)
            cbar.set_label(str(cbar_label))
            ax.tick_params(labelsize=8)
        fig.tight_layout()
        fig.savefig(str(path), dpi=300, bbox_inches="tight")

    def _export_diff_csv(self, path):
        """Save the current A-B image as a CSV table with axis coordinates expressed in the current unit."""
        state = dict(self._render_state or {})
        diff = np.asarray(state.get("diff"), dtype=float)
        extent = state.get("extent")
        shape = state.get("shape")
        axis_unit = str(state.get("axis_unit") or "nm")
        if diff.ndim != 2 or extent is None or not shape:
            raise ValueError("The current difference image is unavailable.")
        x_coords = _axis_coords_for_size(extent[0], extent[1], shape[1])
        y_coords = _axis_coords_for_size(extent[2], extent[3], shape[0])
        with open(path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow([f"y_{axis_unit}/x_{axis_unit}", *[f"{val:.9g}" for val in x_coords]])
            for y_val, row in zip(y_coords, diff):
                writer.writerow([f"{float(y_val):.9g}", *[f"{float(val):.9g}" if np.isfinite(val) else "" for val in row]])

    def _on_pick_landmarks_toggled(self, _checked):
        """Toggle alternating A/B point picking and keep it separate from profile drawing."""
        if self.pick_landmarks_btn.isChecked():
            if self.profile_mode_btn.isChecked():
                self.profile_mode_btn.blockSignals(True)
                self.profile_mode_btn.setChecked(False)
                self.profile_mode_btn.blockSignals(False)
            if not self._landmark_pairs:
                self._set_transform_controls(0.0, 0.0, 0.0)
                self._schedule_update(immediate=True)
            self._show_landmark_hint("Click a reference point in A, then the matching point in B.")
        self._update_landmark_status()

    def _fit_from_landmarks(self):
        """Solve the exact landmark-defined transform from the currently completed pairs."""
        if self._base_a is None or self._base_b is None:
            return
        points_a, points_b = self._complete_landmark_arrays()
        mode = str(self.landmark_mode_combo.currentText() or "Rigid (rotate + shift)")
        rigid_hint = None
        try:
            if _is_rigid_mode(mode):
                rotation, shift_x, shift_y, rmse = _fit_rigid_from_points(points_a, points_b, self._base_b.shape)
            else:
                rotation, shift_x, shift_y, rmse = _fit_translation_from_points(points_a, points_b)
                if int(points_a.shape[0]) >= 2:
                    try:
                        rigid_rotation, rigid_dx, rigid_dy, rigid_rmse = _fit_rigid_from_points(
                            points_a,
                            points_b,
                            self._base_b.shape,
                        )
                        if np.isfinite(rigid_rmse) and rigid_rmse + 1e-6 < (0.75 * max(rmse, 1e-6)):
                            rigid_hint = (
                                "Translate-only fit keeps rotation fixed. "
                                f"Rigid fit would reduce point RMSE to {rigid_rmse:.2f} px "
                                f"(rot={rigid_rotation:.2f} deg, dx={rigid_dx:.2f} px, dy={rigid_dy:.2f} px)."
                            )
                    except Exception:
                        rigid_hint = None
            self._set_transform_controls(rotation, shift_x, shift_y)
            self._last_landmark_rmse = float(rmse)
            self._update_landmark_status()
            self._schedule_update(immediate=True)
            if rigid_hint:
                self._show_landmark_hint(rigid_hint)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to fit the selected landmarks.\n{exc}")

    def _undo_landmark(self):
        """Remove the last picked point so a mis-click can be corrected quickly."""
        if not self._landmark_pairs:
            return
        last = self._landmark_pairs[-1]
        if last.get("b") is not None:
            last["b"] = None
        else:
            self._landmark_pairs.pop()
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _clear_landmarks(self):
        """Drop all manually picked points while leaving the current transform untouched."""
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _update_landmark_status(self):
        """Refresh the compact status text and enablement for landmark-related controls."""
        complete_pairs = sum(1 for pair in self._landmark_pairs if pair.get("a") is not None and pair.get("b") is not None)
        total_pairs = len(self._landmark_pairs)
        expected = self._expected_landmark_target()
        next_index = total_pairs + 1 if expected == "a" else total_pairs
        mode_text = str(self.landmark_mode_combo.currentText() or "Rigid (rotate + shift)")
        required_pairs = 2 if _is_rigid_mode(mode_text) else 1
        if self._base_a is None or self._base_b is None:
            text = "Landmarks unavailable until both compare slots are set."
        elif self.pick_landmarks_btn.isChecked():
            text = f"{complete_pairs} pair(s) stored. Next click: {expected.upper()}{max(1, next_index)}."
        elif complete_pairs:
            text = f"{complete_pairs} pair(s) stored."
        else:
            text = "Toggle 'Pick landmarks' to add matching A/B reference points."
        if self._base_a is not None and complete_pairs < required_pairs:
            fit_kind = "rigid" if _is_rigid_mode(mode_text) else "translate-only"
            text += f" {required_pairs} complete pair(s) required for {fit_kind} fit."
        if self._last_landmark_rmse is not None and np.isfinite(self._last_landmark_rmse):
            text += f" Fit RMSE: {self._last_landmark_rmse:.2f} px."
        self.landmark_status.setText(text)
        self.fit_landmarks_btn.setEnabled(
            complete_pairs >= required_pairs and self._base_a is not None and self._base_b is not None
        )
        self.undo_landmarks_btn.setEnabled(bool(self._landmark_pairs))
        self.clear_landmarks_btn.setEnabled(bool(self._landmark_pairs))

    def _expected_landmark_target(self):
        """Return which panel should receive the next click in the alternating A/B workflow."""
        if not self._landmark_pairs or self._landmark_pairs[-1].get("b") is not None:
            return "a"
        return "b"

    def _complete_landmark_arrays(self):
        """Return completed landmark pairs as two Nx2 float arrays."""
        points_a = []
        points_b = []
        for pair in self._landmark_pairs:
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is None or point_b is None:
                continue
            points_a.append(point_a)
            points_b.append(point_b)
        if not points_a:
            raise ValueError("Pick at least one complete A/B landmark pair first.")
        return np.asarray(points_a, dtype=float), np.asarray(points_b, dtype=float)

    def _event_to_landmark_point(self, role, x_val, y_val):
        """Convert a click into A-grid or raw-B pixel coordinates in the shared compare frame."""
        point = self._axis_to_pixel(x_val, y_val)
        if role == "a":
            if not self._point_in_bounds(point, self._base_a.shape):
                raise ValueError("A landmark click must stay inside the visible image.")
            return tuple(point)
        base_point = _inverse_transform_points(
            [point],
            self._base_b.shape,
            rotation_deg=float(self.rotation_spin.value()),
            shift_x=float(self.shift_x_spin.value()),
            shift_y=float(self.shift_y_spin.value()),
        )[0]
        if not self._point_in_bounds(base_point, self._base_b.shape):
            raise ValueError("The clicked B point maps outside the source image. Reset or refine the transform first.")
        return tuple(base_point)

    def _clear_landmark_overlays(self):
        """Remove existing landmark markers before re-drawing them on the latest axes."""
        for artist in self._landmark_overlay_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._landmark_overlay_artists = []

    def _draw_landmark_overlays(self):
        """Repaint numbered landmark markers on the A and aligned-B image panels."""
        self._clear_landmark_overlays()
        if not self._landmark_pairs or not self._image_axes:
            self.canvas.draw_idle()
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        ax_a = self._image_axes.get("a")
        ax_b = self._image_axes.get("b")
        if ax_a is None or ax_b is None:
            return
        for index, pair in enumerate(self._landmark_pairs, start=1):
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is not None:
                self._add_landmark_marker(ax_a, point_a, index, color="#ffd24d")
            if point_b is not None:
                aligned_point = _transform_points(
                    [point_b],
                    self._base_b.shape,
                    rotation_deg=rotation,
                    shift_x=shift_x,
                    shift_y=shift_y,
                )[0]
                self._add_landmark_marker(ax_b, aligned_point, index, color="#7ee0ff")
        self.canvas.draw_idle()

    def _add_landmark_marker(self, ax, point, index, *, color):
        """Draw a numbered landmark marker on one image axis."""
        coords = self._pixel_to_axis(point)
        if coords is None:
            return
        x_val, y_val = coords
        scatter = ax.scatter(
            [x_val],
            [y_val],
            s=64,
            c=color,
            edgecolors="black",
            linewidths=0.9,
            zorder=30,
        )
        text = ax.text(
            x_val,
            y_val,
            f" {index}",
            color="white",
            fontsize=9,
            fontweight="bold",
            ha="left",
            va="bottom",
            zorder=31,
        )
        self._landmark_overlay_artists.extend([scatter, text])

    def _on_profile_mode_toggled(self, checked):
        """Switch into line-profile drawing mode without interfering with landmark picking."""
        if checked and self.pick_landmarks_btn.isChecked():
            self.pick_landmarks_btn.blockSignals(True)
            self.pick_landmarks_btn.setChecked(False)
            self.pick_landmarks_btn.blockSignals(False)
            self._update_landmark_status()
        cursor = QtCore.Qt.CrossCursor if checked else QtCore.Qt.ArrowCursor
        self.canvas.setCursor(QtGui.QCursor(cursor))

    def _clear_profile_overlay(self):
        """Remove the current ROI line overlay from all image panels."""
        for artist in self._profile_overlay_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._profile_overlay_artists = []

    def _draw_profile_overlay(self):
        """Show the active line ROI across all four image panels simultaneously."""
        self._clear_profile_overlay()
        if self._profile_line is None or not self._image_axes:
            self.canvas.draw_idle()
            return
        start = np.asarray(self._profile_line["start"], dtype=float)
        end = np.asarray(self._profile_line["end"], dtype=float)
        for ax in self._image_axes.values():
            line, = ax.plot(
                [start[0], end[0]],
                [start[1], end[1]],
                color="#ffd24d",
                linewidth=1.5,
                zorder=28,
            )
            start_marker = ax.scatter([start[0]], [start[1]], s=20, c="#ffd24d", edgecolors="black", zorder=29)
            end_marker = ax.scatter([end[0]], [end[1]], s=20, c="#ffd24d", edgecolors="black", zorder=29)
            self._profile_overlay_artists.extend([line, start_marker, end_marker])
        self.canvas.draw_idle()

    def _set_profile_line(self, start_axis, end_axis):
        """Store a shared axis-space profile ROI and refresh its overlays and line plot."""
        self._profile_line = {
            "start": (float(start_axis[0]), float(start_axis[1])),
            "end": (float(end_axis[0]), float(end_axis[1])),
        }
        self._draw_profile_overlay()
        self._plot_profile_panel()
        self.canvas.draw_idle()

    def _ensure_crosshair_artists(self):
        """Create hidden crosshair artists lazily so hover updates only move line objects."""
        if self._crosshair_artists or not self._image_axes:
            return
        for role, ax in self._image_axes.items():
            vline = ax.axvline(0.0, color="white", linewidth=0.8, alpha=0.8, visible=False, zorder=24)
            hline = ax.axhline(0.0, color="white", linewidth=0.8, alpha=0.8, visible=False, zorder=24)
            self._crosshair_artists[role] = (vline, hline)

    def _update_crosshair(self, x_val, y_val, *, redraw=True):
        """Link a hover crosshair across all image panels and update the status readout."""
        if not self._render_state or not self._image_axes:
            return
        point = self._axis_to_pixel(x_val, y_val)
        if not self._point_in_bounds(point, self._render_state.get("shape")):
            self._clear_crosshair(redraw=redraw)
            return
        self._ensure_crosshair_artists()
        for vline, hline in self._crosshair_artists.values():
            vline.set_xdata([float(x_val), float(x_val)])
            hline.set_ydata([float(y_val), float(y_val)])
            vline.set_visible(True)
            hline.set_visible(True)
        self._hover_axis_point = (float(x_val), float(y_val))
        self._update_hover_label(point, float(x_val), float(y_val))
        if redraw:
            self.canvas.draw_idle()

    def _clear_crosshair(self, *, redraw=True):
        """Hide the linked crosshair and clear the hover readout."""
        for vline, hline in self._crosshair_artists.values():
            vline.set_visible(False)
            hline.set_visible(False)
        self._hover_axis_point = None
        self.hover_label.setText("")
        if redraw:
            self.canvas.draw_idle()

    def _update_hover_label(self, point_px, x_val, y_val):
        """Format the current hover location and interpolated A/B/diff values for the status line."""
        axis_unit = str(self._render_state.get("axis_unit") or "nm")
        value_unit = str(self._render_state.get("value_unit") or "nm")
        a_val = _sample_image_value(self._render_state.get("a_raw"), point_px)
        b_val = _sample_image_value(self._render_state.get("b_raw"), point_px)
        diff_val = _sample_image_value(self._render_state.get("diff"), point_px)

        def _fmt(value):
            return "--" if value is None or not np.isfinite(value) else f"{float(value):.4g}"

        self.hover_label.setText(
            f"x: {x_val:.4g} {axis_unit} | y: {y_val:.4g} {axis_unit} | "
            f"A: {_fmt(a_val)} {value_unit} | B: {_fmt(b_val)} {value_unit} | A-B: {_fmt(diff_val)} {value_unit}"
        )

    def _on_canvas_press(self, event):
        """Handle landmark picking and profile-line starts on the compare figure."""
        if getattr(event, "button", None) != 1 or event.xdata is None or event.ydata is None:
            return
        role = self._axis_role(getattr(event, "inaxes", None))
        if role not in {"a", "b", "diff", "abs"}:
            return
        if self.profile_mode_btn.isChecked():
            self._profile_drag_active = True
            self._set_profile_line((event.xdata, event.ydata), (event.xdata, event.ydata))
            return
        if not self.pick_landmarks_btn.isChecked() or role not in {"a", "b"}:
            return
        expected = self._expected_landmark_target()
        if role != expected:
            self._show_landmark_hint(f"Click the next reference point in {expected.upper()} first.")
            return
        try:
            point = self._event_to_landmark_point(role, event.xdata, event.ydata)
        except ValueError as exc:
            self._show_landmark_hint(str(exc))
            return
        if role == "a":
            self._landmark_pairs.append({"a": point, "b": None})
        else:
            self._landmark_pairs[-1]["b"] = point
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _on_canvas_motion(self, event):
        """Update the linked crosshair and live profile while the user moves across image panels."""
        role = self._axis_role(getattr(event, "inaxes", None))
        if self._profile_drag_active:
            if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
                self._set_profile_line(self._profile_line["start"], (event.xdata, event.ydata))
                self._update_crosshair(event.xdata, event.ydata, redraw=False)
                self.canvas.draw_idle()
            return
        if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
            self._update_crosshair(event.xdata, event.ydata)
        else:
            self._clear_crosshair()

    def _on_canvas_release(self, event):
        """Finish a profile drag without disturbing normal hover crosshair updates."""
        if not self._profile_drag_active:
            return
        self._profile_drag_active = False
        role = self._axis_role(getattr(event, "inaxes", None))
        if role in {"a", "b", "diff", "abs"} and event.xdata is not None and event.ydata is not None:
            self._set_profile_line(self._profile_line["start"], (event.xdata, event.ydata))
            self._update_crosshair(event.xdata, event.ydata)
        else:
            self.canvas.draw_idle()

    def _on_canvas_leave(self, _event):
        """Clear the linked crosshair when the pointer leaves the compare figure."""
        if not self._profile_drag_active:
            self._clear_crosshair()

    def _show_landmark_hint(self, text):
        """Show lightweight point-picking feedback without interrupting the dialog workflow."""
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self, self.rect(), 2200)
        except Exception:
            pass

    def _default_export_stem(self):
        """Build a deterministic export stem from the current compare filenames plus a timestamp."""
        source_a = Path(str((self._snapshot_a or {}).get("path") or _safe_label(self._snapshot_a)))
        source_b = Path(str((self._snapshot_b or {}).get("path") or _safe_label(self._snapshot_b)))
        stem_a = _sanitize_filename_token(source_a.stem or source_a.name or "A")
        stem_b = _sanitize_filename_token(source_b.stem or source_b.name or "B")
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        return f"{stem_a}__vs__{stem_b}_{timestamp}"

    def _export_comparison(self):
        """Save a 300 DPI four-panel PNG and a CSV of the current A-B difference image."""
        state = dict(self._render_state or {})
        if not state:
            QtWidgets.QMessageBox.information(self, "Export", "No comparison data is available to export.")
            return
        default_dir = None
        path_a = str((self._snapshot_a or {}).get("path") or "").strip()
        if path_a:
            try:
                default_dir = Path(path_a).resolve().parent
            except Exception:
                default_dir = None
        if default_dir is None:
            viewer_dir = getattr(self.viewer, "last_dir", None)
            if viewer_dir:
                try:
                    default_dir = Path(str(viewer_dir))
                except Exception:
                    default_dir = None
        if default_dir is None:
            default_dir = Path.cwd()
        stem = self._default_export_stem()
        default_path = default_dir / f"{stem}.png"
        png_path, _selected = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export comparison figure",
            str(default_path),
            "PNG Files (*.png)",
        )
        if not png_path:
            return
        png_path = Path(png_path)
        if png_path.suffix.lower() != ".png":
            png_path = png_path.with_suffix(".png")
        csv_path = png_path.with_name(f"{png_path.stem}_AminusB.csv")
        try:
            self._export_png(png_path)
            self._export_diff_csv(csv_path)
            try:
                self.viewer.last_dir = str(png_path.parent)
            except Exception:
                pass
            QtWidgets.QMessageBox.information(
                self,
                "Export",
                f"Saved\n{png_path}\nand\n{csv_path}",
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Export", f"Unable to export the comparison.\n{exc}")

    def _export_png(self, path):
        """Create a standalone 2x2 export figure with current display ranges and colorbars."""
        state = dict(self._render_state or {})
        fig = Figure(figsize=(9.0, 8.2))
        axes = fig.subplots(2, 2)
        panels = [
            (axes[0][0], state["a_display"], state["title_a"], state["topo_cmap"], state["a_clim"], state["display_unit"]),
            (axes[0][1], state["b_display"], state["title_b"], state["topo_cmap"], state["b_clim"], state["display_unit"]),
            (axes[1][0], state["diff"], "A - B", "RdBu_r", state["diff_clim"], f"A-B ({state['value_unit']})"),
            (axes[1][1], state["abs_diff"], "|A - B|", "magma", state["abs_clim"], f"|A-B| ({state['value_unit']})"),
        ]
        for ax, arr, title, cmap, clim, cbar_label in panels:
            image = ax.imshow(
                np.asarray(arr, dtype=float),
                cmap=str(cmap),
                vmin=float(clim[0]),
                vmax=float(clim[1]),
                origin=self.IMAGE_ORIGIN,
                extent=state["extent"],
                interpolation="nearest",
                aspect="equal",
            )
            ax.set_title(str(title), fontsize=10)
            cbar = fig.colorbar(image, ax=ax, fraction=0.046, pad=0.03)
            cbar.set_label(str(cbar_label))
            ax.tick_params(labelsize=8)
        fig.tight_layout()
        fig.savefig(str(path), dpi=300, bbox_inches="tight")

    def _export_diff_csv(self, path):
        """Save the current A-B image as a CSV table with axis coordinates expressed in the current unit."""
        state = dict(self._render_state or {})
        diff = np.asarray(state.get("diff"), dtype=float)
        extent = state.get("extent")
        shape = state.get("shape")
        axis_unit = str(state.get("axis_unit") or "nm")
        if diff.ndim != 2 or extent is None or not shape:
            raise ValueError("The current difference image is unavailable.")
        x_coords = _axis_coords_for_size(extent[0], extent[1], shape[1])
        y_coords = _axis_coords_for_size(extent[2], extent[3], shape[0])
        with open(path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow([f"y_{axis_unit}/x_{axis_unit}", *[f"{val:.9g}" for val in x_coords]])
            for y_val, row in zip(y_coords, diff):
                writer.writerow([f"{float(y_val):.9g}", *[f"{float(val):.9g}" if np.isfinite(val) else "" for val in row]])

    def _clear_landmark_overlays(self):
        """Remove existing landmark markers before re-drawing them on the latest axes."""
        for artist in self._landmark_overlay_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._landmark_overlay_artists = []

    def _draw_landmark_overlays(self):
        """Repaint numbered landmark markers on the A and aligned-B panels."""
        self._clear_landmark_overlays()
        if not self._landmark_pairs:
            try:
                self.canvas.draw_idle()
            except Exception:
                pass
            return
        self._refresh_compare_axes()
        ax_a = self._compare_axes.get("a")
        ax_b = self._compare_axes.get("b")
        if ax_a is None or ax_b is None:
            return
        view_a = self.canvas._ax_view_map.get(ax_a)
        view_b = self.canvas._ax_view_map.get(ax_b)
        if view_a is None or view_b is None:
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        for index, pair in enumerate(self._landmark_pairs, start=1):
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is not None:
                self._add_landmark_marker(ax_a, view_a, point_a, index, color="#ffd24d")
            if point_b is not None:
                aligned_point = _transform_points(
                    [point_b],
                    self._base_b.shape,
                    rotation_deg=rotation,
                    shift_x=shift_x,
                    shift_y=shift_y,
                )[0]
                self._add_landmark_marker(ax_b, view_b, aligned_point, index, color="#7ee0ff")
        try:
            self.canvas.draw_idle()
        except Exception:
            pass

    def _add_landmark_marker(self, ax, view, point, index, *, color):
        """Draw a single numbered landmark marker on the provided compare axis."""
        coords = self._pixel_to_axis_coords(ax, view, point)
        if coords is None:
            return
        x_val, y_val = coords
        scatter = ax.scatter(
            [x_val],
            [y_val],
            s=64,
            c=color,
            edgecolors="black",
            linewidths=0.9,
            zorder=30,
        )
        text = ax.text(
            x_val,
            y_val,
            f" {index}",
            color="white",
            fontsize=9,
            fontweight="bold",
            ha="left",
            va="bottom",
            zorder=31,
        )
        self._landmark_overlay_artists.extend([scatter, text])

    def _display_meta(self, ax, view):
        """Return the rendered extent/origin/shape tuple used by matplotlib for this compare axis."""
        meta = dict(getattr(self.canvas, "_image_meta", {}).get(ax) or {})
        extent = meta.get("extent")
        origin = str(meta.get("origin", "upper") or "upper").lower()
        shape = meta.get("shape")
        if not shape:
            arr = np.asarray(view.get("arr")) if view and view.get("arr") is not None else None
            if arr is not None and arr.ndim >= 2:
                shape = arr.shape[:2]
        if extent is None:
            try:
                extent = self.canvas._view_extent(view)
            except Exception:
                extent = None
        try:
            shape = tuple(shape) if shape is not None else None
        except Exception:
            shape = None
        return extent, origin, shape

    def _axis_to_display_pixel(self, ax, view, x_val, y_val):
        """Map displayed axis coordinates onto the rendered image pixel grid for this axis."""
        extent, origin, shape = self._display_meta(ax, view)
        if shape and len(shape) >= 2 and extent is not None and len(extent) == 4:
            height = int(shape[0])
            width = int(shape[1])
            cols = max(width - 1, 1)
            rows = max(height - 1, 1)
            xmin, xmax, ymin, ymax = (float(v) for v in extent)
            span_x = float(xmax - xmin)
            span_y = float(ymax - ymin)
            col = 0.0 if abs(span_x) <= 1e-12 else ((float(x_val) - xmin) / span_x) * float(cols)
            if abs(span_y) <= 1e-12:
                row_use = 0.0
            elif origin == "upper":
                row_use = ((ymax - float(y_val)) / span_y) * float(rows)
            else:
                row_use = ((float(y_val) - ymin) / span_y) * float(rows)
            row = float(rows) - row_use if origin == "lower" and rows > 0 else row_use
            return np.array(
                [
                    float(np.clip(col, 0.0, max(0.0, float(width - 1)))),
                    float(np.clip(row, 0.0, max(0.0, float(height - 1)))),
                ],
                dtype=float,
            )
        arr = np.asarray(view.get("arr")) if view is not None and view.get("arr") is not None else None
        if arr is None or arr.ndim < 2 or arr.size == 0:
            return np.zeros((2,), dtype=float)
        height, width = arr.shape[:2]
        return np.array(
            [
                float(self.canvas._axis_coord_to_pixel_float(view, x_val, width, "x", ax=ax)),
                float(self.canvas._axis_coord_to_pixel_float(view, y_val, height, "y", ax=ax)),
            ],
            dtype=float,
        )

    def _display_pixel_to_axis(self, ax, view, point):
        """Map stored image pixel coordinates back onto the rendered axis coordinates."""
        if point is None:
            return None
        extent, origin, shape = self._display_meta(ax, view)
        if shape and len(shape) >= 2 and extent is not None and len(extent) == 4:
            height = int(shape[0])
            width = int(shape[1])
            cols = max(width - 1, 1)
            rows = max(height - 1, 1)
            xmin, xmax, ymin, ymax = (float(v) for v in extent)
            col = float(point[0])
            row = float(point[1])
            row_use = float(rows) - row if origin == "lower" and rows > 0 else row
            x_axis = xmin if cols == 0 else xmin + (col / float(cols)) * (xmax - xmin)
            if rows == 0:
                y_axis = ymax if origin == "upper" else ymin
            elif origin == "upper":
                y_axis = ymax - (row_use / float(rows)) * (ymax - ymin)
            else:
                y_axis = ymin + (row_use / float(rows)) * (ymax - ymin)
            return float(x_axis), float(y_axis)
        arr = np.asarray(view.get("arr")) if view is not None and view.get("arr") is not None else None
        if arr is None or arr.ndim < 2 or arr.size == 0:
            return None
        height, width = arr.shape[:2]
        x_idx = float(point[0])
        y_idx = float(point[1])
        try:
            extent = self.canvas._view_extent(view)
        except Exception:
            extent = None
        if extent is None:
            return x_idx, y_idx
        x_axis = float(self.canvas._index_to_axis_coord(x_idx, extent[0], extent[1], width))
        y_axis = float(self.canvas._index_to_axis_coord(y_idx, extent[2], extent[3], height))
        return x_axis, y_axis

    def _pixel_to_axis_coords(self, ax, view, point):
        """Convert stored pixel-space landmarks back into the current axis coordinate system."""
        return self._display_pixel_to_axis(ax, view, point)

    def _show_landmark_hint(self, text):
        """Show lightweight point-picking feedback without interrupting the dialog workflow."""
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self, self.rect(), 2200)
        except Exception:
            pass


class ImageCompareDialog(QtWidgets.QDialog):
    """Popup that aligns B onto A and renders image, scatter, and profile diagnostics."""

    TOPO_CMAPS = ("viridis", "plasma", "magma", "inferno", "cividis", "afmhot", "gray")
    IMAGE_ORIGIN = "upper"

    def __init__(self, controller, viewer, snapshot_a, snapshot_b):
        super().__init__(viewer)
        self.controller = controller
        self.viewer = viewer
        self._snapshot_a = None
        self._snapshot_b = None
        self._base_a = None
        self._base_b = None
        self._render_state = {}
        self._image_axes = {}
        self._profile_axes = {}
        self._landmark_pairs = []
        self._landmark_overlay_artists = []
        self._profile_overlay_artists = []
        self._crosshair_artists = {}
        self._last_landmark_rmse = None
        self._profile_line = None
        self._profile_drag_active = False
        self._hover_axis_point = None
        self._update_timer = QtCore.QTimer(self)
        self._update_timer.setSingleShot(True)
        self._update_timer.setInterval(40)
        self._update_timer.timeout.connect(self._rebuild_views)
        self.setWindowTitle("Compare A/B")
        self.setMinimumSize(980, 820)
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
            | QtCore.Qt.WindowSystemMenuHint
        )

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_row = QtWidgets.QHBoxLayout()
        header_row.setSpacing(8)
        header_row.addWidget(QtWidgets.QLabel("A:", self))
        self.label_a = QtWidgets.QLabel("", self)
        self.label_a.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        header_row.addWidget(self.label_a, 1)
        header_row.addWidget(QtWidgets.QLabel("B:", self))
        self.label_b = QtWidgets.QLabel("", self)
        self.label_b.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        header_row.addWidget(self.label_b, 1)
        layout.addLayout(header_row)

        controls = QtWidgets.QHBoxLayout()
        controls.setSpacing(8)
        controls.addWidget(QtWidgets.QLabel("Auto mode:", self))
        self.auto_mode_combo = QtWidgets.QComboBox(self)
        self.auto_mode_combo.addItems(["Translate only", "Rigid (rotate + shift)"])
        self.auto_mode_combo.setToolTip(
            "Translate only keeps the current rotation fixed. Rigid solves rotation plus translation."
        )
        controls.addWidget(self.auto_mode_combo)
        controls.addWidget(QtWidgets.QLabel("Intensity:", self))
        self.intensity_combo = QtWidgets.QComboBox(self)
        self.intensity_combo.addItems(["None", "Median match", "Mean match"])
        self.intensity_combo.setToolTip("Match the baseline of B to A before computing difference maps.")
        controls.addWidget(self.intensity_combo)
        controls.addWidget(QtWidgets.QLabel("Rotation:", self))
        self.rotation_spin = QtWidgets.QDoubleSpinBox(self)
        self.rotation_spin.setRange(-180.0, 180.0)
        self.rotation_spin.setDecimals(2)
        self.rotation_spin.setSingleStep(1.0)
        self.rotation_spin.setSuffix(" deg")
        self.rotation_spin.setKeyboardTracking(False)
        self.rotation_spin.setToolTip("Manual rotation applied to B before comparison.")
        controls.addWidget(self.rotation_spin)
        controls.addWidget(QtWidgets.QLabel("Shift X:", self))
        self.shift_x_spin = QtWidgets.QDoubleSpinBox(self)
        self.shift_x_spin.setRange(-4096.0, 4096.0)
        self.shift_x_spin.setDecimals(2)
        self.shift_x_spin.setSingleStep(0.5)
        self.shift_x_spin.setSuffix(" px")
        self.shift_x_spin.setKeyboardTracking(False)
        self.shift_x_spin.setToolTip("Horizontal shift on the resampled A grid.")
        controls.addWidget(self.shift_x_spin)
        controls.addWidget(QtWidgets.QLabel("Shift Y:", self))
        self.shift_y_spin = QtWidgets.QDoubleSpinBox(self)
        self.shift_y_spin.setRange(-4096.0, 4096.0)
        self.shift_y_spin.setDecimals(2)
        self.shift_y_spin.setSingleStep(0.5)
        self.shift_y_spin.setSuffix(" px")
        self.shift_y_spin.setKeyboardTracking(False)
        self.shift_y_spin.setToolTip("Vertical shift on the resampled A grid.")
        controls.addWidget(self.shift_y_spin)
        self.auto_btn = QtWidgets.QPushButton("Auto fit", self)
        self.auto_btn.clicked.connect(self._auto_fit)
        controls.addWidget(self.auto_btn)
        self.reset_btn = QtWidgets.QPushButton("Reset", self)
        self.reset_btn.clicked.connect(self._reset_transform)
        controls.addWidget(self.reset_btn)
        self.swap_btn = QtWidgets.QPushButton("Swap A/B", self)
        self.swap_btn.clicked.connect(self._swap_slots)
        controls.addWidget(self.swap_btn)
        controls.addStretch(1)
        layout.addLayout(controls)

        display_row = QtWidgets.QHBoxLayout()
        display_row.setSpacing(8)
        display_row.addWidget(QtWidgets.QLabel("Colormap:", self))
        self.cmap_combo = QtWidgets.QComboBox(self)
        self.cmap_combo.addItems(list(self.TOPO_CMAPS))
        self.cmap_combo.setCurrentText("viridis")
        self.cmap_combo.setToolTip("Topography colormap for A and B.")
        display_row.addWidget(self.cmap_combo)
        display_row.addWidget(QtWidgets.QLabel("Stretch:", self))
        self.stretch_combo = QtWidgets.QComboBox(self)
        self.stretch_combo.addItems(["linear", "histeq"])
        self.stretch_combo.setToolTip("Display-only stretch for A and B. Does not modify stored data.")
        display_row.addWidget(self.stretch_combo)
        self.lock_range_btn = QtWidgets.QPushButton("Lock A/B range", self)
        self.lock_range_btn.setCheckable(True)
        self.lock_range_btn.setToolTip("Use one shared vmin/vmax for A and B so the colormaps are directly comparable.")
        display_row.addWidget(self.lock_range_btn)
        self.profile_mode_btn = QtWidgets.QPushButton("Profile", self)
        self.profile_mode_btn.setCheckable(True)
        self.profile_mode_btn.setToolTip("Drag a line on any image panel to compare A, B, and A-B cross-sections.")
        self.profile_mode_btn.toggled.connect(self._on_profile_mode_toggled)
        display_row.addWidget(self.profile_mode_btn)
        self.export_btn = QtWidgets.QPushButton("Export", self)
        self.export_btn.clicked.connect(self._export_comparison)
        display_row.addWidget(self.export_btn)
        display_row.addStretch(1)
        layout.addLayout(display_row)

        landmark_row = QtWidgets.QHBoxLayout()
        landmark_row.setSpacing(8)
        self.pick_landmarks_btn = QtWidgets.QPushButton("Pick landmarks", self)
        self.pick_landmarks_btn.setCheckable(True)
        self.pick_landmarks_btn.setToolTip(
            "Alternate clicks between the A and B panels to define matching reference points."
        )
        self.pick_landmarks_btn.toggled.connect(self._on_pick_landmarks_toggled)
        landmark_row.addWidget(self.pick_landmarks_btn)
        landmark_row.addWidget(QtWidgets.QLabel("Point fit:", self))
        self.landmark_mode_combo = QtWidgets.QComboBox(self)
        self.landmark_mode_combo.addItems(["Rigid (rotate + shift)", "Translate only"])
        self.landmark_mode_combo.setToolTip(
            "Rigid uses two or more pairs to solve rotation plus shift. Translate only keeps rotation fixed."
        )
        self.landmark_mode_combo.currentIndexChanged.connect(lambda *_args: self._update_landmark_status())
        landmark_row.addWidget(self.landmark_mode_combo)
        self.fit_landmarks_btn = QtWidgets.QPushButton("Fit from points", self)
        self.fit_landmarks_btn.clicked.connect(self._fit_from_landmarks)
        landmark_row.addWidget(self.fit_landmarks_btn)
        self.undo_landmarks_btn = QtWidgets.QPushButton("Undo point", self)
        self.undo_landmarks_btn.clicked.connect(self._undo_landmark)
        landmark_row.addWidget(self.undo_landmarks_btn)
        self.clear_landmarks_btn = QtWidgets.QPushButton("Clear points", self)
        self.clear_landmarks_btn.clicked.connect(self._clear_landmarks)
        landmark_row.addWidget(self.clear_landmarks_btn)
        self.landmark_status = QtWidgets.QLabel("", self)
        self.landmark_status.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        landmark_row.addWidget(self.landmark_status, 1)
        layout.addLayout(landmark_row)

        self.metrics_label = QtWidgets.QLabel("", self)
        self.metrics_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        layout.addWidget(self.metrics_label)
        self.hover_label = QtWidgets.QLabel("", self)
        self.hover_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        layout.addWidget(self.hover_label)

        self._figure = Figure(figsize=(9.8, 8.6))
        self.canvas = FigureCanvas(self._figure)
        layout.addWidget(self.canvas, 1)
        self._press_cid = self.canvas.mpl_connect("button_press_event", self._on_canvas_press)
        self._motion_cid = self.canvas.mpl_connect("motion_notify_event", self._on_canvas_motion)
        self._release_cid = self.canvas.mpl_connect("button_release_event", self._on_canvas_release)
        self._leave_cid = self.canvas.mpl_connect("figure_leave_event", self._on_canvas_leave)

        self.intensity_combo.currentIndexChanged.connect(lambda *_args: self._schedule_update())
        self.rotation_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.shift_x_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.shift_y_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.cmap_combo.currentIndexChanged.connect(lambda *_args: self._schedule_update(immediate=True))
        self.stretch_combo.currentIndexChanged.connect(lambda *_args: self._schedule_update(immediate=True))
        self.lock_range_btn.toggled.connect(lambda *_args: self._schedule_update(immediate=True))

        self.set_snapshots(snapshot_a, snapshot_b)
        self._update_landmark_status()
        self.resize(1220, 930)

    def _on_pick_landmarks_toggled(self, _checked):
        if self.pick_landmarks_btn.isChecked():
            if self.profile_mode_btn.isChecked():
                self.profile_mode_btn.blockSignals(True)
                self.profile_mode_btn.setChecked(False)
                self.profile_mode_btn.blockSignals(False)
            if not self._landmark_pairs:
                self._set_transform_controls(0.0, 0.0, 0.0)
                self._schedule_update(immediate=True)
            self._show_landmark_hint("Click a reference point in A, then the matching point in B.")
        self._update_landmark_status()

    def _fit_from_landmarks(self):
        if self._base_a is None or self._base_b is None:
            return
        points_a, points_b = self._complete_landmark_arrays()
        mode = str(self.landmark_mode_combo.currentText() or "Rigid (rotate + shift)")
        rigid_hint = None
        try:
            if _is_rigid_mode(mode):
                rotation, shift_x, shift_y, rmse = _fit_rigid_from_points(points_a, points_b, self._base_b.shape)
            else:
                rotation, shift_x, shift_y, rmse = _fit_translation_from_points(points_a, points_b)
                if int(points_a.shape[0]) >= 2:
                    try:
                        rigid_rotation, rigid_dx, rigid_dy, rigid_rmse = _fit_rigid_from_points(
                            points_a,
                            points_b,
                            self._base_b.shape,
                        )
                        if np.isfinite(rigid_rmse) and rigid_rmse + 1e-6 < (0.75 * max(rmse, 1e-6)):
                            rigid_hint = (
                                "Translate-only fit keeps rotation fixed. "
                                f"Rigid fit would reduce point RMSE to {rigid_rmse:.2f} px "
                                f"(rot={rigid_rotation:.2f} deg, dx={rigid_dx:.2f} px, dy={rigid_dy:.2f} px)."
                            )
                    except Exception:
                        rigid_hint = None
            self._set_transform_controls(rotation, shift_x, shift_y)
            self._last_landmark_rmse = float(rmse)
            self._update_landmark_status()
            self._schedule_update(immediate=True)
            if rigid_hint:
                self._show_landmark_hint(rigid_hint)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to fit the selected landmarks.\n{exc}")

    def _undo_landmark(self):
        if not self._landmark_pairs:
            return
        last = self._landmark_pairs[-1]
        if last.get("b") is not None:
            last["b"] = None
        else:
            self._landmark_pairs.pop()
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _clear_landmarks(self):
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _update_landmark_status(self):
        complete_pairs = sum(1 for pair in self._landmark_pairs if pair.get("a") is not None and pair.get("b") is not None)
        total_pairs = len(self._landmark_pairs)
        expected = self._expected_landmark_target()
        next_index = total_pairs + 1 if expected == "a" else total_pairs
        mode_text = str(self.landmark_mode_combo.currentText() or "Rigid (rotate + shift)")
        required_pairs = 2 if _is_rigid_mode(mode_text) else 1
        if self._base_a is None or self._base_b is None:
            text = "Landmarks unavailable until both compare slots are set."
        elif self.pick_landmarks_btn.isChecked():
            text = f"{complete_pairs} pair(s) stored. Next click: {expected.upper()}{max(1, next_index)}."
        elif complete_pairs:
            text = f"{complete_pairs} pair(s) stored."
        else:
            text = "Toggle 'Pick landmarks' to add matching A/B reference points."
        if self._base_a is not None and complete_pairs < required_pairs:
            fit_kind = "rigid" if _is_rigid_mode(mode_text) else "translate-only"
            text += f" {required_pairs} complete pair(s) required for {fit_kind} fit."
        if self._last_landmark_rmse is not None and np.isfinite(self._last_landmark_rmse):
            text += f" Fit RMSE: {self._last_landmark_rmse:.2f} px."
        self.landmark_status.setText(text)
        self.fit_landmarks_btn.setEnabled(
            complete_pairs >= required_pairs and self._base_a is not None and self._base_b is not None
        )
        self.undo_landmarks_btn.setEnabled(bool(self._landmark_pairs))
        self.clear_landmarks_btn.setEnabled(bool(self._landmark_pairs))

    def _expected_landmark_target(self):
        if not self._landmark_pairs or self._landmark_pairs[-1].get("b") is not None:
            return "a"
        return "b"

    def _complete_landmark_arrays(self):
        points_a = []
        points_b = []
        for pair in self._landmark_pairs:
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is None or point_b is None:
                continue
            points_a.append(point_a)
            points_b.append(point_b)
        if not points_a:
            raise ValueError("Pick at least one complete A/B landmark pair first.")
        return np.asarray(points_a, dtype=float), np.asarray(points_b, dtype=float)

    def _event_to_landmark_point(self, role, x_val, y_val):
        point = self._axis_to_pixel(x_val, y_val)
        if role == "a":
            if not self._point_in_bounds(point, self._base_a.shape):
                raise ValueError("A landmark click must stay inside the visible image.")
            return tuple(point)
        base_point = _inverse_transform_points(
            [point],
            self._base_b.shape,
            rotation_deg=float(self.rotation_spin.value()),
            shift_x=float(self.shift_x_spin.value()),
            shift_y=float(self.shift_y_spin.value()),
        )[0]
        if not self._point_in_bounds(base_point, self._base_b.shape):
            raise ValueError("The clicked B point maps outside the source image. Reset or refine the transform first.")
        return tuple(base_point)

    def _clear_landmark_overlays(self):
        for artist in self._landmark_overlay_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._landmark_overlay_artists = []

    def _draw_landmark_overlays(self):
        self._clear_landmark_overlays()
        if not self._landmark_pairs or not self._image_axes:
            self.canvas.draw_idle()
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        ax_a = self._image_axes.get("a")
        ax_b = self._image_axes.get("b")
        if ax_a is None or ax_b is None:
            return
        for index, pair in enumerate(self._landmark_pairs, start=1):
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is not None:
                self._add_landmark_marker(ax_a, point_a, index, color="#ffd24d")
            if point_b is not None:
                aligned_point = _transform_points(
                    [point_b],
                    self._base_b.shape,
                    rotation_deg=rotation,
                    shift_x=shift_x,
                    shift_y=shift_y,
                )[0]
                self._add_landmark_marker(ax_b, aligned_point, index, color="#7ee0ff")
        self.canvas.draw_idle()

    def _add_landmark_marker(self, ax, point, index, *, color):
        coords = self._pixel_to_axis(point)
        if coords is None:
            return
        x_val, y_val = coords
        scatter = ax.scatter(
            [x_val],
            [y_val],
            s=64,
            c=color,
            edgecolors="black",
            linewidths=0.9,
            zorder=30,
        )
        text = ax.text(
            x_val,
            y_val,
            f" {index}",
            color="white",
            fontsize=9,
            fontweight="bold",
            ha="left",
            va="bottom",
            zorder=31,
        )
        self._landmark_overlay_artists.extend([scatter, text])

    def _on_profile_mode_toggled(self, checked):
        return _icd_on_profile_mode_toggled(self, checked)

    def _clear_profile_overlay(self):
        return _icd_clear_profile_overlay(self)

    def _draw_profile_overlay(self):
        return _icd_draw_profile_overlay(self)

    def _set_profile_line(self, start_axis, end_axis):
        return _icd_set_profile_line(self, start_axis, end_axis)

    def _ensure_crosshair_artists(self):
        return _icd_ensure_crosshair_artists(self)

    def _update_crosshair(self, x_val, y_val, *, redraw=True):
        return _icd_update_crosshair(self, x_val, y_val, redraw=redraw)

    def _clear_crosshair(self, *, redraw=True):
        return _icd_clear_crosshair(self, redraw=redraw)

    def _update_hover_label(self, point_px, x_val, y_val):
        return _icd_update_hover_label(self, point_px, x_val, y_val)

    def _on_canvas_press(self, event):
        return _icd_on_canvas_press(self, event)

    def _on_canvas_motion(self, event):
        return _icd_on_canvas_motion(self, event)

    def _on_canvas_release(self, event):
        return _icd_on_canvas_release(self, event)

    def _on_canvas_leave(self, event):
        return _icd_on_canvas_leave(self, event)

    def _show_landmark_hint(self, text):
        return _icd_show_landmark_hint(self, text)

    def _default_export_stem(self):
        return _icd_default_export_stem(self)

    def _export_comparison(self):
        return _icd_export_comparison(self)

    def _export_png(self, path):
        return _icd_export_png(self, path)

    def _export_diff_csv(self, path):
        return _icd_export_diff_csv(self, path)

    def set_slots_pending(self, snapshot_a, snapshot_b):
        """Keep the dialog open but show that one of the slots is still missing."""
        self._snapshot_a = snapshot_a
        self._snapshot_b = snapshot_b
        self._base_a = None
        self._base_b = None
        self._render_state = {}
        self._image_axes = {}
        self._profile_axes = {}
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._profile_line = None
        self._profile_drag_active = False
        self._hover_axis_point = None
        self.label_a.setText(_safe_label(snapshot_a))
        self.label_b.setText(_safe_label(snapshot_b))
        self.metrics_label.setText("Both compare slots must be populated to render the comparison.")
        self.hover_label.setText("")
        self._set_transform_controls(0.0, 0.0, 0.0)
        self._draw_placeholder("Both compare slots must be populated to render the comparison.")
        self._update_landmark_status()

    def set_snapshots(self, snapshot_a, snapshot_b):
        """Replace A/B sources and rebuild the comparison figure."""
        self._snapshot_a = snapshot_a
        self._snapshot_b = snapshot_b
        self.label_a.setText(_safe_label(snapshot_a))
        self.label_b.setText(_safe_label(snapshot_b))
        if snapshot_a is None or snapshot_b is None:
            self.set_slots_pending(snapshot_a, snapshot_b)
            return
        self._base_a = np.array(snapshot_a.get("arr"), copy=True, dtype=float)
        self._base_b = _resample_to_reference_grid(
            snapshot_b.get("arr"),
            snapshot_b.get("extent"),
            self._base_a.shape,
            snapshot_a.get("extent"),
        )
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._profile_line = None
        self._profile_drag_active = False
        self._hover_axis_point = None
        span = float(max(self._base_a.shape[:2]))
        for spin in (self.shift_x_spin, self.shift_y_spin):
            spin.blockSignals(True)
            spin.setRange(-span, span)
            spin.blockSignals(False)
        self._set_transform_controls(0.0, 0.0, 0.0)
        self.setWindowTitle(f"Compare A/B - {_safe_label(snapshot_a)} vs {_safe_label(snapshot_b)}")
        self._update_landmark_status()
        self._schedule_update(immediate=True)

    def _schedule_update(self, immediate=False):
        """Batch rapid control changes into one redraw so dragging spinboxes stays responsive."""
        if immediate:
            self._update_timer.stop()
            self._rebuild_views()
            return
        self._update_timer.start()

    def _set_transform_controls(self, rotation_deg, shift_x, shift_y):
        """Update the transform widgets atomically so new slots do not inherit stale alignment."""
        self.rotation_spin.blockSignals(True)
        self.shift_x_spin.blockSignals(True)
        self.shift_y_spin.blockSignals(True)
        self.rotation_spin.setValue(float(rotation_deg))
        self.shift_x_spin.setValue(float(shift_x))
        self.shift_y_spin.setValue(float(shift_y))
        self.rotation_spin.blockSignals(False)
        self.shift_x_spin.blockSignals(False)
        self.shift_y_spin.blockSignals(False)

    def _reset_transform(self):
        """Reset manual alignment without clearing the current landmark pairs."""
        self._set_transform_controls(0.0, 0.0, 0.0)
        self._schedule_update(immediate=True)

    def _swap_slots(self):
        """Delegate slot swapping to the controller so menus and popup state stay in sync."""
        self.controller.swap_slots()

    def _auto_fit(self):
        """Estimate a transform from the image content instead of user-supplied landmarks."""
        if self._base_a is None or self._base_b is None:
            return
        mode = str(self.auto_mode_combo.currentText() or "Translate only")
        try:
            if _is_rigid_mode(mode):
                rotation, shift_x, shift_y = self._estimate_rigid_transform(self._base_a, self._base_b)
                self.rotation_spin.setValue(rotation)
            else:
                rotation = float(self.rotation_spin.value())
                shift_x, shift_y = self._estimate_translation(self._base_a, self._base_b, rotation)
            self.shift_x_spin.setValue(shift_x)
            self.shift_y_spin.setValue(shift_y)
            self._schedule_update(immediate=True)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to estimate alignment.\n{exc}")

    def _estimate_translation(self, reference, moving, rotation_deg):
        rotated, _mask = _transform_with_mask(moving, rotation_deg, 0.0, 0.0)
        shift_x, shift_y, _score = _phase_correlation_shift(reference, rotated)
        return float(shift_x), float(shift_y)

    def _estimate_rigid_transform(self, reference, moving):
        if not _HAS_SCIPY or ndimage is None:
            raise RuntimeError("Rigid auto-alignment requires scipy.")
        ref_ds, mov_ds, scale = self._downsample_pair(reference, moving)
        coarse_angles = np.arange(-180.0, 180.0, 15.0, dtype=float)
        best = None
        for angles in (coarse_angles, None):
            if angles is None and best is not None:
                angles = np.arange(best[0] - 15.0, best[0] + 15.01, 3.0, dtype=float)
            for angle in angles:
                try:
                    shift_x, shift_y = self._estimate_translation(ref_ds, mov_ds, float(angle))
                    aligned, _mask = _transform_with_mask(mov_ds, float(angle), shift_x, shift_y)
                    metrics = _alignment_metrics(ref_ds, aligned)
                    score = metrics.get("corr")
                    if score is None or not np.isfinite(score):
                        continue
                    if best is None or score > best[1]:
                        best = (float(angle), float(score), float(shift_x), float(shift_y))
                except Exception:
                    continue
        if best is None:
            raise RuntimeError("No valid rigid alignment candidate was found.")
        angle, _score, shift_x, shift_y = best
        if abs(scale - 1.0) > 1e-9:
            shift_x /= scale
            shift_y /= scale
        return float(angle), float(shift_x), float(shift_y)

    def _downsample_pair(self, reference, moving, max_size=256):
        """Shrink large arrays before auto-fit so rigid searches stay interactive."""
        ref = np.asarray(reference, dtype=float)
        mov = np.asarray(moving, dtype=float)
        longest = float(max(ref.shape[:2]))
        if longest <= float(max_size):
            return ref, mov, 1.0
        scale = float(max_size) / longest
        new_shape = (
            max(48, int(round(ref.shape[0] * scale))),
            max(48, int(round(ref.shape[1] * scale))),
        )
        return _resize_to_shape(ref, new_shape), _resize_to_shape(mov, new_shape), float(new_shape[1]) / float(ref.shape[1])

    def _rebuild_views(self):
        """Recompute aligned arrays, statistics, and figure artists from the current control state."""
        if self._snapshot_a is None or self._snapshot_b is None or self._base_a is None or self._base_b is None:
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        try:
            aligned_b, _mask = _transform_with_mask(self._base_b, rotation, shift_x, shift_y)
        except Exception as exc:
            self.metrics_label.setText(str(exc))
            self._draw_placeholder(str(exc))
            return
        aligned_b, offset = _match_intensity(self._base_a, aligned_b, self.intensity_combo.currentText())
        diff = np.array(self._base_a, copy=True) - np.array(aligned_b, copy=True)
        abs_diff = np.abs(diff)
        metrics = _alignment_metrics(self._base_a, aligned_b)
        self.metrics_label.setText(
            "Overlap: {coverage:.1%} | Corr: {corr:.4f} | RMSE: {rmse:.4g} | Mean(A-B): {mean_delta:.4g} | "
            "Offset(B): {offset:.4g} | rot={rotation:.2f} deg, dx={shift_x:.2f} px, dy={shift_y:.2f} px".format(
                coverage=metrics.get("coverage", 0.0),
                corr=metrics.get("corr", float("nan")),
                rmse=metrics.get("rmse", float("nan")),
                mean_delta=metrics.get("mean_delta", float("nan")),
                offset=offset,
                rotation=rotation,
                shift_x=shift_x,
                shift_y=shift_y,
            )
        )

        display_a = _apply_display_stretch(self._base_a, self.stretch_combo.currentText())
        display_b = _apply_display_stretch(aligned_b, self.stretch_combo.currentText())
        if self.lock_range_btn.isChecked():
            combined = np.concatenate(
                [
                    np.asarray(display_a, dtype=float)[np.isfinite(display_a)],
                    np.asarray(display_b, dtype=float)[np.isfinite(display_b)],
                ]
            )
            shared_clim = _finite_minmax(combined, fallback=(0.0, 1.0))
            clim_a = shared_clim
            clim_b = shared_clim
        else:
            clim_a = _finite_minmax(display_a)
            clim_b = _finite_minmax(display_b)
        finite_diff = np.asarray(diff, dtype=float)
        finite_diff = finite_diff[np.isfinite(finite_diff)]
        diff_span = max(float(np.max(np.abs(finite_diff))) if finite_diff.size else 1.0, 1e-9)
        finite_abs = np.asarray(abs_diff, dtype=float)
        finite_abs = finite_abs[np.isfinite(finite_abs)]
        abs_span = max(float(np.max(finite_abs)) if finite_abs.size else 1.0, 1e-9)
        extent = self._snapshot_a.get("extent")
        axis_unit = str(self._snapshot_a.get("axis_unit") or "nm")
        value_unit = str(self._snapshot_a.get("unit") or "").strip() or "nm"
        stretch_mode = str(self.stretch_combo.currentText() or "linear").strip().lower()
        display_unit = value_unit if not stretch_mode.startswith("hist") else "equalized intensity"
        self._render_state = {
            "extent": extent,
            "origin": self.IMAGE_ORIGIN,
            "shape": tuple(self._base_a.shape[:2]),
            "axis_unit": axis_unit,
            "value_unit": value_unit,
            "display_unit": display_unit,
            "rotation": rotation,
            "shift_x": shift_x,
            "shift_y": shift_y,
            "offset": float(offset),
            "metrics": metrics,
            "a_raw": np.array(self._base_a, copy=True),
            "b_raw": np.array(aligned_b, copy=True),
            "diff": diff,
            "abs_diff": abs_diff,
            "a_display": display_a,
            "b_display": display_b,
            "a_clim": clim_a,
            "b_clim": clim_b,
            "diff_clim": (-diff_span, diff_span),
            "abs_clim": (0.0, abs_span),
            "topo_cmap": str(self.cmap_combo.currentText() or "viridis"),
            "overlap_mask": np.isfinite(self._base_a) & np.isfinite(aligned_b),
            "title_a": f"A: {_safe_label(self._snapshot_a)}",
            "title_b": f"B aligned: {_safe_label(self._snapshot_b)}",
        }
        self._render_figure()

    def _draw_placeholder(self, text):
        """Render a simple placeholder message when the compare figure cannot be drawn."""
        self._figure.clear()
        ax = self._figure.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, str(text), ha="center", va="center", fontsize=11)
        self.canvas.draw_idle()

    def _render_figure(self):
        """Render the four image panels plus scatter and profile diagnostics from the current state."""
        state = dict(self._render_state or {})
        if not state:
            self._draw_placeholder("No comparison data available.")
            return
        self._figure.clear()
        grid = self._figure.add_gridspec(
            3,
            3,
            width_ratios=[1.0, 1.0, 1.1],
            height_ratios=[1.0, 1.0, 0.82],
            wspace=0.34,
            hspace=0.34,
        )
        ax_a = self._figure.add_subplot(grid[0, 0])
        ax_b = self._figure.add_subplot(grid[0, 1])
        ax_diff = self._figure.add_subplot(grid[1, 0])
        ax_abs = self._figure.add_subplot(grid[1, 1])
        ax_scatter = self._figure.add_subplot(grid[:2, 2])
        ax_profile = self._figure.add_subplot(grid[2, :])
        ax_profile_diff = ax_profile.twinx()

        self._image_axes = {
            "a": ax_a,
            "b": ax_b,
            "diff": ax_diff,
            "abs": ax_abs,
        }
        self._profile_axes = {"profile": ax_profile, "diff": ax_profile_diff, "scatter": ax_scatter}
        self._crosshair_artists = {}

        self._plot_image_panel(
            ax_a,
            state["a_display"],
            title=state["title_a"],
            cmap=state["topo_cmap"],
            clim=state["a_clim"],
            extent=state["extent"],
            colorbar_label=state["display_unit"],
        )
        self._plot_image_panel(
            ax_b,
            state["b_display"],
            title=state["title_b"],
            cmap=state["topo_cmap"],
            clim=state["b_clim"],
            extent=state["extent"],
            colorbar_label=state["display_unit"],
        )
        self._plot_image_panel(
            ax_diff,
            state["diff"],
            title="A - B",
            cmap="RdBu_r",
            clim=state["diff_clim"],
            extent=state["extent"],
            colorbar_label=f"A-B ({state['value_unit']})",
        )
        self._plot_image_panel(
            ax_abs,
            state["abs_diff"],
            title="|A - B|",
            cmap="magma",
            clim=state["abs_clim"],
            extent=state["extent"],
            colorbar_label=f"|A-B| ({state['value_unit']})",
        )
        self._plot_scatter_panel(ax_scatter)
        self._plot_profile_panel()
        self._draw_landmark_overlays()
        self._draw_profile_overlay()
        if self._hover_axis_point is not None:
            self._update_crosshair(self._hover_axis_point[0], self._hover_axis_point[1], redraw=False)
        else:
            self.hover_label.setText("")
        self.canvas.draw_idle()

    def _plot_image_panel(self, ax, arr, *, title, cmap, clim, extent, colorbar_label):
        """Plot one imshow panel plus its colorbar using the current display state."""
        image = ax.imshow(
            np.asarray(arr, dtype=float),
            cmap=str(cmap),
            vmin=float(clim[0]),
            vmax=float(clim[1]),
            origin=self.IMAGE_ORIGIN,
            extent=extent,
            interpolation="nearest",
            aspect="equal",
        )
        ax.set_title(str(title), fontsize=10)
        cbar = self._figure.colorbar(image, ax=ax, fraction=0.046, pad=0.03)
        cbar.set_label(str(colorbar_label))
        ax.tick_params(labelsize=8)

    def _plot_scatter_panel(self, ax):
        """Render A-vs-B overlap scatter with a 1:1 guide and fit annotations."""
        state = self._render_state
        mask = np.asarray(state.get("overlap_mask"), dtype=bool)
        a_vals = np.asarray(state.get("a_raw"), dtype=float)[mask]
        b_vals = np.asarray(state.get("b_raw"), dtype=float)[mask]
        diff_vals = np.asarray(state.get("abs_diff"), dtype=float)[mask]
        ax.cla()
        ax.set_title("A vs B overlap", fontsize=10)
        ax.set_xlabel(f"A ({state.get('value_unit', 'nm')})")
        ax.set_ylabel(f"B ({state.get('value_unit', 'nm')})")
        if a_vals.size <= 1:
            ax.text(0.5, 0.5, "No overlapping pixels", ha="center", va="center", transform=ax.transAxes)
            return
        if a_vals.size > 30000:
            idx = np.linspace(0, a_vals.size - 1, 30000, dtype=int)
            a_vals = a_vals[idx]
            b_vals = b_vals[idx]
            diff_vals = diff_vals[idx]
        scatter = ax.scatter(
            a_vals,
            b_vals,
            c=diff_vals,
            cmap="magma",
            s=7,
            alpha=0.55,
            linewidths=0.0,
            rasterized=True,
        )
        line_min = float(np.nanmin(np.concatenate([a_vals, b_vals])))
        line_max = float(np.nanmax(np.concatenate([a_vals, b_vals])))
        ax.plot([line_min, line_max], [line_min, line_max], linestyle="--", color="black", linewidth=1.0)
        corr = self._render_state.get("metrics", {}).get("corr", float("nan"))
        try:
            slope = float(np.polyfit(a_vals, b_vals, 1)[0]) if np.nanstd(a_vals) > 1e-12 else float("nan")
        except Exception:
            slope = float("nan")
        ax.text(
            0.03,
            0.97,
            f"r = {corr:.4f}\nslope = {slope:.4f}\nN = {a_vals.size}",
            transform=ax.transAxes,
            va="top",
            ha="left",
            bbox={"facecolor": "white", "alpha": 0.85, "edgecolor": "0.7"},
        )
        cbar = self._figure.colorbar(scatter, ax=ax, fraction=0.05, pad=0.03)
        cbar.set_label("|A-B|")
        ax.tick_params(labelsize=8)

    def _plot_profile_panel(self):
        """Render the profile panel from the current ROI line, or a hint when no line exists yet."""
        ax_profile = self._profile_axes.get("profile")
        ax_diff = self._profile_axes.get("diff")
        if ax_profile is None or ax_diff is None:
            return
        ax_profile.cla()
        ax_diff.cla()
        ax_profile.set_title("Profile cross-section", fontsize=10)
        axis_unit = str(self._render_state.get("axis_unit") or "nm")
        value_unit = str(self._render_state.get("value_unit") or "nm")
        if self._profile_line is None:
            ax_profile.text(
                0.5,
                0.5,
                "Enable Profile mode, then drag a line on any image panel.",
                ha="center",
                va="center",
                transform=ax_profile.transAxes,
            )
            ax_profile.set_xticks([])
            ax_profile.set_yticks([])
            ax_diff.set_yticks([])
            return
        start_axis = np.asarray(self._profile_line["start"], dtype=float)
        end_axis = np.asarray(self._profile_line["end"], dtype=float)
        start_px = self._axis_to_pixel(start_axis[0], start_axis[1])
        end_px = self._axis_to_pixel(end_axis[0], end_axis[1])
        a_profile, _cols_a, _rows_a = _sample_line_profile(self._render_state["a_raw"], start_px, end_px)
        b_profile, _cols_b, _rows_b = _sample_line_profile(self._render_state["b_raw"], start_px, end_px)
        diff_profile, _cols_d, _rows_d = _sample_line_profile(self._render_state["diff"], start_px, end_px)
        total_distance = float(np.hypot(*(end_axis - start_axis)))
        distances = np.linspace(0.0, total_distance, a_profile.size, dtype=float)
        line_a, = ax_profile.plot(distances, a_profile, color="tab:blue", linewidth=1.6, label="A")
        line_b, = ax_profile.plot(
            distances,
            b_profile,
            color="tab:green",
            linewidth=1.6,
            linestyle="--",
            label="B",
        )
        line_d, = ax_diff.plot(distances, diff_profile, color="tab:red", linewidth=1.4, label="A-B")
        ax_profile.set_xlabel(f"Distance ({axis_unit})")
        ax_profile.set_ylabel(f"A / B ({value_unit})")
        ax_diff.set_ylabel(f"A-B ({value_unit})", color="tab:red")
        ax_diff.tick_params(axis="y", colors="tab:red")
        ax_profile.grid(True, alpha=0.25)
        legend_lines = [line_a, line_b, line_d]
        ax_profile.legend(legend_lines, [ln.get_label() for ln in legend_lines], loc="upper right")
        ax_profile.tick_params(labelsize=8)
        ax_diff.tick_params(labelsize=8)

    def _axis_role(self, ax):
        """Return which logical compare panel owns a matplotlib axis."""
        for role, target_ax in self._image_axes.items():
            if ax is target_ax:
                return role
        return None

    def _axis_to_pixel(self, x_val, y_val):
        """Convert axis coordinates into pixel coordinates on the shared compare grid."""
        extent = self._render_state.get("extent")
        shape = self._render_state.get("shape")
        if extent is None or not shape:
            raise ValueError("Comparison grid is unavailable.")
        return _axis_to_display_pixel(x_val, y_val, extent, shape, origin=self.IMAGE_ORIGIN)

    def _pixel_to_axis(self, point):
        """Convert compare-grid pixel coordinates back into displayed axis coordinates."""
        extent = self._render_state.get("extent")
        shape = self._render_state.get("shape")
        if extent is None or not shape:
            return None
        return _display_pixel_to_axis(point, extent, shape, origin=self.IMAGE_ORIGIN)

    def _point_in_bounds(self, point, shape):
        """Return True when a pixel-space landmark lies inside the target image."""
        if point is None or shape is None or len(shape) < 2:
            return False
        x_val = float(point[0])
        y_val = float(point[1])
        height = int(shape[0])
        width = int(shape[1])
        return 0.0 <= x_val <= float(max(0, width - 1)) and 0.0 <= y_val <= float(max(0, height - 1))
