"""Per-image quarter-turn display orientation (Qt-free).

An image's orientation is ``k`` clockwise quarter turns (0..3) applied at
display time: the raw channel data, headers and every spectroscopy
position stay in the acquisition raster frame, and only what is drawn is
turned. Anything that places raster-frame coordinates onto a turned image
(spectroscopy markers, thumbnail hit-testing) must push them through the
same turn - see CLAUDE.md "Coordinate frames & orientation".

Coordinate conventions (row 0 = top, as displayed with origin='upper'):

* *Index* coordinates (``rotate_index_coords``) are pixel centres in
  ``[0, w-1] x [0, h-1]`` - what ``spec_mapping.map_spec_to_pixels``
  returns.
* *Continuous* coordinates (``rotate_continuous_coords``) are positions in
  a ``w x h`` box such as a thumbnail pixmap, in ``[0, w] x [0, h]``.
"""
from __future__ import annotations

import numpy as np


def normalize_turns(k) -> int:
    """Clamp any integer-ish turn count to 0..3 (clockwise quarter turns)."""
    try:
        return int(k) % 4
    except Exception:
        return 0


def turned_shape(w, h, k):
    """``(w, h)`` of a ``w x h`` image after ``k`` clockwise quarter turns."""
    return (h, w) if normalize_turns(k) % 2 else (w, h)


def rotate_array(arr, k):
    """Turn a 2D array (row 0 = top) clockwise by ``k`` quarter turns."""
    k = normalize_turns(k)
    if not k:
        return arr
    return np.ascontiguousarray(np.rot90(np.asarray(arr), -k))


def rotate_extent(extent, k):
    """Extent of the turned image, keeping its centre and axis directions.

    ``extent`` is ``[x0, x1, y1, y0]`` (any orientation of each pair). For
    odd ``k`` the x and y spans swap, so ticks, scale bar and colorbar of
    the turned image describe its turned dimensions; the per-axis sign of
    each pair is kept so the extent convention callers rely on is intact.
    """
    k = normalize_turns(k)
    if extent is None or not k % 2:
        return extent
    try:
        a, b, c, d = (float(v) for v in extent)
    except Exception:
        return extent
    cx = 0.5 * (a + b)
    cy = 0.5 * (c + d)
    span_x = b - a
    span_y = d - c
    half_x = 0.5 * abs(span_y) * (1.0 if span_x >= 0 else -1.0)
    half_y = 0.5 * abs(span_x) * (1.0 if span_y >= 0 else -1.0)
    out = [cx - half_x, cx + half_x, cy - half_y, cy + half_y]
    return type(extent)(out) if isinstance(extent, (list, tuple)) else out


def rotate_index_coords(col, row, w, h, k):
    """Map a pixel-centre ``(col, row)`` of a ``w x h`` raster through ``k`` turns."""
    k = normalize_turns(k)
    wm = max(0.0, float(w) - 1.0)
    hm = max(0.0, float(h) - 1.0)
    return _turn(float(col), float(row), wm, hm, k)


def rotate_continuous_coords(x, y, w, h, k):
    """Map a position in a ``w x h`` box through ``k`` clockwise turns."""
    return _turn(float(x), float(y), float(w), float(h), normalize_turns(k))


def unrotate_continuous_coords(x, y, w, h, k):
    """Inverse of ``rotate_continuous_coords``; ``w x h`` is the *unturned* box."""
    k = normalize_turns(k)
    if not k:
        return float(x), float(y)
    tw, th = turned_shape(float(w), float(h), k)
    return _turn(float(x), float(y), tw, th, (4 - k) % 4)


def _turn(x, y, w, h, k):
    if k == 1:
        return h - y, x
    if k == 2:
        return w - x, h - y
    if k == 3:
        return y, w - x
    return x, y
