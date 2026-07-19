"""Spectroscopy position <-> image raster mapping (Qt-free).

Extracted from ``SXMGridViewer`` so the most bug-prone geometry in the
codebase is a set of pure functions that can be unit-tested without a
display. See CLAUDE.md "Spectroscopy position mapping" for the full bug
history; the load-bearing rules are restated at the functions that own
them so they cannot drift away from the code again.

**The invariant that matters most** (``map_spec_to_pixels``): rotate the
spec's real-nm offset by the header's scan angle FIRST, then normalize by
width/height SEPARATELY per axis. Doing it the other way round - normalize
each axis by its own span, then rotate - mixes non-uniformly-scaled
components inside the rotation matrix and **shears** instead of rotating.
That is invisible on square scans, which is why it shipped, and only
showed up as a smeared point cloud on an elongated (2.5 x 7.5 nm) scan.
"""
from __future__ import annotations

import math

# Extent convention used throughout: [x0, x1, y1, y0] in nm, matching
# matplotlib's imshow(extent=...) for an origin='upper' image, i.e. y1 is
# the *bottom* edge value and y0 the top.
UNIT_EXTENT = [0.0, 1.0, 0.0, 1.0]

_X_RANGE_KEYS = ("XScanRange", "XRange", "ScanRange")
_Y_RANGE_KEYS = ("YScanRange", "YRange", "ScanRange")
_X_CENTER_KEYS = ("xCenter", "XCenter", "XOffset", "OffsetX", "XPosition", "XPos")
_Y_CENTER_KEYS = ("yCenter", "YCenter", "YOffset", "OffsetY", "YPosition", "YPos")
_ANGLE_KEYS = ("Angle", "ScanAngle", "scan_angle", "Scan_Angle")

# Fallback placement for specs with no coordinates: a 3x3 spread so
# several markers on one image stay distinguishable.
_FALLBACK_SLOTS = (
    (0.15, 0.15), (0.50, 0.15), (0.85, 0.15),
    (0.15, 0.50), (0.50, 0.50), (0.85, 0.50),
    (0.15, 0.85), (0.50, 0.85), (0.85, 0.85),
)


def _first_float(mapping, keys, default=0.0):
    for key in keys:
        if key in mapping and mapping.get(key) not in (None, ""):
            try:
                return float(mapping.get(key))
            except (TypeError, ValueError):
                continue
    return default


def header_extent(header):
    """``[x0, x1, y1, y0]`` in nm for a scan header.

    Falls back to a unit square when the header carries no usable scan
    range, which callers treat as "degenerate extent" and route into the
    cloud-bounds fallback.
    """
    if not header:
        return list(UNIT_EXTENT)
    try:
        x_range = _first_float(header, _X_RANGE_KEYS)
        y_range = _first_float(header, _Y_RANGE_KEYS)
        if x_range == 0.0 or y_range == 0.0:
            return list(UNIT_EXTENT)
        x_center = _first_float(header, _X_CENTER_KEYS)
        y_center = _first_float(header, _Y_CENTER_KEYS)
        if x_center == 0.0 and y_center == 0.0:
            # No center recorded - assume the scan starts at the origin.
            x_center = 0.5 * x_range
            y_center = 0.5 * y_range
        x0 = x_center - 0.5 * x_range
        x1 = x_center + 0.5 * x_range
        y0 = y_center - 0.5 * y_range
        y1 = y_center + 0.5 * y_range
        return [x0, x1, y1, y0]
    except Exception:
        return list(UNIT_EXTENT)


def header_scan_angle(header):
    """Scan angle in degrees, 0.0 when absent/unparseable."""
    if not header:
        return 0.0
    for key in _ANGLE_KEYS:
        if key in header and header.get(key) not in (None, ""):
            try:
                return float(header.get(key))
            except (TypeError, ValueError):
                continue
    return 0.0


def fallback_spec_coords(idx, xpix, ypix):
    """Placement for a spec with no coordinates, spread over a 3x3 grid."""
    try:
        idx = int(idx)
    except (TypeError, ValueError):
        idx = 1
    frac_x, frac_y = _FALLBACK_SLOTS[(idx - 1) % len(_FALLBACK_SLOTS)]
    return frac_x * max(1, int(xpix) - 1), frac_y * max(1, int(ypix) - 1)


def apply_thumb_crop(col, row, thumb_crop):
    """Shift a raster row into a cropped thumbnail's own row space.

    ``detect_valid_scan_region`` trims blank/aborted rows off a thumbnail
    before display, so anything drawing onto the *rendered* thumbnail (not
    the header's raster) must subtract the crop offset. Only rows are
    cropped, never columns - callers rescaling into the final pixmap need
    separate x/y scale factors for the same reason.
    """
    if not thumb_crop:
        return col, row
    try:
        r0 = int(thumb_crop.get("r0"))
        r1 = int(thumb_crop.get("r1"))
    except (TypeError, ValueError, AttributeError):
        return col, row
    if r1 <= r0:
        return col, row
    try:
        row = float(row) - float(r0)
    except (TypeError, ValueError):
        return col, row
    crop_rows = r1 - r0 + 1
    return col, min(max(row, 0.0), max(0.0, crop_rows - 1))


def map_spec_by_grid(spec, xpix, ypix):
    """Position from the spec's own grid indices, ignoring nm coordinates."""
    grid_cols = spec.get("grid_cols")
    grid_rows = spec.get("grid_rows")
    if not grid_cols or not grid_rows:
        return None
    try:
        col_idx = int(spec.get("grid_col", 0))
        row_idx = int(spec.get("grid_row", 0))
        grid_cols = int(grid_cols)
        grid_rows = int(grid_rows)
    except (TypeError, ValueError):
        return None
    if grid_cols <= 0 or grid_rows <= 0:
        return None
    cols = max(1, grid_cols - 1)
    rows = max(1, grid_rows - 1)
    col = (col_idx / cols if cols > 0 else 0.0) * max(1, int(xpix) - 1)
    row = (row_idx / rows if rows > 0 else 0.0) * max(1, int(ypix) - 1)
    return col, row


def map_spec_by_cloud_bounds(spec, bounds, xpix, ypix):
    """Position within the bounding box of all specs on this image.

    Only meaningful when the image's own header extent is degenerate. For
    a genuinely off-frame spec on a *valid* image this is actively wrong -
    it stretches the box to include the outlier and yields a
    plausible-looking interior position indistinguishable from an in-frame
    point. ``map_spec_to_pixels`` gates on that; see its docstring.
    """
    if not bounds:
        return None
    try:
        xmin, xmax, ymin, ymax = (float(v) for v in bounds)
        x = float(spec.get("x"))
        y = float(spec.get("y"))
    except (TypeError, ValueError):
        return None
    span_x = (xmax - xmin) or 1.0
    span_y = (ymax - ymin) or 1.0
    frac_x = min(max((x - xmin) / span_x, 0.0), 1.0)
    frac_y = min(max((ymax - y) / span_y, 0.0), 1.0)
    return (frac_x * max(1, int(xpix) - 1),
            frac_y * max(1, int(ypix) - 1))


def _resolve_bounds(cloud_bounds):
    """``cloud_bounds`` may be a tuple or a zero-arg callable.

    The callable form keeps the (potentially expensive) spec-cloud bounding
    box lazy: it is only needed on fallback paths, and computing it eagerly
    for every marker was an O(n^2) scan on large grids.
    """
    if cloud_bounds is None:
        return None
    if callable(cloud_bounds):
        try:
            return cloud_bounds()
        except Exception:
            return None
    return cloud_bounds


def map_spec_to_pixels(spec, extent, angle_deg, xpix, ypix,
                       thumb_crop=None, cloud_bounds=None):
    """Map a spec's absolute (x, y) nm to fractional (col, row) pixels.

    ``extent`` is ``[x0, x1, y1, y0]`` (see ``header_extent``);
    ``angle_deg`` the image's scan angle; ``cloud_bounds`` an optional
    ``(xmin, xmax, ymin, ymax)`` tuple *or callable* used only on fallback
    paths.

    Rotation order is load-bearing - see the module docstring.

    **Off-frame handling.** A spec whose off-frame status was already
    evaluated at assignment time is clamped to the true nearest edge rather
    than remapped through ``map_spec_by_cloud_bounds``. The gate checks key
    **presence**, not truthiness: ``off_frame_direction`` is legitimately
    ``None`` for a point sitting exactly on the edge (a normal acquisition
    pattern - a grid built to exactly cover its anchor image), and treating
    that identically to "never checked" sent an entire boundary row through
    the cloud fallback, producing a visibly warped edge. Only fall back
    when the key is absent entirely.
    """
    # No coordinates at all: deterministic spread placement.
    try:
        x = float(spec.get("x"))
        y = float(spec.get("y"))
    except (TypeError, ValueError):
        try:
            idx = int(spec.get("order_idx", 1))
        except (TypeError, ValueError):
            idx = 1
        return fallback_spec_coords(idx, xpix, ypix)

    try:
        x0, x1, y1, y0 = (float(v) for v in extent)
    except (TypeError, ValueError):
        x0, x1, y1, y0 = UNIT_EXTENT
    xspan = x1 - x0
    yspan = y1 - y0

    if xspan <= 0 or yspan <= 0:
        # Degenerate header extent - routine for grid files.
        pt = map_spec_by_cloud_bounds(spec, _resolve_bounds(cloud_bounds),
                                      xpix, ypix)
        if pt is None:
            pt = map_spec_by_grid(spec, xpix, ypix)
        if pt is None:
            return None
        return apply_thumb_crop(pt[0], pt[1], thumb_crop)

    dx = x - 0.5 * (x0 + x1)
    dy = y - 0.5 * (y0 + y1)
    try:
        angle_deg = float(angle_deg or 0.0)
    except (TypeError, ValueError):
        angle_deg = 0.0
    if angle_deg:
        theta = math.radians(angle_deg)
        cos_t, sin_t = math.cos(theta), math.sin(theta)
        # Rotate in real nm (isotropic) BEFORE normalizing per axis.
        u_nm = dx * cos_t - dy * sin_t
        v_nm = dx * sin_t + dy * cos_t
    else:
        u_nm, v_nm = dx, dy

    frac_x = (u_nm / xspan) + 0.5
    frac_y = 0.5 - (v_nm / yspan)

    if not (0.0 <= frac_x <= 1.0 and 0.0 <= frac_y <= 1.0):
        if "off_frame_direction" not in spec:
            pt = map_spec_by_cloud_bounds(spec, _resolve_bounds(cloud_bounds),
                                          xpix, ypix)
            if pt is None:
                pt = map_spec_by_grid(spec, xpix, ypix)
            if pt is not None:
                return apply_thumb_crop(pt[0], pt[1], thumb_crop)
        frac_x = min(max(frac_x, 0.0), 1.0)
        frac_y = min(max(frac_y, 0.0), 1.0)

    col = frac_x * max(1, int(xpix) - 1)
    row = frac_y * max(1, int(ypix) - 1)
    return apply_thumb_crop(col, row, thumb_crop)


# NOTE: `_matrix_bbox_pixels` deliberately stays on the GUI side. It
# returns a QRectF and applies Qt-specific presentation rules (a minimum
# badge size for degenerate boxes, clamping to the drawn pixmap), so it is
# not pure geometry. It calls map_spec_to_pixels above for the actual
# coordinate work.
