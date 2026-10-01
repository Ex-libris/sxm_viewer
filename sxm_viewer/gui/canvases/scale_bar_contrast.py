"""Pick a black or white scale-bar color that contrasts with the image.

Qt-free (matplotlib/numpy only) so the preview canvas and the export
figure builders share one rule.  The image pixels actually under the scale
bar are pushed through the artist's own norm + colormap (``to_rgba``), so
the decision follows the live colormap, CLIM/histogram state, and the
display-time amber override - not the raw data values.

Contrast follows WCAG relative luminance: white wins when the median
luminance under the bar is below ~0.179 (the point where black and white
have equal contrast ratios), black otherwise.
"""
from __future__ import annotations

import numpy as np

LIGHT = '#f5f5f5'
DARK = '#111111'

# Luminance where contrast(L, white) == contrast(L, black):
# (1.05) / (L + 0.05) == (L + 0.05) / 0.05  ->  L ~= 0.179
_CROSSOVER_LUMINANCE = 0.179


def _relative_luminance(rgb):
    rgb = np.clip(np.asarray(rgb, dtype=float), 0.0, 1.0)
    lin = np.where(rgb <= 0.04045, rgb / 12.92, ((rgb + 0.055) / 1.055) ** 2.4)
    return lin[..., 0] * 0.2126 + lin[..., 1] * 0.7152 + lin[..., 2] * 0.0722


def estimate_region(ax, anchor, loc, bar_length, font_size_pt):
    """Approximate the scale bar's footprint as an axes-fraction box.

    Returns ``(x0, x1, y0, y1)``.  ``loc`` is the AnchoredSizeBar loc
    ('lower right' on the live preview, 'center' in export figures).
    Exact window extents need a renderer; an estimate is plenty for
    sampling the background.
    """
    try:
        xlim = ax.get_xlim()
        span = abs(float(xlim[1]) - float(xlim[0]))
        bar_frac = abs(float(bar_length)) / span if span > 0 else 0.2
    except Exception:
        bar_frac = 0.2
    bar_frac = min(max(bar_frac, 0.02), 1.0)
    try:
        dpi = float(ax.figure.dpi)
        ax_h = max(1.0, float(ax.bbox.height))
        ax_w = max(1.0, float(ax.bbox.width))
    except Exception:
        dpi, ax_h, ax_w = 100.0, 400.0, 400.0
    font_px = max(1.0, float(font_size_pt)) * dpi / 72.0
    # pad (0.4 font) on both sides + label line + sep + bar thickness.
    h_frac = (font_px * 0.8 + font_px * 1.25 + 3.0 + 0.01 * ax_h) / ax_h
    pad_w = font_px * 0.4 / ax_w
    ax_x, ax_y = float(anchor[0]), float(anchor[1])
    if loc == 'center':
        return (ax_x - bar_frac / 2 - pad_w, ax_x + bar_frac / 2 + pad_w,
                ax_y - h_frac / 2, ax_y + h_frac / 2)
    # 'lower right': anchor is the bottom-right corner, bar extends leftward.
    return (ax_x - bar_frac - 2 * pad_w, ax_x, ax_y, ax_y + h_frac)


def _primary_image(ax):
    images = [im for im in (getattr(ax, 'images', None) or []) if im.get_visible()]
    if not images:
        return None
    return images[0]


def _sample_rgba(ax, image, region):
    """RGBA (N, 4) floats of the image pixels inside the axes-fraction box."""
    arr = image.get_array()
    if arr is None:
        return None
    shape = np.shape(arr)
    if len(shape) < 2 or shape[0] < 1 or shape[1] < 1:
        return None
    rows, cols = int(shape[0]), int(shape[1])
    x0, x1, y0, y1 = region
    to_data = ax.transAxes + ax.transData.inverted()
    corners = to_data.transform([(x0, y0), (x1, y1)])
    left, right, bottom, top = image.get_extent()
    # Row 0 sits at the "top" extent edge for origin='upper', the "bottom"
    # one for origin='lower'; dividing by the signed span handles inverted
    # extents/axes either way.
    if image.origin == 'upper':
        y_row0, y_rowN = top, bottom
    else:
        y_row0, y_rowN = bottom, top
    x_span = right - left
    y_span = y_rowN - y_row0
    if x_span == 0 or y_span == 0:
        return None
    c = (corners[:, 0] - left) / x_span * cols
    r = (corners[:, 1] - y_row0) / y_span * rows
    c0, c1 = sorted(c)
    r0, r1 = sorted(r)
    c0, c1 = int(np.floor(max(c0, 0))), int(np.ceil(min(c1, cols)))
    r0, r1 = int(np.floor(max(r0, 0))), int(np.ceil(min(r1, rows)))
    if c1 <= c0 or r1 <= r0:
        return None  # scale bar sits entirely outside the image
    sub = arr[r0:r1, c0:c1]
    if np.ndim(sub) == 3:
        rgba = np.asarray(sub, dtype=float)
        if rgba.max(initial=0.0) > 1.0:
            rgba = rgba / 255.0
        if rgba.shape[-1] == 3:
            rgba = np.concatenate([rgba, np.ones(rgba.shape[:-1] + (1,))], axis=-1)
    else:
        rgba = np.asarray(image.to_rgba(sub), dtype=float)
    return rgba.reshape(-1, 4)


def contrast_color(ax, region, fallback):
    """Return LIGHT or DARK for the scale bar drawn over ``region``.

    ``fallback`` is returned when nothing image-like sits under the bar
    (no image, bar outside the image, sampling failed).
    """
    try:
        image = _primary_image(ax)
        if image is None:
            return fallback
        rgba = _sample_rgba(ax, image, region)
        if rgba is None or not len(rgba):
            return fallback
        # Composite transparent pixels (NaN/"bad" color) over the axes face.
        try:
            from matplotlib.colors import to_rgb
            face = np.asarray(to_rgb(ax.get_facecolor()), dtype=float)
        except Exception:
            face = np.ones(3)
        alpha = rgba[:, 3:4] * float(image.get_alpha() if image.get_alpha() is not None else 1.0)
        rgb = rgba[:, :3] * alpha + face * (1.0 - alpha)
        lum = float(np.median(_relative_luminance(rgb)))
    except Exception:
        return fallback
    if not np.isfinite(lum):
        return fallback
    return DARK if lum > _CROSSOVER_LUMINANCE else LIGHT
