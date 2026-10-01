"""Keep text drawn *inside* a colorbar (the in-bar label, publication endpoint
values) inside the bar, at every font scale, preset and window size.

Two independent mechanisms, because there are two ways text can escape:

1. **Thickness tracks the font.** A divider-attached colorbar used to be a
   pure percentage of the image ("4%"/"5%"), so Ctrl+wheel font growth made
   the in-bar label taller than the bar itself and it spilled over the tick
   labels. `append_colorbar_axes` sizes the bar as a small percentage of the
   image *plus* a fixed physical size derived from the font, and keeps a
   handle to that fixed part so `set_colorbar_thickness` can update it in
   place on a font-scale change without rebuilding the figure.
2. **Length is fitted after layout.** A long label ("Z (Forward) [pm]") can
   still be longer than a short bar in a small popup. `fit_colorbar_texts`
   measures the in-bar texts with the real renderer against the bar's real
   located box and shrinks them (uniformly, never growing past the requested
   size) until they fit. Call it after anything that changes layout: font
   scale, tight_layout, resize, label text change.

Qt-free; works on any matplotlib Figure with an Agg-capable canvas.
"""
from __future__ import annotations

from mpl_toolkits.axes_grid1 import axes_size

# Bar thickness = _REL_FRACTION of the parent image + font height * _FONT_FACTOR.
_REL_FRACTION = 0.015
_FONT_FACTOR = 1.45
# Fraction of the bar that text may occupy (leaves a visible margin).
_LEN_FILL = 0.94
_THICK_FILL = 0.88
_MIN_FONT_PT = 4.0


def _thickness_inches(font_pt):
    return max(0.08, float(font_pt) * _FONT_FACTOR / 72.0)


def append_colorbar_axes(divider, ax, orientation, font_pt, pad):
    """Append a colorbar axes whose thickness grows with `font_pt`."""
    horizontal = str(orientation).lower() == "horizontal"
    ref = axes_size.AxesY(ax) if horizontal else axes_size.AxesX(ax)
    fixed = axes_size.Fixed(_thickness_inches(font_pt))
    size = axes_size.Fraction(_REL_FRACTION, ref) + fixed
    cax = divider.append_axes("bottom" if horizontal else "right", size=size, pad=pad)
    cax._sxm_thickness_fixed = fixed
    return cax


def set_colorbar_thickness(cbar, font_pt):
    """Update a bar created by `append_colorbar_axes` for a new font size."""
    fixed = getattr(getattr(cbar, "ax", None), "_sxm_thickness_fixed", None)
    if fixed is not None:
        fixed.fixed_size = _thickness_inches(font_pt)


def _renderer(fig):
    try:
        return fig.canvas.get_renderer()
    except Exception:
        pass
    try:
        return fig._get_renderer()
    except Exception:
        return None


def _inner_texts(cbar):
    """(text, requested_pt) pairs for everything drawn inside the bar."""
    pub_label = getattr(cbar, "_publication_label_artist", None)
    endpoints = list(getattr(cbar, "_publication_endpoint_artists", None) or [])
    if pub_label is not None or endpoints:
        items = [] if pub_label is None else [pub_label]
        items += endpoints
    else:
        vertical = str(getattr(cbar, "orientation", "vertical")).lower() == "vertical"
        label = cbar.ax.yaxis.label if vertical else cbar.ax.xaxis.label
        # Only an axis label moved *into* the bar (set_label_coords(0.5, 0.5))
        # is constrained; a conventional outside label is left alone.
        try:
            inside = label.get_transform() == cbar.ax.transAxes and all(
                0.0 < float(c) < 1.0 for c in label.get_position()
            )
        except Exception:
            inside = False
        items = [label] if inside else []
        if inside:
            # Colorbar re-applies YAxis.set_label_position('right') on every
            # update, which forces va='top' (rotation_mode 'anchor'): the
            # "centered" label then hangs off the bar's midline sideways onto
            # the tick labels. Re-center it on both axes every fit.
            label.set_horizontalalignment("center")
            label.set_verticalalignment("center")
    out = []
    for text in items:
        if text is None or not text.get_visible() or not str(text.get_text() or "").strip():
            continue
        requested = getattr(text, "_sxm_requested_pt", None)
        if requested is None:
            requested = float(text.get_fontsize())
        out.append((text, float(requested)))
    return out


def request_font_size(text, size_pt):
    """Set a text's *requested* size; `fit_colorbar_texts` may render it smaller."""
    if text is None:
        return
    text._sxm_requested_pt = float(size_pt)
    text.set_fontsize(float(size_pt))


def fit_colorbar_texts(cbar, renderer=None):
    """Shrink in-bar texts of `cbar` so none of them leaves the bar."""
    if cbar is None or getattr(cbar, "ax", None) is None:
        return
    cax = cbar.ax
    fig = cax.figure
    renderer = renderer or _renderer(fig)
    if renderer is None:
        return
    texts = _inner_texts(cbar)
    if not texts:
        return
    try:
        locator = cax.get_axes_locator()
        if locator is not None:
            cax.apply_aspect(locator(cax, renderer))
        bar = cax.get_window_extent(renderer)
    except Exception:
        return
    vertical = str(getattr(cbar, "orientation", "vertical")).lower() == "vertical"
    bar_len = bar.height if vertical else bar.width
    bar_thick = bar.width if vertical else bar.height
    if bar_len <= 1 or bar_thick <= 1:
        return
    lengths, thicks = [], []
    for text, requested in texts:
        text.set_fontsize(requested)
        try:
            ext = text.get_window_extent(renderer)
        except Exception:
            return
        lengths.append(ext.height if vertical else ext.width)
        thicks.append(ext.width if vertical else ext.height)
    # Texts sit side by side along the bar (low | label | high); budget a
    # half-em gap between neighbours so they never touch.
    gap = 0.5 * max(t[1] for t in texts) * fig.dpi / 72.0
    total_len = sum(lengths) + gap * max(0, len(texts) - 1)
    ratio = 1.0
    if total_len > 0:
        ratio = min(ratio, bar_len * _LEN_FILL / total_len)
    if max(thicks) > 0:
        ratio = min(ratio, bar_thick * _THICK_FILL / max(thicks))
    if ratio >= 0.999:
        return
    for text, requested in texts:
        text.set_fontsize(max(_MIN_FONT_PT, requested * ratio))


def fit_all_colorbar_texts(cbars, fig=None):
    cbars = [c for c in (cbars or []) if c is not None]
    if not cbars:
        return
    renderer = _renderer(fig if fig is not None else cbars[0].ax.figure)
    for cbar in cbars:
        try:
            fit_colorbar_texts(cbar, renderer)
        except Exception:
            continue
