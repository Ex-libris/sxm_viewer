"""Axes/image/colorbar sync helpers shared by MultiPreviewCanvas's full
(_redraw, fresh axes/image) and fast (_fast_update_single_view, reused
axes/image) update paths, so both apply the same matplotlib invalidation
rules instead of drifting independently.

Reusing an Axes/AxesImage/Colorbar across updates (the fast path) exposes a
few non-obvious matplotlib behaviors that a fresh axes (the full path,
recreated via fig.clf()/add_subplot()/imshow() every time) never runs into:

- `image.set_extent()` does NOT update `ax.dataLim` (that only happens
  automatically when an artist is first added via the public API). Left
  alone, `dataLim` keeps accumulating the bounding box of every extent ever
  set on a reused axes, so anything that consults it (aspect/autoscale/
  zoom-clamp logic) can end up using a stale, much larger historical bbox
  instead of the current image's real extent.
- With `aspect='equal'`, the axes' physical box (and any divider-attached
  colorbar) is only resized to match a new image's width:height ratio when
  `Axes.apply_aspect()` runs, which normally happens automatically inside a
  full draw. A reused axes can lag behind an extent change with a very
  different aspect ratio without an explicit nudge.
- `ax.set_xticks(ax.get_xticks())` is a trap: the first call freezes the
  axis onto a `FixedLocator` using whatever tick positions were valid at
  that moment. Every later `get_xticks()` call then just replays those same
  frozen values regardless of the current view, and re-applying them via
  `set_xticks()` drags `xlim`/`ylim` back out to whatever range the very
  first image on that axes needed.
"""
from __future__ import annotations

from matplotlib.ticker import AutoLocator

from ..._shared import np
from ..plot_typography import apply_text_style


def sync_axes_to_view(ax, image, view, *, flip, origin, display_extent, title, show_title,
                       font_family, plot_style_kwargs, show_ticks):
    """Configure `image`/`ax` in place to display `view`.

    `display_extent` must already be resolved by the caller (e.g. via
    `MultiPreviewCanvas._display_extent_for_view`), since that resolution
    depends on instance-level relative-axes state this function doesn't own.

    Returns the image-meta dict `{"extent", "origin", "shape"}` for the
    caller to store in its own bookkeeping (e.g. `self._image_meta[ax]`).
    """
    arr = np.asarray(view.get("arr"))
    arr_plot = np.flipud(arr) if flip else arr

    image.set_data(arr_plot)
    image.set_cmap(view.get("cmap", "viridis"))

    if display_extent is not None:
        image.set_extent(display_extent)
        try:
            x0, x1, y0, y1 = image.get_extent()
            ax.set_xlim(x0, x1)
            ax.set_ylim(y0, y1)
            ax.dataLim.set_points(np.array([
                [min(x0, x1), min(y0, y1)],
                [max(x0, x1), max(y0, y1)],
            ]))
        except Exception:
            pass
    else:
        try:
            ax.set_xlim(-0.5, max(arr_plot.shape[1] - 0.5, 0.5))
            if origin == "upper":
                ax.set_ylim(max(arr_plot.shape[0] - 0.5, 0.5), -0.5)
            else:
                ax.set_ylim(-0.5, max(arr_plot.shape[0] - 0.5, 0.5))
        except Exception:
            pass

    try:
        ax.apply_aspect()
    except Exception:
        pass

    clim = view.get("clim")
    if clim:
        try:
            image.set_clim(*clim)
        except Exception:
            pass
    else:
        try:
            image.autoscale()
        except Exception:
            pass

    if title and show_title:
        ax.set_title(title, fontsize=9)
        apply_text_style(ax.title, family=font_family, **plot_style_kwargs)
    else:
        ax.set_title("")

    ax.tick_params(labelsize=8)
    if not show_ticks:
        ax.set_xticks([])
        ax.set_yticks([])
    else:
        try:
            ax.xaxis.set_major_locator(AutoLocator())
            ax.yaxis.set_major_locator(AutoLocator())
        except Exception:
            pass
    for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
        apply_text_style(lbl, family=font_family, **plot_style_kwargs)

    try:
        extent = image.get_extent()
    except Exception:
        extent = display_extent
    return {"extent": extent, "origin": origin, "shape": arr_plot.shape}


def style_colorbar(cbar, image, label, *, font_family, plot_style_kwargs, show_ticks):
    """Refresh an existing colorbar's normalization, label, ticks, and font styling."""
    cbar.update_normal(image)
    cbar.set_label(label)
    if not show_ticks:
        cbar.set_ticks([])
    try:
        apply_text_style(cbar.ax.xaxis.label, family=font_family, **plot_style_kwargs)
        apply_text_style(cbar.ax.yaxis.label, family=font_family, **plot_style_kwargs)
        for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
            apply_text_style(lbl, family=font_family, **plot_style_kwargs)
    except Exception:
        pass
