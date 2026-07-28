"""Export-quality figure rendering for the preview canvas.

Builds a **throwaway** matplotlib Figure for saving/printing/copying - it
never touches the live canvas's axes, images or colorbars. That
separation is what made this safe to lift out of the 12k-line
`detail_preview_canvas.py`, and it must be preserved.

These functions deliberately re-derive their own axis configuration,
because an export figure has different requirements from the interactive
one (fixed DPI, no interactive artists, its own colorbar styling). The
live render path shares its setup through `preview_axes_sync.py` instead.
CLAUDE.md records three separate bugs caused by a second hand-rolled copy
of that setup drifting out of sync - so if you change extent/aspect/tick
handling here, check whether `preview_axes_sync` needs the same change.

Module functions taking the canvas as `canvas`, matching the convention
used elsewhere in `gui/canvases/`.
"""
from __future__ import annotations

import math

import numpy as np
from matplotlib.figure import Figure
from mpl_toolkits.axes_grid1 import make_axes_locatable
from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar

from ... import cmap_registry
from ..plot_typography import apply_text_style


def render_view_figure(canvas, view):
    # Known related risk: this (and _render_views_grid below) builds a
    # standalone export/print figure and independently re-derives extent/
    # clim/title setup rather than calling preview_axes_sync.sync_axes_to_view,
    # since it targets a fresh, throwaway Figure/Axes rather than the
    # interactive canvas's reused ones. If a display bug like the ones
    # documented in preview_axes_sync.py's module docstring resurfaces in
    # exported/printed images, check here for the same class of drift.
    fig = Figure(figsize=(6, 6))
    ax = fig.add_subplot(1, 1, 1)
    arr = np.asarray(view.get('arr'))
    flip = canvas._use_relative_axes(view)
    if flip:
        arr_plot = np.flipud(arr)
    else:
        arr_plot = arr
    raw_extent = view.get('extent_raw')
    if raw_extent is None:
        raw_extent = view.get('extent')
    cmap = view.get('cmap', 'viridis')
    origin = 'lower' if flip else 'upper'
    display_extent = canvas._display_extent_for_view(view, raw_extent)
    if display_extent is None:
        im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
    else:
        im = ax.imshow(
            arr_plot,
            extent=display_extent,
            origin=origin,
            interpolation='nearest',
            aspect='equal',
            cmap=cmap,
        )
    # Ensure axes limits reflect the current extent (important when toggling relative axes)
    try:
        ext = display_extent if display_extent is not None else im.get_extent()
        if ext is not None:
            x0, x1, y1, y0 = ext
            ax.set_xlim(x0, x1)
            if flip:
                ax.set_ylim(y0, y1)
            else:
                ax.set_ylim(y1, y0)
    except Exception:
        pass
    ax.set_autoscale_on(False)
    cbar_label = view.get('colorbar_label') or view.get('unit', '')
    cbar = None
    if cbar_label and canvas._show_colorbar:
        try:
            divider = make_axes_locatable(ax)
            if canvas._colorbar_orientation == 'horizontal':
                cax = divider.append_axes("bottom", size="5%", pad=0.08)
                cbar = fig.colorbar(im, cax=cax, orientation='horizontal')
                cbar.set_label(cbar_label)
                cbar.ax.xaxis.set_label_coords(0.5, 0.5)
                cbar.ax.xaxis.label.set_horizontalalignment('center')
                cbar.ax.xaxis.label.set_verticalalignment('center')
            else:
                cax = divider.append_axes("right", size="4%", pad=0.02)
                cbar = fig.colorbar(im, cax=cax, orientation='vertical')
                cbar.set_label(cbar_label)
                cbar.ax.yaxis.set_label_coords(0.5, 0.5)
                cbar.ax.yaxis.label.set_horizontalalignment('center')
                cbar.ax.yaxis.label.set_verticalalignment('center')
        except Exception:
            cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02, orientation=canvas._colorbar_orientation)
            cbar.set_label(cbar_label)
        if not canvas._show_ticks:
            cbar.set_ticks([])
    try:
        canvas._draw_outlines(ax, view)
    except Exception:
        pass
    title = view.get('title', '')
    if title and canvas._show_title:
        ax.set_title(title, fontsize=9)
        apply_text_style(ax.title, family=canvas._font_family, **canvas._plot_style_state())
    canvas._draw_acquisition_overlay(ax, view)
    ax.tick_params(labelsize=8)
    for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
        apply_text_style(lbl, family=canvas._font_family, **canvas._plot_style_state())

    if canvas.scale_bar_enabled:
        extent_for_scale = display_extent if display_extent is not None else raw_extent
        if extent_for_scale is None:
            h, w = np.shape(view['arr'])
            width = w
            unit = 'px'
        else:
            width = abs(extent_for_scale[1] - extent_for_scale[0])
            unit = view.get('axis_unit') or 'nm'
        
        size, label = canvas._calculate_best_scale_bar(width, unit)
        # Hide unit text if blank to avoid default "nm" showing up when unset
        label = label if label and label.strip() else None
        font_scale = getattr(canvas, '_view_font_scale', 1.0)
        
        dark = bool(canvas._detail_dark)
        default_color = '#f5f5f5' if dark else '#111111'
        sb_settings = getattr(canvas, '_scale_bar_settings', {})
        sb_text_col = sb_settings.get('text_color') or default_color
        sb_bar_col = sb_settings.get('bar_color') or default_color
        font_family = sb_settings.get('font_family', 'sans-serif')

        sb = AnchoredSizeBar(ax.transData, size, label, loc='center',
                             pad=0.4, borderpad=0, sep=3, frameon=False,
                             size_vertical=width*0.004*font_scale, color=sb_bar_col,
                             label_top=True,
                             bbox_to_anchor=canvas._scale_bar_pos, bbox_transform=ax.transAxes)
        sb.size_bar.get_children()[0].set_linewidth(0)
        text = sb.txt_label.get_children()[0]
        text.set_color(sb_text_col)
        text.set_fontfamily(font_family)
        text.set_fontsize(10 * font_scale)
        text.set_fontweight('bold')
        ax.add_artist(sb)

    canvas._draw_image_size_overlay(ax, view)

    if not canvas._show_ticks:
        ax.set_xticks([])
        ax.set_yticks([])

    canvas._draw_molecules(ax)

    style_export_figure(canvas, fig, ax, cbar)
    try:
        fig.tight_layout()
    except Exception:
        pass
    return fig

def render_views_grid(canvas, views):
    """Render multiple views into a single figure grid.

    See the "known related risk" note on _render_view_figure above — this
    independently re-derives the same kind of axis setup as
    preview_axes_sync.sync_axes_to_view for a throwaway export figure.
    """
    views = views or []
    total = len(views)
    if total == 0:
        return render_view_figure(canvas, {})
    cols = int(math.ceil(math.sqrt(total)))
    rows = int(math.ceil(total / cols))
    fig = Figure(figsize=(6 * cols, 6 * rows))
    dark = bool(canvas._detail_dark)
    fig_face = '#111217' if dark else '#ffffff'
    fig.set_facecolor(fig_face)
    text_color = '#f5f5f5' if dark else '#111111'
    font_scale = getattr(canvas, '_view_font_scale', 1.0)
    for i, view in enumerate(views, 1):
        ax = fig.add_subplot(rows, cols, i)
        arr = np.asarray(view.get('arr'))
        flip = canvas._use_relative_axes(view)
        arr_plot = np.flipud(arr) if flip else arr
        raw_extent = view.get('extent_raw')
        if raw_extent is None:
            raw_extent = view.get('extent')
        cmap = view.get('cmap', 'viridis')
        origin = 'lower' if flip else 'upper'
        display_extent = canvas._display_extent_for_view(view, raw_extent)
        if display_extent is None:
            im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmap)
        else:
            im = ax.imshow(
                arr_plot,
                extent=display_extent,
                origin=origin,
                interpolation='nearest',
                aspect='equal',
                cmap=cmap,
            )
        try:
            ext = display_extent if display_extent is not None else im.get_extent()
            if ext is not None:
                x0, x1, y1, y0 = ext
                ax.set_xlim(x0, x1)
                if flip:
                    ax.set_ylim(y0, y1)
                else:
                    ax.set_ylim(y1, y0)
        except Exception:
            pass
        # record base limits for reset before any restore
        try:
            canvas._zoom_reset_limits[ax] = (ax.get_xlim(), ax.get_ylim())
        except Exception:
            pass
        ax.set_autoscale_on(False)
        if not canvas._show_ticks:
            ax.set_xticks([])
            ax.set_yticks([])
        ax.tick_params(labelsize=8 * font_scale, colors=text_color, labelcolor=text_color)
        for spine in ax.spines.values():
            spine.set_color(text_color)
        cbar_label = view.get('colorbar_label') or view.get('unit', '')
        if cbar_label and canvas._show_colorbar:
            try:
                divider = make_axes_locatable(ax)
                cax = divider.append_axes("right", size="5%", pad=0.05)
                cbar = fig.colorbar(im, cax=cax, orientation='vertical')
                cbar.set_label(cbar_label, size=10 * font_scale)
                cbar.ax.yaxis.label.set_color(text_color)
                cbar.ax.tick_params(colors=text_color, labelcolor=text_color, labelsize=8 * font_scale)
                if not canvas._show_ticks:
                    cbar.set_ticks([])
                cbar.outline.set_edgecolor(text_color)
                apply_text_style(cbar.ax.yaxis.label, family=canvas._font_family, **canvas._plot_style_state())
                for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
                    apply_text_style(lbl, family=canvas._font_family, **canvas._plot_style_state())
            except Exception:
                pass
        try:
            canvas._draw_outlines(ax, view)
        except Exception:
            pass
        title = canvas._compose_view_title(view) or view.get('label', '')
        if title and canvas._show_title:
            ax.set_title(title, fontsize=9 * font_scale, color=text_color)
            apply_text_style(ax.title, family=canvas._font_family, **canvas._plot_style_state())
        canvas._draw_acquisition_overlay(ax, view)
        canvas._draw_filter_summary_overlay(ax, view)
        canvas._draw_image_size_overlay(ax, view)
        for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
            apply_text_style(lbl, family=canvas._font_family, **canvas._plot_style_state())
    fig.tight_layout()
    return fig

def style_export_figure(canvas, fig, ax, cbar):
    dark = bool(canvas._detail_dark)
    fig_face = '#111217' if dark else '#ffffff'
    ax_face = '#14161c' if dark else '#ffffff'
    text_color = '#f5f5f5' if dark else '#111111'
    grid_color = '#4f5a64' if dark else '#9a9a9a'
    try:
        fig.set_facecolor(fig_face)
    except Exception:
        pass
    try:
        ax.set_facecolor(ax_face)
        ax.tick_params(colors=text_color, labelcolor=text_color)
        ax.xaxis.label.set_color(text_color)
        ax.yaxis.label.set_color(text_color)
        for spine in ax.spines.values():
            spine.set_color(text_color)
        if canvas._detail_grid:
            ax.grid(True, color=grid_color, alpha=0.3, linewidth=0.6)
        else:
            ax.grid(False)
    except Exception:
        pass
    if cbar is not None:
        try:
            cbar.ax.tick_params(colors=text_color, labelcolor=text_color)
            cbar.ax.yaxis.label.set_color(text_color)
            cbar.ax.xaxis.label.set_color(text_color)
            cbar.outline.set_edgecolor(text_color)
        except Exception:
            pass
    scale = max(0.6, min(2.5, getattr(canvas, '_view_font_scale', 1.0)))
    tick_size = 8 * scale
    label_size = 10 * scale
    title_size = 9 * scale
    try:
        ax.tick_params(labelsize=tick_size)
        ax.xaxis.label.set_fontsize(label_size)
        ax.yaxis.label.set_fontsize(label_size)
        ax.title.set_fontsize(title_size)
        apply_text_style(ax.xaxis.label, family=canvas._font_family, **canvas._plot_style_state())
        apply_text_style(ax.yaxis.label, family=canvas._font_family, **canvas._plot_style_state())
        apply_text_style(ax.title, family=canvas._font_family, **canvas._plot_style_state())
        for lbl in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
            apply_text_style(lbl, family=canvas._font_family, **canvas._plot_style_state())
    except Exception:
        pass
    if cbar is not None:    
        try:
            cbar.ax.tick_params(labelsize=tick_size)
            cbar.ax.yaxis.label.set_fontsize(label_size)
            cbar.ax.xaxis.label.set_fontsize(label_size)
            apply_text_style(cbar.ax.yaxis.label, family=canvas._font_family, **canvas._plot_style_state())
            apply_text_style(cbar.ax.xaxis.label, family=canvas._font_family, **canvas._plot_style_state())
            for lbl in list(cbar.ax.get_xticklabels()) + list(cbar.ax.get_yticklabels()):
                apply_text_style(lbl, family=canvas._font_family, **canvas._plot_style_state())
        except Exception:
            pass
