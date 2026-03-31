"""Matplotlib rendering helpers for canvas tiles."""
from __future__ import annotations

from ..._shared import QtGui, colormaps, np


def _format_colorbar_value(value):
    magnitude = abs(value)
    if magnitude == 0:
        return "0"
    if magnitude < 0.01 or magnitude >= 1000:
        return f"{value:.2e}"
    if magnitude < 1:
        return f"{value:.3f}"
    return f"{value:.2f}"


def _normalize_cbar_label(label: str) -> str:
    if not label:
        return ""
    lbl = label.strip()
    low = lbl.lower()
    if low.startswith("df"):
        return f"\\u0394f{lbl[2:]}"
    if low.startswith("delta f"):
        return f"\\u0394f{lbl[6:]}"
    return lbl


def _format_scale_bar_label(length: float, unit: str) -> str:
    if length is None:
        return ""
    if abs(length) < 1e-6:
        return f"0 {unit}".strip()
    unit_norm = (unit or "").strip().lower()
    if unit_norm == "nm":
        if abs(length) < 1.0:
            return f"{length * 1000:.0f} pm"
        if abs(length) >= 1000.0:
            return f"{length / 1000.0:.2g} um"
        return f"{length:g} nm"
    if unit_norm in ("um", "µm"):
        if abs(length) < 1.0:
            return f"{length * 1000:.0f} nm"
        return f"{length:g} um"
    if abs(length) >= 100 or abs(length) < 0.01:
        return f"{length:.2g} {unit}".strip()
    if abs(length) < 1:
        return f"{length:.2f} {unit}".strip()
    if abs(length) < 10:
        return f"{length:.2f} {unit}".strip()
    return f"{length:.1f} {unit}".strip()


def _normalized_value(norm, value):
    try:
        norm_val = norm(value)
        norm_val = float(norm_val)
    except Exception:
        return 0.5
    return float(np.clip(norm_val, 0.0, 1.0))


def _text_color_for_frame(frame_color):
    color = QtGui.QColor(frame_color or "#070707")
    if not color.isValid():
        color = QtGui.QColor("#070707")
    lum = (0.299 * color.redF()) + (0.587 * color.greenF()) + (0.114 * color.blueF())
    return "#101010" if lum > 0.55 else "#f5f5f5"


def _annotate_colorbar(cb, vmin, vmax, scale, orientation, show_ticks, text_color):
    if cb is None or vmin is None or vmax is None:
        return
    if not show_ticks:
        axis = cb.ax.xaxis if orientation == "horizontal" else cb.ax.yaxis
        axis.set_ticks([])
        axis.set_ticklabels([])
        axis.set_tick_params(length=0)
        for spine in cb.ax.spines.values():
            spine.set_visible(False)
        return
    ticks = [float(vmin), float(vmax)]
    cb.set_ticks(ticks)
    axis = cb.ax.xaxis if orientation == "horizontal" else cb.ax.yaxis
    labels = [_format_colorbar_value(val) for val in ticks]
    axis.set_ticklabels(labels)
    label_size = max(6.0, 8.0 * scale)
    axis.set_tick_params(labelsize=label_size, length=4, colors=text_color, width=0.8)
    for label in axis.get_ticklabels():
        label.set_color(text_color)
        label.set_fontsize(label_size)
        for spine in cb.ax.spines.values():
            spine.set_visible(False)


def _normalize_extent(extent):
    if not extent or len(extent) != 4:
        return None
    try:
        x0, x1, y1, y0 = [float(v) for v in extent]
    except Exception:
        return None
    ymin = min(y0, y1)
    ymax = max(y0, y1)
    return (x0, x1, ymin, ymax)


def _draw_molecules(ax, molecules, palette, show_hydrogens=True):
    if not molecules:
        return
    try:
        from matplotlib.patches import Circle
        from .molecular_overlay import Molecule, get_atom_color, get_atom_radius
    except Exception:
        return

    for entry in molecules:
        try:
            mol = entry if isinstance(entry, Molecule) else Molecule.from_dict(entry)
            coords = mol.get_transformed_coordinates()
        except Exception:
            continue
        if coords is None or len(coords) == 0:
            continue

        xs_full = np.asarray(coords[:, 0], dtype=float)
        ys_full = np.asarray(coords[:, 1], dtype=float)
        zs_full = np.asarray(coords[:, 2], dtype=float) if coords.shape[1] >= 3 else np.zeros(len(coords))
        xs = xs_full
        ys = ys_full
        zs = zs_full
        if not show_hydrogens:
            keep = np.array([(str(el).strip().upper() != "H") for el in getattr(mol, "elements", [])], dtype=bool)
            if keep.size == len(xs) and np.any(keep):
                xs = xs[keep]
                ys = ys[keep]
                zs = zs[keep]
                elements = [el for el, k in zip(getattr(mol, "elements", []), keep) if k]
            else:
                elements = list(getattr(mol, "elements", []))
        else:
            elements = list(getattr(mol, "elements", []))
        order = np.argsort(zs)
        display_mode = str(getattr(mol, "display_mode", "Atoms + Bonds") or "Atoms + Bonds").lower()
        bond_color = getattr(mol, "bond_color_override", None) or "#e8edf4"
        bond_style = str(getattr(mol, "bond_style", "default") or "default").lower()
        line_width = 0.8 if bond_style == "thin" else 2.0 if bond_style == "thick" else 1.2

        if display_mode != "atoms only":
            for bond in getattr(mol, "bonds", []) or []:
                try:
                    i, j = int(bond[0]), int(bond[1])
                    if not show_hydrogens:
                        ei = (mol.elements[i] if i < len(mol.elements) else "") or ""
                        ej = (mol.elements[j] if j < len(mol.elements) else "") or ""
                        if str(ei).strip().upper() == "H" or str(ej).strip().upper() == "H":
                            continue
                    ax.plot([xs_full[i], xs_full[j]], [ys_full[i], ys_full[j]], color=bond_color, linewidth=line_width, alpha=0.85, zorder=5)
                except Exception:
                    continue

        if display_mode == "bonds only":
            continue

        for idx in order:
            try:
                element = (elements[idx] if idx < len(elements) else "C") or "C"
                color = getattr(mol, "atom_color_override", None) or get_atom_color(element, palette)
                radius = get_atom_radius(element, getattr(mol, "radius_mode", "covalent"))
                radius *= float(getattr(mol, "scale", 0.1)) * float(getattr(mol, "radius_scale", 1.0))
                radius = max(0.03, float(radius) * 0.33)
                patch = Circle((xs[idx], ys[idx]), radius=radius, facecolor=color, edgecolor="#101316", linewidth=0.5, alpha=0.92, zorder=6 + (idx / max(1, len(order))))
                ax.add_patch(patch)
            except Exception:
                continue


def render_tile_figure_mpl(
    data,
    *,
    cmap,
    vmin,
    vmax,
    title,
    colorbar_label,
    width_px,
    height_px,
    dpi=200,
    show_colorbar=True,
    show_colorbar_ticks=True,
    show_title=False,
    show_metadata=True,
    metadata_left="",
    metadata_right="",
    show_overlay_main=False,
    overlay_main="",
    show_overlay_file=False,
    overlay_file="",
    cbar_position="bottom",
    metadata_height=0.0,
    frame_color="#070707",
    text_scale=None,
    text_color=None,
    show_scale_bar=False,
    scale_bar_length=None,
    scale_bar_unit="",
    scale_bar_width=None,
    extent=None,
    show_molecules=False,
    molecules=None,
    molecule_palette="pymol",
    show_hydrogens=True,
):
    """Build a Matplotlib figure for a canvas tile, including annotations."""
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib.figure import Figure
    from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar

    width_px = max(2, int(round(width_px)))
    height_px = max(2, int(round(height_px)))
    normalized_position = (cbar_position or "bottom").lower()
    if normalized_position == "hidden":
        normalized_position = "none"
    if normalized_position not in ("bottom", "top", "left", "right", "inset", "none"):
        normalized_position = "bottom"
    rendered_colorbar = show_colorbar and normalized_position != "none"

    metadata_ratio = 0.0
    if show_metadata and metadata_height > 0:
        metadata_ratio = min(0.35, metadata_height / height_px)
    bottom_margin = 0.002 + metadata_ratio
    if text_scale is None:
        text_scale = 1.0
    text_scale = max(0.002, min(2.4, float(text_scale)))
    text_color = text_color or _text_color_for_frame(frame_color)
    cbar_label_text = _normalize_cbar_label(colorbar_label or "")

    min_title_px = 2.0
    min_tick_px = 2.0
    min_label_px = 2.0
    min_overlay_px = 2.0
    min_meta_px = 2.0

    title_fs = max(min_title_px, 11.0 * text_scale)
    tick_fs = max(min_tick_px, 9.5 * text_scale)
    label_fs = max(min_label_px, 9.5 * text_scale)
    overlay_fs = max(min_overlay_px, 8.5 * text_scale)
    meta_fs = max(min_meta_px, 8.5 * text_scale)
    tick_pad = max(2.0, 0.25 * tick_fs)

    extra_tick_margin = 0.0
    if rendered_colorbar and show_colorbar_ticks and normalized_position in ("top", "bottom"):
        extra_tick_margin = min(0.18, max(0.06, (tick_fs * 2.2) / max(1.0, height_px)))
    top_margin = 0.98 - (extra_tick_margin if normalized_position == "top" else 0.0)
    fig = Figure(figsize=(width_px / dpi, height_px / dpi), dpi=dpi, facecolor=frame_color or "#070707")
    side_margin = 0.02
    if rendered_colorbar and show_colorbar_ticks and normalized_position in ("top", "bottom"):
        side_margin = min(0.12, max(0.055, (tick_fs * 1.7) / max(1.0, width_px)))
    fig.subplots_adjust(
        left=side_margin,
        right=1.0 - side_margin,
        top=top_margin,
        bottom=bottom_margin + (extra_tick_margin if normalized_position == "bottom" else 0.0),
    )

    ax = None
    cax = None
    orientation = "horizontal"
    if rendered_colorbar and normalized_position in ("top", "bottom"):
        cbar_height = max(10.0, 1.5 * tick_fs)
        cbar_ratio = min(0.45, max(0.08, cbar_height / max(1.0, height_px)))
        ratios = [cbar_ratio, 1] if normalized_position == "top" else [1, cbar_ratio]
        gs = fig.add_gridspec(2, 1, height_ratios=ratios, hspace=0.04)
        if normalized_position == "top":
            cax = fig.add_subplot(gs[0])
            ax = fig.add_subplot(gs[1])
        else:
            ax = fig.add_subplot(gs[0])
            cax = fig.add_subplot(gs[1])
        orientation = "horizontal"
    elif rendered_colorbar and normalized_position in ("left", "right"):
        cbar_width = max(10.0, 1.5 * tick_fs)
        cbar_ratio = min(0.45, max(0.08, cbar_width / max(1.0, width_px)))
        ratios = [cbar_ratio, 1] if normalized_position == "left" else [1, cbar_ratio]
        gs = fig.add_gridspec(1, 2, width_ratios=ratios, wspace=0.04)
        if normalized_position == "left":
            cax = fig.add_subplot(gs[0])
            ax = fig.add_subplot(gs[1])
        else:
            ax = fig.add_subplot(gs[0])
            cax = fig.add_subplot(gs[1])
        orientation = "vertical"
    else:
        ax = fig.add_subplot(1, 1, 1)

    try:
        cmap_obj = colormaps.get(cmap) if cmap else colormaps.get("viridis")
    except Exception:
        cmap_obj = colormaps.get("viridis")

    im = ax.imshow(
        data,
        cmap=cmap_obj,
        vmin=vmin,
        vmax=vmax,
        origin="lower",
        interpolation="nearest",
        extent=_normalize_extent(extent),
    )

    actual_vmin = None
    actual_vmax = None
    if vmin is not None and vmax is not None:
        actual_vmin = float(vmin)
        actual_vmax = float(vmax)
    else:
        try:
            actual_vmin = float(np.nanmin(data))
            actual_vmax = float(np.nanmax(data))
        except Exception:
            actual_vmin = None
            actual_vmax = None

    ax.set_facecolor("#070707")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    # Title display removed; keep colorbar label inside the bar instead.

    cb = None
    if rendered_colorbar and normalized_position in ("top", "bottom", "left", "right"):
        cb = fig.colorbar(im, cax=cax, orientation=orientation)
        cb.set_label("")  # place label manually inside
        cb.ax.tick_params(labelsize=tick_fs, length=3, width=0.7, colors=text_color, pad=tick_pad)
        cb.outline.set_edgecolor("#797979")
        cb.outline.set_linewidth(0.6)
        if orientation == "horizontal":
            cb.ax.xaxis.set_ticks_position("bottom")
            cb.ax.xaxis.set_label_position("bottom")
            cb.ax.text(0.5, 0.5, cbar_label_text, color=text_color, fontsize=label_fs, ha="center", va="center", transform=cb.ax.transAxes)
        else:
            cb.ax.yaxis.set_ticks_position("left")
            cb.ax.yaxis.set_label_position("left")
            cb.ax.text(0.5, 0.5, cbar_label_text, color=text_color, fontsize=label_fs, ha="center", va="center", rotation=90, rotation_mode="anchor", transform=cb.ax.transAxes)
    elif rendered_colorbar and normalized_position == "inset":
        inset_pos = [0.62, bottom_margin + 0.02, 0.3, 0.035]
        cax = fig.add_axes(inset_pos)
        cb = fig.colorbar(im, cax=cax, orientation="horizontal")
        cb.set_label("")
        cb.ax.tick_params(labelsize=tick_fs, length=3, width=0.5, colors=text_color, pad=tick_pad)
        cb.outline.set_edgecolor("#797979")
        cb.outline.set_linewidth(0.6)
        cb.ax.text(0.5, 0.5, cbar_label_text, color=text_color, fontsize=label_fs, ha="center", va="center", transform=cb.ax.transAxes)

    if cb is not None:
        annotate_orientation = orientation if normalized_position != "inset" else "horizontal"
        _annotate_colorbar(
            cb,
            actual_vmin,
            actual_vmax,
            tick_fs / 8.0,
            annotate_orientation,
            show_colorbar_ticks,
            text_color,
        )

    if ax is not None and cax is not None and normalized_position in ("top", "bottom"):
        ax_pos = ax.get_position()
        cax_pos = cax.get_position()
        cax.set_position([ax_pos.x0, cax_pos.y0, ax_pos.width, cax_pos.height])
    if ax is not None and cax is not None and normalized_position in ("left", "right"):
        ax_pos = ax.get_position()
        cax_pos = cax.get_position()
        cax.set_position([cax_pos.x0, ax_pos.y0, cax_pos.width, ax_pos.height])

    if show_scale_bar and scale_bar_length and scale_bar_width:
        try:
            bar_size = float(scale_bar_length)
            bar_width = float(scale_bar_width)
        except Exception:
            bar_size = 0.0
            bar_width = 0.0
        if bar_size > 0.0 and bar_width > 0.0:
            try:
                sb = AnchoredSizeBar(
                    ax.transData,
                    bar_size,
                    _format_scale_bar_label(bar_size, scale_bar_unit),
                    loc="center",
                    pad=0.35,
                    borderpad=0,
                    sep=3,
                    frameon=False,
                    size_vertical=max(bar_width * 0.004 * text_scale, bar_width * 0.0015),
                    color=text_color,
                    label_top=True,
                    bbox_to_anchor=(0.88, 0.08),
                    bbox_transform=ax.transAxes,
                )
                try:
                    sb.size_bar.get_children()[0].set_linewidth(0)
                except Exception:
                    pass
                text = sb.txt_label.get_children()[0]
                text.set_color(text_color)
                text.set_fontsize(max(2.0, 8.5 * text_scale))
                text.set_fontweight("bold")
                sb.set_zorder(6)
                ax.add_artist(sb)
            except Exception:
                pass

    if show_molecules and molecules:
        _draw_molecules(ax, molecules, molecule_palette, show_hydrogens=show_hydrogens)

    overlay_lines = []
    if show_overlay_main and overlay_main:
        overlay_lines.append(overlay_main)
    if show_overlay_file and overlay_file:
        overlay_lines.append(overlay_file)
    if overlay_lines:
        overlay_face = "#0b1424" if text_color == "#f5f5f5" else "#f5f5f5"
        ax.text(
            0.02,
            0.96,
            "\n".join(overlay_lines),
            fontsize=overlay_fs,
            color=text_color,
            weight="bold",
            ha="left",
            va="top",
            transform=ax.transAxes,
            bbox=dict(facecolor=overlay_face, alpha=0.75, edgecolor="none", boxstyle="round,pad=0.2"),
        )

    if show_metadata and (metadata_left or metadata_right):
        text_y = bottom_margin / 2
        meta_face = "#050505" if text_color == "#f5f5f5" else "#f5f5f5"
        bbox = dict(facecolor=meta_face, alpha=0.9, edgecolor="none", boxstyle="round,pad=0.2")
        font_size = meta_fs
        if metadata_left:
            fig.text(
                0.02,
                text_y,
                metadata_left,
                fontsize=font_size,
                color=text_color,
                ha="left",
                va="center",
                bbox=bbox,
            )
        if metadata_right:
            fig.text(
                0.98,
                text_y,
                metadata_right,
                fontsize=font_size,
                color=text_color,
                ha="right",
                va="center",
                bbox=bbox,
            )
    return fig


def render_tile_mpl(
    data,
    *,
    cmap,
    vmin,
    vmax,
    title,
    colorbar_label,
    width_px,
    height_px,
    dpi=200,
    show_colorbar=True,
    show_colorbar_ticks=True,
    show_title=False,
    show_metadata=True,
    metadata_left="",
    metadata_right="",
    show_overlay_main=False,
    overlay_main="",
    show_overlay_file=False,
    overlay_file="",
    cbar_position="bottom",
    metadata_height=0.0,
    frame_color="#070707",
    text_scale=None,
    text_color=None,
    show_scale_bar=False,
    scale_bar_length=None,
    scale_bar_unit="",
    scale_bar_width=None,
    extent=None,
    show_molecules=False,
    molecules=None,
    molecule_palette="pymol",
    show_hydrogens=True,
):
    """Render a canvas tile through Matplotlib, including annotations."""
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib.backends.backend_agg import FigureCanvasAgg

    fig = render_tile_figure_mpl(
        data,
        cmap=cmap,
        vmin=vmin,
        vmax=vmax,
        title=title,
        colorbar_label=colorbar_label,
        width_px=width_px,
        height_px=height_px,
        dpi=dpi,
        show_colorbar=show_colorbar,
        show_colorbar_ticks=show_colorbar_ticks,
        show_title=show_title,
        show_metadata=show_metadata,
        metadata_left=metadata_left,
        metadata_right=metadata_right,
        show_overlay_main=show_overlay_main,
        overlay_main=overlay_main,
        show_overlay_file=show_overlay_file,
        overlay_file=overlay_file,
        cbar_position=cbar_position,
        metadata_height=metadata_height,
        frame_color=frame_color,
        text_scale=text_scale,
        text_color=text_color,
        show_scale_bar=show_scale_bar,
        scale_bar_length=scale_bar_length,
        scale_bar_unit=scale_bar_unit,
        scale_bar_width=scale_bar_width,
        extent=extent,
        show_molecules=show_molecules,
        molecules=molecules,
        molecule_palette=molecule_palette,
        show_hydrogens=show_hydrogens,
    )
    canvas = FigureCanvasAgg(fig)
    canvas.draw()
    buf = np.asarray(canvas.buffer_rgba())
    h, w, _ = buf.shape
    qimg = QtGui.QImage(buf.data, w, h, QtGui.QImage.Format_RGBA8888)
    return QtGui.QPixmap.fromImage(qimg.copy())



