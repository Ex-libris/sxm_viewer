"""Render a folder report model (see model.build_report_model) to PDF.

Pure matplotlib/Agg - every page is a throwaway Figure written through
``PdfPages``; no Qt anywhere (see package docstring). Page builders are
also exposed individually so tests can render single pages to PNG.
"""
from __future__ import annotations

import math
from datetime import datetime

import numpy as np
import matplotlib
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.colors import ListedColormap
from matplotlib.figure import Figure
from matplotlib import patches, patheffects

from .. import cmap_registry
from .channels import classify_channel

PAGE_W, PAGE_H = 11.69, 8.27  # A4 landscape, inches
_FALLBACK_CYCLE = [matplotlib.colors.to_hex(matplotlib.colormaps["tab10"](i)) for i in range(10)]
_MAP_CMAP = "viridis"
# Default colormap per channel class (user convention): current -> Blues,
# frequency shift -> gray, topography -> Blues; anything else gray.
_CLASS_CMAPS = {"current": "Blues", "freq": "gray", "z": "Blues", "other": "gray"}


def _prefs(model):
    return (model.get("payload") or {}).get("prefs") or {}


def _cycle_colors(model):
    """The user's persisted color cycle, as concrete colors."""
    colors = _prefs(model).get("color_cycle")
    return list(colors) if colors else list(_FALLBACK_CYCLE)


def _class_cmap(cls, override=None):
    return override or _CLASS_CMAPS.get(str(cls or "other"), "gray")


def _panel_cmap(model, panel):
    """Colormap for an image panel: the user's starred preview favourite
    for topography, the per-class convention otherwise. The favourite is
    resolved through the central registry so a name from the optional
    extra-colormaps package degrades to the class default (instead of
    crashing report generation) when that package isn't installed."""
    cls = str((panel or {}).get("cls") or "z")
    if cls == "z":
        fav = _prefs(model).get("preview_cmap")
        if fav:
            return cmap_registry.get_cmap(fav, fallback=_class_cmap(cls)).name
    return _class_cmap(cls)


def _nice_scale_bar_nm(width_nm):
    """A round 1/2/5x10^k bar length, roughly a quarter of the image."""
    target = width_nm / 4.0
    if not math.isfinite(target) or target <= 0:
        return None
    base = 10.0 ** math.floor(math.log10(target))
    for mult in (5.0, 2.0, 1.0):
        if base * mult <= target:
            return base * mult
    return base


def _draw_scale_bar(ax, shape, width_nm, fontsize=5):
    """White scale bar with a black outline, bottom-left, for an image
    drawn in raster pixel coordinates."""
    try:
        width_nm = float(width_nm or 0.0)
    except Exception:
        return
    bar_nm = _nice_scale_bar_nm(width_nm)
    if bar_nm is None:
        return
    h, w = shape[:2]
    bar_px = bar_nm / width_nm * w
    x0, y0 = 0.05 * w, 0.93 * h
    ax.plot([x0, x0 + bar_px], [y0, y0], color="white", linewidth=2,
            solid_capstyle="butt",
            path_effects=[patheffects.withStroke(linewidth=3.2, foreground="black")])
    label = f"{bar_nm:g} nm" if bar_nm >= 1 else f"{bar_nm * 10:g} Å"
    ax.text(x0 + bar_px / 2.0, y0 - 0.025 * h, label, color="white",
            fontsize=fontsize, ha="center", va="bottom",
            path_effects=[patheffects.withStroke(linewidth=1.6, foreground="black")])


def _image_width_nm(image):
    extent = (image or {}).get("extent")
    if not extent or len(extent) != 4:
        return None
    try:
        return abs(float(extent[1]) - float(extent[0]))
    except Exception:
        return None


def _fmt_time(t, fmt="%Y-%m-%d %H:%M"):
    return t.strftime(fmt) if isinstance(t, datetime) else "?"


def _new_page(title=None):
    fig = Figure(figsize=(PAGE_W, PAGE_H), dpi=150)
    if title:
        fig.suptitle(title, fontsize=13, fontweight="bold", y=0.97)
    return fig


def _auto_clim(arr):
    finite = np.asarray(arr, dtype=float)
    finite = finite[np.isfinite(finite)]
    if not finite.size:
        return None
    lo, hi = np.percentile(finite, [2.0, 98.0])
    if lo == hi:
        lo, hi = float(finite.min()), float(finite.max()) or 1.0
    return lo, hi


def _imshow_topo(ax, image, cmap="Blues"):
    """Auto-contrast image panel in raster (pixel) coordinates; returns True
    when there was something to draw."""
    arr = image.get("topo")
    if arr is None:
        ax.text(0.5, 0.5, "(no image data)", ha="center", va="center",
                transform=ax.transAxes, fontsize=8, color="0.5")
        ax.set_xticks([]), ax.set_yticks([])
        return False
    arr = np.asarray(arr, dtype=float)
    clim = _auto_clim(arr)
    ax.imshow(arr, cmap=cmap, origin="upper", aspect="equal",
              vmin=clim[0] if clim else None, vmax=clim[1] if clim else None,
              interpolation="nearest")
    ax.set_xticks([]), ax.set_yticks([])
    return True


def _rotated_footprint(extent, angle_deg):
    """Corner xy list of an image's rotated footprint in absolute nm.
    Extent convention [x0, x1, y1, y0] per _header_extent; rotation about
    the frame center by +angle to match _map_spec_to_pixels' convention."""
    x0, x1, y1, y0 = [float(v) for v in extent]
    cx, cy = 0.5 * (x0 + x1), 0.5 * (y0 + y1)
    theta = math.radians(angle_deg or 0.0)
    cos_t, sin_t = math.cos(theta), math.sin(theta)
    corners = []
    for px, py in ((x0, y0), (x1, y0), (x1, y1), (x0, y1)):
        dx, dy = px - cx, py - cy
        # Inverse of the spec->pixel rotation: place the frame's own
        # corners back into the absolute frame (rotate by -theta).
        corners.append((cx + dx * cos_t + dy * sin_t,
                        cy - dx * sin_t + dy * cos_t))
    return corners


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

def page_cover(model):
    s = model["summary"]
    fig = _new_page()
    ax = fig.add_axes([0.08, 0.08, 0.84, 0.84])
    ax.axis("off")
    lines = [
        ("SXM Folder Report", 22, "bold", 0.92),
        (str(s["folder"]), 11, "normal", 0.84),
        (f"Acquired {_fmt_time(s['t0'])}  to  {_fmt_time(s['t1'])}", 10, "normal", 0.78),
        (f"{s['n_images']} scans   |   {s['n_specs']} single spectra   |   "
         f"{s['n_grids']} grid datasets ({s['n_grid_points']} grid points)",
         11, "normal", 0.70),
        (f"{len(model['regions'])} sample regions (scans grouped by position + acquisition time)",
         10, "normal", 0.64),
        (f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} by SXM Viewer",
         8, "normal", 0.06),
    ]
    sequences = model.get("sequences") or []
    if sequences:
        lines.insert(5, (
            f"{len(sequences)} image data-set(s): same frame re-scanned "
            "stepping bias / tip height", 10, "normal", 0.58))
    for text, size, weight, y in lines:
        ax.text(0.02, y, text, fontsize=size, fontweight=weight, transform=ax.transAxes)

    # Contact sheet of up to 4x8 image thumbnails.
    images = [i for i in model["payload"]["images"] if i.get("topo") is not None]
    n = min(len(images), 32)
    if n:
        ncols = 8
        for i in range(n):
            r, c = divmod(i, ncols)
            sub = fig.add_axes([0.08 + c * 0.105, 0.46 - r * 0.105, 0.095, 0.095])
            panels = images[i].get("panels") or []
            _imshow_topo(sub, images[i],
                         cmap=_panel_cmap(model, panels[0] if panels else None))
            width_nm = _image_width_nm(images[i])
            if width_nm:
                _draw_scale_bar(sub, np.asarray(images[i]["topo"]).shape, width_nm,
                                fontsize=3.2)
            sub.set_title(images[i].get("name", ""), fontsize=4, pad=1)
    return fig


def page_overview(model):
    fig = _new_page("Session overview - image footprints, spectra and grids by sample region")
    ax = fig.add_axes([0.07, 0.09, 0.68, 0.82])
    colors = _cycle_colors(model)
    regions = model["regions"]
    image_region = model["image_region"]
    payload = model["payload"]

    def _region_color(idx):
        return colors[idx % len(colors)]

    for image in payload["images"]:
        extent = image.get("extent")
        if not extent:
            continue
        region_idx = image_region.get(str(image["key"]), 0)
        corners = _rotated_footprint(extent, image.get("angle") or 0.0)
        ax.add_patch(patches.Polygon(corners, closed=True, fill=False,
                                     edgecolor=_region_color(region_idx),
                                     linewidth=0.7, alpha=0.75))
    # Batch points into one artist per (region, kind) - a grid alone can
    # contribute thousands of points, one Line2D each would crawl.
    buckets = {}
    for spec in payload["specs"]:
        x, y = spec.get("x"), spec.get("y")
        if x is None or y is None:
            continue
        key = str(spec.get("image_key") or "")
        region_idx = image_region.get(key, 0)
        kind = "grid" if spec.get("is_matrix") else "single"
        xs, ys = buckets.setdefault((region_idx, kind), ([], []))
        xs.append(x), ys.append(y)
    for (region_idx, kind), (xs, ys) in buckets.items():
        color = _region_color(region_idx)
        if kind == "grid":
            ax.plot(xs, ys, ".", color=color, ms=1.2, alpha=0.5, linestyle="none")
        else:
            ax.plot(xs, ys, "o", mfc="none", mec=color, ms=5, mew=1.2, linestyle="none")
    ax.set_xlabel("x (nm)")
    ax.set_ylabel("y (nm)")
    ax.set_aspect("equal", adjustable="datalim")
    ax.autoscale_view()
    ax.grid(True, linewidth=0.3, alpha=0.4)

    legend_ax = fig.add_axes([0.78, 0.09, 0.2, 0.82])
    legend_ax.axis("off")
    legend_ax.text(0, 1.0, "Sample regions", fontsize=10, fontweight="bold",
                   va="top", transform=legend_ax.transAxes)
    legend_ax.text(0, 0.965, "(same area + same time window)", fontsize=6.5,
                   color="0.45", va="top", transform=legend_ax.transAxes)
    y = 0.92
    for idx, region in enumerate(regions):
        legend_ax.add_patch(patches.Rectangle((0.0, y - 0.012), 0.06, 0.024,
                                              color=_region_color(idx),
                                              transform=legend_ax.transAxes))
        legend_ax.text(0.09, y, f"{region['label']}  ({len(region['image_keys'])} scans)",
                       fontsize=8, va="center", transform=legend_ax.transAxes)
        # Sessions routinely span days - show dates, not just clock times.
        legend_ax.text(0.09, y - 0.028,
                       f"{_fmt_time(region['t0'], '%m-%d %H:%M')} - {_fmt_time(region['t1'], '%m-%d %H:%M')}",
                       fontsize=7, color="0.4", va="center", transform=legend_ax.transAxes)
        y -= 0.075
        if y < 0.02:
            legend_ax.text(0.09, y, "...", fontsize=8, transform=legend_ax.transAxes)
            break
    legend_ax.text(0, 0.0,
                   "o  single spectrum\n.  grid point",
                   fontsize=7, color="0.35", transform=legend_ax.transAxes)
    return fig


def page_image_specs(model, image, specs, region_label=""):
    title = f"{region_label}  -  {image.get('name', '')}" if region_label else image.get("name", "")
    if image.get("tag"):
        title += f"   [{image['tag']}]"
    fig = _new_page(title)
    # Constant-height images carry two signal panels (current + frequency
    # shift / lock-in); everything else a single topography panel.
    panels = [p for p in (image.get("panels") or []) if p.get("data") is not None]
    if not panels and image.get("topo") is not None:
        panels = [{"label": "", "unit": image.get("topo_unit", ""), "data": image["topo"]}]
    panel_axes = []
    if len(panels) >= 2:
        panel_axes = [(fig.add_axes([0.05, 0.52, 0.42, 0.38]), panels[0]),
                      (fig.add_axes([0.05, 0.10, 0.42, 0.38]), panels[1])]
    elif panels:
        panel_axes = [(fig.add_axes([0.05, 0.12, 0.42, 0.76]), panels[0])]
    else:
        ax = fig.add_axes([0.05, 0.12, 0.42, 0.76])
        _imshow_topo(ax, {"topo": None})
    colors = _cycle_colors(model)
    width_nm = _image_width_nm(image)

    for ax_img, panel in panel_axes:
        _imshow_topo(ax_img, {"topo": panel["data"]}, cmap=_panel_cmap(model, panel))
        if width_nm:
            _draw_scale_bar(ax_img, np.asarray(panel["data"]).shape, width_nm)
        if panel.get("label"):
            ax_img.set_title(panel["label"], fontsize=7, pad=2)

    ax_curves = fig.add_axes([0.55, 0.30, 0.40, 0.58])
    any_curve = False
    for idx, spec in enumerate(specs):
        color = colors[idx % len(colors)]
        frac = spec.get("marker_frac")
        if frac is not None:
            # Markers come as raster fractions; scale by each *displayed*
            # (possibly downsampled) array's own shape.
            for ax_img, panel in panel_axes:
                disp_h, disp_w = np.asarray(panel["data"]).shape[:2]
                if disp_w < 2 or disp_h < 2:
                    continue
                fx, fy = frac
                marker = "s" if spec.get("off_frame_direction") else "o"
                px, py = fx * (disp_w - 1), fy * (disp_h - 1)
                ax_img.plot(px, py, marker, mfc="none", mec=color, ms=7, mew=1.5)
                ax_img.annotate(str(idx + 1), (px, py),
                                textcoords="offset points", xytext=(5, 5),
                                fontsize=6, color=color)
        v, c = spec.get("V"), spec.get("curve")
        if v is not None and c is not None and len(v) == len(c):
            label = f"{idx + 1}: {spec.get('name', '')}"
            off = spec.get("off_frame_direction")
            if off:
                label += f" (off-frame {off})"
            ax_curves.plot(v, c, color=color, linewidth=0.9, label=label)
            any_curve = True
    if any_curve:
        first = next((s for s in specs if s.get("curve") is not None), {})
        ax_curves.set_xlabel(first.get("x_label") or "Bias (V)", fontsize=8)
        ax_curves.set_ylabel(first.get("curve_channel") or "signal", fontsize=8)
        ax_curves.tick_params(labelsize=7)
        ax_curves.legend(fontsize=5.5, loc="best", framealpha=0.6)
        ax_curves.grid(True, linewidth=0.3, alpha=0.4)
    else:
        ax_curves.axis("off")
        ax_curves.text(0.5, 0.5, "(no curve data)", ha="center", va="center",
                       transform=ax_curves.transAxes, fontsize=8, color="0.5")

    meta_ax = fig.add_axes([0.55, 0.08, 0.40, 0.16])
    meta_ax.axis("off")
    meta = image.get("meta") or {}
    meta_lines = [f"{k}: {v}" for k, v in list(meta.items())[:8]]
    meta_lines.append(f"time: {_fmt_time(image.get('time'))}")
    meta_ax.text(0, 1, "\n".join(meta_lines), fontsize=7, va="top",
                 family="monospace", transform=meta_ax.transAxes)

    img_meta_ax = fig.add_axes([0.05, 0.03, 0.42, 0.06])
    img_meta_ax.axis("off")
    img_meta_ax.text(0, 0.5, f"{len(specs)} spectra on this image "
                     "(squares = off-frame, clamped to nearest edge)",
                     fontsize=7, color="0.35", transform=img_meta_ax.transAxes)
    return fig


def _fmt_bias_v(bias_v):
    if bias_v is None:
        return "?"
    if abs(bias_v) < 1.0:
        return f"{bias_v * 1e3:.4g} mV"
    return f"{bias_v:.4g} V"


def _fmt_z_rel_pm(z_rel_pm):
    if z_rel_pm is None:
        return "z rel = ?"
    if abs(z_rel_pm) < 0.05:
        return "z rel = 0 (ref)"
    return f"z rel = {z_rel_pm:+.3g} pm"


def _sequence_member_annotation(seq, member):
    """The stepped parameter for one member, e.g. 'z rel = +20 pm'."""
    kind = seq.get("kind")
    parts = []
    if kind in ("height", "bias+height"):
        parts.append(_fmt_z_rel_pm(member.get("z_rel_pm")))
    if kind in ("bias", "bias+height"):
        parts.append(f"bias = {_fmt_bias_v(member.get('bias_v'))}")
    if not parts:
        parts.append(_fmt_time(member.get("time"), "%H:%M:%S"))
    return "   ".join(parts)


# Sequence pages pin the channel colormaps explicitly (user convention):
# current in reversed Blues, frequency shift / lock-in in gray.
_SEQ_CLASS_CMAPS = {"current": "Blues_r", "freq": "gray", "lockin": "gray"}


def _sequence_member_pair(image):
    """(current_panel, signal_panel) for one sequence member: the current
    channel and the frequency-shift (or lock-in) channel, either None when
    the member doesn't carry it."""
    cur = sig = None
    for panel in (image or {}).get("panels") or []:
        cls = str(panel.get("cls") or "")
        if cls == "current" and cur is None:
            cur = panel
        elif cls in ("freq", "lockin") and sig is None:
            sig = panel
    return cur, sig


def _seq_panel_cmap(model, panel):
    cls = str((panel or {}).get("cls") or "")
    return _SEQ_CLASS_CMAPS.get(cls) or _panel_cmap(model, panel)


def pages_sequence(model, seq, images_by_key, max_members=None):
    """Figures for one image sequence (data-set): the same frame re-scanned
    while stepping bias and/or tip-sample distance. Panels are annotated
    with each member's value of the stepped parameter (z rel is relative to
    the sequence's first image); the first page adds a member table.

    Constant-height members carrying both channels render in *paired rows*
    - per column one member, current on the top row (Blues_r), frequency
    shift on the bottom row (gray) - so the two channels of the same file
    always correspond vertically. Members with a single display channel use
    the plain 2x4 panel grid. Long sequences (drift-corrected movies can
    run to dozens of frames) are capped with a note instead of flooding
    the report."""
    members = seq.get("members") or []
    paired = any(all(_sequence_member_pair(images_by_key.get(str(m.get("key") or ""))))
                 for m in members)
    per_page = 4 if paired else 8
    if max_members is None:
        max_members = 16 if paired else 24
    shown = members[:max_members]
    n_hidden = len(members) - len(shown)
    n_pages = max(1, math.ceil(len(shown) / per_page))
    for page_idx in range(n_pages):
        chunk = shown[page_idx * per_page:(page_idx + 1) * per_page]
        title = (f"{seq.get('label', '')}  -  {seq.get('kind_label', '')}   "
                 f"({len(members)} images)")
        if seq.get("region") is not None:
            title += f"   -   Region {int(seq['region']) + 1}"
        if n_pages > 1:
            title += f"   [{page_idx + 1}/{n_pages}]"
        fig = _new_page(title)
        row_labels = [None, None]
        for i, member in enumerate(chunk):
            idx = page_idx * per_page + i
            image = images_by_key.get(str(member.get("key") or ""))
            width_nm = _image_width_nm(image)
            member_title = (f"{idx}: {member.get('name', '')}\n"
                            f"{_sequence_member_annotation(seq, member)}")
            if paired:
                cur, sig = _sequence_member_pair(image)
                cells = ((0, cur), (1, sig))
            else:
                panels = (image or {}).get("panels") or []
                cells = ((i // 4, panels[0] if panels else None),)
            for row, panel in cells:
                col = i if paired else i % 4
                ax = fig.add_axes([0.05 + col * 0.235, 0.56 - row * 0.325, 0.20, 0.28])
                have = _imshow_topo(
                    ax, {"topo": panel.get("data") if panel else None},
                    cmap=_seq_panel_cmap(model, panel))
                if have and width_nm:
                    _draw_scale_bar(ax, np.asarray(panel["data"]).shape, width_nm,
                                    fontsize=4)
                if paired and panel and row_labels[row] is None:
                    row_labels[row] = str(panel.get("label") or "")
                if not paired or row == 0:
                    ax.set_title(member_title, fontsize=6, pad=2)
        if paired:
            for row, label in enumerate(row_labels):
                if label:
                    fig.text(0.035, 0.56 - row * 0.325 + 0.14, label,
                             fontsize=6.5, color="0.35", rotation=90,
                             ha="center", va="center")
        if page_idx == 0:
            info_ax = fig.add_axes([0.05, 0.015, 0.9, 0.17])
            info_ax.axis("off")
            head = [f"{'#':>2}  {'image':<36} {'time':<17} {'bias':>10} {'z rel':>12}"]
            rows = []
            for idx, member in enumerate(members[:5]):
                z_txt = (f"{member['z_rel_pm']:+.4g} pm"
                         if member.get("z_rel_pm") is not None else "-")
                rows.append(f"{idx:>2}  {str(member.get('name', '')):<36.36} "
                            f"{_fmt_time(member.get('time'), '%m-%d %H:%M:%S'):<17} "
                            f"{_fmt_bias_v(member.get('bias_v')):>10} {z_txt:>12}")
            if len(members) > 5:
                extra = (f"(first {len(shown)} shown as panels)"
                         if n_hidden else "(all shown as panels)")
                rows.append(f"    ... and {len(members) - 5} more {extra}")
            info_ax.text(0, 1, "\n".join(head + rows), fontsize=6, va="top",
                         family="monospace", transform=info_ax.transAxes)
        if page_idx == n_pages - 1 and n_hidden:
            fig.text(0.97, 0.02, f"... {n_hidden} more image(s) not shown",
                     fontsize=7, ha="right", color="0.4")
        yield fig


def page_grid(model, grid, images_by_key):
    fig = _new_page(f"Grid map  -  {grid.get('name', '')}   "
                    f"({grid['rows']}x{grid['cols']}, {grid['n_points']} points)")
    rows, cols = grid["rows"], grid["cols"]
    dx, dy = grid["local_dx"], grid["local_dy"]
    # Local/relative frame per the Grid Map Explorer convention: extent
    # (0, cols*dx, rows*dy, 0) with origin='upper' on already-flipped arrays.
    local_extent = (0.0, cols * dx, rows * dy, 0.0)

    # 1. Anchor image + footprint.
    ax_anchor = fig.add_axes([0.045, 0.36, 0.27, 0.5])
    image = images_by_key.get(str(grid.get("image_key") or ""))
    if image is not None:
        anchor_panels = image.get("panels") or []
        have_topo = _imshow_topo(
            ax_anchor, image,
            cmap=_panel_cmap(model, anchor_panels[0] if anchor_panels else None))
        if have_topo:
            anchor_width_nm = _image_width_nm(image)
            if anchor_width_nm:
                _draw_scale_bar(ax_anchor, np.asarray(image["topo"]).shape,
                                anchor_width_nm)
        X, Y = grid.get("X"), grid.get("Y")
        extent = image.get("extent")
        if have_topo and X is not None and extent:
            # Project grid corner points into raster pixels the same way
            # _map_spec_to_pixels does (rotate-then-normalize).
            x0, x1, y1, y0 = [float(v) for v in extent]
            cx, cy = 0.5 * (x0 + x1), 0.5 * (y0 + y1)
            xspan, yspan = (x1 - x0), (y1 - y0)
            theta = math.radians(image.get("angle") or 0.0)
            cos_t, sin_t = math.cos(theta), math.sin(theta)
            # Coordinates of the *displayed* (downsampled) array, not the
            # header's raster - same rescaling rule as page_image_specs.
            disp_h, disp_w = np.asarray(image["topo"]).shape[:2]
            pts_x, pts_y = [], []
            for gx, gy in zip(np.asarray(X).ravel(), np.asarray(Y).ravel()):
                if not (np.isfinite(gx) and np.isfinite(gy)) or xspan <= 0 or yspan <= 0:
                    continue
                ddx, ddy = gx - cx, gy - cy
                u = (ddx * cos_t - ddy * sin_t) / xspan
                v = (ddx * sin_t + ddy * cos_t) / yspan
                pts_x.append((u + 0.5) * (disp_w - 1))
                pts_y.append((0.5 - v) * (disp_h - 1))
            if pts_x:
                ax_anchor.plot(pts_x, pts_y, ".", color="#ff5050", ms=0.8, alpha=0.5)
                ax_anchor.set_xlim(min(-0.5, min(pts_x)), max(disp_w - 0.5, max(pts_x)))
                ax_anchor.set_ylim(max(disp_h - 0.5, max(pts_y)), min(-0.5, min(pts_y)))
        ax_anchor.set_title(f"anchor: {image.get('name', '')}", fontsize=7)
    else:
        ax_anchor.axis("off")
        ax_anchor.text(0.5, 0.5, "(no anchor image)", ha="center", va="center",
                       transform=ax_anchor.transAxes, fontsize=8, color="0.5")

    # 2. Slice map (local frame). Colormap: the user's starred grid-map
    # favourite, else the channel-class convention (freq -> gray,
    # current -> Blues, ...).
    ax_slice = fig.add_axes([0.37, 0.36, 0.27, 0.5])
    slice_map = grid.get("slice_map")
    slice_cmap = _class_cmap(classify_channel(grid.get("curve_channel")),
                             override=_prefs(model).get("grid_metric_cmap"))
    if slice_map is not None and np.isfinite(slice_map).any():
        clim = _auto_clim(slice_map)
        im = ax_slice.imshow(slice_map, cmap=slice_cmap, origin="upper",
                             extent=local_extent, aspect="equal",
                             vmin=clim[0] if clim else None,
                             vmax=clim[1] if clim else None,
                             interpolation="nearest")
        fig.colorbar(im, ax=ax_slice, fraction=0.046, pad=0.03).ax.tick_params(labelsize=6)
        bias = grid.get("bias_value")
        bias_txt = f" @ {bias:.3g} V" if bias is not None else ""
        ax_slice.set_title(f"{grid.get('curve_channel') or 'signal'}{bias_txt}", fontsize=7)
        ax_slice.set_xlabel("nm", fontsize=6)
        ax_slice.tick_params(labelsize=6)
    else:
        ax_slice.axis("off")
        ax_slice.text(0.5, 0.5, "(no slice data)", ha="center", va="center",
                      transform=ax_slice.transAxes, fontsize=8, color="0.5")

    # 3. Cluster-label map (local frame), colored with the user's color
    # cycle so the map and the curves below stay visually linked.
    colors = _cycle_colors(model)
    cluster_cmap = ListedColormap(colors[:10])
    ax_clu = fig.add_axes([0.70, 0.36, 0.27, 0.5])
    cluster_map = grid.get("cluster_map")
    curves = grid.get("cluster_curves") or []
    if cluster_map is not None and np.isfinite(cluster_map).any() and curves:
        ax_clu.imshow(cluster_map, cmap=cluster_cmap, origin="upper",
                      extent=local_extent, aspect="equal",
                      vmin=-0.5, vmax=len(colors[:10]) - 0.5,
                      interpolation="nearest")
        ax_clu.set_title(f"curve-shape clusters (k={len(curves)})", fontsize=7)
        ax_clu.set_xlabel("nm", fontsize=6)
        ax_clu.tick_params(labelsize=6)
    else:
        ax_clu.axis("off")
        ax_clu.text(0.5, 0.5, "(no clusters)", ha="center", va="center",
                    transform=ax_clu.transAxes, fontsize=8, color="0.5")

    # 4. Representative spectrum per cluster - a real member curve (nearest
    # the cluster centroid), deliberately not an average: averaging grid
    # spectra washes out the physics.
    ax_rep = fig.add_axes([0.08, 0.06, 0.55, 0.22])
    if curves:
        for entry in curves:
            color = colors[int(entry["label"]) % len(colors[:10])]
            ax_rep.plot(entry["V"], entry["curve"], color=color, linewidth=1.1,
                        label=f"cluster {entry['label']} (n={entry['count']})")
        ax_rep.set_xlabel(grid.get("x_label") or "Bias (V)", fontsize=7)
        ax_rep.set_ylabel(grid.get("curve_channel") or "signal", fontsize=7)
        ax_rep.tick_params(labelsize=6)
        ax_rep.legend(fontsize=6, loc="best", framealpha=0.6)
        ax_rep.grid(True, linewidth=0.3, alpha=0.4)
        ax_rep.set_title("representative spectrum per cluster (nearest centroid)", fontsize=7)
    else:
        ax_rep.axis("off")

    info_ax = fig.add_axes([0.68, 0.06, 0.29, 0.22])
    info_ax.axis("off")
    info_lines = [
        f"dataset: {grid.get('dataset') or ''}",
        f"time: {_fmt_time(grid.get('time'))}",
        f"pitch: {grid['local_dx']:.3g} x {grid['local_dy']:.3g} nm",
        f"local-frame flips: row={grid['row_flip']} col={grid['col_flip']}",
        "maps shown in the grid's local (relative-axes) frame,",
        "matching the Grid Map Explorer's default view",
    ]
    info_ax.text(0, 1, "\n".join(info_lines), fontsize=7, va="top",
                 family="monospace", transform=info_ax.transAxes)
    return fig


def page_grid_kpfm(grid):
    """2x2 KPFM parabola-fit maps (LCPD and c, each with 1-sigma error) for
    a bias-swept grid whose frequency-shift channel varies with bias."""
    kpfm = grid.get("kpfm")
    if not kpfm:
        return None
    fig = _new_page(f"KPFM fit  -  {grid.get('name', '')}   "
                    f"({kpfm['n_ok']}/{kpfm['n_total']} pixels fit, "
                    f"channel: {kpfm.get('channel') or 'frequency shift'})")
    rows, cols = grid["rows"], grid["cols"]
    dx, dy = grid["local_dx"], grid["local_dy"]
    local_extent = (0.0, cols * dx, rows * dy, 0.0)
    # Colormap conventions (user-specified, matching MatrixFitDialog's
    # PARAM_INFO): LCPD diverging bwr, c in gray (more negative = darker),
    # errors in viridis.
    panels = (
        ("LCPD (V)", kpfm["lcpd"], "bwr", [0.07, 0.52, 0.38, 0.38]),
        ("LCPD error (V)", kpfm["lcpd_err"], "viridis", [0.55, 0.52, 0.38, 0.38]),
        ("c coefficient", kpfm["c"], "gray", [0.07, 0.06, 0.38, 0.38]),
        ("c error", kpfm["c_err"], "viridis", [0.55, 0.06, 0.38, 0.38]),
    )
    for label, data, cmap_name, rect in panels:
        ax = fig.add_axes(rect)
        if data is not None and np.isfinite(data).any():
            clim = _auto_clim(data)
            im = ax.imshow(data, cmap=cmap_name, origin="upper",
                           extent=local_extent, aspect="equal",
                           vmin=clim[0] if clim else None,
                           vmax=clim[1] if clim else None,
                           interpolation="nearest")
            fig.colorbar(im, ax=ax, fraction=0.046, pad=0.03).ax.tick_params(labelsize=6)
        else:
            ax.text(0.5, 0.5, "(no data)", ha="center", va="center",
                    transform=ax.transAxes, fontsize=8, color="0.5")
        ax.set_title(label, fontsize=8)
        ax.tick_params(labelsize=6)
        ax.set_xlabel("nm", fontsize=6)
    return fig


def page_flagged(model):
    flagged = model["flagged"]
    if not any(flagged.values()):
        return None
    fig = _new_page("Needs attention")
    ax = fig.add_axes([0.06, 0.05, 0.9, 0.86])
    ax.axis("off")
    y = 1.0

    def _section(title, items, fmt):
        nonlocal y
        if not items:
            return
        ax.text(0, y, title, fontsize=10, fontweight="bold", va="top",
                transform=ax.transAxes)
        y -= 0.045
        for spec in items[:30]:
            ax.text(0.02, y, fmt(spec), fontsize=7, va="top",
                    family="monospace", transform=ax.transAxes)
            y -= 0.028
        if len(items) > 30:
            ax.text(0.02, y, f"... and {len(items) - 30} more", fontsize=7,
                    va="top", transform=ax.transAxes)
            y -= 0.028
        y -= 0.03

    _section("Off-frame positions (clamped to nearest image edge)",
             flagged["off_frame"],
             lambda s: f"{s.get('name', ''):<40} {str(s.get('off_frame_direction') or ''):<6} "
                       f"{(s.get('off_frame_distance_nm') or 0):.1f} nm outside "
                       f"{('[grid] ' if s.get('is_matrix') else '')}")
    _section("Low-confidence image assignments",
             flagged["low_confidence"],
             lambda s: f"{s.get('name', ''):<40} {s.get('assignment_reason_label') or ''}")
    _section("Unassigned spectra (no matching image)",
             flagged["unassigned"],
             lambda s: f"{s.get('name', '')}")
    return fig


# ---------------------------------------------------------------------------
# Top level
# ---------------------------------------------------------------------------

def iter_report_figures(model):
    """Yield (description, Figure) for every report page, in order."""
    payload = model["payload"]
    images_by_key = {str(i["key"]): i for i in payload["images"]}

    yield "cover", page_cover(model)
    if payload["images"]:
        yield "overview", page_overview(model)

    for seq in model.get("sequences") or []:
        for fig in pages_sequence(model, seq, images_by_key):
            yield f"sequence {seq.get('label', '')}", fig

    specs_by_image = model["specs_by_image"]
    for region in model["regions"]:
        for key in region["image_keys"]:
            specs = specs_by_image.get(str(key))
            if not specs:
                continue
            image = images_by_key.get(str(key))
            if image is None:
                continue
            yield (f"image {image.get('name', '')}",
                   page_image_specs(model, image, specs, region_label=region["label"]))

    for grid in model["grids"]:
        yield f"grid {grid.get('name', '')}", page_grid(model, grid, images_by_key)
        kpfm_fig = page_grid_kpfm(grid)
        if kpfm_fig is not None:
            yield f"kpfm {grid.get('name', '')}", kpfm_fig

    flagged_fig = page_flagged(model)
    if flagged_fig is not None:
        yield "needs attention", flagged_fig


def render_report_pdf(model, out_path, progress_cb=None):
    """Write the full report to ``out_path``. ``progress_cb(done, total,
    description)`` is invoked per page when given. Returns the page count."""
    pages = list(iter_report_figures(model))
    total = len(pages)
    count = 0
    with PdfPages(str(out_path)) as pdf:
        for desc, fig in pages:
            pdf.savefig(fig)
            count += 1
            if progress_cb is not None:
                progress_cb(count, total, desc)
        meta = pdf.infodict()
        meta["Title"] = f"SXM folder report - {model['summary'].get('folder', '')}"
        meta["Creator"] = "SXM Viewer"
    return count
