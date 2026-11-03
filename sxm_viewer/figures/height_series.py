"""Height-series figure recipe (Qt-free).

A constant-height (CH) height series is one frame re-scanned at a series of
tip heights. The figure is fully determined by the data:

* columns = the series' images, ordered by tip height (closest approach
  left, so Δz >= 0 grows to the right with the default reference);
* rows    = channel kinds (current, then frequency shift / lock-in), picked
  with the same heuristics as the folder report
  (``reporting.channels.pick_image_channels``);
* one colour range per row (pooled over the row), one colourbar per row,
  one scale bar, Δz column headers, panel letters.

Pipeline: ``find_height_series`` (detection, cheap - header data + the CH
tag's stored height) -> ``build_spec`` (columns/rows/labels) ->
``render`` (one matplotlib renderer for the screen *and* vector export).
The GUI side (``gui/figures/``) only gathers plain inputs and arrays.

Heights: ``z_nm`` must be the same absolute tip height the preview HUD
shows (``ch_height.ch_abs_z_nm``), so captions can never contradict the
viewer. Flatness comes from the CH tag itself (the auto-tagger's topography
variance rule) - only CH-tagged frames take part.
"""
from __future__ import annotations

import hashlib
import math
import re
from datetime import datetime

import numpy as np

from ..reporting.channels import classify_channel
from ..reporting.model import build_image_sequences
from ..utils.units import _auto_display_unit

# Row kinds in display order; "df" groups frequency shift and lock-in so a
# member that only has one of them still lands in the second row.
ROW_KINDS = ("current", "df", "z", "other")
_ROW_OF_CLASS = {"current": "current", "freq": "df", "lockin": "df", "z": "z"}
ROW_CMAPS = {"current": "Blues_r", "df": "gray", "z": "afmhot", "other": "viridis"}
_ROW_SYMBOL = {"current": "I", "freq": "Δf", "lockin": "LIA", "z": "z"}

LETTERS = "abcdefghijklmnopqrstuvwxyz"

_BACKWARD_RE = re.compile(r"(?i)(bwd|backward|\bbk\b|_bk|retrace|\(b\)|<-|←)")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def _finite(value):
    try:
        f = float(value)
    except (TypeError, ValueError):
        return None
    return f if math.isfinite(f) else None


def series_id(keys):
    """Stable id for a set of image keys (order-independent)."""
    blob = "\n".join(sorted(str(k) for k in keys)).encode("utf-8", "replace")
    return hashlib.sha1(blob).hexdigest()[:12]


def is_backward(label):
    return bool(_BACKWARD_RE.search(str(label or "")))


def dz_text(z_nm, ref_nm):
    """``Δz = +20 pm`` (same rounding as ``format_ch_height_label``)."""
    z, ref = _finite(z_nm), _finite(ref_nm)
    if z is None or ref is None:
        return ""
    dz_pm = (z - ref) * 1e3
    if abs(dz_pm) < 0.5:
        return "Δz = 0 pm"
    if abs(dz_pm) < 1000.0:
        return f"Δz = {dz_pm:+.0f} pm"
    return f"Δz = {dz_pm / 1e3:+.3f} nm"


def bias_text(bias_v):
    b = _finite(bias_v)
    if b is None:
        return ""
    if abs(b) < 1.0:
        return f"{b * 1e3:.0f} mV"
    return f"{b:.2f} V"


def nice_length(target):
    """Round ``target`` down to 1/2/5 x 10^n."""
    if not target or target <= 0 or not math.isfinite(target):
        return None
    exp = math.floor(math.log10(target))
    base = 10.0 ** exp
    for m in (5.0, 2.0, 1.0):
        if m * base <= target:
            return m * base
    return base


def length_text(nm):
    if nm >= 1000:
        return f"{nm / 1000:g} µm"
    if nm < 1:
        return f"{nm * 10:g} Å"
    return f"{nm:g} nm"


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------

def find_height_series(images, min_members=2):
    """Height series among ``images``.

    ``images``: dicts ``{key, name, time, extent, angle, tag, bias_v, z_nm}``
    (``extent`` = header extent ``[x0, x1, y1, y0]`` in nm, ``z_nm`` = HUD
    absolute tip height). Only constant-height frames with a height take
    part. Frames re-scanning the same footprint are grouped by the folder
    report's own detector, so report and figure always agree on what a
    series is.

    Returns ``[{id, keys, kind, z_span_nm, bias_varies}]`` with ``keys``
    ordered by tip height (lowest z first).
    """
    inputs = []
    for img in images or []:
        if img.get("tag") != "constant-height":
            continue
        z = _finite(img.get("z_nm"))
        if z is None:
            continue
        entry = dict(img)
        entry["z_level_nm"] = z
        # z_range_nm = the CH tagger's topo max-min. A CH tag only means
        # "some rows are flat"; a frame whose Z stepped mid-scan has two
        # heights and no single Δz, so the report's flat gate drops its z
        # (it stays a member, reported as excluded). Unknown range (manual
        # tag without one) = trust the tag.
        rng = _finite(img.get("z_range_nm"))
        entry["z_spread_nm"] = rng if rng is not None else 0.0
        inputs.append(entry)
    if len(inputs) < min_members:
        return []
    by_key = {str(i["key"]): i for i in inputs}
    out = []
    for seq in build_image_sequences(inputs):
        if seq["kind"] not in ("height", "bias+height"):
            continue
        keys = [m["key"] for m in seq["members"] if m.get("z_level_nm") is not None]
        if len(keys) < min_members:
            continue
        keys.sort(key=lambda k: (by_key[k]["z_level_nm"], _time_sort(by_key[k])))
        excluded = [m["key"] for m in seq["members"] if m.get("z_level_nm") is None]
        out.append({
            "id": series_id(keys),
            "keys": keys,
            "excluded": [{"key": k, "reason": "z_not_flat",
                          "z_range_nm": _finite(by_key[k].get("z_range_nm"))}
                         for k in excluded],
            "kind": seq["kind"],
            "z_span_nm": float(seq["z_span_nm"]),
            "bias_varies": seq["kind"] == "bias+height",
        })
    return out


def _time_sort(img):
    t = img.get("time")
    return t if isinstance(t, datetime) else datetime.max


# ---------------------------------------------------------------------------
# Spec
# ---------------------------------------------------------------------------

def _row_kind_for(cls):
    return _ROW_OF_CLASS.get(cls, "other")


def _row_symbol(cls, label):
    sym = _ROW_SYMBOL.get(cls) or (str(label).split("(")[0].strip() or "?")
    return f"{sym} (bwd)" if is_backward(label) else sym


def build_spec(members, *, mode="height", reference_key=None, dropped=(),
               row_choices=None):
    """Columns/rows for the figure.

    ``members``: dicts ``{key, name, time, z_nm, bias_v, tag, channels,
    picks}`` where ``channels`` = ``[{idx, label, cls}]`` (every channel of
    the file) and ``picks`` = ``[(idx, cls)]`` from ``pick_image_channels``.

    ``mode``: ``"height"`` (columns by z, Δz headers) or ``"selection"``
    (fallback grid for an arbitrary selection: by z when every member has
    one, else by time; headers Δz or the file name).
    ``reference_key``: Δz zero (default: lowest z = closest approach).
    ``dropped``: keys removed by the user. ``row_choices``: ``{row_kind:
    channel label}`` overriding the default channel of a row (cycling
    forward/backward, freq/lock-in).
    """
    dropped = {str(k) for k in (dropped or ())}
    row_choices = dict(row_choices or {})
    cols_in = [m for m in members if str(m["key"]) not in dropped]
    has_z = bool(cols_in) and all(_finite(m.get("z_nm")) is not None for m in cols_in)
    if mode == "height" or has_z:
        cols_in.sort(key=lambda m: (_finite(m.get("z_nm")) if _finite(m.get("z_nm")) is not None else math.inf,
                                    _time_sort(m), str(m.get("name"))))
    else:
        cols_in.sort(key=lambda m: (_time_sort(m), str(m.get("name"))))

    keys = [str(m["key"]) for m in cols_in]
    ref_key = str(reference_key) if reference_key is not None and str(reference_key) in keys else None
    if ref_key is None and has_z and cols_in:
        ref_key = keys[0]
    ref_z = None
    if ref_key is not None:
        ref_z = _finite(next(m for m in cols_in if str(m["key"]) == ref_key).get("z_nm"))

    biases = [_finite(m.get("bias_v")) for m in cols_in]
    finite_b = [b for b in biases if b is not None]
    bias_varies = len(finite_b) >= 2 and (max(finite_b) - min(finite_b)) > 2e-3

    columns = []
    for i, m in enumerate(cols_in):
        key = str(m["key"])
        if has_z and ref_z is not None:
            header = dz_text(m.get("z_nm"), ref_z)
            dz_pm = (_finite(m["z_nm"]) - ref_z) * 1e3
        else:
            header = str(m.get("name") or key)
            if len(header) > 24:
                header = "…" + header[-23:]
            dz_pm = None
        sub = bias_text(m.get("bias_v")) if bias_varies else ""
        columns.append({
            "key": key,
            "name": str(m.get("name") or ""),
            "z_nm": _finite(m.get("z_nm")),
            "dz_pm": dz_pm,
            "bias_v": _finite(m.get("bias_v")),
            "is_ref": key == ref_key,
            "header": header,
            "subheader": sub,
            "letter": LETTERS[i] if i < len(LETTERS) else "",
        })

    # Rows: union of each member's default picks, in ROW_KINDS order.
    row_kinds = []
    default_label = {}   # row kind -> label votes
    for m in cols_in:
        labels = {int(c["idx"]): c for c in (m.get("channels") or [])}
        for idx, cls in (m.get("picks") or []):
            kind = _row_kind_for(cls)
            if kind not in row_kinds:
                row_kinds.append(kind)
            ch = labels.get(int(idx))
            if ch is not None:
                default_label.setdefault(kind, []).append(ch["label"])
    row_kinds.sort(key=lambda k: ROW_KINDS.index(k) if k in ROW_KINDS else len(ROW_KINDS))

    rows = []
    for kind in row_kinds:
        # Alternatives: every channel label of this row kind across members,
        # in first-seen order (forward before backward as files list them).
        alternatives = []
        for m in cols_in:
            for c in m.get("channels") or []:
                if _row_kind_for(c["cls"]) == kind and c["label"] not in alternatives:
                    alternatives.append(c["label"])
        votes = default_label.get(kind) or alternatives[:1]
        default = max(votes, key=votes.count) if votes else None
        choice = row_choices.get(kind)
        if choice not in alternatives:
            choice = default
        cells = {}
        cls_of_choice = None
        for m in cols_in:
            chans = m.get("channels") or []
            hit = next((c for c in chans if c["label"] == choice), None)
            if hit is None:
                # This member names the channel differently - fall back to
                # its own default pick for the row, then any of the kind.
                pick = next((idx for idx, cls in (m.get("picks") or [])
                             if _row_kind_for(cls) == kind), None)
                if pick is not None:
                    hit = next((c for c in chans if int(c["idx"]) == int(pick)), None)
                if hit is None:
                    hit = next((c for c in chans if _row_kind_for(c["cls"]) == kind), None)
            if hit is not None:
                cells[str(m["key"])] = int(hit["idx"])
                if cls_of_choice is None or hit["label"] == choice:
                    cls_of_choice = hit["cls"]
        if not cells:
            continue
        rows.append({
            "kind": kind,
            "cls": cls_of_choice or kind,
            "choice": choice,
            "alternatives": alternatives,
            "label": _row_symbol(cls_of_choice or kind, choice or ""),
            "cmap": ROW_CMAPS.get(kind, "viridis"),
            "cells": cells,
        })

    return {
        "mode": mode,
        "columns": columns,
        "rows": rows,
        "reference_key": ref_key,
        "has_z": has_z,
        "bias_varies": bias_varies,
    }


# ---------------------------------------------------------------------------
# Ranges / units
# ---------------------------------------------------------------------------

def _sample(arr, limit=250_000):
    a = np.asarray(arr, dtype=float).ravel()
    a = a[np.isfinite(a)]
    if a.size > limit:
        a = a[:: int(math.ceil(a.size / limit))]
    return a


def robust_range(values, pct=(1.0, 99.0)):
    v = _sample(values)
    if not v.size:
        return (0.0, 1.0)
    lo, hi = (float(x) for x in np.percentile(v, pct))
    if not hi > lo:
        span = abs(lo) * 1e-3 or 1e-12
        lo, hi = lo - span, hi + span
    return (lo, hi)


def row_display_unit(unit, arrays):
    """One display unit for a whole row (the viewer picks per image, which
    would put 4 pA and 0.4 nA next to each other on a 'shared' range)."""
    if not arrays:
        return unit or "", 1.0
    pooled = np.concatenate([_sample(a, 50_000) for a in arrays]) if arrays else np.array([])
    label, factor = _auto_display_unit(unit or "", pooled)
    return label or unit or "", float(factor)


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

# Figure geometry in inches (publication scale: ~1.6" panels, 7 pt text).
_PANEL_W = 1.55
_GAP = 0.05
_LEFT = 0.42
_TOP_BASE = 0.30
_TOP_SUB = 0.14
_RIGHT_CBAR = 0.80
_RIGHT_NONE = 0.12
_BOTTOM = 0.10
_CBAR_W = 0.07
_FONT = 7.0


MAX_COLS = 8         # longer series wrap into balanced blocks of columns
_BLOCK_GAP = 0.10


def column_blocks(n_cols, max_cols=MAX_COLS):
    """Column indices per block: <= max_cols per block, blocks balanced
    (16 -> 8+8, 13 -> 7+6) so the last block is never a lonely straggler."""
    if n_cols <= max_cols:
        return [list(range(n_cols))]
    n_blocks = int(math.ceil(n_cols / max_cols))
    per = int(math.ceil(n_cols / n_blocks))
    return [list(range(i, min(n_cols, i + per))) for i in range(0, n_cols, per)]


def figure_layout(spec, panels, *, range_mode="row"):
    """Panel boxes in inches; deterministic so the GUI can hit-test without
    asking matplotlib. Boxes are ``(x0, y0, x1, y1)`` in inches from the
    figure's *top-left*. ``row_labels``/``colorbars`` are ``[(row, box)]``
    (one per row per block once a long series wraps)."""
    n_cols = max(1, len(spec["columns"]))
    n_rows = max(1, len(spec["rows"]))
    aspect = 1.0
    for (_, _), p in panels.items():
        w, h = _finite(p.get("width_nm")), _finite(p.get("height_nm"))
        if w and h:
            aspect = h / w
            break
        shp = np.shape(p.get("data"))
        if len(shp) == 2 and shp[1]:
            aspect = shp[0] / shp[1]
            break
    aspect = min(max(aspect, 0.25), 4.0)
    pw = _PANEL_W
    ph = pw * aspect
    top = _TOP_BASE + (_TOP_SUB if any(c.get("subheader") for c in spec["columns"]) else 0.0)
    right = _RIGHT_CBAR if range_mode == "row" else _RIGHT_NONE
    blocks = column_blocks(n_cols)
    per = max(len(b) for b in blocks)
    block_h = top + n_rows * ph + (n_rows - 1) * _GAP
    fig_w = _LEFT + per * pw + (per - 1) * _GAP + right
    fig_h = len(blocks) * block_h + (len(blocks) - 1) * _BLOCK_GAP + _BOTTOM
    cells, headers, row_labels, cbars = {}, [None] * n_cols, [], []
    block_of = {}
    for b, cols in enumerate(blocks):
        y_base = b * (block_h + _BLOCK_GAP)
        for j, c in enumerate(cols):
            block_of[c] = b
            x0 = _LEFT + j * (pw + _GAP)
            headers[c] = (x0, y_base, x0 + pw, y_base + top)
            for r in range(n_rows):
                y0 = y_base + top + r * (ph + _GAP)
                cells[(r, c)] = (x0, y0, x0 + pw, y0 + ph)
        x_end = _LEFT + len(cols) * pw + (len(cols) - 1) * _GAP
        for r in range(n_rows):
            y0 = y_base + top + r * (ph + _GAP)
            row_labels.append((r, (0.0, y0, _LEFT, y0 + ph)))
            if range_mode == "row":
                cbars.append((r, (x_end + 0.06, y0, x_end + 0.06 + _CBAR_W, y0 + ph)))
    return {"size": (fig_w, fig_h), "cells": cells, "headers": headers,
            "row_labels": row_labels, "colorbars": cbars, "panel": (pw, ph),
            "blocks": blocks, "block_of": block_of}


def _ax_rect(box, size):
    fig_w, fig_h = size
    x0, y0, x1, y1 = box
    return [x0 / fig_w, 1.0 - y1 / fig_h, (x1 - x0) / fig_w, (y1 - y0) / fig_h]


def render(fig, spec, panels, *, range_mode="row", clims=None, letters=True,
           scale_bar=True, get_cmap=None, text_color="black"):
    """Draw the figure into ``fig`` (cleared first).

    ``panels``: ``{(row_index, key): {data, unit, width_nm, height_nm}}``
    with ``data`` displayed values in the row's *normalised* unit and row 0
    = top of the scan (as the preview shows it). ``clims``: ``{row_index:
    (lo, hi)}`` user overrides (row mode). Returns the layout dict plus the
    effective per-row ``ranges`` and display ``units``.
    """
    import matplotlib.patheffects as pe

    clims = dict(clims or {})
    layout = figure_layout(spec, panels, range_mode=range_mode)
    size = layout["size"]
    fig.clear()
    fig.set_size_inches(*size, forward=False)
    fig.patch.set_alpha(0.0)
    fig.set_facecolor("none")
    if get_cmap is None:
        import matplotlib
        get_cmap = matplotlib.colormaps.get_cmap
    stroke = [pe.withStroke(linewidth=1.6, foreground="black")]

    ranges, units, images, cbars = {}, {}, {}, {}
    own_scale = {}   # (row, key) -> "offset" | "flat"
    n_rows = len(spec["rows"])
    for r, row in enumerate(spec["rows"]):
        # Topography rows are relative: each frame offset to its own floor
        # (absolute z of different frames differs by the tip/sample drift,
        # so a range pooled over absolute heights shows nothing).
        relative = row["kind"] == "z"
        row_data = {c["key"]: (_relative(panels[(r, c["key"])]["data"]) if relative
                               else np.asarray(panels[(r, c["key"])]["data"], dtype=float))
                    for c in spec["columns"] if (r, c["key"]) in panels}
        unit = next((panels[(r, c["key"])].get("unit") for c in spec["columns"]
                     if (r, c["key"]) in panels), "")
        # Frames on a different baseline (e.g. a Δf reference that changed
        # between frames by kHz) would wreck a pooled range - they are drawn
        # on their own scale and tagged, the rest share the row range.
        outliers = baseline_outliers(row_data) if range_mode == "row" else {}
        for k, why in outliers.items():
            own_scale[(r, k)] = why
        inliers = [a for k, a in row_data.items() if k not in outliers] or list(row_data.values())
        raw_lo, raw_hi = pooled_range(inliers)
        d_unit, factor = row_display_unit(unit, [np.array([raw_lo, raw_hi])])
        units[r] = d_unit
        if range_mode == "row":
            if r in clims and clims[r] is not None:
                lo, hi = clims[r]
            else:
                lo, hi = raw_lo * factor, raw_hi * factor
            ranges[r] = (float(lo), float(hi))
        cmap = get_cmap(row["cmap"])
        im = None
        for c, col in enumerate(spec["columns"]):
            box = layout["cells"][(r, c)]
            ax = fig.add_axes(_ax_rect(box, size))
            ax.set_axis_off()
            p = panels.get((r, col["key"]))
            if p is None:
                ax.add_patch(_placeholder(ax))
                continue
            data = row_data[col["key"]] * factor
            own = (r, col["key"]) in own_scale
            if range_mode == "row" and not own:
                lo, hi = ranges[r]
            else:
                lo, hi = robust_range(_sample(data))
            w = _finite(p.get("width_nm")) or float(data.shape[1])
            h = _finite(p.get("height_nm")) or float(data.shape[0])
            im = ax.imshow(data, cmap=cmap, vmin=lo, vmax=hi, origin="upper",
                           interpolation="nearest", extent=(0.0, w, 0.0, h),
                           aspect="auto")
            if own:
                tag = "own scale"
                if own_scale[(r, col["key"])] == "flat":
                    # No signal: dim it so amplified noise can't pass as
                    # contrast, and say what the level actually is.
                    im.set_alpha(0.35)
                    med = _sample(data)
                    level = f" · {_tick(float(np.median(med)))} {d_unit}" if med.size else ""
                    tag = "flat signal" + level
                ax.text(0.96 * w, 0.04 * h, tag, color="white", fontsize=_FONT - 1,
                        ha="right", va="bottom", path_effects=stroke)
            else:
                images.setdefault(r, []).append(im)
            ax.set_xlim(0.0, w)
            ax.set_ylim(0.0, h)
            # True aspect inside the cell (a selection can mix frame sizes;
            # series frames share one footprint, so this is a no-op there).
            ax.set_aspect("equal", adjustable="box", anchor="C")
            if letters and r == 0 and col.get("letter"):
                ax.text(0.04 * w, 0.96 * h, col["letter"], color="white",
                        fontsize=_FONT + 2, fontweight="bold", va="top", ha="left",
                        path_effects=stroke)
            if (scale_bar and r == n_rows - 1 and c == layout["blocks"][-1][0]
                    and _finite(p.get("width_nm"))):
                _draw_scale_bar(ax, w, h, stroke)
        if range_mode == "row" and im is not None:
            for rr, cbox in layout["colorbars"]:
                if rr != r:
                    continue
                cax = fig.add_axes(_ax_rect(cbox, size))
                cb = fig.colorbar(im, cax=cax)
                set_colorbar_range(cb, *ranges[r])
                cbars.setdefault(r, []).append(cb)
                cb.ax.tick_params(labelsize=_FONT - 0.5, length=2, width=0.5, pad=1.5,
                                  colors=text_color)
                cb.outline.set_linewidth(0.5)
                cb.set_label(f"{row['label']} ({d_unit})" if d_unit else row["label"],
                             fontsize=_FONT, labelpad=1.0, color=text_color)

    # Column headers and row labels as figure text.
    fig_w, fig_h = size
    for c, col in enumerate(spec["columns"]):
        x0, y0, x1, y1 = layout["headers"][c]
        xc = 0.5 * (x0 + x1) / fig_w
        has_sub = bool(col.get("subheader"))
        y_main = 1.0 - (y1 - 0.06 - (_TOP_SUB if has_sub else 0.0)) / fig_h
        fig.text(xc, y_main, col["header"], ha="center", va="bottom",
                 fontsize=_FONT, color=text_color,
                 fontweight="bold" if col.get("is_ref") else "normal")
        if has_sub:
            fig.text(xc, 1.0 - (y1 - 0.05) / fig_h, col["subheader"], ha="center",
                     va="bottom", fontsize=_FONT - 1, color=text_color)
    for r, (x0, y0, x1, y1) in layout["row_labels"]:
        row = spec["rows"][r]
        fig.text((x1 - 0.08) / fig_w, 1.0 - 0.5 * (y0 + y1) / fig_h, row["label"],
                 ha="right", va="center", rotation=90 if len(row["label"]) > 4 else 0,
                 fontsize=_FONT + 1, color=text_color)
    layout["ranges"] = ranges
    layout["units"] = units
    layout["images"] = images
    layout["colorbar_artists"] = cbars
    layout["own_scale"] = own_scale
    return layout


def baseline_outliers(arrays_by_key, factor=3.0):
    """Frames that must not share the row range: ``{key: "offset"|"flat"}``.

    * ``"offset"`` - the frame's median sits more than ``factor`` typical
      spans from the row's median (a Δf reference that changed by kHz, a
      PLL that lost lock): pooled with the rest it would flatten everyone.
    * ``"flat"`` - additionally (almost) no signal of its own (< 5 % of the
      typical span): drawn on its own scale it would only show noise, so
      it is tagged "flat signal".

    Centres are frame medians and the typical span is the 75th percentile of
    the frames' p1..p99 spans, so a few dead frames can't shrink it and a
    frame that jumps mid-scan can't define it. Needs >= 3 frames.
    """
    if len(arrays_by_key) < 3:
        return {}
    stats = {}
    for k, a in arrays_by_key.items():
        v = _sample(a, 40_000)
        if v.size:
            lo, med, hi = np.percentile(v, (1.0, 50.0, 99.0))
            stats[k] = (float(med), float(hi - lo))
    if len(stats) < 3:
        return {}
    centres = np.array([c for c, _ in stats.values()])
    spans = np.array([sp for _, sp in stats.values()])
    m = float(np.median(centres))
    live = spans[spans > 0]
    typical = float(np.percentile(live, 75.0)) if live.size else 0.0
    if typical <= 0:
        return {}
    out = {}
    for k, (c, sp) in stats.items():
        if abs(c - m) > factor * typical:
            out[k] = "flat" if sp < 0.05 * typical else "offset"
    return out


def pooled_range(arrays, pct=(1.0, 99.0), per_frame=20_000):
    """p1..p99 over several frames, each frame weighted equally (a 240 px
    frame must not outvote a 64 px one)."""
    parts = []
    for a in arrays:
        v = _sample(a, 10 * per_frame)
        if v.size > per_frame:
            v = v[np.linspace(0, v.size - 1, per_frame).astype(int)]
        if v.size:
            parts.append(v)
    if not parts:
        return (0.0, 1.0)
    return robust_range(np.concatenate(parts), pct)


def _relative(arr):
    a = np.asarray(arr, dtype=float)
    v = _sample(a)
    return a - float(np.percentile(v, 1.0)) if v.size else a


def set_colorbar_range(cb, lo, hi):
    """Two end ticks only - a figure colourbar, not an axis."""
    span = abs(hi - lo)
    cb.set_ticks([lo, hi])
    cb.set_ticklabels([_tick(lo, span), _tick(hi, span)])


def set_row_range(layout, r, lo, hi):
    """Cheap in-place range update for one row (colourbar drag)."""
    for im in layout.get("images", {}).get(r, []):
        im.set_clim(lo, hi)
    for cb in layout.get("colorbar_artists", {}).get(r, []):
        set_colorbar_range(cb, lo, hi)
    layout.setdefault("ranges", {})[r] = (lo, hi)


def _tick(v, span=None):
    a = abs(v)
    if a == 0 or (span and a < 5e-3 * span):
        return "0"
    if a >= 100:
        return f"{v:.0f}"
    if a >= 10:
        return f"{v:.1f}"
    if a >= 0.01:
        return f"{v:.2f}"
    return f"{v:.1e}"


def _placeholder(ax):
    from matplotlib.patches import Rectangle
    return Rectangle((0, 0), 1, 1, transform=ax.transAxes, facecolor="#8882",
                     edgecolor="#8886", linewidth=0.5, hatch="///")


def _draw_scale_bar(ax, w, h, stroke):
    length = nice_length(0.3 * w)
    if not length:
        return
    x0 = 0.06 * w
    y0 = 0.07 * h
    bar_h = max(0.018 * h, 1e-9)
    from matplotlib.patches import Rectangle
    ax.add_patch(Rectangle((x0, y0), length, bar_h, facecolor="white",
                           edgecolor="black", linewidth=0.4, zorder=5))
    ax.text(x0 + 0.5 * length, y0 + bar_h + 0.02 * h, length_text(length),
            color="white", fontsize=_FONT, ha="center", va="bottom", zorder=6,
            path_effects=stroke)
