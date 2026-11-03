"""Height-series figure recipe (Qt-free).

A constant-height (CH) height series is one frame re-scanned at a series of
tip heights. The figure is determined by the data:

* frames   = the series' images, ordered by tip height (closest approach
  first, so Δz >= 0 grows along the frame axis with the default reference);
* channels = what each frame shows (current, frequency shift, lock-in,
  topography, ... any simultaneously acquired channel). Defaults come from
  the folder report's heuristics (``reporting.channels.pick_image_channels``);
  the user can add, remove, reorder and switch them;
* frames run along columns and channels along rows - or the other way round
  (``style["transpose"]``).

Pipeline: ``find_height_series`` (detection, cheap - header data + the CH
tag's stored height) -> ``build_spec`` (frames/channels/labels) -> ``render``
(one matplotlib renderer for the screen *and* vector export), driven by one
``style`` dict (``STYLE_DEFAULTS``). The GUI side only gathers plain inputs
and arrays.

Naming inside the code: ``spec["columns"]`` are the *frames* and
``spec["rows"]`` the *channels*, whatever the on-page orientation is.

Heights: ``z_nm`` must be the same absolute tip height the preview HUD
shows (``ch_height.ch_abs_z_nm``), so captions can never contradict the
viewer.
"""
from __future__ import annotations

import hashlib
import math
import re
from datetime import datetime

import numpy as np

from ..reporting.model import build_image_sequences
from ..utils.units import _auto_display_unit

# Channel kinds in default display order; "df" groups frequency shift and
# lock-in so a member that only has one of them still lands in that row.
ROW_KINDS = ("current", "df", "z", "other")
_ROW_OF_CLASS = {"current": "current", "freq": "df", "lockin": "df", "z": "z"}
ROW_CMAPS = {"current": "Blues_r", "df": "gray", "z": "afmhot", "other": "viridis"}
_CLASS_CMAPS = {"lockin": "RdBu_r", "phase": "twilight", "amplitude": "viridis",
                "excitation": "magma"}
_ROW_SYMBOL = {"current": "I", "freq": "Δf", "lockin": "LIA", "z": "z",
               "amplitude": "A", "phase": "φ", "excitation": "Drive"}

LETTERS = "abcdefghijklmnopqrstuvwxyz"

_BACKWARD_RE = re.compile(r"(?i)(bwd|backward|\bbk\b|_bk|retrace|\(b\)|<-|←)")

STYLE_DEFAULTS = {
    "transpose": False,        # False: frames = columns, channels = rows
    "reverse": False,          # reverse the frame order (z / bias / time)
    "range_mode": "row",       # "row": one range per channel | "image"
    "colorbars": True,
    "cbar_ticks": 2,           # 2 (ends) | 3 | 5
    "cbar_label": True,
    "letters": True,
    "letter_style": "a",       # "a" | "A" | "(a)" | "a)" | "a."
    "letter_scope": "frame",   # "frame": one per frame | "panel": every panel
    "letter_pos": "tl",        # tl | tr | bl | br
    "letter_color": "white",   # white | black
    "letter_size": 9.0,
    "letter_bold": True,
    "scale_bar": "one",        # "one" | "all" | "off"
    "scale_len_nm": None,      # None = automatic
    "scale_pos": "bl",
    "scale_color": "white",
    "scale_label": True,
    "header": "auto",          # "auto" | "dz" | "zabs" | "bias" | "name"
    "header_name": False,      # file name as an extra header line
    "font": 7.0,
    "panel_w": 1.55,           # inches
}

# Which style keys are meaningful in each range mode (for the UI).
LETTER_STYLES = ("a", "A", "(a)", "a)", "a.")
CORNERS = ("tl", "tr", "bl", "br")


def resolve_style(style):
    out = dict(STYLE_DEFAULTS)
    for k, v in (style or {}).items():
        if k in out and v is not None or k == "scale_len_nm":
            out[k] = v
    return out


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


def bias_header(bias_v):
    """``V = +300 mV`` / ``V = -1.20 V`` - always signed (contrast inversion
    across zero is usually the point of a bias series)."""
    b = _finite(bias_v)
    if b is None:
        return ""
    if abs(b) < 0.5e-3:
        return "V = 0 mV"
    if abs(b) < 1.0:
        return f"V = {b * 1e3:+.0f} mV"
    return f"V = {b:+.2f} V"


_SI_PREFIXES = ((1e-15, "f"), (1e-12, "p"), (1e-9, "n"), (1e-6, "µ"), (1e-3, "m"),
                (1.0, ""), (1e3, "k"), (1e6, "M"))


def si_text(value, unit):
    """``5e-11, 'A'`` -> ``50 pA`` (unit given as a base SI unit)."""
    v = _finite(value)
    if v is None:
        return ""
    unit = str(unit or "")
    if v == 0:
        return f"0 {unit}".strip()
    for scale, prefix in reversed(_SI_PREFIXES):
        if abs(v) >= scale * 0.9995:
            return f"{v / scale:.3g} {prefix}{unit}".strip()
    scale, prefix = _SI_PREFIXES[0]
    return f"{v / scale:.3g} {prefix}{unit}".strip()


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


def letter_text(i, style):
    if i >= len(LETTERS):
        return ""
    ch = LETTERS[i]
    if style == "A":
        return ch.upper()
    if style == "(a)":
        return f"({ch})"
    if style == "a)":
        return f"{ch})"
    if style == "a.":
        return f"{ch}."
    return ch


def _short(text, n=24):
    text = str(text or "")
    return text if len(text) <= n else "…" + text[-(n - 1):]


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------

def find_height_series(images, min_members=2):
    """Height series among ``images``.

    ``images``: dicts ``{key, name, time, extent, angle, tag, bias_v, z_nm,
    z_range_nm}`` (``extent`` = header extent ``[x0, x1, y1, y0]`` in nm,
    ``z_nm`` = HUD absolute tip height, ``z_range_nm`` = the CH tagger's
    topo max-min). Only constant-height frames with a height take part.
    Frames re-scanning the same footprint are grouped by the folder report's
    own detector, so report and figure always agree on what a series is.

    Returns ``[{id, keys, excluded, kind, z_span_nm, bias_varies}]`` with
    ``keys`` ordered by tip height (lowest z first).
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
        # A CH tag only means "some rows are flat"; a frame whose Z stepped
        # mid-scan has two heights and no single Δz, so the report's flat
        # gate drops its z (it stays a member, reported as excluded).
        # Unknown range (manual tag without one) = trust the tag.
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
        levels = len({condition_key(by_key[k]) for k in keys})
        if levels < min_members:
            continue  # only re-scans of one height: not a series
        excluded = [m["key"] for m in seq["members"] if m.get("z_level_nm") is None]
        out.append({
            "id": series_id(keys),
            "keys": keys,
            "excluded": [{"key": k, "reason": "z_not_flat",
                          "z_range_nm": _finite(by_key[k].get("z_range_nm"))}
                         for k in excluded],
            "kind": seq["kind"],
            "z_span_nm": float(seq["z_span_nm"]),
            "levels": levels,
            "bias_varies": seq["kind"] == "bias+height",
            "recipe": "height",
        })
    return out


def find_bias_series(images, min_members=2):
    """Bias series among ``images``: the same frame re-scanned while only the
    sample bias steps (STM topography vs V in constant current, or a
    constant-height frame at a ladder of biases).

    Constant-height frames and constant-current / untagged frames are
    grouped *separately*, so a CC bias ladder and a CH height ladder of the
    same spot don't merge into one mixed "bias+height" group (that mix is
    a height series with a varying bias, handled by ``find_height_series``).
    CH frames take part with their flat-gated height (as in
    ``find_height_series``); CC frames have no controlled height.

    Returns ``[{id, keys, excluded, kind, bias_span_v, recipe}]`` with
    ``keys`` ordered by bias (most negative first).
    """
    ch, other = [], []
    for img in images or []:
        entry = dict(img)
        if img.get("tag") == "constant-height" and _finite(img.get("z_nm")) is not None:
            rng = _finite(img.get("z_range_nm"))
            entry["z_level_nm"] = _finite(img.get("z_nm"))
            entry["z_spread_nm"] = rng if rng is not None else 0.0
            ch.append(entry)
        else:
            entry["z_level_nm"] = None
            entry["z_spread_nm"] = None
            other.append(entry)
    out = []
    for group in (ch, other):
        if len(group) < min_members:
            continue
        by_key = {str(i["key"]): i for i in group}
        for seq in build_image_sequences(group):
            if seq["kind"] != "bias":
                continue
            keys = [m["key"] for m in seq["members"] if _finite(m.get("bias_v")) is not None]
            if len(keys) < min_members:
                continue
            keys.sort(key=lambda k: (_finite(by_key[k].get("bias_v")), _time_sort(by_key[k])))
            levels = len({condition_key(by_key[k]) for k in keys})
            if levels < min_members:
                continue
            out.append({
                "id": "bias-" + series_id(keys),
                "keys": keys,
                "excluded": [],
                "kind": "bias",
                "bias_span_v": float(seq["bias_span_v"]),
                "levels": levels,
                "recipe": "bias",
            })
    return out


def condition_key(m):
    """Acquisition conditions of a frame: bias (mV), tip height (pm, CH
    only) and setpoint (3 significant digits). Same key = a repeat."""
    b = _finite(m.get("bias_v"))
    z = _finite(m.get("z_nm"))
    sp = _finite(m.get("setpoint"))
    return (None if b is None else int(round(b * 1e3)),
            None if z is None else int(round(z * 1e3)),
            None if sp is None else float(f"{sp:.3g}"))


def collapse_repeats(members):
    """One frame per acquisition condition - the latest one (re-scans of the
    same frame are usually taken until it is stable). Returns ``(kept,
    hidden_keys)`` keeping the input order of the kept frames."""
    latest = {}
    for m in members:
        k = condition_key(m)
        cur = latest.get(k)
        if cur is None or _time_sort(m) >= _time_sort(cur) and _time_sort(m) != datetime.max:
            latest[k] = m
    kept_ids = {id(m) for m in latest.values()}
    kept = [m for m in members if id(m) in kept_ids]
    hidden = [str(m["key"]) for m in members if id(m) not in kept_ids]
    return kept, hidden


def _time_sort(img):
    t = img.get("time")
    return t if isinstance(t, datetime) else datetime.max


# ---------------------------------------------------------------------------
# Spec
# ---------------------------------------------------------------------------

def row_kind_for(cls):
    return _ROW_OF_CLASS.get(cls, "other")


def _symbol(cls, label):
    sym = _ROW_SYMBOL.get(cls)
    if not sym:
        sym = re.sub(r"(?i)(fwd|bwd|forward|backward)$", "", str(label or "")).strip(" _()") or "?"
    return f"{sym} (bwd)" if is_backward(label) else sym


def _varies(values, rel=0.01, abs_tol=0.0):
    vals = [v for v in values if v is not None]
    if len(vals) < 2:
        return False
    span = max(vals) - min(vals)
    return span > max(abs_tol, rel * max(abs(v) for v in vals))


def resolve_header(header, mode, has_z, bias_varies):
    """``"auto"`` -> the recipe's natural header: Δz for a height series,
    the bias for a bias series; a selection takes Δz, else bias, else name."""
    if header and header != "auto":
        return header
    if mode == "bias":
        return "bias"
    if mode == "height" or has_z:
        return "dz"
    return "bias" if bias_varies else "name"


def build_spec(members, *, mode="height", reference_key=None, dropped=(),
               rows=None, row_choices=None, header="auto", header_name=False,
               row_cmaps=None, reverse=False, show_repeats=False):
    """Frames (``columns``) and channels (``rows``) for the figure.

    ``members``: dicts ``{key, name, time, z_nm, bias_v, setpoint,
    setpoint_unit, tag, channels, picks}`` - ``channels`` = ``[{idx, label,
    cls}]`` (every channel of the file), ``picks`` = ``[(idx, cls)]`` from
    ``pick_image_channels``.

    ``mode``: ``"height"`` (frames by tip height), ``"bias"`` (frames by
    sample bias) or ``"selection"`` (by z when every member has one, else by
    bias when it varies, else by time). ``reverse`` flips the frame order.
    ``reference_key``: Δz zero (default: lowest z = closest approach).
    ``dropped``: frame keys removed by the user.
    ``rows``: channel labels to show, in order (None = defaults);
    ``row_choices`` (legacy ``{kind: label}``) is only used for defaults.
    ``header``: ``"auto"`` | ``"dz"`` | ``"zabs"`` | ``"bias"`` | ``"name"``;
    ``header_name`` adds the file name as an extra line.
    ``row_cmaps``: ``{label: cmap}``.
    """
    dropped = {str(k) for k in (dropped or ())}
    row_choices = dict(row_choices or {})
    row_cmaps = dict(row_cmaps or {})
    cols_in = [m for m in members if str(m["key"]) not in dropped]
    hidden_repeats = []
    if mode != "selection" and not show_repeats:
        # Re-scans under identical conditions add columns, not information:
        # keep the latest of each (``show_repeats`` brings them back).
        cols_in, hidden_repeats = collapse_repeats(cols_in)
    has_z = bool(cols_in) and all(_finite(m.get("z_nm")) is not None for m in cols_in)
    biases = [_finite(m.get("bias_v")) for m in cols_in]
    bias_varies = _varies(biases, rel=0.0, abs_tol=2e-3)
    has_bias = bool(cols_in) and all(b is not None for b in biases)

    def by_value(field):
        return lambda m: (_finite(m.get(field)) if _finite(m.get(field)) is not None else math.inf,
                          _time_sort(m), str(m.get("name")))

    if mode == "bias" or (mode == "selection" and not has_z and has_bias and bias_varies):
        cols_in.sort(key=by_value("bias_v"))
        axis = "bias"
    elif mode == "height" or has_z:
        cols_in.sort(key=by_value("z_nm"))
        axis = "z"
    else:
        cols_in.sort(key=lambda m: (_time_sort(m), str(m.get("name"))))
        axis = "time"
    if reverse:
        cols_in.reverse()
    header = resolve_header(header, mode, has_z, bias_varies)

    keys = [str(m["key"]) for m in cols_in]
    ref_key = str(reference_key) if reference_key is not None and str(reference_key) in keys else None
    if ref_key is None and has_z and cols_in:
        # Closest approach (lowest z), whatever the display order.
        ref_key = str(min(cols_in, key=by_value("z_nm"))["key"])
    ref_z = None
    if ref_key is not None:
        ref_z = _finite(next(m for m in cols_in if str(m["key"]) == ref_key).get("z_nm"))

    setpoints = [_finite(m.get("setpoint")) for m in cols_in]
    setpoint_varies = _varies(setpoints, rel=0.01)

    columns = []
    for m in cols_in:
        key = str(m["key"])
        name = str(m.get("name") or key)
        z = _finite(m.get("z_nm"))
        b = _finite(m.get("bias_v"))
        dz_pm = (z - ref_z) * 1e3 if (z is not None and ref_z is not None) else None
        if header == "zabs" and z is not None:
            text = f"z = {z:.3f} nm"
        elif header == "bias" and b is not None:
            text = bias_header(b)
        elif header == "dz" and has_z and ref_z is not None:
            text = dz_text(z, ref_z)
        else:
            text = _short(name)
        subs = []
        if header_name and text != _short(name):
            subs.append(_short(name, 28))
        if bias_varies and header != "bias":
            subs.append(bias_text(b))
        if setpoint_varies:
            subs.append(si_text(m.get("setpoint"), m.get("setpoint_unit") or ""))
        if header == "bias" and has_z and _varies([_finite(c.get("z_nm")) for c in cols_in], 0.0, 5e-3):
            # A bias series whose tip height also moved: say by how much.
            subs.append(dz_text(z, ref_z))
        columns.append({
            "key": key, "name": name, "z_nm": z, "dz_pm": dz_pm, "bias_v": b,
            "is_ref": key == ref_key and header == "dz",
            "header": text,
            "subheader": "  ·  ".join(s for s in subs if s),
        })

    # Every channel the frames have, in first-seen order (menus / Add row).
    channels, cls_of = [], {}
    for m in cols_in:
        for c in m.get("channels") or []:
            if c["label"] not in cls_of:
                cls_of[c["label"]] = c["cls"]
                channels.append({"label": c["label"], "cls": c["cls"],
                                 "kind": row_kind_for(c["cls"])})

    labels = list(rows) if rows else _default_rows(cols_in, row_choices)
    out_rows = []
    for label in labels:
        if label not in cls_of:
            continue
        cls = cls_of[label]
        kind = row_kind_for(cls)
        cells = {}
        for m in cols_in:
            hit = next((c for c in (m.get("channels") or []) if c["label"] == label), None)
            if hit is None:
                # This member names the channel differently - same class,
                # same direction (fwd/bwd) as the row's channel.
                hit = next((c for c in (m.get("channels") or [])
                            if c["cls"] == cls and is_backward(c["label"]) == is_backward(label)
                            and cls != "other"), None)
            if hit is not None:
                cells[str(m["key"])] = int(hit["idx"])
        if not cells:
            continue
        out_rows.append({
            "label": label,          # channel identity (caption)
            "kind": kind,
            "cls": cls,
            "symbol": _symbol(cls, label),
            "cmap": row_cmaps.get(label) or _CLASS_CMAPS.get(cls) or ROW_CMAPS.get(kind, "viridis"),
            "alternatives": [c["label"] for c in channels if c["kind"] == kind and c["label"] != label],
            "cells": cells,
        })
    # Two rows sharing a symbol (I from It_to_PC and It_ext) show their caption.
    counts = {}
    for r in out_rows:
        counts[r["symbol"]] = counts.get(r["symbol"], 0) + 1
    for r in out_rows:
        r["display"] = r["symbol"] if counts[r["symbol"]] == 1 else _short(r["label"], 16)

    return {
        "mode": mode,
        "columns": columns,
        "rows": out_rows,
        "channels": channels,
        "reference_key": ref_key,
        "has_z": has_z,
        "bias_varies": bias_varies,
        "header": header,
        "axis": axis,
        "hidden_repeats": hidden_repeats,
    }


def _default_rows(cols_in, row_choices):
    """Majority default pick per channel kind, in ROW_KINDS order."""
    votes = {}
    for m in cols_in:
        by_idx = {int(c["idx"]): c for c in (m.get("channels") or [])}
        for idx, cls in (m.get("picks") or []):
            ch = by_idx.get(int(idx))
            if ch is not None:
                votes.setdefault(row_kind_for(ch["cls"]), []).append(ch["label"])
    kinds = sorted(votes, key=lambda k: ROW_KINDS.index(k) if k in ROW_KINDS else len(ROW_KINDS))
    out = []
    for kind in kinds:
        choice = row_choices.get(kind)
        out.append(choice if choice else max(votes[kind], key=votes[kind].count))
    return out


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
    pooled = np.concatenate([_sample(a, 50_000) for a in arrays])
    label, factor = _auto_display_unit(unit or "", pooled)
    return label or unit or "", float(factor)


def baseline_outliers(arrays_by_key, factor=3.0):
    """Frames that must not share the row range: ``{key: "offset"|"flat"}``.

    * ``"offset"`` - the frame's median sits more than ``factor`` typical
      spans from the row's median (a Δf reference that changed by kHz, a
      PLL that lost lock): pooled with the rest it would flatten everyone.
    * ``"flat"`` - additionally (almost) no signal of its own (< 5 % of the
      typical span): drawn on its own scale it would only show noise, so
      it is dimmed and tagged "flat signal".

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


def _flat_row_level(arrays_by_key):
    """Median level if *every* frame of a channel is flat (p1..p99 span
    ~0 relative to its magnitude), else None."""
    if not arrays_by_key:
        return None
    levels = []
    for a in arrays_by_key.values():
        v = _sample(a, 40_000)
        if not v.size:
            continue
        lo, med, hi = np.percentile(v, (1.0, 50.0, 99.0))
        if (hi - lo) > 1e-9 * max(abs(med), 1e-12) and (hi - lo) > 0:
            return None
        levels.append(float(med))
    return float(np.median(levels)) if levels else None


def typical_range(arrays, pct=(2.0, 98.0)):
    """Median of the frames' own robust ranges."""
    los, his = [], []
    for a in arrays:
        v = _sample(a, 40_000)
        if v.size:
            lo, hi = np.percentile(v, pct)
            los.append(float(lo))
            his.append(float(hi))
    if not los:
        return (0.0, 1.0)
    lo, hi = float(np.median(los)), float(np.median(his))
    if not hi > lo:
        return robust_range(np.concatenate([_sample(a, 40_000) for a in arrays]))
    return (lo, hi)


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


def clim_key(row_label, frame_key=None):
    """Key of a user colour-range override (per channel, or per panel)."""
    return str(row_label) if frame_key is None else f"{row_label}|{frame_key}"


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------

MAX_FRAMES_PER_BLOCK = 8   # longer series wrap into balanced blocks
_GAP = 0.05
_BLOCK_GAP = 0.12
_CBAR_W = 0.07
_PIMG_CBAR = 0.46          # extra width per panel for a per-image colourbar
_BOTTOM = 0.10


def column_blocks(n, max_per=MAX_FRAMES_PER_BLOCK):
    """Frame indices per block: <= max_per per block, blocks balanced
    (16 -> 8+8, 13 -> 7+6) so the last block is never a lonely straggler."""
    if n <= max_per:
        return [list(range(n))]
    n_blocks = int(math.ceil(n / max_per))
    per = int(math.ceil(n / n_blocks))
    return [list(range(i, min(n, i + per))) for i in range(0, n, per)]


def _panel_aspect(panels):
    for p in panels.values():
        w, h = _finite(p.get("width_nm")), _finite(p.get("height_nm"))
        if w and h:
            return min(max(h / w, 0.25), 4.0)
        shp = np.shape(p.get("data"))
        if len(shp) == 2 and shp[1]:
            return min(max(shp[0] / shp[1], 0.25), 4.0)
    return 1.0


def figure_layout(spec, panels, style=None):
    """Element boxes in inches, ``(x0, y0, x1, y1)`` from the figure's
    *top-left*; deterministic so the GUI hit-tests without matplotlib.

    Returns ``size``, ``cells`` {(row, frame): box}, ``frame_labels``
    [(frame, box)], ``channel_labels`` [(row, box)], ``colorbars``
    [{row, frame|None, box, orient}] and ``transpose``.
    """
    st = resolve_style(style)
    font = float(st["font"])
    n_f = max(1, len(spec["columns"]))
    n_ch = max(1, len(spec["rows"]))
    pw = float(st["panel_w"])
    ph = pw * _panel_aspect(panels)
    per_img_cb = st["colorbars"] and st["range_mode"] == "image"
    shared_cb = st["colorbars"] and st["range_mode"] == "row"
    slot = pw + (_PIMG_CBAR if per_img_cb else 0.0)
    has_sub = any(c.get("subheader") for c in spec["columns"])
    line = font / 72.0 * 1.45
    blocks = column_blocks(n_f)
    per = max(len(b) for b in blocks)
    cells, frame_labels, channel_labels, cbars = {}, [], [], []

    if not st["transpose"]:
        top = 0.12 + line * (2 if has_sub else 1)
        longest = max((len(r.get("display") or "") for r in spec["rows"]), default=1)
        left = 0.42 if longest <= 4 else 0.30
        right = (0.06 + _CBAR_W + 0.68) if shared_cb else 0.12
        block_h = top + n_ch * ph + (n_ch - 1) * _GAP
        fig_w = left + per * slot + (per - 1) * _GAP + right
        fig_h = len(blocks) * block_h + (len(blocks) - 1) * _BLOCK_GAP + _BOTTOM
        for b, frames in enumerate(blocks):
            y_base = b * (block_h + _BLOCK_GAP)
            for j, f in enumerate(frames):
                x0 = left + j * (slot + _GAP)
                frame_labels.append((f, (x0, y_base, x0 + pw, y_base + top)))
                for r in range(n_ch):
                    y0 = y_base + top + r * (ph + _GAP)
                    cells[(r, f)] = (x0, y0, x0 + pw, y0 + ph)
                    if per_img_cb:
                        cbars.append({"row": r, "frame": f, "orient": "v",
                                      "box": (x0 + pw + 0.05, y0, x0 + pw + 0.05 + _CBAR_W, y0 + ph)})
            x_end = left + len(frames) * slot + (len(frames) - 1) * _GAP
            for r in range(n_ch):
                y0 = y_base + top + r * (ph + _GAP)
                channel_labels.append((r, (0.0, y0, left, y0 + ph)))
                if shared_cb:
                    cbars.append({"row": r, "frame": None, "orient": "v",
                                  "box": (x_end + 0.06, y0, x_end + 0.06 + _CBAR_W, y0 + ph)})
    else:
        top = 0.10 + line
        left = 0.20 + max(0.75, font / 72.0 * 0.62 * max(
            (len(c.get("header") or "") for c in spec["columns"]), default=8))
        bottom = (0.06 + _CBAR_W + 0.40) if shared_cb else _BOTTOM
        block_w = left + n_ch * slot + (n_ch - 1) * _GAP + 0.12
        fig_w = len(blocks) * block_w + (len(blocks) - 1) * _BLOCK_GAP
        fig_h = top + per * ph + (per - 1) * _GAP + bottom
        for b, frames in enumerate(blocks):
            x_base = b * (block_w + _BLOCK_GAP)
            y_end = top + len(frames) * ph + (len(frames) - 1) * _GAP
            for r in range(n_ch):
                x0 = x_base + left + r * (slot + _GAP)
                channel_labels.append((r, (x0, 0.0, x0 + pw, top)))
                if shared_cb:
                    cbars.append({"row": r, "frame": None, "orient": "h",
                                  "box": (x0, y_end + 0.06, x0 + pw, y_end + 0.06 + _CBAR_W)})
            for j, f in enumerate(frames):
                y0 = top + j * (ph + _GAP)
                frame_labels.append((f, (x_base, y0, x_base + left - 0.06, y0 + ph)))
                for r in range(n_ch):
                    x0 = x_base + left + r * (slot + _GAP)
                    cells[(r, f)] = (x0, y0, x0 + pw, y0 + ph)
                    if per_img_cb:
                        cbars.append({"row": r, "frame": f, "orient": "v",
                                      "box": (x0 + pw + 0.05, y0, x0 + pw + 0.05 + _CBAR_W, y0 + ph)})
    return {"size": (fig_w, fig_h), "cells": cells, "frame_labels": frame_labels,
            "channel_labels": channel_labels, "colorbars": cbars,
            "transpose": bool(st["transpose"]), "blocks": blocks, "panel": (pw, ph)}


def _ax_rect(box, size):
    fig_w, fig_h = size
    x0, y0, x1, y1 = box
    return [x0 / fig_w, 1.0 - y1 / fig_h, (x1 - x0) / fig_w, (y1 - y0) / fig_h]


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def render(fig, spec, panels, *, style=None, clims=None, get_cmap=None,
           text_color="black"):
    """Draw the figure into ``fig`` (cleared first).

    ``panels``: ``{(row_index, key): {data, unit, width_nm, height_nm}}``
    with ``data`` in the channel's *normalised* unit and row 0 = top of the
    scan (as the preview shows it). ``clims``: user overrides keyed by
    ``clim_key(row_label)`` (shared mode) or ``clim_key(row_label, key)``
    (per image), values in the row's display unit.
    Returns the layout dict plus ``ranges`` {row: (lo, hi)} (shared mode),
    ``panel_ranges`` {(row, frame): (lo, hi)}, ``units``, ``own_scale``.
    """
    import matplotlib.patheffects as pe

    st = resolve_style(style)
    font = float(st["font"])
    clims = dict(clims or {})
    layout = figure_layout(spec, panels, st)
    size = layout["size"]
    fig.clear()
    fig.set_size_inches(*size, forward=False)
    fig.patch.set_alpha(0.0)
    fig.set_facecolor("none")
    if get_cmap is None:
        import matplotlib
        get_cmap = matplotlib.colormaps.get_cmap

    def stroke(color):
        other = "black" if color == "white" else "white"
        return [pe.withStroke(linewidth=1.6, foreground=other)]

    shared = st["range_mode"] == "row"
    columns = spec["columns"]
    n_rows = len(spec["rows"])

    # Letter / scale-bar placement in on-page reading order.
    order = sorted(layout["cells"], key=lambda rf: (layout["cells"][rf][1], layout["cells"][rf][0]))
    letter_of = {}
    if st["letters"]:
        if st["letter_scope"] == "panel":
            for i, rf in enumerate(order):
                letter_of[rf] = letter_text(i, st["letter_style"])
        else:
            for f in range(len(columns)):
                letter_of[(0, f)] = letter_text(f, st["letter_style"])
    scale_cells = set()
    if st["scale_bar"] == "all":
        scale_cells = set(layout["cells"])
    elif st["scale_bar"] == "one" and layout["cells"]:
        # bottom-left panel of the page (last block's first panel row)
        scale_cells = {max(layout["cells"], key=lambda rf: (layout["cells"][rf][1], -layout["cells"][rf][0]))}

    ranges, panel_ranges, units, images, panel_images = {}, {}, {}, {}, {}
    own_scale = {}
    cbar_artists = []
    for r, row in enumerate(spec["rows"]):
        relative = row["kind"] == "z"
        row_data = {c["key"]: (_relative(panels[(r, c["key"])]["data"]) if relative
                               else np.asarray(panels[(r, c["key"])]["data"], dtype=float))
                    for c in columns if (r, c["key"]) in panels}
        unit = next((panels[(r, c["key"])].get("unit") for c in columns
                     if (r, c["key"]) in panels), "")
        # A channel with no contrast anywhere (topography of CH frames: Z is
        # held constant) gets "flat signal" tags with its actual level
        # instead of an orange block and 1e-12 colourbar ticks.
        flat_level = _flat_row_level(
            {c["key"]: panels[(r, c["key"])]["data"] for c in columns if (r, c["key"]) in panels})
        if flat_level is not None:
            row_data = {k: np.asarray(panels[(r, k)]["data"], dtype=float) for k in row_data}
            outliers = {k: "flat" for k in row_data}
        else:
            outliers = baseline_outliers(row_data) if shared else {}
        for k, why in outliers.items():
            own_scale[(r, k)] = why
        inliers = [a for k, a in row_data.items() if k not in outliers] or list(row_data.values())
        if relative and flat_level is None:
            # Topography: one height scale that fits the *typical* frame, so
            # a tip crash or step in one frame saturates there instead of
            # flattening every other frame (signal rows pool instead: a frame
            # with more contrast - closer tip, other bias - is the point).
            raw_lo, raw_hi = typical_range(inliers)
        else:
            raw_lo, raw_hi = pooled_range(inliers)
        d_unit, factor = row_display_unit(unit, [np.array([raw_lo, raw_hi])])
        units[r] = d_unit
        user = clims.get(clim_key(row["label"]))
        ranges[r] = (float(user[0]), float(user[1])) if (shared and user) else (raw_lo * factor, raw_hi * factor)
        cmap = get_cmap(row["cmap"])
        last_im = None
        for f, col in enumerate(columns):
            box = layout["cells"][(r, f)]
            ax = fig.add_axes(_ax_rect(box, size))
            ax.set_axis_off()
            if col["key"] not in row_data:
                ax.add_patch(_placeholder(ax))
                continue
            p = panels[(r, col["key"])]
            data = row_data[col["key"]] * factor
            own = (r, col["key"]) in own_scale
            if shared and not own:
                lo, hi = ranges[r]
            else:
                user_p = clims.get(clim_key(row["label"], col["key"]))
                lo, hi = (float(user_p[0]), float(user_p[1])) if user_p else robust_range(_sample(data))
            panel_ranges[(r, f)] = (lo, hi)
            w = _finite(p.get("width_nm")) or float(data.shape[1])
            h = _finite(p.get("height_nm")) or float(data.shape[0])
            im = ax.imshow(data, cmap=cmap, vmin=lo, vmax=hi, origin="upper",
                           interpolation="nearest", extent=(0.0, w, 0.0, h), aspect="auto")
            panel_images[(r, f)] = im
            ax.set_xlim(0.0, w)
            ax.set_ylim(0.0, h)
            # True aspect inside the cell (a selection can mix frame sizes).
            ax.set_aspect("equal", adjustable="box", anchor="C")
            if own:
                tag = "own scale"
                if own_scale[(r, col["key"])] == "flat":
                    # No signal: dim it so amplified noise can't pass as
                    # contrast, and say what the level actually is.
                    im.set_alpha(0.35)
                    v = _sample(data)
                    tag = "flat signal" + (f" · {float(np.median(v)):.5g} {d_unit}" if v.size else "")
                # First corner not taken by this panel's letter / scale bar.
                taken = set()
                if letter_of.get((r, f)):
                    taken.add(st["letter_pos"])
                if (r, f) in scale_cells:
                    taken.add(st["scale_pos"])
                corner = next((c for c in ("br", "bl", "tr", "tl") if c not in taken), "br")
                tx, ty, tha, tva = _corner(w, h, corner, inset=0.03)
                ax.text(tx, ty, tag, color="white", fontsize=font - 1, ha=tha, va=tva,
                        path_effects=stroke("white"), zorder=7)
            else:
                images.setdefault(r, []).append(im)
                last_im = im
            letter = letter_of.get((r, f))
            if letter:
                _draw_corner_text(ax, w, h, letter, st["letter_pos"], st["letter_color"],
                                  float(st["letter_size"]), bool(st["letter_bold"]),
                                  stroke(st["letter_color"]))
            if (r, f) in scale_cells and _finite(p.get("width_nm")):
                _draw_scale_bar(ax, w, h, st, font, stroke(st["scale_color"]))
            if st["colorbars"] and not shared and own_scale.get((r, col["key"])) != "flat":
                for entry in layout["colorbars"]:
                    if entry["row"] == r and entry["frame"] == f:
                        cb = _colorbar(fig, im, entry, size, lo, hi, st, font, text_color,
                                       label=None)
                        cbar_artists.append((entry, cb))
        if st["colorbars"] and shared:
            ref_im = last_im
            if ref_im is None and images.get(r):
                ref_im = images[r][-1]
            if ref_im is not None:
                for entry in layout["colorbars"]:
                    if entry["row"] == r and entry["frame"] is None:
                        label = None
                        if st["cbar_label"]:
                            label = f"{row['display']} ({d_unit})" if d_unit else row["display"]
                        cb = _colorbar(fig, ref_im, entry, size, *ranges[r], st, font, text_color, label)
                        cbar_artists.append((entry, cb))

    # Frame headers and channel labels as figure text.
    fig_w, fig_h = size
    for f, (x0, y0, x1, y1) in layout["frame_labels"]:
        col = columns[f]
        weight = "bold" if col.get("is_ref") else "normal"
        if not layout["transpose"]:
            xc = 0.5 * (x0 + x1) / fig_w
            line = font / 72.0 * 1.45
            n_lines = 2 if col.get("subheader") else 1
            base = y1 - 0.05 - (line if n_lines == 2 else 0.0)
            fig.text(xc, 1.0 - base / fig_h, col["header"], ha="center", va="bottom",
                     fontsize=font, color=text_color, fontweight=weight)
            if col.get("subheader"):
                fig.text(xc, 1.0 - (y1 - 0.05) / fig_h, col["subheader"], ha="center",
                         va="bottom", fontsize=font - 1, color=text_color)
        else:
            yc = 0.5 * (y0 + y1)
            dy = 0.5 * font / 72.0 * 1.4 if col.get("subheader") else 0.0
            fig.text(x1 / fig_w, 1.0 - (yc - dy) / fig_h, col["header"], ha="right",
                     va="center", fontsize=font, color=text_color, fontweight=weight)
            if col.get("subheader"):
                fig.text(x1 / fig_w, 1.0 - (yc + dy) / fig_h, col["subheader"], ha="right",
                         va="center", fontsize=font - 1, color=text_color)
    for r, (x0, y0, x1, y1) in layout["channel_labels"]:
        row = spec["rows"][r]
        text = row["display"]
        if not layout["transpose"]:
            fig.text((x1 - 0.08) / fig_w, 1.0 - 0.5 * (y0 + y1) / fig_h, text,
                     ha="right", va="center", rotation=90 if len(text) > 4 else 0,
                     fontsize=font + 1, color=text_color)
        else:
            fig.text(0.5 * (x0 + x1) / fig_w, 1.0 - (y1 - 0.04) / fig_h, text,
                     ha="center", va="bottom", fontsize=font + 1, color=text_color)

    layout.update({"ranges": ranges, "panel_ranges": panel_ranges, "units": units,
                   "images": images, "panel_images": panel_images,
                   "colorbar_artists": cbar_artists, "own_scale": own_scale,
                   "style": st, "n_rows": n_rows})
    return layout


def _colorbar(fig, im, entry, size, lo, hi, st, font, text_color, label):
    cax = fig.add_axes(_ax_rect(entry["box"], size))
    orient = "horizontal" if entry["orient"] == "h" else "vertical"
    import matplotlib.cm as mcm
    mappable = mcm.ScalarMappable(norm=im.norm, cmap=im.get_cmap())
    cb = fig.colorbar(mappable, cax=cax, orientation=orient)
    cb._sxm_n_ticks = int(st["cbar_ticks"])
    set_colorbar_range(cb, lo, hi)
    cb.ax.tick_params(labelsize=font - 0.5, length=2, width=0.5, pad=1.5, colors=text_color)
    cb.outline.set_linewidth(0.5)
    if label:
        cb.set_label(label, fontsize=font, labelpad=1.0, color=text_color)
    return cb


def set_colorbar_range(cb, lo, hi):
    """Few ticks only - a figure colourbar, not an axis."""
    n = max(2, int(getattr(cb, "_sxm_n_ticks", 2) or 2))
    span = abs(hi - lo)
    cb.mappable.set_clim(lo, hi)
    ticks = list(np.linspace(lo, hi, n))
    cb.set_ticks(ticks)
    cb.set_ticklabels([_tick(t, span) for t in ticks])
    # End labels sit inside the bar's own span, so stacked / side-by-side
    # colourbars never print "0.08" on top of the neighbour's "10.9".
    horizontal = cb.orientation == "horizontal"
    axis = cb.ax.xaxis if horizontal else cb.ax.yaxis
    major = axis.get_major_ticks()
    for i, tick in enumerate(major):
        for lab in (tick.label1, tick.label2):
            if horizontal:
                lab.set_horizontalalignment("left" if i == 0 else "right" if i == len(major) - 1 else "center")
            else:
                lab.set_verticalalignment("bottom" if i == 0 else "top" if i == len(major) - 1 else "center")


def set_row_range(layout, r, lo, hi):
    """Cheap in-place shared-range update for one channel (colourbar drag)."""
    for im in layout.get("images", {}).get(r, []):
        im.set_clim(lo, hi)
    for entry, cb in layout.get("colorbar_artists", []):
        if entry["row"] == r and entry["frame"] is None:
            set_colorbar_range(cb, lo, hi)
    layout.setdefault("ranges", {})[r] = (lo, hi)


def set_panel_range(layout, r, f, lo, hi):
    """Cheap in-place per-image range update (per-image colourbar drag)."""
    im = layout.get("panel_images", {}).get((r, f))
    if im is not None:
        im.set_clim(lo, hi)
    for entry, cb in layout.get("colorbar_artists", []):
        if entry["row"] == r and entry["frame"] == f:
            set_colorbar_range(cb, lo, hi)
    layout.setdefault("panel_ranges", {})[(r, f)] = (lo, hi)


def _tick(v, span=None):
    a = abs(v)
    if a == 0 or (span and a < 5e-3 * span):
        return "0"
    ref = max(a, span or 0.0)
    if ref >= 100:
        return f"{v:.0f}"
    if ref >= 10:
        return f"{v:.1f}"
    if ref >= 0.01:
        return f"{v:.2f}"
    return f"{v:.1e}"


def _placeholder(ax):
    from matplotlib.patches import Rectangle
    return Rectangle((0, 0), 1, 1, transform=ax.transAxes, facecolor="#8882",
                     edgecolor="#8886", linewidth=0.5, hatch="///")


def _corner(w, h, pos, inset=0.04):
    x = inset * w if pos in ("tl", "bl") else (1 - inset) * w
    y = (1 - inset) * h if pos in ("tl", "tr") else inset * h
    ha = "left" if pos in ("tl", "bl") else "right"
    va = "top" if pos in ("tl", "tr") else "bottom"
    return x, y, ha, va


def _draw_corner_text(ax, w, h, text, pos, color, size, bold, effects):
    x, y, ha, va = _corner(w, h, pos)
    ax.text(x, y, text, color=color, fontsize=size, fontweight="bold" if bold else "normal",
            va=va, ha=ha, path_effects=effects, zorder=7)


def _draw_scale_bar(ax, w, h, st, font, effects):
    length = _finite(st.get("scale_len_nm")) or nice_length(0.3 * w)
    if not length or length > w:
        return
    pos = st["scale_pos"]
    color = st["scale_color"]
    edge = "black" if color == "white" else "white"
    bar_h = max(0.018 * h, 1e-9)
    x0 = 0.06 * w if pos in ("tl", "bl") else 0.94 * w - length
    label_h = 0.07 * h if st["scale_label"] else 0.0
    if pos in ("bl", "br"):
        y0 = 0.07 * h
        ty, tva = y0 + bar_h + 0.02 * h, "bottom"
    else:
        y0 = 0.93 * h - bar_h
        ty, tva = y0 - 0.02 * h, "top"
    from matplotlib.patches import Rectangle
    ax.add_patch(Rectangle((x0, y0), length, bar_h, facecolor=color,
                           edgecolor=edge, linewidth=0.4, zorder=5))
    if label_h:
        ax.text(x0 + 0.5 * length, ty, length_text(length), color=color, fontsize=font,
                ha="center", va=tva, zorder=6, path_effects=effects)
