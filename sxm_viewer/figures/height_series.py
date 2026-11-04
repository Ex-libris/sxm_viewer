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
BEST = "@best"   # pseudo channel: each frame's own best channel


def channel_text(label):
    """How a channel (or the pseudo channel) is named in menus."""
    return "Best channel per frame" if label == BEST else str(label)

_BACKWARD_RE = re.compile(r"(?i)(bwd|backward|\bbk\b|_bk|retrace|\(b\)|<-|←)")

STYLE_DEFAULTS = {
    "transpose": False,        # False: frames = columns, channels = rows
    "reverse": False,          # reverse the frame order (z / bias / time)
    "range_mode": "row",       # "row": one range per channel | "image"
    "colorbars": True,
    "cbar_ticks": 2,           # 0 (none) | 2 (ends) | 3 | 5
    "cbar_label": True,
    "letters": True,
    "letter_style": "a",       # "a" | "A" | "(a)" | "a)" | "a."
    "letter_scope": "frame",   # "frame": one per frame | "panel": every panel
    "letter_pos": "tl",        # tl | tr | bl | br (inside) | out_tl | out_tr (outside, above)
    "letter_color": "auto",    # auto (contrast with the image under it) | white | black
    "letter_size": 9.0,
    "letter_bold": True,
    "scale_bar": "all",        # "all" (every image) | "one" | "off"
    "scale_len_nm": None,      # None = automatic
    "scale_pos": "bl",
    "scale_color": "auto",     # auto | white | black
    "scale_label": "first",    # length label: "first" (once per length) | "all" | "none"
    "scale_font": None,        # length label size (pt); None = text size
    "header": "auto",          # "auto" | "dz" | "zabs" | "bias" | "name"
    "header_name": False,      # file name as an extra header line
    "ref_bold": False,         # bold header for the Δz = 0 reference frame
    "channel_labels": True,    # channel names (I, Δf ...) beside the rows
    "font": 7.0,
    "font_family": None,       # None = Arial when installed (journal standard), else DejaVu Sans
    "cbar_scope": "auto",      # auto (follows the range) | row (one bar per channel, min/max)
    "panel_w": 1.55,           # inches (solved from the page width if set)
    "page": "free",            # free | 1col | 15col | 2col | custom
    "page_mm": 85.0,           # custom page width
    "letter_outline": True,    # contrasting outline around letters
    "scale_outline": True,     # ... around the scale bar and its label
    "scale_thick": 0.018,      # scale bar thickness, fraction of panel height
    "cbar_w": 0.07,            # colourbar thickness (inches)
    "cbar_style": "outside",   # outside | inside (numbers) | minmax | minmax_in ("min"/"max" words)
    "cbar_text": "rotated",    # rotated | horizontal (vertical colourbars)
    "zoom_color": "color",     # zoom outlines: color (one per zoom) | white | black
    "zoom_labels": True,       # letter of each zoom next to its outline
    "zoom_lw": 1.0,            # outline width (pt)
    "cbar_label_pos": "inside",  # channel name: centred inside the bar | outside
    "cbar_outline": True,      # thin frame around colour bars
    "rotate_tall": "off",      # turn tall images to landscape: off | cw | ccw
    "molecules_on": True,      # show the figure's molecular model(s)
    "mol_scope": "all",        # all panels | first (first channel row only)
    "mol_size": 1.0,           # atom/bond size x the automatic panel-width size
}

# Keys whose sensible default depends on the recipe (a zoom figure shows
# frames of different sizes: own ranges, a scale bar on every frame, no
# colourbars). The user's last choice is remembered *per recipe* for these
# and globally for everything else (letters, fonts, outlines, page ...).
RECIPE_KEYS = ("range_mode", "colorbars", "scale_bar", "header", "transpose", "reverse")
RECIPE_DEFAULTS = {
    "zoom": {"range_mode": "image", "colorbars": False, "scale_bar": "all"},
}

# Which style keys are meaningful in each range mode (for the UI).
LETTER_STYLES = ("a", "A", "(a)", "a)", "a.")
CORNERS = ("tl", "tr", "bl", "br")


def resolve_style(style):
    out = dict(STYLE_DEFAULTS)
    for k, v in (style or {}).items():
        if k in out and v is not None or k in ("scale_len_nm", "scale_font"):
            out[k] = v
    # Figures saved before the label choice existed stored a bool.
    if out.get("scale_label") is True:
        out["scale_label"] = "all"
    elif out.get("scale_label") is False:
        out["scale_label"] = "none"
    return out


_JOURNAL_FONTS = ("Arial", "Helvetica", "Calibri", "Segoe UI", "Liberation Sans", "DejaVu Sans",
                  "Times New Roman", "Cambria", "Georgia", "DejaVu Serif", "Courier New")


def available_fonts():
    """Font families matplotlib can render: journal-friendly ones first."""
    try:
        from matplotlib import font_manager
        names = sorted({f.name for f in font_manager.fontManager.ttflist})
    except Exception:
        names = ["DejaVu Sans"]
    first = [n for n in _JOURNAL_FONTS if n in names]
    return first + [n for n in names if n not in first]


def default_font():
    fonts = available_fonts()
    return "Arial" if "Arial" in fonts else ("DejaVu Sans" if "DejaVu Sans" in fonts else fonts[0])


def font_rc(st):
    """matplotlib rc for the figure's font (falls back glyph-wise to DejaVu
    for anything the chosen font lacks, e.g. Δ in some families)."""
    family = (st or {}).get("font_family") or default_font()
    return {"font.family": [family, "DejaVu Sans"], "mathtext.fontset": "dejavusans",
            "svg.fonttype": "none", "pdf.fonttype": 42, "ps.fonttype": 42}


def pt(size):
    """Smallest legible size for any printed text (journals ask >= 5 pt)."""
    return max(5.0, float(size))


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


def _collapse_same_footprint(members):
    """Zooms re-scanning the same footprint (any conditions) show once - the
    latest; returns ``(kept, hidden_keys)``."""
    from ..reporting.model import _same_footprint

    groups = []
    for m in members:
        g = _frame_geometry(m)
        geo = (g["cx"], g["cy"], g["w"], g["h"], g["angle"]) if g else None
        for grp in groups:
            if geo is not None and grp[0] is not None and _same_footprint(geo, grp[0]):
                grp[1].append(m)
                break
        else:
            groups.append((geo, [m]))
    kept, hidden = [], []
    for _geo, ms in groups:
        latest = max(ms, key=lambda m: (_time_sort(m) != datetime.max, _time_sort(m)))
        kept.append(latest)
        hidden.extend(str(m["key"]) for m in ms if m is not latest)
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
    the bias for a bias series, the frame size for overview + zooms; a
    selection takes Δz, else bias, else name."""
    if header and header != "auto":
        return header
    if mode == "zoom":
        return "size"
    if mode == "bias":
        return "bias"
    if mode == "height" or has_z:
        return "dz"
    return "bias" if bias_varies else "name"


def build_spec(members, *, mode="height", reference_key=None, dropped=(),
               rows=None, row_choices=None, header="auto", header_name=False,
               row_cmaps=None, reverse=False, show_repeats=False, root_key=None):
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
    root = None
    if mode == "zoom":
        root = next((m for m in cols_in if str(m["key"]) == str(root_key)), None) or (cols_in[0] if cols_in else None)
        zooms = [m for m in cols_in if m is not root]
        if not show_repeats:
            zooms, hidden_repeats = _collapse_same_footprint(zooms)
        zooms.sort(key=lambda m: (_time_sort(m), str(m.get("name"))))
        if reverse:
            zooms.reverse()
        cols_in = ([root] if root is not None else []) + zooms
    elif mode != "selection" and not show_repeats:
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

    if mode == "zoom":
        axis = "zoom"
    elif mode == "bias" or (mode == "selection" and not has_z and has_bias and bias_varies):
        cols_in.sort(key=by_value("bias_v"))
        axis = "bias"
    elif mode == "height" or has_z:
        cols_in.sort(key=by_value("z_nm"))
        axis = "z"
    else:
        cols_in.sort(key=lambda m: (_time_sort(m), str(m.get("name"))))
        axis = "time"
    if reverse and mode != "zoom":
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
        if header == "none":
            text = ""
        elif header == "size" and size_text(m):
            text = size_text(m)
        elif header == "zabs" and z is not None:
            text = f"z = {z:.3f} nm"
        elif header == "bias" and b is not None:
            text = bias_header(b)
        elif header == "dz" and has_z and ref_z is not None:
            text = dz_text(z, ref_z)
        else:
            text = _short(name)
        subs = []
        if header == "none":
            pass
        elif header_name and text != _short(name):
            subs.append(_short(name, 28))
        if header != "none" and bias_varies and header != "bias":
            subs.append(bias_text(b))
        if header != "none" and setpoint_varies:
            subs.append(si_text(m.get("setpoint"), m.get("setpoint_unit") or ""))
        if header != "none" and header == "bias" and has_z and _varies([_finite(c.get("z_nm")) for c in cols_in], 0.0, 5e-3):
            # A bias series whose tip height also moved: say by how much.
            subs.append(dz_text(z, ref_z))
        col = {
            "key": key, "name": name, "z_nm": z, "dz_pm": dz_pm, "bias_v": b,
            "is_ref": key == ref_key and header == "dz",
            "header": text,
            "subheader": "  ·  ".join(s for s in subs if s),
        }
        if mode == "zoom" and root is not None:
            if m is root:
                col["is_overview"] = True
            else:
                col["parent"] = str(root["key"])
                col["box"] = zoom_box(m, root)
        columns.append(col)

    # Every channel the frames have, in first-seen order (menus / Add row).
    channels, cls_of = [], {}
    for m in cols_in:
        for c in m.get("channels") or []:
            if c["label"] not in cls_of:
                cls_of[c["label"]] = c["cls"]
                channels.append({"label": c["label"], "cls": c["cls"],
                                 "kind": row_kind_for(c["cls"])})

    # "Best channel per frame": each frame's own first pick (topography of
    # a CC overview next to current of its CH zooms) - a pseudo channel.
    channels.insert(0, {"label": BEST, "cls": "mixed", "kind": "mixed"})
    labels = list(rows) if rows else _default_rows(cols_in, row_choices)
    if not rows and mode == "zoom" and root is not None and labels:
        # One band by default (a zoom figure is about *where*; a band per
        # channel triples it): the overview's main channel, or each frame's
        # own when the frames were taken in different modes.
        kinds = {row_kind_for(m["picks"][0][1]) for m in cols_in if m.get("picks")}
        if len(kinds) > 1:
            labels = [BEST]
        else:
            by_idx = {int(c["idx"]): c for c in (root.get("channels") or [])}
            first = next((by_idx[int(i)]["label"] for i, _cls in (root.get("picks") or [])
                          if int(i) in by_idx), None)
            labels = [first] if first else labels[:1]
    out_rows = []
    for label in labels:
        if label == BEST:
            cells, cell_kind, cell_cmap = {}, {}, {}
            for m in cols_in:
                if not m.get("picks"):
                    continue
                idx, cls = m["picks"][0]
                key = str(m["key"])
                cells[key] = int(idx)
                cell_kind[key] = row_kind_for(cls)
                cell_cmap[key] = _CLASS_CMAPS.get(cls) or ROW_CMAPS.get(row_kind_for(cls), "viridis")
            if cells:
                out_rows.append({"label": BEST, "kind": "mixed", "cls": "mixed", "symbol": "",
                                 "cmap": "gray", "alternatives": [], "cells": cells,
                                 "cell_kind": cell_kind, "cell_cmap": cell_cmap})
            continue
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
        if r["label"] == BEST:
            r["display"] = ""
            continue
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
_MIN_PANEL_IN = 0.55       # ~14 mm: below this a page layout wraps instead
_GAP = 0.05
_BLOCK_GAP = 0.12
_BOTTOM = 0.10

# Journal column widths (mm). Fonts stay in absolute points, so choosing a
# page width shrinks/grows the *panels*, never the text - what you see is
# what prints at 100 %.
# key: (menu name, width mm, text pt, panel-letter pt, letter style or None)
# Text sizes follow the journals' figure guides: Nature 5-7 pt text with
# 8 pt bold lowercase panel letters; Science 6-8 pt with bold capitals;
# ACS >= 4.5 pt (7-8 pt recommended) with (a)-style letters. Every printed
# text size derives from the text size (headers = text, sub-headers and
# colour-bar values a step smaller, channel names a step larger; never
# below 5 pt), so choosing a page sets the whole figure.
PAGE_PRESETS = {
    "1col": ("Single column (85 mm)", 85.0, 7.0, 8.5, None),
    "15col": ("1.5 column (120 mm)", 120.0, 7.0, 9.0, None),
    "2col": ("Double column (178 mm)", 178.0, 8.0, 10.0, None),
    "nature1": ("Nature - single column (89 mm)", 89.0, 7.0, 8.0, "a"),
    "nature2": ("Nature - double column (183 mm)", 183.0, 7.0, 8.0, "a"),
    "science1": ("Science - 1 column (55 mm)", 55.0, 6.0, 8.0, "A"),
    "science2": ("Science - 2 columns (120 mm)", 120.0, 7.0, 9.0, "A"),
    "science3": ("Science - 3 columns (183 mm)", 183.0, 7.0, 9.0, "A"),
    "acs1": ("ACS - single column (3.25 in)", 82.55, 7.0, 8.0, "(a)"),
    "acs2": ("ACS - double column (7 in)", 177.8, 8.0, 9.0, "(a)"),
}
PAGE_WIDTHS_MM = {k: v[1] for k, v in PAGE_PRESETS.items()}
PAGE_NAMES = dict([("free", "Free size")] + [(k, v[0]) for k, v in PAGE_PRESETS.items()]
                  + [("custom", "Custom width")])


def page_text_style(page):
    """Style keys a page preset sets (text sizes, letter style)."""
    pre = PAGE_PRESETS.get(page)
    if not pre:
        return {}
    out = {"font": pre[2], "letter_size": pre[3], "scale_font": None}
    if pre[4]:
        out["letter_style"] = pre[4]
        out["letter_bold"] = True
    return out


def page_width_in(st):
    page = st.get("page") or "free"
    if page == "free":
        return None
    mm = PAGE_WIDTHS_MM.get(page) if page != "custom" else _finite(st.get("page_mm"))
    return (mm / 25.4) if mm else None


def column_blocks(n, max_per=MAX_FRAMES_PER_BLOCK):
    """Frame indices per block: <= max_per per block, blocks balanced
    (16 -> 8+8, 13 -> 7+6) so the last block is never a lonely straggler."""
    if n <= max_per:
        return [list(range(n))]
    n_blocks = int(math.ceil(n / max_per))
    per = int(math.ceil(n / n_blocks))
    return [list(range(i, min(n, i + per))) for i in range(0, n, per)]


def _aspect_of(p):
    if not p:
        return None
    # width/height are already the displayed (turned) frame size.
    w, h = _finite(p.get("width_nm")), _finite(p.get("height_nm"))
    if w and h:
        return min(max(h / w, 0.25), 4.0)
    shp = np.shape(p.get("data"))
    if len(shp) == 2 and shp[1]:
        return min(max(shp[0] / shp[1], 0.25), 4.0)
    return None


def _panel_aspect(panels, keys=None):
    for (r, k), p in panels.items():
        if keys is not None and k not in keys:
            continue
        a = _aspect_of(p)
        if a:
            return a
    return 1.0


def cbar_dims(st, orient="v"):
    """(bar thickness, extra room for its tick values + label) in inches -
    shared by the layout and the renderer so they always agree."""
    font = float(st["font"])
    pt = font / 72.0
    style = st.get("cbar_style") or "outside"
    rotated = (st.get("cbar_text") or "rotated") == "rotated"
    bar = float(st.get("cbar_w") or 0.07)
    no_ticks = int(st.get("cbar_ticks", 2) or 0) == 0
    if style in ("inside", "minmax_in"):
        # values ("12.5" or "max") centred inside the bar's ends
        bar = max(bar, 1.7 * pt) if not no_ticks else bar
        extra = 0.0
    elif no_ticks:
        extra = 0.0
    elif style == "minmax":
        # A real bar whose ends read "min" / "max" (no numbers).
        extra = (1.3 * pt if rotated else 0.30 * font / 7.0) if orient == "v" else 1.5 * pt
    else:
        extra = (1.3 * pt if rotated else 0.36 * font / 7.0) if orient == "v" else 1.5 * pt
    if st.get("cbar_label"):
        if (st.get("cbar_label_pos") or "inside") == "inside":
            bar = max(bar, 1.7 * pt)   # the name is centred in the bar
        else:
            extra += (1.5 * pt if rotated else 0.0) if orient == "v" else 1.5 * pt
    return bar, extra


def figure_layout(spec, panels, style=None):
    """Element boxes in inches, ``(x0, y0, x1, y1)`` from the figure's
    *top-left*; deterministic so the GUI hit-tests without matplotlib.

    With a page width (``style["page"]``) the panel width is solved so the
    whole figure is exactly that wide - every width in the layout is affine
    in the panel width, so two trial layouts give the answer.
    """
    st = resolve_style(style)
    target = page_width_in(st)
    if target:
        # Too many frames for the page at a legible panel size: wrap them
        # into more rows (fewer frames per row) until the panels fit.
        n_f = max(1, len(spec["columns"]))
        for max_per in range(min(MAX_FRAMES_PER_BLOCK, n_f), 0, -1):
            trial = dict(st, _max_per=max_per)
            a = _layout(spec, panels, dict(trial, panel_w=1.0))["size"][0]
            b = _layout(spec, panels, dict(trial, panel_w=2.0))["size"][0]
            slope = b - a
            pw = (target - (a - slope)) / slope if slope > 1e-6 else st["panel_w"]
            if pw >= _MIN_PANEL_IN or max_per == 1 or spec.get("mode") == "zoom":
                st = dict(trial, panel_w=max(0.25, pw))
                break
    lay = _layout(spec, panels, st)
    lay["_letter_h"] = _outside_letter_h(st)
    lay["panel_w"] = float(st["panel_w"])
    lay["page_in"] = target
    return lay


def _layout(spec, panels, st):
    if spec.get("mode") == "zoom":
        return _layout_zoom(spec, panels, st)
    return _layout_grid(spec, panels, st)


def _letter_line_y(layout, f, cell):
    """Baseline (inches from the top) of an outside letter: on the top line
    of the frame's header band when the panel has one above it (so it never
    collides with the header text), else just above the panel."""
    for ff, hb in layout.get("frame_labels", []):
        if ff == f and abs(hb[3] - cell[1]) < 1e-6 and not layout.get("transpose"):
            return hb[1] + float(layout.get("_letter_h") or 0.0) - 0.02
    return cell[1] - 0.015


def _outside_letter_h(st):
    """Room an outside panel letter needs above its panel (0 if inside)."""
    if not st.get("letters") or not str(st.get("letter_pos") or "").startswith("out"):
        return 0.0
    return float(st.get("letter_size") or 9.0) * 1.2 / 72.0 + 0.04


def _header_height(spec, st):
    line = float(st["font"]) / 72.0 * 1.45
    if not any(c.get("header") or c.get("subheader") for c in spec["columns"]):
        return max(0.04, _outside_letter_h(st))   # headers off: panels start at the top
    has_sub = any(c.get("subheader") for c in spec["columns"])
    # An outside letter sits on its own line above the header text.
    return 0.12 + line * (2 if has_sub else 1) + _outside_letter_h(st)


def _left_band(spec, st):
    """Width of the channel-name band left of the rows (0 when hidden)."""
    if not st.get("channel_labels", True):
        return 0.06
    longest = max((len(r.get("display") or "") for r in spec["rows"]), default=1)
    return 0.42 if longest <= 4 else 0.30


def _layout_grid(spec, panels, st):
    font = float(st["font"])
    n_f = max(1, len(spec["columns"]))
    n_ch = max(1, len(spec["rows"]))
    pw = float(st["panel_w"])
    ph = pw * _panel_aspect(panels)
    one_bar = st.get("cbar_scope") == "row"
    per_img_cb = st["colorbars"] and st["range_mode"] == "image" and not one_bar
    shared_cb = st["colorbars"] and (st["range_mode"] == "row" or one_bar)
    blocks = column_blocks(n_f, int(st.get("_max_per") or MAX_FRAMES_PER_BLOCK))
    per = max(len(b) for b in blocks)
    cells, frame_labels, channel_labels, cbars = {}, [], [], []

    if not st["transpose"]:
        bar, extra = cbar_dims(st, "v")
        slot = pw + ((0.05 + bar + extra) if per_img_cb else 0.0)
        top = _header_height(spec, st)
        left = _left_band(spec, st)
        right = (0.06 + bar + extra + 0.06) if shared_cb else 0.12
        rgap = _GAP + (_outside_letter_h(st) if st.get("letter_scope") == "panel" else 0.0)
        block_h = top + n_ch * ph + (n_ch - 1) * rgap
        fig_w = left + per * slot + (per - 1) * _GAP + right
        fig_h = len(blocks) * block_h + (len(blocks) - 1) * _BLOCK_GAP + _BOTTOM
        for b, frames in enumerate(blocks):
            y_base = b * (block_h + _BLOCK_GAP)
            for j, f in enumerate(frames):
                x0 = left + j * (slot + _GAP)
                frame_labels.append((f, (x0, y_base, x0 + pw, y_base + top)))
                for r in range(n_ch):
                    y0 = y_base + top + r * (ph + rgap)
                    cells[(r, f)] = (x0, y0, x0 + pw, y0 + ph)
                    if per_img_cb:
                        cbars.append({"row": r, "frame": f, "orient": "v",
                                      "box": (x0 + pw + 0.05, y0, x0 + pw + 0.05 + max(bar, 0.01), y0 + ph)})
            x_end = left + len(frames) * slot + (len(frames) - 1) * _GAP
            for r in range(n_ch):
                y0 = y_base + top + r * (ph + rgap)
                channel_labels.append((r, (0.0, y0, left, y0 + ph)))
                if shared_cb:
                    cbars.append({"row": r, "frame": None, "orient": "v",
                                  "box": (x_end + 0.06, y0, x_end + 0.06 + max(bar, 0.01), y0 + ph)})
    else:
        bar_v, extra_v = cbar_dims(st, "v")
        bar_h, extra_h = cbar_dims(st, "h")
        slot = pw + ((0.05 + bar_v + extra_v) if per_img_cb else 0.0)
        top = (0.10 + font / 72.0 * 1.45) if st.get("channel_labels", True) else 0.04
        longest_h = max((max(len(c.get("header") or ""), len(c.get("subheader") or ""))
                         for c in spec["columns"]), default=0)
        left = (0.20 + max(0.75, font / 72.0 * 0.62 * longest_h)) if longest_h else 0.06
        bottom = (0.06 + bar_h + extra_h + 0.04) if shared_cb else _BOTTOM
        block_w = left + n_ch * slot + (n_ch - 1) * _GAP + 0.12
        fig_w = len(blocks) * block_w + (len(blocks) - 1) * _BLOCK_GAP
        rgap = _GAP + _outside_letter_h(st)
        top = max(top, _outside_letter_h(st))
        fig_h = top + per * ph + (per - 1) * rgap + bottom
        for b, frames in enumerate(blocks):
            x_base = b * (block_w + _BLOCK_GAP)
            y_end = top + len(frames) * ph + (len(frames) - 1) * rgap
            for r in range(n_ch):
                x0 = x_base + left + r * (slot + _GAP)
                channel_labels.append((r, (x0, 0.0, x0 + pw, top)))
                if shared_cb:
                    cbars.append({"row": r, "frame": None, "orient": "h",
                                  "box": (x0, y_end + 0.06, x0 + pw, y_end + 0.06 + max(bar_h, 0.01))})
            for j, f in enumerate(frames):
                y0 = top + j * (ph + rgap)
                frame_labels.append((f, (x_base, y0, x_base + left - 0.06, y0 + ph)))
                for r in range(n_ch):
                    x0 = x_base + left + r * (slot + _GAP)
                    cells[(r, f)] = (x0, y0, x0 + pw, y0 + ph)
                    if per_img_cb:
                        cbars.append({"row": r, "frame": f, "orient": "v",
                                      "box": (x0 + pw + 0.05, y0, x0 + pw + 0.05 + max(bar_v, 0.01), y0 + ph)})
    return {"size": (fig_w, fig_h), "cells": cells, "frame_labels": frame_labels,
            "channel_labels": channel_labels, "colorbars": cbars,
            "transpose": bool(st["transpose"]), "blocks": blocks, "panel": (pw, ph)}


def zoom_grid_shape(n_zooms):
    """(rows, cols) of the zoom panels next to the overview."""
    if n_zooms <= 0:
        return 0, 0
    rows = 1 if n_zooms <= 3 else 2 if n_zooms <= 8 else 3
    return rows, int(math.ceil(n_zooms / rows))


def _layout_zoom(spec, panels, st):
    """Overview on the left (as tall as the zoom grid), zooms in a grid on
    the right; one such band per channel."""
    n_f = len(spec["columns"])
    n_ch = max(1, len(spec["rows"]))
    pw = float(st["panel_w"])
    keys = [c["key"] for c in spec["columns"]]
    ov_aspect = _panel_aspect(panels, {keys[0]}) if keys else 1.0
    z_aspect = _panel_aspect(panels, set(keys[1:])) if n_f > 1 else 1.0
    ph = pw * z_aspect
    one_bar = st.get("cbar_scope") == "row"
    per_img_cb = st["colorbars"] and st["range_mode"] == "image" and not one_bar
    shared_cb = st["colorbars"] and (st["range_mode"] == "row" or one_bar)
    bar, extra = cbar_dims(st, "v")
    pimg = (0.05 + bar + extra) if per_img_cb else 0.0
    top = _header_height(spec, st)
    n_z = max(0, n_f - 1)
    z_rows, z_cols = zoom_grid_shape(n_z)
    if n_z:
        band_h = z_rows * (top + ph) + (z_rows - 1) * _GAP
    else:
        band_h = top + 2.0 * pw * ov_aspect
    ov_h = band_h - top
    ov_w = ov_h / ov_aspect
    left = _left_band(spec, st) if n_ch > 1 else 0.06
    x_z = left + ov_w + pimg + 3 * _GAP
    zoom_w = z_cols * (pw + pimg) + max(0, z_cols - 1) * _GAP
    x_end = x_z + zoom_w if n_z else left + ov_w + pimg
    right = (0.06 + bar + extra + 0.06) if shared_cb else 0.12
    fig_w = x_end + right
    fig_h = n_ch * band_h + (n_ch - 1) * _BLOCK_GAP + _BOTTOM
    cells, frame_labels, channel_labels, cbars = {}, [], [], []
    for r in range(n_ch):
        yb = r * (band_h + _BLOCK_GAP)
        if n_f:
            frame_labels.append((0, (left, yb, left + ov_w, yb + top)))
            cells[(r, 0)] = (left, yb + top, left + ov_w, yb + top + ov_h)
            if per_img_cb:
                cbars.append({"row": r, "frame": 0, "orient": "v",
                              "box": (left + ov_w + 0.05, yb + top,
                                      left + ov_w + 0.05 + max(bar, 0.01), yb + top + ov_h)})
        for j in range(n_z):
            f = j + 1
            gr, gc = divmod(j, z_cols)
            x0 = x_z + gc * (pw + pimg + _GAP)
            y0 = yb + gr * (top + ph + _GAP)
            frame_labels.append((f, (x0, y0, x0 + pw, y0 + top)))
            cells[(r, f)] = (x0, y0 + top, x0 + pw, y0 + top + ph)
            if per_img_cb:
                cbars.append({"row": r, "frame": f, "orient": "v",
                              "box": (x0 + pw + 0.05, y0 + top, x0 + pw + 0.05 + max(bar, 0.01), y0 + top + ph)})
        channel_labels.append((r, (0.0, yb + top, left, yb + band_h)))
        if shared_cb:
            cbars.append({"row": r, "frame": None, "orient": "v",
                          "box": (x_end + 0.06, yb + top, x_end + 0.06 + max(bar, 0.01), yb + band_h)})
    return {"size": (fig_w, fig_h), "cells": cells, "frame_labels": frame_labels,
            "channel_labels": channel_labels, "colorbars": cbars, "transpose": False,
            "blocks": [list(range(n_f))], "panel": (pw, ph)}


def _ax_rect(box, size):
    fig_w, fig_h = size
    x0, y0, x1, y1 = box
    return [x0 / fig_w, 1.0 - y1 / fig_h, (x1 - x0) / fig_w, (y1 - y0) / fig_h]


# ---------------------------------------------------------------------------
# Zoom boxes (overview + zooms)
# ---------------------------------------------------------------------------

ZOOM_COLORS = ("#ffd23f", "#3bceac", "#ff5d73", "#5ab0ff", "#c792ea", "#ff9f1c",
               "#a3e048", "#f78fb3")


def _frame_geometry(m):
    ext = m.get("extent")
    if not ext or len(ext) != 4:
        return None
    try:
        x0, x1, y1, y0 = (float(v) for v in ext)
    except (TypeError, ValueError):
        return None
    xs, ys = x1 - x0, y1 - y0
    if not (xs > 0 and ys > 0):
        return None
    return {"cx": 0.5 * (x0 + x1), "cy": 0.5 * (y0 + y1), "w": xs, "h": ys,
            "angle": float(m.get("angle") or 0.0)}


def _local_to_abs(g, fx, fy):
    """Frame fraction (fx right, fy down from the top-left) -> absolute nm.
    Same convention as ``geometry.spec_mapping.pixel_to_nm``."""
    u = (fx - 0.5) * g["w"]
    v = (0.5 - fy) * g["h"]
    t = math.radians(g["angle"])
    c, s = math.cos(t), math.sin(t)
    return g["cx"] + u * c + v * s, g["cy"] - u * s + v * c


def _abs_to_local(g, x, y):
    """Absolute nm -> frame fraction; the rotate-then-normalise order of
    ``geometry.spec_mapping.map_spec_to_pixels`` (what the preview draws)."""
    dx, dy = x - g["cx"], y - g["cy"]
    t = math.radians(g["angle"])
    c, s = math.cos(t), math.sin(t)
    u = dx * c - dy * s
    v = dx * s + dy * c
    return u / g["w"] + 0.5, 0.5 - v / g["h"]


def zoom_box(child, parent):
    """The child frame's four corners as fractions of the parent's raster
    (fx, fy from the top-left), or None."""
    gc, gp = _frame_geometry(child), _frame_geometry(parent)
    if not gc or not gp:
        return None
    return [_abs_to_local(gp, *_local_to_abs(gc, fx, fy))
            for fx, fy in ((0, 0), (1, 0), (1, 1), (0, 1))]


def _turn_frac(fx, fy, turns):
    """Fraction in the acquired raster -> after ``turns`` clockwise quarter
    turns (the display orientation of ``geometry.orientation``)."""
    for _ in range(int(turns or 0) % 4):
        fx, fy = 1.0 - fy, fx
    return fx, fy


def size_text(m):
    g = _frame_geometry(m)
    if not g:
        return ""
    w, h = g["w"], g["h"]

    def fmt(v):
        return f"{v / 1000:.3g} µm" if v >= 1000 else f"{v:.3g} nm"

    return fmt(w) if abs(w - h) <= 0.02 * max(w, h) else f"{fmt(w).split()[0]} × {fmt(h)}"


def find_zoom_families(images, min_ratio=2.5, tol=0.04):
    """Overview -> zooms: every frame that sits inside a larger one.

    A frame's parent is the *smallest* frame with >= ``min_ratio`` times its
    area that contains all four of its corners (rotation-aware, same
    convention as the spectroscopy markers; ``tol`` = allowed overhang as a
    fraction of the parent). Families are one level deep - an overview and
    its direct zooms; a zoom with zooms of its own gets its own family.
    Re-scans of the same zoom footprint count once (the latest is shown).

    Returns ``[{id, keys, root, recipe:"zoom", levels, excluded}]``,
    ``keys`` = root first, then the zooms by acquisition time.
    """
    from ..reporting.model import _same_footprint

    geoms = {}
    for img in images or []:
        g = _frame_geometry(img)
        if g:
            geoms[str(img["key"])] = (img, g)
    parent_of = {}
    items = list(geoms.items())
    for ck, (cimg, gc) in items:
        area_c = gc["w"] * gc["h"]
        best = None
        for pk, (pimg, gp) in items:
            if pk == ck:
                continue
            area_p = gp["w"] * gp["h"]
            if area_p < min_ratio * area_c:
                continue
            corners = [_abs_to_local(gp, *_local_to_abs(gc, fx, fy))
                       for fx, fy in ((0, 0), (1, 0), (1, 1), (0, 1))]
            if all(-tol <= fx <= 1 + tol and -tol <= fy <= 1 + tol for fx, fy in corners):
                if best is None or area_p < best[0]:
                    best = (area_p, pk)
        if best:
            parent_of[ck] = best[1]
    children = {}
    for ck, pk in parent_of.items():
        children.setdefault(pk, []).append(ck)
    out = []
    for pk, kids in children.items():
        kids.sort(key=lambda k: (_time_sort(geoms[k][0]), k))
        # distinct footprints among the zooms
        groups = []
        for k in kids:
            g = geoms[k][1]
            geo = (g["cx"], g["cy"], g["w"], g["h"], g["angle"])
            for grp in groups:
                if _same_footprint(geo, grp[0]):
                    grp[1].append(k)
                    break
            else:
                groups.append((geo, [k]))
        keys = [pk] + kids
        out.append({
            "id": "zoom-" + series_id(keys),
            "keys": keys,
            "root": pk,
            "recipe": "zoom",
            "levels": len(groups),
            "excluded": [],
        })
    return out


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def rotate_tall_panels(panels, mode):
    """Turn portrait panels (h > 1.15 w) a quarter turn to landscape, for
    this figure only: data rotated, width/height swapped, ``turns`` bumped
    so zoom outlines follow; scale bars follow from the swapped size."""
    if mode not in ("cw", "ccw"):
        return panels
    out = {}
    for k, p in panels.items():
        a = _aspect_of(p)
        if p is None or not a or a <= 1.15:
            out[k] = p
            continue
        q = dict(p)
        q["data"] = np.rot90(np.asarray(p["data"]), -1 if mode == "cw" else 1)
        q["width_nm"], q["height_nm"] = p.get("height_nm"), p.get("width_nm")
        q["turns"] = (int(p.get("turns") or 0) + (1 if mode == "cw" else 3)) % 4
        q["rotated_tall"] = True
        out[k] = q
    return out


def _is_light(color):
    try:
        from matplotlib.colors import to_rgb
        r, g, b = to_rgb(color)
    except Exception:
        return True
    return (0.299 * r + 0.587 * g + 0.114 * b) > 0.5


def _effects(enabled, color):
    if not enabled:
        return []
    import matplotlib.patheffects as pe
    # Outline in the opposite tone of the text (works for any colour).
    other = "black" if _is_light(color) else "white"
    return [pe.withStroke(linewidth=1.6, foreground=other)]


def auto_color(ax, region, fallback="white"):
    """White or black for text drawn over ``region`` (axes fractions
    ``(x0, x1, y0, y1)``) - judged on the image pixels actually there,
    through the live colormap and range. The same WCAG rule as the preview's
    scale bar (``gui/canvases/scale_bar_contrast.py``, Qt-free)."""
    try:
        from ..gui.canvases import scale_bar_contrast as sbc
    except Exception:
        return fallback
    c = sbc.contrast_color(ax, region, None)
    if c is None:
        return fallback
    return "white" if c == sbc.LIGHT else "black"


def _letter_region(pos, cell, size_pt, n_chars, inset=0.04):
    """Axes-fraction box covered by a corner letter (for its contrast)."""
    x0, y0, x1, y1 = cell
    cw, ch = max(x1 - x0, 1e-6), max(y1 - y0, 1e-6)
    fw = min(1.0, max(0.12, 0.62 * size_pt / 72.0 * max(1, n_chars)) / cw)
    fh = min(1.0, 1.25 * size_pt / 72.0 / ch)
    left = inset if pos in ("tl", "bl") else 1.0 - inset - fw
    bottom = 1.0 - inset - fh if pos in ("tl", "tr") else inset
    return (left, left + fw, bottom, bottom + fh)


def render(fig, spec, panels, *, style=None, clims=None, get_cmap=None,
           text_color="black", letter_colors=None, draw_molecules=None):
    """Draw the figure (see ``_render``) in the style's font family.
    ``letter_colors``: per-letter overrides ``{letter_key(row_label,
    frame_key): colour}``; everything else follows ``style["letter_color"]``
    ("auto" = contrast with the image under each letter)."""
    import matplotlib
    with matplotlib.rc_context(font_rc(resolve_style(style))):
        return _render(fig, spec, panels, style=style, clims=clims, get_cmap=get_cmap,
                       text_color=text_color, letter_colors=letter_colors,
                       draw_molecules=draw_molecules)


def letter_key(row_label, frame_key):
    return f"{row_label}|{frame_key}"


def _render(fig, spec, panels, *, style=None, clims=None, get_cmap=None,
            text_color="black", letter_colors=None, draw_molecules=None):
    """Draw the figure into ``fig`` (cleared first).

    ``panels``: ``{(row_index, key): {data, unit, width_nm, height_nm,
    turns}}`` with ``data`` in the channel's *normalised* unit and row 0 =
    top of the scan (as the preview shows it). ``clims``: user overrides
    keyed by ``clim_key(row_label)`` (shared mode) or ``clim_key(row_label,
    key)`` (per image), values in the row's display unit.
    Returns the layout dict plus ``ranges`` {row: (lo, hi)} (shared mode),
    ``panel_ranges`` {(row, frame): (lo, hi)}, ``units``, ``own_scale`` and
    hit boxes for letters / scale bars.
    """
    st = resolve_style(style)
    font = float(st["font"])
    clims = dict(clims or {})
    panels = rotate_tall_panels(panels, st.get("rotate_tall"))
    layout = figure_layout(spec, panels, st)
    st = dict(st, panel_w=layout["panel_w"])
    size = layout["size"]
    fig.clear()
    fig.set_size_inches(*size, forward=False)
    fig.patch.set_alpha(0.0)
    fig.set_facecolor("none")
    if get_cmap is None:
        import matplotlib
        get_cmap = matplotlib.colormaps.get_cmap

    shared = st["range_mode"] == "row"
    columns = spec["columns"]
    n_rows = len(spec["rows"])
    is_zoom = spec.get("mode") == "zoom"

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
    frame_letter = {f: letter_text(f, st["letter_style"]) for f in range(len(columns))}
    scale_cells = set()
    if st["scale_bar"] == "all":
        scale_cells = set(layout["cells"])
    elif st["scale_bar"] == "one" and layout["cells"]:
        if is_zoom:
            # every frame has its own size: one bar per frame (first channel)
            scale_cells = {(0, f) for f in range(len(columns))}
        else:
            scale_cells = {max(layout["cells"], key=lambda rf: (layout["cells"][rf][1], -layout["cells"][rf][0]))}

    # Length labels: on every bar, none, or once per distinct length (in
    # reading order) - frames of one size share a single "2 nm".
    label_cells = set()
    where = st.get("scale_label") or "first"
    seen_lengths = []
    for rf in order:
        if rf not in scale_cells or where == "none":
            continue
        pp = panels.get((rf[0], columns[rf[1]]["key"])) if rf[1] < len(columns) else None
        wn = _finite((pp or {}).get("width_nm"))
        if not wn:
            continue
        length = _finite(st.get("scale_len_nm")) or nice_length(0.3 * wn)
        if where == "all" or not any(abs(length - L) <= 1e-9 * max(1.0, L) for L in seen_lengths):
            label_cells.add(rf)
            seen_lengths.append(length)

    ranges, panel_ranges, units, images, panel_images = {}, {}, {}, {}, {}
    own_scale, axes_of, dims_of = {}, {}, {}
    cbar_artists, letter_boxes, scale_boxes = [], [], []
    data_boxes = {}   # (row, frame) -> (inch box of the drawn image, w_nm, h_nm)
    for r, row in enumerate(spec["rows"]):
        mixed = row["kind"] == "mixed"
        cell_kind = row.get("cell_kind") or {}
        relative = row["kind"] == "z"

        def is_rel(k):
            return cell_kind.get(k) == "z" if mixed else relative

        row_data = {c["key"]: (_relative(panels[(r, c["key"])]["data"]) if is_rel(c["key"])
                               else np.asarray(panels[(r, c["key"])]["data"], dtype=float))
                    for c in columns if (r, c["key"]) in panels}
        unit = next((panels[(r, c["key"])].get("unit") for c in columns
                     if (r, c["key"]) in panels), "")
        # A channel with no contrast anywhere (topography of CH frames: Z is
        # held constant) gets "flat signal" tags with its actual level.
        flat_level = _flat_row_level(
            {c["key"]: panels[(r, c["key"])]["data"] for c in columns if (r, c["key"]) in panels})
        if flat_level is not None:
            row_data = {k: np.asarray(panels[(r, k)]["data"], dtype=float) for k in row_data}
            outliers = {k: "flat" for k in row_data}
        else:
            outliers = baseline_outliers(row_data) if (shared and not mixed) else {}
            # A single frame with no contrast of its own (CH topography in a
            # zoom figure, a dead channel) is "flat signal" in any mode.
            for k, a in row_data.items():
                if k not in outliers and _flat_row_level({k: panels[(r, k)]["data"]}) is not None:
                    outliers[k] = "flat"
                    row_data[k] = np.asarray(panels[(r, k)]["data"], dtype=float)  # its real level
        for k, why in outliers.items():
            own_scale[(r, k)] = why
        inliers = [a for k, a in row_data.items() if k not in outliers] or list(row_data.values())
        if relative and flat_level is None:
            # Topography: one height scale fitted to the *typical* frame, so
            # a tip crash in one frame saturates there instead of darkening
            # every other frame (signal rows pool: more contrast is content).
            raw_lo, raw_hi = typical_range(inliers)
        else:
            raw_lo, raw_hi = pooled_range(inliers)
        d_unit, factor = row_display_unit(unit, [np.array([raw_lo, raw_hi])])
        units[r] = d_unit if not mixed else ""
        user = clims.get(clim_key(row["label"]))
        ranges[r] = (float(user[0]), float(user[1])) if (shared and user) else (raw_lo * factor, raw_hi * factor)
        cmap = get_cmap(row["cmap"])
        # Mixed rows (different channels per frame) have no common range or
        # unit: every panel is scaled on its own, in its own unit.
        row_shared = shared and not mixed
        last_im = None
        for f, col in enumerate(columns):
            box = layout["cells"].get((r, f))
            if box is None:
                continue
            ax = fig.add_axes(_ax_rect(box, size))
            ax.set_axis_off()
            ax.set_gid(f"panel_{r}_{f}")
            axes_of[(r, f)] = ax
            if col["key"] not in row_data:
                ax.add_patch(_placeholder(ax))
                continue
            p = panels[(r, col["key"])]
            c_unit, c_factor, c_cmap = d_unit, factor, cmap
            if mixed:
                c_unit, c_factor = row_display_unit(p.get("unit") or "", [row_data[col["key"]]])
                c_cmap = get_cmap((row.get("cell_cmap") or {}).get(col["key"]) or "gray")
            data = row_data[col["key"]] * c_factor
            own = (r, col["key"]) in own_scale
            if row_shared and not own:
                lo, hi = ranges[r]
            else:
                user_p = clims.get(clim_key(row["label"], col["key"]))
                lo, hi = (float(user_p[0]), float(user_p[1])) if user_p else robust_range(_sample(data))
            panel_ranges[(r, f)] = (lo, hi)
            w = _finite(p.get("width_nm")) or float(data.shape[1])
            h = _finite(p.get("height_nm")) or float(data.shape[0])
            dims_of[(r, f)] = (w, h, p.get("turns", 0))
            im = ax.imshow(data, cmap=c_cmap, vmin=lo, vmax=hi, origin="upper",
                           interpolation="nearest", extent=(0.0, w, 0.0, h), aspect="auto")
            im.set_gid(f"image_{r}_{f}")
            panel_images[(r, f)] = im
            ax.set_xlim(0.0, w)
            ax.set_ylim(0.0, h)
            ax.set_aspect("equal", adjustable="box", anchor="C")
            try:
                ax.apply_aspect()
                pos = ax.get_position()
                data_boxes[(r, f)] = ((pos.x0 * size[0], (1.0 - pos.y1) * size[1],
                                       pos.x1 * size[0], (1.0 - pos.y0) * size[1]), w, h)
            except Exception:
                data_boxes[(r, f)] = (box, w, h)
            # Molecular model(s): panel nm coordinates, so one placement is
            # the same spot on every channel / frame of the figure.
            if draw_molecules is not None and st.get("molecules_on", True) and                     (st.get("mol_scope") != "first" or r == 0):
                try:
                    draw_molecules(ax, r, f, w, h)
                except Exception:
                    pass
            taken = set()
            letter = letter_of.get((r, f))
            if letter and not str(st["letter_pos"]).startswith("out"):
                taken.add(st["letter_pos"])
            if (r, f) in scale_cells:
                taken.add(st["scale_pos"])
            if own:
                tag = "own scale"
                if own_scale[(r, col["key"])] == "flat":
                    im.set_alpha(0.35)
                    v = _sample(data)
                    tag = "flat signal" + (f" · {float(np.median(v)):.5g} {c_unit}" if v.size else "")
                corner = next((c for c in ("br", "bl", "tr", "tl") if c not in taken), "br")
                taken.add(corner)
                tx, ty, tha, tva = _corner(w, h, corner, inset=0.03)
                ax.text(tx, ty, tag, color="white", fontsize=pt(font - 1), ha=tha, va=tva,
                        path_effects=_effects(True, "white"), zorder=7, gid=f"note_{r}_{f}")
            else:
                images.setdefault(r, []).append(im)
                last_im = im
            override = (letter_colors or {}).get(letter_key(row["label"], col["key"]))
            if letter and str(st["letter_pos"]).startswith("out"):
                lx = box[0] if st["letter_pos"] == "out_tl" else box[2]
                # On the white page: dark unless the user chose a colour.
                color = override or (st["letter_color"] if st["letter_color"] not in ("white", "auto")
                                     else text_color)
                ly = _letter_line_y(layout, f, box)
                t = fig.text(lx / size[0], 1.0 - ly / size[1], letter, color=color,
                             fontsize=float(st["letter_size"]),
                             fontweight="bold" if st["letter_bold"] else "normal",
                             ha="left" if st["letter_pos"] == "out_tl" else "right", va="bottom",
                             zorder=7, gid=f"letter_{r}_{f}")
                lw_in = max(0.12, 0.62 * float(st["letter_size"]) / 72.0 * len(letter))
                lh_in = float(st["letter_size"]) * 1.2 / 72.0
                bx0 = lx if st["letter_pos"] == "out_tl" else lx - lw_in
                letter_boxes.append({"row": r, "frame": f,
                                     "box": (bx0, ly - lh_in, bx0 + lw_in, ly)})
            elif letter:
                color = override or st["letter_color"]
                if color == "auto":
                    # Each letter judged on its own background: a grey row
                    # and a dark row get different letter colours.
                    color = auto_color(ax, _letter_region(st["letter_pos"], box, float(st["letter_size"]),
                                                          len(letter)))
                t = _draw_corner_text(ax, w, h, letter, st["letter_pos"], color,
                                      float(st["letter_size"]), bool(st["letter_bold"]),
                                      _effects(st["letter_outline"], color))
                t.set_gid(f"letter_{r}_{f}")
                letter_boxes.append({"row": r, "frame": f, "box": _corner_box(
                    box, st["letter_pos"], float(st["letter_size"]), len(letter))})
            if (r, f) in scale_cells and _finite(p.get("width_nm")):
                sb = _draw_scale_bar(ax, w, h, st, font, None,
                                     gid=f"scalebar_{r}_{f}", show_label=(r, f) in label_cells)
                if sb:
                    scale_boxes.append({"row": r, "frame": f, "box": _frac_box(box, sb)})
            if (st["colorbars"] and not shared and st.get("cbar_scope") != "row"
                    and own_scale.get((r, col["key"])) != "flat"):
                for entry in layout["colorbars"]:
                    if entry["row"] == r and entry["frame"] == f:
                        cb = _colorbar(fig, im, entry, size, lo, hi, st, text_color, None, c_unit)
                        cbar_artists.append((entry, cb))
        generic_bar = st["colorbars"] and not row_shared and st.get("cbar_scope") == "row"
        if st["colorbars"] and (row_shared or generic_bar):
            ref_im = last_im if last_im is not None else (images.get(r) or [None])[-1]
            if ref_im is None:
                ref_im = next((panel_images[(r, f)] for f in range(len(columns)) if (r, f) in panel_images), None)
            if ref_im is not None:
                for entry in layout["colorbars"]:
                    if entry["row"] == r and entry["frame"] is None:
                        label = None
                        if generic_bar:
                            # Each image has its own range: one bar for the
                            # channel says only where min and max sit.
                            if st["cbar_label"]:
                                label = row["display"] or None
                            cb = _colorbar(fig, ref_im, entry, size, 0.0, 1.0,
                                           dict(st, cbar_style="minmax_in" if st.get("cbar_style") in
                                                ("inside", "minmax_in") else "minmax"),
                                           text_color, label, "")
                        else:
                            if st["cbar_label"]:
                                label = f"{row['display']} ({d_unit})" if d_unit else row["display"]
                            cb = _colorbar(fig, ref_im, entry, size, *ranges[r], st, text_color, label, d_unit)
                        cbar_artists.append((entry, cb))

    if is_zoom:
        _draw_zoom_boxes(spec, st, axes_of, dims_of, frame_letter, font)

    # Frame headers and channel labels as figure text.
    fig_w, fig_h = size
    seen_headers = set()
    for f, (x0, y0, x1, y1) in layout["frame_labels"]:
        col = columns[f]
        weight = "bold" if (col.get("is_ref") and st.get("ref_bold")) else "normal"
        if not col.get("header") and not col.get("subheader"):
            continue
        if not layout["transpose"]:
            xc = 0.5 * (x0 + x1) / fig_w
            line = font / 72.0 * 1.45
            n_lines = 2 if col.get("subheader") else 1
            base = y1 - 0.05 - (line if n_lines == 2 else 0.0)
            fig.text(xc, 1.0 - base / fig_h, col["header"], ha="center", va="bottom",
                     fontsize=font, color=text_color, fontweight=weight, gid=f"header_{f}")
            if col.get("subheader"):
                fig.text(xc, 1.0 - (y1 - 0.05) / fig_h, col["subheader"], ha="center",
                         va="bottom", fontsize=pt(font - 1), color=text_color, gid=f"header_{f}_sub")
        else:
            yc = 0.5 * (y0 + y1)
            dy = 0.5 * font / 72.0 * 1.4 if col.get("subheader") else 0.0
            fig.text(x1 / fig_w, 1.0 - (yc - dy) / fig_h, col["header"], ha="right",
                     va="center", fontsize=font, color=text_color, fontweight=weight, gid=f"header_{f}")
            if col.get("subheader"):
                fig.text(x1 / fig_w, 1.0 - (yc + dy) / fig_h, col["subheader"], ha="right",
                         va="center", fontsize=pt(font - 1), color=text_color, gid=f"header_{f}_sub")
        seen_headers.add(f)
    for r, (x0, y0, x1, y1) in layout["channel_labels"]:
        if (n_rows < 2 and is_zoom) or not st.get("channel_labels", True):
            continue  # hidden (or one zoom band: the colourbar says what it is)
        row = spec["rows"][r]
        text = row["display"]
        if not layout["transpose"]:
            fig.text((x1 - 0.08) / fig_w, 1.0 - 0.5 * (y0 + y1) / fig_h, text,
                     ha="right", va="center", rotation=90 if len(text) > 4 else 0,
                     fontsize=font + 1, color=text_color, gid=f"channel_{r}")
        else:
            fig.text(0.5 * (x0 + x1) / fig_w, 1.0 - (y1 - 0.04) / fig_h, text,
                     ha="center", va="bottom", fontsize=font + 1, color=text_color, gid=f"channel_{r}")

    layout.update({"ranges": ranges, "panel_ranges": panel_ranges, "units": units,
                   "images": images, "panel_images": panel_images,
                   "colorbar_artists": cbar_artists, "own_scale": own_scale,
                   "letter_boxes": letter_boxes, "scale_boxes": scale_boxes,
                   "data_boxes": data_boxes,
                   "style": st, "n_rows": n_rows})
    return layout


def _draw_zoom_boxes(spec, st, axes_of, dims_of, frame_letter, font):
    """Outline every zoom on its parent's panel, in the zoom's colour,
    labelled with the zoom's letter; colour-frame the zoom panels."""
    from matplotlib.patches import Polygon, Rectangle

    columns = spec["columns"]
    index = {c["key"]: f for f, c in enumerate(columns)}
    mode = st.get("zoom_color") or "color"
    lw = float(st.get("zoom_lw") or 1.0)
    for f, col in enumerate(columns):
        parent = col.get("parent")
        corners = col.get("box")
        if not parent or not corners or parent not in index:
            continue
        pf = index[parent]
        color = (ZOOM_COLORS[(f - 1) % len(ZOOM_COLORS)] if mode == "color"
                 else "white" if mode == "white" else "black")
        for (r, ff), ax in axes_of.items():
            if ff != pf or (r, ff) not in dims_of:
                continue
            w, h, turns = dims_of[(r, ff)]
            pts = []
            for fx, fy in corners:
                tx, ty = _turn_frac(fx, fy, turns)
                pts.append((tx * w, (1.0 - ty) * h))
            ax.add_patch(Polygon(pts, closed=True, fill=False, edgecolor=color, linewidth=lw,
                                 zorder=6, gid=f"zoombox_{r}_{f}"))
            if st.get("zoom_labels", True) and frame_letter.get(f):
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                ax.text(min(xs) + 0.01 * w, max(ys) - 0.01 * h, frame_letter[f], color=color,
                        fontsize=max(5.0, font), fontweight="bold", ha="left", va="top", zorder=7,
                        path_effects=_effects(st.get("letter_outline", True),
                                              "white" if mode != "black" else "black"),
                        gid=f"zoomlabel_{r}_{f}")
        if mode == "color":
            for (r, ff), ax in axes_of.items():
                if ff == f:
                    ax.add_patch(Rectangle((0, 0), 1, 1, transform=ax.transAxes, fill=False,
                                           edgecolor=color, linewidth=lw * 1.6, zorder=8,
                                           clip_on=False, gid=f"zoomframe_{r}_{f}"))


def _contrast_text_color(cmap, at):
    try:
        r, g, b, _a = cmap(at)
    except Exception:
        return "black"
    return "black" if (0.299 * r + 0.587 * g + 0.114 * b) > 0.55 else "white"


def _colorbar(fig, im, entry, size, lo, hi, st, text_color, label, unit):
    import matplotlib.cm as mcm

    font = float(st["font"])
    style = st.get("cbar_style") or "outside"
    rotated = (st.get("cbar_text") or "rotated") == "rotated"
    horizontal = entry["orient"] == "h"
    cax = fig.add_axes(_ax_rect(entry["box"], size))
    cax.set_gid(f"colorbar_{entry['row']}_{entry['frame'] if entry['frame'] is not None else 'row'}")
    import matplotlib.colors as mcolors
    norm = (mcolors.Normalize(lo, hi) if (st.get("cbar_style") in ("minmax", "minmax_in")
                                           and lo == 0.0 and hi == 1.0) else im.norm)
    mappable = mcm.ScalarMappable(norm=norm, cmap=im.get_cmap())
    cb = fig.colorbar(mappable, cax=cax, orientation="horizontal" if horizontal else "vertical")
    cb._sxm_n_ticks = int(st["cbar_ticks"])
    cb._sxm_style = style
    cb._sxm_rotated = rotated
    cb._sxm_font = font
    cb.outline.set_linewidth(0.5)
    cb.outline.set_visible(bool(st.get("cbar_outline", True)))
    cb.ax.tick_params(labelsize=pt(font - 0.5), length=2, width=0.5, pad=1.5, colors=text_color)
    if style in ("inside", "minmax_in"):
        cb._sxm_inside_colors = (_contrast_text_color(im.get_cmap(), 0.04),
                                 _contrast_text_color(im.get_cmap(), 0.96))
    set_colorbar_range(cb, lo, hi)
    if label:
        x0, y0, x1, y1 = entry["box"]
        length_in = (x1 - x0) if horizontal else (y1 - y0)
        fs_in = pt(font - 0.5) / 72.0 * 0.62
        ends = 0 if int(st.get("cbar_ticks", 2) or 0) == 0 or style in ("outside", "minmax") else 10
        # A short bar can't hold the name and both end values: name outside.
        fits = (len(label) + ends) * fs_in * 1.25 <= length_in
        if (st.get("cbar_label_pos") or "inside") == "inside" and fits:
            # The channel name sits centred *in* the bar - along it for a
            # vertical bar, flat for a horizontal one - in a colour that
            # reads on the middle of the colormap.
            cax.text(0.5, 0.5, label, transform=cax.transAxes, ha="center", va="center",
                     rotation=0 if horizontal else 90, fontsize=pt(font - 0.5),
                     color=_contrast_text_color(im.get_cmap(), 0.5), zorder=5,
                     gid=f"cbarlabel_{entry['row']}")
        elif horizontal or rotated:
            cb.set_label(label, fontsize=font, labelpad=1.0, color=text_color)
        else:
            cax.set_title(label, fontsize=font, color=text_color, loc="left", pad=2.0)
    return cb


class _MinMax:
    """Stand-in for a colourbar when only the range is printed."""

    orientation = "minmax"

    def __init__(self, cax, lo, hi, unit, label, horizontal, font, color):
        self.ax = cax
        self.unit = unit
        self.label = label
        self.horizontal = horizontal
        self.font = font
        self.color = color
        self.mappable = None


def set_colorbar_range(cb, lo, hi):
    """Few ticks only - a figure colourbar, not an axis."""
    span = abs(hi - lo)
    if isinstance(cb, _MinMax):
        cax = cb.ax
        for t in list(cax.texts):
            t.remove()
        if cb.horizontal:
            txt = f"{_tick(lo, span)} … {_tick(hi, span)}" + (f" {cb.unit}" if cb.unit else "")
        else:
            txt = f"{_tick(hi, span)}\n{_tick(lo, span)}" + (f"\n{cb.unit}" if cb.unit else "")
            if cb.label:
                txt = f"{cb.label.split(' (')[0]}\n" + txt
        cax.text(0.0, 0.5, txt, transform=cax.transAxes, fontsize=cb.font - 0.5, color=cb.color,
                 ha="left", va="center", linespacing=1.15)
        return
    cb.mappable.set_clim(lo, hi)
    horizontal = cb.orientation == "horizontal"
    style = getattr(cb, "_sxm_style", "outside")
    if int(getattr(cb, "_sxm_n_ticks", 2) or 0) == 0:
        # No ticks at all: just the gradient (and its name).
        cb.set_ticks([])
        return
    if style in ("inside", "minmax_in"):
        # Values live *inside* the bar (Publication style), centred at its
        # ends, in a colour that reads on that end of the colormap.
        cb.set_ticks([])
        for t in getattr(cb, "_sxm_inside", []):
            try:
                t.remove()
            except Exception:
                pass
        words = style == "minmax_in"
        lo_txt = "min" if words else _tick(lo, span)
        hi_txt = "max" if words else _tick(hi, span)
        lo_c, hi_c = getattr(cb, "_sxm_inside_colors", ("white", "black"))
        fs = getattr(cb, "_sxm_font", 7.0) - 0.5
        if horizontal:
            a = cb.ax.text(0.02, 0.5, lo_txt, transform=cb.ax.transAxes, ha="left",
                           va="center", fontsize=fs, color=lo_c)
            b = cb.ax.text(0.98, 0.5, hi_txt, transform=cb.ax.transAxes, ha="right",
                           va="center", fontsize=fs, color=hi_c)
        else:
            rot = 90 if getattr(cb, "_sxm_rotated", True) else 0
            a = cb.ax.text(0.5, 0.02, lo_txt, transform=cb.ax.transAxes, ha="center",
                           va="bottom", rotation=rot, fontsize=fs, color=lo_c)
            b = cb.ax.text(0.5, 0.98, hi_txt, transform=cb.ax.transAxes, ha="center",
                           va="top", rotation=rot, fontsize=fs, color=hi_c)
        cb._sxm_inside = [a, b]
        return
    minmax = getattr(cb, "_sxm_style", "outside") == "minmax"
    n = 2 if minmax else max(2, int(getattr(cb, "_sxm_n_ticks", 2) or 2))
    ticks = list(np.linspace(lo, hi, n))
    cb.set_ticks(ticks)
    cb.set_ticklabels(["min", "max"] if minmax else [_tick(t, span) for t in ticks])
    # End labels sit inside the bar's own span, so stacked / side-by-side
    # colourbars never print "0.08" on top of the neighbour's "10.9".
    axis = cb.ax.xaxis if horizontal else cb.ax.yaxis
    rotated = getattr(cb, "_sxm_rotated", True)
    major = axis.get_major_ticks()
    for i, tick in enumerate(major):
        for lab in (tick.label1, tick.label2):
            if horizontal:
                lab.set_horizontalalignment("left" if i == 0 else "right" if i == len(major) - 1 else "center")
            elif rotated:
                lab.set_rotation(90)
                lab.set_verticalalignment("center")
                lab.set_horizontalalignment("left")
                lab.set_verticalalignment("bottom" if i == 0 else "top" if i == len(major) - 1 else "center")
            else:
                lab.set_rotation(0)
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
    if a >= 1e5 or a < 1e-3:
        return f"{v:.1e}"
    # Precision follows the *range* (3 significant digits of the span) so a
    # narrow range around a large value never prints one number twice.
    ref = span if span else a
    d = max(0, min(6, 2 - int(math.floor(math.log10(ref))))) if ref > 0 else 2
    txt = f"{v:.{d}f}"
    if "." in txt:
        txt = txt.rstrip("0").rstrip(".")
    return txt


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


def _corner_box(cell, pos, size_pt, n_chars, inset=0.04):
    """Approximate inch box of a corner label (for hit-testing)."""
    x0, y0, x1, y1 = cell
    w_in = max(0.12, 0.62 * size_pt / 72.0 * max(1, n_chars))
    h_in = 1.25 * size_pt / 72.0
    cw, ch = x1 - x0, y1 - y0
    left = x0 + inset * cw if pos in ("tl", "bl") else x1 - inset * cw - w_in
    top = y0 + inset * ch if pos in ("tl", "tr") else y1 - inset * ch - h_in
    return (left, top, left + w_in, top + h_in)


def _frac_box(cell, frac):
    """Data-fraction box (fx0, fy0, fx1, fy1 with fy up) -> inch box."""
    x0, y0, x1, y1 = cell
    fx0, fy0, fx1, fy1 = frac
    cw, ch = x1 - x0, y1 - y0
    return (x0 + fx0 * cw, y1 - fy1 * ch, x0 + fx1 * cw, y1 - fy0 * ch)


def _draw_corner_text(ax, w, h, text, pos, color, size, bold, effects):
    x, y, ha, va = _corner(w, h, pos)
    return ax.text(x, y, text, color=color, fontsize=size, fontweight="bold" if bold else "normal",
                   va=va, ha=ha, path_effects=effects, zorder=7)


def _draw_scale_bar(ax, w, h, st, font, effects, gid="scalebar", show_label=True):
    """Draw the bar; returns its (fx0, fy0, fx1, fy1) data fractions incl.
    the label, for hit-testing (None if not drawn)."""
    length = _finite(st.get("scale_len_nm")) or nice_length(0.3 * w)
    if not length or length > w:
        return None
    pos = st["scale_pos"]
    color = st["scale_color"]
    if color == "auto":
        fx0 = 0.06 if pos in ("tl", "bl") else 0.94 - length / w
        fy = (0.07, 0.22) if pos in ("bl", "br") else (0.78, 0.93)
        color = auto_color(ax, (fx0, fx0 + length / w, fy[0], fy[1]))
    edge = "black" if _is_light(color) else "white"
    if effects is None:
        effects = _effects(st.get("scale_outline", True), color)
    # Thickness from the panel's shorter side, so a tall frame doesn't get
    # a slab (and the bar looks the same on every aspect ratio).
    bar_h = max(min(float(st.get("scale_thick") or 0.018), 0.05) * min(w, h), 1e-9)
    x0 = 0.06 * w if pos in ("tl", "bl") else 0.94 * w - length
    label_pt = pt(st.get("scale_font") or font)
    label_h = (0.09 * h * label_pt / max(font, 1e-6)) if show_label else 0.0
    if pos in ("bl", "br"):
        y0 = 0.07 * h
        ty, tva = y0 + bar_h + 0.02 * h, "bottom"
        box_y = (y0, y0 + bar_h + label_h)
    else:
        y0 = 0.93 * h - bar_h
        ty, tva = y0 - 0.02 * h, "top"
        box_y = (y0 - label_h, y0 + bar_h)
    from matplotlib.patches import Rectangle
    ax.add_patch(Rectangle((x0, y0), length, bar_h, facecolor=color,
                           edgecolor=edge if st.get("scale_outline", True) else color,
                           linewidth=0.4, zorder=5, gid=gid))
    if label_h:
        ax.text(x0 + 0.5 * length, ty, length_text(length), color=color, fontsize=label_pt,
                ha="center", va=tva, zorder=6, path_effects=effects, gid=gid + "_label")
    return (x0 / w, box_y[0] / h, (x0 + length) / w, box_y[1] / h)
