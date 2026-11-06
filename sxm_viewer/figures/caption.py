"""Draft figure captions from what the figure already knows (Qt-free).

The figure window knows every frame's acquisition parameters (mode, bias,
setpoint, size, tip height) and every panel's processing (filters, Zero,
crop, colour range), and the letters exactly as drawn
(``height_series.panel_letters``). ``build_caption`` turns that into a
draft the user edits rather than writes - two paragraphs:

* the caption: what the panels show, imaging parameters, scale bars;
* the processing statement journals ask for (filters per panel, Zero,
  crops, how the colour scales were set).

Values that are the same for every panel are said once ("V = 5 mV");
values that differ are listed with their letters ("V = +0.3 V (a, b),
+0.5 V (c)"). Letters run as ranges ("a-c").
"""
from __future__ import annotations

import math

from . import height_series as hs

_CHANNEL_WORDS = {
    "current": "tunnelling current I",
    "freq": "frequency shift Δf",
    "lockin": "lock-in signal",
    "z": "topography z",
    "phase": "phase",
    "amplitude": "amplitude",
}
_MODE_WORDS = {"constant-height": "constant-height", "constant-current": "constant-current"}


def _finite(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    return v if math.isfinite(v) else None


def letters_text(letters):
    """``["a", "b", "c", "e"]`` -> ``"a-c, e"`` (only plain single letters
    are merged into ranges)."""
    letters = [str(x) for x in letters if x]
    if not letters:
        return ""
    plain = all(len(x) == 1 and x.isalpha() for x in letters)
    if not plain:
        return ", ".join(letters)
    out, run = [], [letters[0]]
    for x in letters[1:]:
        if ord(x) == ord(run[-1]) + 1:
            run.append(x)
        else:
            out.append(run)
            run = [x]
    out.append(run)
    parts = []
    for run in out:
        parts.append(f"{run[0]}–{run[-1]}" if len(run) >= 3 else ", ".join(run))
    return ", ".join(parts)


def _grouped(pairs, fmt):
    """``[(letter, value)]`` -> "value" when all equal, else "v1 (a, b), v2
    (c)" in first-seen order. ``None`` values are left out."""
    pairs = [(lt, v) for lt, v in pairs if v is not None]
    if not pairs:
        return ""
    order, by = [], {}
    for lt, v in pairs:
        key = fmt(v)
        if key not in by:
            by[key] = []
            order.append(key)
        by[key].append(lt)
    if len(order) == 1:
        return order[0]
    return ", ".join(f"{k} ({letters_text(by[k])})" for k in order)


def _size_text(w, h):
    if not (w and h):
        return None
    unit, f = ("µm", 1e-3) if max(w, h) >= 1000 else ("nm", 1.0)

    def num(v):
        v *= f
        return f"{v:.3g}" if v < 100 else f"{v:.0f}"
    return f"{num(w)} × {num(h)} {unit}²"


def _bias_text(v):
    v = _finite(v)
    if v is None:
        return None
    return hs.bias_text(v).replace("V = ", "")


def _setpoint_text(value, unit):
    value = _finite(value)
    if value is None:
        return None
    return hs.si_text(value, unit or "")


def _upper_first(text):
    return text[:1].upper() + text[1:] if text else text


def _channel_phrase(row):
    if row.get("kind") == "mixed":
        return "the main channel of each frame"
    word = _CHANNEL_WORDS.get(row.get("cls"))
    return word or str(row.get("display") or row.get("label") or "signal")


def build_caption(spec, *, letters, members, panels, style=None, processing=None,
                  crops=None, scale_lengths=None):
    """Draft caption + processing statement.

    ``letters``: ``{(row, frame): letter}`` as drawn (empty = letters off);
    ``members``: ``{key: {tag, bias_v, setpoint, setpoint_unit, z_nm}}``;
    ``panels``: ``{(row, key): {width_nm, height_nm}}`` as drawn (cropped);
    ``processing``: ``{(row, frame): {"filters": text, "zero": bool,
    "own_range": bool}}``; ``crops``: ``{key | "*": box}``;
    ``scale_lengths``: ``{frame: nm}`` of the drawn scale bars.
    Returns ``{"caption": str, "processing": str}``."""
    st = hs.resolve_style(style)
    cols = spec.get("columns") or []
    rows = spec.get("rows") or []
    processing = processing or {}
    crops = crops or {}
    mode = spec.get("mode")
    n_f, n_r = len(cols), len(rows)

    # A letter for every frame (its first lettered panel); with letters off,
    # frames are named by position.
    frame_letter = {}
    for f in range(n_f):
        lt = next((letters[(r, f)] for r in range(n_r) if (r, f) in letters), None)
        frame_letter[f] = lt.strip("()).") if lt else None
    have_letters = all(frame_letter.get(f) for f in range(n_f))

    def ref(f):
        return frame_letter.get(f) if have_letters else (cols[f].get("name") if f < n_f else "")

    def ref_list(frames):
        return letters_text([ref(f) for f in frames]) if have_letters else \
            ", ".join(str(ref(f)) for f in frames)

    sentences = []
    all_frames = list(range(n_f))
    lead = f"({ref_list(all_frames)}) " if have_letters and n_f > 1 else ""

    # 1. What is shown.
    tags = {(members.get(c["key"]) or {}).get("tag") for c in cols}
    tags.discard(None)
    acq = _MODE_WORDS.get(next(iter(tags))) if len(tags) == 1 else None
    kinds = [_channel_phrase(r) for r in rows]
    what = " and ".join(kinds) if len(kinds) <= 2 else ", ".join(kinds[:-1]) + " and " + kinds[-1]
    instrument = "nc-AFM" if any(r.get("cls") == "freq" for r in rows) else "STM"
    if instrument == "nc-AFM" and any(r.get("cls") == "current" for r in rows):
        instrument = "STM/nc-AFM"
    noun = f"{acq + ' ' if acq else ''}{instrument} images"
    if mode == "height":
        first = f"{lead}{_upper_first(noun)} of the same area at different tip heights"
    elif mode == "bias":
        first = f"{lead}{_upper_first(noun)} of the same area at different sample biases"
    elif mode == "zoom" and n_f > 1:
        first = (f"({ref(0)}) Overview {noun.replace('images', 'image')} and "
                 f"({ref_list(all_frames[1:])}) close-ups of the areas outlined in ({ref(0)})")
    else:
        first = f"{lead}{_upper_first(noun)}"
    if n_r > 1:
        side = "columns" if st.get("transpose") else "rows"
        first += f"; {side}: {what}"
    else:
        first += f" ({what})"
    sentences.append(first + ".")

    # 2. Imaging parameters (each said once when shared).
    params = []
    if mode == "height" or any(c.get("dz_pm") is not None for c in cols):
        dz = [(ref(f), c.get("dz_pm")) for f, c in enumerate(cols)]
        dz_txt = _grouped(dz, lambda v: f"{v:+.0f} pm" if abs(v) >= 0.5 else "0 pm")
        refc = next((f for f, c in enumerate(cols) if c.get("is_ref")), None)
        if dz_txt:
            tail = f" relative to ({ref(refc)})" if refc is not None and have_letters else ""
            params.append(f"Δz = {dz_txt}{tail} (positive = tip retracted)")
    bias = _grouped([(ref(f), (members.get(c["key"]) or {}).get("bias_v")) for f, c in enumerate(cols)],
                    _bias_text)
    if bias:
        params.append(f"V = {bias}")
    sp = _grouped([(ref(f), ((members.get(c["key"]) or {}).get("setpoint"),
                             (members.get(c["key"]) or {}).get("setpoint_unit")))
                   for f, c in enumerate(cols) if (members.get(c["key"]) or {}).get("setpoint") is not None],
                  lambda v: _setpoint_text(*v))
    if sp:
        word = "I" if all(((members.get(c["key"]) or {}).get("setpoint_unit") == "A") for c in cols) else "setpoint"
        params.append(f"{word} = {sp}")
    sizes = []
    for f, c in enumerate(cols):
        p = next((panels.get((r, c["key"])) for r in range(n_r) if panels.get((r, c["key"]))), None)
        if p:
            sizes.append((ref(f), (_finite(p.get("width_nm")), _finite(p.get("height_nm")))))
    size_txt = _grouped(sizes, lambda v: _size_text(*v))
    if size_txt:
        params.append(size_txt)
    if params:
        sentences.append("; ".join(params) + ".")

    # 3. Scale bars.
    if scale_lengths and st.get("scale_bar") != "off":
        bars = _grouped([(ref(f), _finite(v)) for f, v in sorted(scale_lengths.items())],
                        lambda v: hs.length_text(v))
        if bars:
            sentences.append(f"Scale bar{'s' if len(scale_lengths) > 1 else ''}: {bars}.")

    caption = " ".join(sentences)

    # Processing statement.
    proc = []

    def cell_ref(r, f):
        base = ref(f) if have_letters else cols[f].get("name")
        if n_r > 1:
            return f"{base} ({rows[r].get('display') or rows[r].get('label')})"
        return str(base)

    # Filters: say what was applied where; unfiltered panels in one phrase.
    filt, order, plain = {}, [], []
    for f in range(n_f):
        for r in range(n_r):
            if (r, f) not in processing:
                continue
            text = (processing[(r, f)].get("filters") or "").strip()
            if not text:
                plain.append((r, f))
                continue
            if text not in filt:
                filt[text] = []
                order.append(text)
            filt[text].append((r, f))

    def where(cells):
        by_row = {}
        for r, f in cells:
            by_row.setdefault(r, set()).add(f)
        if n_r > 1 and all(len(fs) == n_f for fs in by_row.values()):
            return ", ".join(f"{rows[r].get('display') or rows[r].get('label')} row" for r in sorted(by_row))
        frames = sorted({f for _r, f in cells})
        if n_r == 1 or all(by_row.get(r, set()) == set(frames) for r in range(n_r)):
            return ref_list(frames)
        return ", ".join(cell_ref(r, f) for r, f in sorted(cells, key=lambda x: (x[1], x[0])))

    if not order:
        proc.append("No filtering (raw data).")
    elif len(order) == 1 and not plain:
        proc.append(f"All images: {order[0][:1].lower() + order[0][1:]}.")
    else:
        parts = [f"{_upper_first(text)}: {where(filt[text])}" for text in order]
        proc.append("; ".join(parts) + ".")
        if plain:
            proc.append("Other images: no filtering.")
    zeroed = [(r, f) for (r, f), i in processing.items() if i.get("zero")]
    if zeroed:
        if len(zeroed) == n_r * n_f:
            proc.append("Values are relative to the bottom of each image's colour range (Zero).")
        else:
            proc.append("Values relative to the bottom of the colour range (Zero): "
                        + ", ".join(cell_ref(r, f) for r, f in sorted(zeroed, key=lambda x: (x[1], x[0]))) + ".")
    cropped = [f for f, c in enumerate(cols) if crops.get(c["key"]) or crops.get("*")]
    if cropped:
        sources = []
        for f in cropped:
            c = cols[f]
            box = crops.get(c["key"]) or crops.get("*")
            p = next((panels.get((r, c["key"])) for r in range(n_r) if panels.get((r, c["key"]))), None)
            w, h = (_finite((p or {}).get("width_nm")), _finite((p or {}).get("height_nm")))
            fw, fh = box[2] - box[0], box[3] - box[1]
            full = _size_text(w / fw, h / fh) if (w and h and fw > 0 and fh > 0) else None
            sources.append((ref(f), full))
        txt = _grouped(sources, lambda v: v)
        noun = "scan" if len(cropped) == 1 else "scans"
        what = f"({ref_list(cropped)}) cropped from " + (f"{txt} {noun}" if txt else f"larger {noun}")
        proc.append(_upper_first(what) + ".")
    if st.get("equal_scale") and mode != "zoom":
        proc.append("All images are shown at the same scale.")
    own = [(r, f) for (r, f), i in processing.items() if i.get("own_range")]
    if st.get("range_mode") == "row":
        scale = "Colour scales are shared within each channel"
        if own:
            scale += " except " + ", ".join(cell_ref(r, f) for r, f in sorted(own, key=lambda x: (x[1], x[0])))
        proc.append(scale + ".")
    else:
        proc.append("Each image has its own colour scale.")
    return {"caption": caption, "processing": " ".join(t for t in proc if t)}
