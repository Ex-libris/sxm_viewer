"""Channel-name classification heuristics for report content selection.

Names arrive in two flavors: sanitized spectroscopy channel labels
("Current_A", "Frequency_Shift_Hz", "LI_Demod_1_X_A", "Z_rel_m") and image
channel captions from headers ("Current", "Frequency Shift", "Topography").
Classification is deliberately permissive substring matching over both.

Report heuristics encoded here (user-specified):
- constant-height images are best represented by current + frequency-shift
  (or lock-in / dI/dV-like) channels, not topography;
- constant-current / topographic images by topography only;
- spectroscopy curves are most meaningful as frequency shift, current, or
  lock-in signal vs the sweep axis, in that order of preference.
"""
from __future__ import annotations

import re

import numpy as np

# Order matters below: first match wins.
_CLASS_PATTERNS = (
    ("bias", (r"bias",)),
    ("freq", (r"freq", r"shift", r"\bdf\b", r"omega", r"resonan")),
    ("lockin", (r"\bli[_\s]?\d*\s*demod", r"\bli\d*[xy]\b", r"\blia", r"lock[\s_-]?in",
                r"demod", r"l\d+l?[xy]\b", r"didv", r"di_dv")),
    ("current", (r"current", r"\bit\b", r"^i[_\s(]", r"^i$")),
    ("z", (r"topo", r"height", r"^z\b", r"^z[_\s(]", r"z[_\s]?rel", r"\bz_m\b")),
    ("time", (r"time",)),
    ("amplitude", (r"amplitude", r"\bamp\b")),
    ("phase", (r"phase",)),
    ("excitation", (r"excitation", r"drive")),
)

# Preference order for the "signal" curve of a spectrum / grid point.
CURVE_PRIORITY = ("freq", "current", "lockin", "amplitude", "phase", "excitation")


def classify_channel(name):
    """Coarse class for a channel name: bias/freq/lockin/current/z/time/
    amplitude/phase/excitation/other."""
    low = str(name or "").strip().lower().replace("-", "_")
    if not low:
        return "other"
    for cls, patterns in _CLASS_PATTERNS:
        for pat in patterns:
            if re.search(pat, low):
                return cls
    return "other"


def _varies(values):
    arr = np.asarray(values, dtype=float).ravel()
    finite = arr[np.isfinite(arr)]
    if finite.size < 2:
        return False
    span = float(finite.max() - finite.min())
    scale = max(abs(float(finite.max())), abs(float(finite.min())), 1e-30)
    return span > 1e-9 * scale


def pick_signal_channels(channels, limit=4):
    """Interesting *signal* channels from a spec's channel dict, ordered by
    CURVE_PRIORITY then dict order; axis-ish (bias/z/time) and flat/dead
    channels are skipped. Returns list of (name, class)."""
    if not channels:
        return []
    scored = []
    for order, (name, values) in enumerate(channels.items()):
        cls = classify_channel(name)
        if cls in ("bias", "z", "time"):
            continue
        try:
            if not _varies(values):
                continue
        except Exception:
            continue
        rank = CURVE_PRIORITY.index(cls) if cls in CURVE_PRIORITY else len(CURVE_PRIORITY)
        scored.append((rank, order, name, cls))
    scored.sort()
    return [(name, cls) for _, _, name, cls in scored[:limit]]


def pick_image_channels(fds, tag, topo_idx):
    """Channel indices to render for an image, per the CH/CC heuristics.

    ``fds`` are header channel descriptors (Caption/FileName), ``tag`` the
    CH/CC tag ('constant-height'/'constant-current'/None), ``topo_idx`` the
    detected topography channel index (may be None). Returns an ordered
    list of (index, class) - up to two entries for constant-height scans
    (current + frequency shift / lock-in), one topography entry otherwise.
    """
    fds = list(fds or [])
    if tag == "constant-height":
        by_class = {}
        for idx, fd in enumerate(fds):
            text = f"{fd.get('Caption', '')} {fd.get('FileName', '')}"
            cls = classify_channel(text)
            by_class.setdefault(cls, idx)
        picks = []
        if "current" in by_class:
            picks.append((by_class["current"], "current"))
        for cls in ("freq", "lockin"):
            if cls in by_class and len(picks) < 2:
                picks.append((by_class[cls], cls))
        if picks:
            return picks
    if topo_idx is not None and topo_idx < len(fds):
        return [(topo_idx, "z")]
    return [(0, "other")] if fds else []
