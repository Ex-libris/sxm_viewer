"""Unit formatting helpers."""
from __future__ import annotations

import math
import re
from typing import Tuple

import numpy as np
_NUMERIC_RE = re.compile(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?")

_UNIT_DISPLAY_CHOICES = {
    # order favours the most common STM ranges first (centered on the native unit)
    "nm": [("nm", 1.0), ("pm", 1e3), ("um", 1e-3), ("mm", 1e-6), ("m", 1e-9)],
    "A": [("nA", 1e9), ("pA", 1e12), ("fA", 1e15), ("uA", 1e6), ("mA", 1e3), ("A", 1.0), ("kA", 1e-3)],
    "V": [("mV", 1e3), ("uV", 1e6), ("V", 1.0), ("kV", 1e-3)],
    "Hz": [("kHz", 1e-3), ("Hz", 1.0), ("MHz", 1e-6), ("GHz", 1e-9)],
}

_SI_BASE_UNITS = {
    "nm": ("m", 1e-9),
    "pm": ("m", 1e-12),
    "um": ("m", 1e-6),
    "mm": ("m", 1e-3),
    "m": ("m", 1.0),
    "fA": ("A", 1e-15),
    "A": ("A", 1.0),
    "V": ("V", 1.0),
    "Hz": ("Hz", 1.0),
}


def _auto_display_unit(unit: str, data: np.ndarray) -> Tuple[str, float]:
    """
    Return (label, factor) describing how to scale ``data`` for comfortable display.
    """
    unit_key = (unit or "").strip()
    default = (unit_key, 1.0)
    options = _UNIT_DISPLAY_CHOICES.get(unit_key)
    if not options:
        return default
    arr = np.asarray(data, dtype=float)
    finite = arr[np.isfinite(arr)]
    if finite.size == 0:
        return default
    mag = float(np.nanmax(np.abs(finite)))
    if not np.isfinite(mag) or mag <= 0:
        return default
    def _pick(range_min, range_max):
        for label, factor in options:
            scaled = mag * factor
            if range_min <= scaled < range_max:
                return label, factor
        return None

    found = _pick(1.0, 1000.0)
    if found:
        return found

    best = None
    best_diff = float("inf")
    for label, factor in options:
        scaled = mag * factor
        if scaled <= 0:
            continue
        diff = abs(math.log10(scaled))
        if diff < best_diff:
            best = (label, factor)
            best_diff = diff

    if best:
        return best
    return options[0]


def _safe_float(value, default=None):
    """Best-effort conversion that tolerates unit suffixes like '80 nm'."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    try:
        text = str(value).strip()
    except Exception:
        return default
    if not text:
        return default
    match = _NUMERIC_RE.search(text.replace(",", "."))
    if match:
        try:
            return float(match.group(0))
        except Exception:
            return default
    try:
        return float(text)
    except Exception:
        return default
__all__ = [
    "_auto_display_unit",
    "_safe_float",
]



