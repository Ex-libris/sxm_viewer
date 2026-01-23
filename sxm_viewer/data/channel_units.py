"""Heuristics for assigning physical units to spectroscopy channels."""
from __future__ import annotations

import re
from collections import OrderedDict
from typing import Optional

_CHANNEL_UNIT_HINTS = OrderedDict([
    ("it_to_pc", "A"),
    ("it-pc", "A"),
    ("it", "A"),
    ("current", "A"),
    ("df", "Hz"),
    ("freq_shift", "Hz"),
    ("freq", "Hz"),
    ("frequency", "Hz"),
    ("hz", "Hz"),
    ("drive", "Hz"),
    ("ampl", "pm"),
    ("amp", "pm"),
    ("qplus", "pm"),
    ("dz", "nm"),
    ("height", "nm"),
    ("topo", "nm"),
    ("z", "nm"),
    ("phase", "deg"),
    ("bias", "V"),
    ("voltage", "V"),
])


def guess_channel_unit(name: Optional[str]) -> Optional[str]:
    """
    Return an educated guess for the physical unit of ``name``.

    Many Omicron/Anfatec datasets encode the acquisition channel inside the
    filename (e.g. ``It_to_PC``, ``df``, ``QPlusAmpl``). The viewer uses the
    returned value for inline labelling and axis descriptions when the raw data
    does not explicitly advertise a unit.
    """
    if not name:
        return None
    text = str(name).strip().lower()
    if not text:
        return None
    tokens = [tok for tok in re.split(r"[^a-z0-9]+", text) if tok]
    if tokens:
        for token, unit in _CHANNEL_UNIT_HINTS.items():
            if token in tokens:
                return unit
    for token, unit in _CHANNEL_UNIT_HINTS.items():
        if token in text:
            return unit
    return None


__all__ = ["guess_channel_unit"]
