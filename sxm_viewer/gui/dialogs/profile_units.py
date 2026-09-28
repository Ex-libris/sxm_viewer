"""Pure, Qt/matplotlib-free unit/formatting helpers for the profile dialog.

Kept separate from ``profile_dialog.py`` so the unit-correctness logic that
caused the ``1e-15 Hz`` and ``Current [A] (A)`` bugs can be unit-tested
without booting a QApplication. Deliberately has no matplotlib or Qt import
of its own (just numpy) - this dialog runs on some measurement PCs with a
locked-down, offline Python install, so nothing here should ever become a
reason to add a dependency.

Three responsibilities live here:

* Converting already-normalized channel/distance values (``nm``/``A``/``V``/
  ``Hz`` - the output of ``data.io.normalize_unit_and_data``) into true SI
  BASE units (``m``/``A``/``V``/``Hz``) exactly once, at the point profile
  data enters this dialog. Matplotlib's ``EngFormatter`` (used for axis
  ticks) and :func:`si_format` (used for one-off values - the readout
  strip, cursor position, provenance headers) both only produce correct
  prefixes (``pA``, ``nm``, ``mHz``, ...) when fed true base-unit values
  and told the base unit's symbol; feeding either an already-prefixed unit
  like ``"nm"`` would produce nonsense like ``"n nm"``.
* Composing an axis label from a channel name plus unit exactly once, so a
  channel name that already embeds its unit (``"Current [A]"``) never gets
  the unit appended a second time.
* :func:`si_format`/:func:`si_scale`: a standalone SI-prefix formatter for
  single scalar values, independent of any plotting library, so the same
  "139.3 nm" / "52.0 pA" style formatting is available for the readout
  strip and provenance headers wherever the code is not formatting an
  actual matplotlib axis tick (where ``EngFormatter`` is used instead).
"""
from __future__ import annotations

from typing import Optional, Tuple

import numpy as np

_SI_PREFIXES = {
    -24: "y", -21: "z", -18: "a", -15: "f", -12: "p", -9: "n", -6: "μ", -3: "m",
    0: "", 3: "k", 6: "M", 9: "G", 12: "T", 15: "P", 18: "E", 21: "Z", 24: "Y",
}


def si_scale(value: float) -> Tuple[float, str]:
    """``(scale, prefix)`` such that ``value * scale`` lands in ``[1, 1000)``.

    Same contract as pyqtgraph's ``siScale`` (which this replaces): multiply
    a value by ``scale`` to get the number to display next to
    ``prefix + base_unit``. Two values multiplied by the *same* scale (e.g.
    a measurement and its tolerance in :func:`fmt_tol`-style callers) share
    one SI prefix instead of drifting to different ones.
    """
    value = abs(float(value))
    if value == 0 or not np.isfinite(value):
        return 1.0, ""
    exp = int(np.floor(np.log10(value) / 3.0) * 3)
    exp = max(-24, min(24, exp))
    return 10.0 ** (-exp), _SI_PREFIXES.get(exp, "")


def si_format(value: float, unit: str = "", precision: int = 3) -> str:
    """``"52.0 pA"`` - a scalar with an automatically chosen SI prefix."""
    scale, prefix = si_scale(value)
    text = ("%%.%dg" % max(0, precision)) % (float(value) * scale)
    return f"{text} {prefix}{unit}".rstrip()

# Units already produced by ``data.io.normalize_unit_and_data`` that need one
# more factor to reach a true SI base unit pyqtgraph can auto-prefix.
# ``nm`` is the app-wide convention for distance/topography (see CLAUDE.md's
# "Coordinate frames" section) but is not itself an SI base unit.
_TO_SI_BASE = {
    "nm": ("m", 1e-9),
}

# Units that are already SI base and only need to be recognized so
# ``enableAutoSIPrefix`` is turned on for them.
_ALREADY_SI_BASE = {"m", "a", "v", "hz", "s"}


def to_si_base(values, unit: Optional[str]) -> Tuple["np.ndarray", str, bool]:
    """Convert ``values`` (already in the app's normalized unit) to SI base.

    Returns ``(values_in_base_unit, base_unit_symbol, use_auto_si_prefix)``.
    For an unrecognized/blank unit, values pass through unchanged and
    ``use_auto_si_prefix`` is False (nothing to safely auto-prefix).
    """
    arr = np.asarray(values, dtype=float) if values is not None else np.asarray([], dtype=float)
    key = str(unit or "").strip()
    key_lower = key.lower()
    if key_lower in _TO_SI_BASE:
        base_unit, factor = _TO_SI_BASE[key_lower]
        return arr * factor, base_unit, True
    if key_lower in _ALREADY_SI_BASE:
        return arr, key, True
    return arr, key, False


def from_si_base(value: float, base_unit: str) -> float:
    """Invert :func:`to_si_base` for a single scalar (e.g. a dragged marker).

    ``base_unit`` is the symbol :func:`to_si_base` returned alongside the
    converted array, so this always finds the matching entry (or is a no-op
    for units that were never converted).
    """
    key_lower = str(base_unit or "").strip().lower()
    for original, (converted, factor) in _TO_SI_BASE.items():
        if converted == key_lower:
            return float(value) / factor
    return float(value)


def compose_axis_label(name: Optional[str], unit: Optional[str]) -> Tuple[str, str]:
    """Compose a clean ``(label, unit)`` pair for a pyqtgraph ``setLabel`` call.

    Strips any ``[unit]`` / ``(unit)`` already embedded in ``name`` (e.g. a
    channel caption of ``"Current [A]"`` paired with ``unit="A"``) so the
    caller never produces ``setLabel(..., "Current [A]", units="A")`` -
    which pyqtgraph would render as the doubled ``"Current [A] (A)"``.

    ``name`` may be empty; ``unit`` may be empty. Never raises.
    """
    label = str(name or "").strip()
    unit_str = str(unit or "").strip()
    if unit_str:
        for junk in (f"[{unit_str}]", f"({unit_str})"):
            if junk in label:
                label = label.replace(junk, "").strip()
    return label, unit_str


__all__ = ["to_si_base", "from_si_base", "compose_axis_label", "si_scale", "si_format"]
