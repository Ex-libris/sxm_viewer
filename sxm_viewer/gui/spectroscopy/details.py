"""Human-readable formatting of a spectroscopy spec for the Details panel.

Pure functions over the spec dict - no viewer, no Qt, no widgets - so this
is directly testable. Extracted from ``SXMGridViewer`` (see
``docs/refactor/GOD_CLASS_PLAN.md``); it never used ``self`` at all.

The output deliberately **summarizes** large values rather than printing
them: a spec carries full channel arrays and sweep axes, and dumping those
into a text box freezes the UI on a grid point with thousands of samples.
Arrays render as ``array(shape=..., dtype=...)``, containers as
``list(1024)``.
"""
from __future__ import annotations

from pathlib import Path

import numpy as np

# Shown first, in this order, when present - the fields that identify a
# spectrum and where it sits.
_HEADLINE_KEYS = (
    "path", "source", "time", "file_mtime", "image_key", "primary_image_key",
    "matrix_dataset", "matrix_index", "x", "y",
    "AxisLabel", "AxisUnit", "AltAxisLabel", "AltAxisUnit",
)

# Rendered in their own sections instead of the raw dump.
_STRUCTURED_KEYS = {"channels", "AxisChoices"}


def format_value(value):
    """Compact one-line rendering; never expands a large array."""
    if value is None:
        return "None"
    if isinstance(value, (str, int, float, bool)):
        return str(value)
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return f"dict({len(value)})"
    if isinstance(value, (list, tuple, set)):
        return f"{type(value).__name__}({len(value)})"
    if hasattr(value, "shape"):
        try:
            arr = np.asarray(value)
            return f"array(shape={arr.shape}, dtype={arr.dtype})"
        except Exception:
            return type(value).__name__
    return str(value)


def _section(lines, title, rows):
    """Append a titled block, skipping it entirely when empty."""
    rows = [r for r in rows if r]
    if not rows:
        return
    lines.append("")
    lines.append(title)
    lines.extend(rows)


def metadata_lines(spec):
    """Full Details-panel text for one spec, as a list of lines."""
    if not spec:
        return ["No spectroscopy metadata."]

    lines = ["Spectroscopy details", ""]
    lines.extend(f"{key}: {format_value(spec.get(key))}"
                 for key in _HEADLINE_KEYS if key in spec)

    _section(lines, "Position:", [
        f"  display: {v}" if (v := str(spec.get('site_display') or '').strip()) else "",
        f"  key: {v}" if (v := str(spec.get('site_key') or '').strip()) else "",
        f"  summary: {v}" if (v := str(spec.get('site_summary') or '').strip()) else "",
    ])

    assignment_reason = str(spec.get("assignment_reason_label")
                            or spec.get("assignment_reason") or "").strip()
    assignment_conf = str(spec.get("assignment_confidence") or "").strip()
    assignment_summary = str(spec.get("assignment_summary") or "").strip()
    # shared_image_keys is reported only *within* an assignment block - it
    # is meaningless without one, and emitting a bare "Assignment:" header
    # for it alone would differ from the original behaviour.
    if assignment_summary or assignment_conf or assignment_reason:
        shared_keys = spec.get("shared_image_keys") or []
        _section(lines, "Assignment:", [
            f"  confidence: {assignment_conf}" if assignment_conf else "",
            f"  reason: {assignment_reason}" if assignment_reason else "",
            f"  summary: {assignment_summary}" if assignment_summary else "",
            f"  shared_image_keys: {format_value(shared_keys)}" if shared_keys else "",
        ])

    channels = spec.get("channels") or {}
    if channels:
        rows = []
        for name, values in channels.items():
            try:
                shape = np.asarray(values).shape
            except Exception:
                shape = "?"
            rows.append(f"  - {name}: shape={shape}")
        _section(lines, f"Channels ({len(channels)}):", rows)

    axis_choices = spec.get("AxisChoices") or []
    if axis_choices:
        rows = []
        for ax in axis_choices:
            key = ax.get("key") or ax.get("label") or "Axis"
            label = ax.get("label") or "Axis"
            unit = ax.get("unit") or ""
            rows.append(f"  - {key}: {label}" + (f" ({unit})" if unit else ""))
        _section(lines, f"Axis choices ({len(axis_choices)}):", rows)

    # Unlike the sections above, this header is emitted unconditionally -
    # it marks the end of the curated view, so its presence is meaningful
    # even when every field happened to be rendered structurally above.
    lines.append("")
    lines.append("Raw fields:")
    lines.extend(
        f"  {key}: {format_value(spec.get(key))}"
        for key in sorted(spec.keys(), key=lambda s: str(s).lower())
        if key not in _STRUCTURED_KEYS
    )
    return lines


def metadata_text(spec):
    """``metadata_lines`` joined for direct display."""
    return "\n".join(metadata_lines(spec))


__all__ = ["metadata_lines", "metadata_text", "format_value"]
