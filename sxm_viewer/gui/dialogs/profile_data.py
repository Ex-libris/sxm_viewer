"""Profile data formatting helpers."""
from __future__ import annotations


def fmt_length(title: str, length_nm) -> str:
    if length_nm is None:
        return f"{title}: N/A"
    return f"{title}: {length_nm:.3f} nm"


def format_stats_text(active, saved) -> str:
    lines = []
    if active:
        lines.append(fmt_length("Active", active.get("length_nm")))
    for idx, data in enumerate(saved, 1):
        lines.append(fmt_length(f"Overlay {idx}", data.get("length_nm")))
    return "\n".join(lines) if lines else "No profile data"


def axis_label(unit: str | None) -> str:
    unit = unit or "px"
    return f"d ({unit})"


def format_marker_delta(axis_delta, axis_scale, display_unit: str | None):
    unit = display_unit or "px"
    if axis_scale is not None:
        return axis_delta * axis_scale, unit or "nm"
    return axis_delta, unit or "px"



