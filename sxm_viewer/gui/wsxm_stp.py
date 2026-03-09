"""Utilities for emitting WSxM-compatible .stp image files."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Mapping, Any

import numpy as np

DEFAULT_STP_HEADER_SIZE = 12288

_BLUE_INFO = {
    "Control Point 0": "(0 , 255)",
    "Control Point 1": "(85 , 255)",
    "Control Point 2": "(160 , 179)",
    "Control Point 3": "(220 , 42)",
    "Control Point 4": "(255 , 4)",
    "Number of Control Points": 5,
}

_GREEN_INFO = {
    "Control Point 0": "(0 , 255)",
    "Control Point 1": "(17 , 255)",
    "Control Point 2": "(47 , 234)",
    "Control Point 3": "(106 , 159)",
    "Control Point 4": "(182 , 47)",
    "Control Point 5": "(232 , 7)",
    "Control Point 6": "(255 , 6)",
    "Number of Control Points": 7,
}

_RED_INFO = {
    "Control Point 0": "(0 , 255)",
    "Control Point 1": "(32 , 136)",
    "Control Point 2": "(94 , 57)",
    "Control Point 3": "(152 , 17)",
    "Control Point 4": "(188 , 6)",
    "Control Point 5": "(255 , 6)",
    "Number of Control Points": 6,
}


def _format_section(name: str, fields: Mapping[str, Any]) -> str:
    lines = [f"[{name}]\r\n\r\n"]
    for key, value in fields.items():
        lines.append(f"    {key}: {value}\r\n")
    lines.append("\r\n")
    return "".join(lines)


def _fmt_value(value: float | None, unit: str | None = None) -> str:
    if value is None or not np.isfinite(value):
        return f"0 {unit}" if unit else "0"
    if unit:
        return f"{value:.6g} {unit}"
    return f"{value:.6g}"


def _escape_comment_block(text: str | None) -> str:
    """Return WSxM-friendly comment text (newlines escaped as ``\\n``)."""
    if not text:
        text = "Exported from SXM Viewer"
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not cleaned:
        cleaned = "Exported from SXM Viewer"
    # WSxM escapes literal backslashes as ``\\`` inside the comment block.
    cleaned = cleaned.replace("\\", "\\\\")
    return cleaned.replace("\n", "\\n")


def save_wsxm_stp(
    path: str | Path,
    data: np.ndarray,
    *,
    channel: str = "Z",
    x_nm: float | None = None,
    y_nm: float | None = None,
    z_unit: str | None = "nm",
    setpoint_pa: float | None = None,
    bias_v: float | None = None,
    angle_deg: float | None = None,
    comment: str | None = None,
    comment_block: str | None = None,
    head_type: str | None = "STM",
    timestamp: str | None = None,
    z_gain: float | None = 1.0,
    name: str | None = None,
    saved_with_version: str = "5.0 Develop 11.2",
    header_size: int | None = None,
) -> Path:
    """
    Write ``data`` to ``path`` using the WSxM .stp container structure.

    Parameters mirror the WSxM metadata fields. ``data`` must be a 2-D array.
    """
    arr = np.asarray(data, dtype=np.float64)
    if arr.ndim != 2:
        raise ValueError("WSxM STP export expects a 2-D array")
    rows, cols = arr.shape
    finite = arr[np.isfinite(arr)]
    if finite.size == 0:
        finite = np.array([0.0], dtype=np.float64)
    fill_value = float(np.nanmedian(finite))
    if not np.isfinite(fill_value):
        fill_value = 0.0
    arr = np.array(arr, copy=True, dtype=np.float64)
    mask = ~np.isfinite(arr)
    if mask.any():
        arr[mask] = fill_value
    arr = np.nan_to_num(arr, nan=fill_value, posinf=fill_value, neginf=fill_value, copy=False)
    if not np.isfinite(arr).all():
        raise ValueError("Non-finite values detected while building WSxM payload")
    z_min = float(np.nanmin(arr))
    z_max = float(np.nanmax(arr))
    z_amp = z_max - z_min
    flipped = np.flip(arr, axis=1)
    payload = np.ascontiguousarray(flipped, dtype="<f8").tobytes()
    width_nm = float(x_nm) if x_nm not in (None, "") else float(cols)
    height_nm = float(y_nm) if y_nm not in (None, "") else float(rows)
    ts = timestamp or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    angle_txt = f"{float(angle_deg):.6g}" if angle_deg is not None else "0"
    z_label = z_unit or "arb."
    if comment_block:
        comment_field = _escape_comment_block(comment_block)
    else:
        comment_text = comment or f"Exported from SXM Viewer on {ts}"
        comment_field = _escape_comment_block(comment_text)
    general_section = _format_section(
        "General Info",
        {
            "Acquisition channel": channel or "Z",
            "Acquisition primary channel": channel or "Z",
            "Head type": head_type or "STM",
            "Image Data Type": "double",
            "Image processes": "converted",
            "Number of columns": cols,
            "Number of rows": rows,
            "X scanning direction": "forward",
            "Y scanning direction": "up",
            "Z Amplitude": _fmt_value(z_amp, z_label),
        },
    )
    control_section = _format_section(
        "Control",
        {
            "Angle": angle_txt,
            "Set Point": _fmt_value(setpoint_pa, "pA"),
            "Topography Bias": _fmt_value(bias_v, "V"),
            "X Amplitude": _fmt_value(width_nm, "nm"),
            "Y Amplitude": _fmt_value(height_nm, "nm"),
            "Z Gain": _fmt_value(z_gain, None),
        },
    )
    misc_section = _format_section(
        "Miscellaneous",
        {
            "Center map": "No",
            "Comments": comment_field,
            "DAC Maximum": 0,
            "DAC Minimum": 0,
            "Maximum": f"{z_max:.6g}",
            "Minimum": f"{z_min:.6g}",
            "Relative Z value": "No",
             "Name": name or "",
            "Saved with version": saved_with_version,
            "Version": "1.0 (April 2000)",
            "View type": "2d",
            "Z Scale Factor": 1,
            "Z Scale Offset": 0,
        },
    )
    head_section = _format_section(
        "Head Settings",
        {"X Calibration": "100 nm/V"},
    )
    graphics_section = _format_section(
        "Graphic Layers",
        {
            "Layer 0": "Name: Base Image; Active: Yes;",
            "Layer 1": "Name: Scalebar; Active: Yes;",
            "Layer 2": "Name: Contours; Number of contours: 10;",
            "Number of Layers": 3,
        },
    )
    palette_section = _format_section(
        "Palette Generation Settings",
        {
            "Derivate Mode for the last blue Point": "Automatic",
            "Derivate Mode for the last green Point": "Automatic",
            "Derivate Mode for the last red Point": "Automatic",
            "Is there a particular palette index colored?": "No",
            "Smooth Blue": "No",
            "Smooth Green": "No",
            "Smooth Red": "No",
        },
    )
    scale_bar_section = _format_section(
        "Scale Bar Settings",
        {
            "Color": "rgb(0, 0, 0)",
            "Default Color": "Yes",
            "Position": "(0.2, 0.9)",
            "Size X": 0.2,
            "Size Y": 0.2,
        },
    )
    rep_section = _format_section(
        "3D Representation Params",
        {
            "Alpha transparency": 1,
            "Ambient light intensity": 0.2,
            "Axes color": "rgb(204, 204, 204)",
            "Background color": "rgb(0, 0, 0)",
            "Base color": "rgb(204, 204, 204)",
            "Diffuse light intensity": 1,
            "Display type": 2,
            "Draw axes": "No",
            "Draw base": "Yes",
            "Draw profiles": "No",
            "Draw text": "No",
            "Interline spacing": 2,
            "Keep proportions between scales": "No",
            "Light rotation angle": 45,
            "Light tilt angle": 45,
            "Offset X": 0,
            "Offset Y": 0,
            "Offset Z": 0,
            "Perspective projection": "No",
            "Printing color": "rgb(0, 255, 0)",
            "Rotation angle": 20,
            "Scale XY": 1,
            "Scale Z": 1,
            "Specular light intensity": 0.5,
            "Specular reflection": "Yes",
            "Text color": "rgb(255, 255, 90)",
            "Texture bitmap": "",
            "Texture bitmap format": "PNG",
            "Texture type": 0,
            "Tilt angle": 45,
            "Z offset auto": "Yes",
        },
    )

    blue_section = _format_section("Blue Info", _BLUE_INFO)
    green_section = _format_section("Green Info", _GREEN_INFO)
    red_section = _format_section("Red Info", _RED_INFO)

    body = (
        rep_section
        + blue_section
        + control_section
        + general_section
        + graphics_section
        + green_section
        + head_section
        + misc_section
        + palette_section
        + red_section
        + scale_bar_section
        + "[Header end]\r\n"
    )
    prefix_base = "WSxM file copyright UAM\r\nSxM Image file\r\nImage header size: "
    suffix = "\r\n\r\n"
    body_bytes = body.encode("latin-1", errors="replace")
    fixed_len = len(prefix_base) + len(suffix) + len(body_bytes)
    if header_size is None:
        size = fixed_len
        while True:
            total = fixed_len + len(str(size))
            if total == size:
                break
            size = total
    else:
        if header_size < fixed_len:
            raise ValueError(f"Provided header_size {header_size} is smaller than required {fixed_len}")
        size = header_size
    prefix = f"{prefix_base}{size}{suffix}"
    header_bytes = prefix.encode("latin-1") + body_bytes
    if header_size is not None and len(header_bytes) < header_size:
        header_bytes += b"\x00" * (header_size - len(header_bytes))
    output_path = Path(path)
    output_path.write_bytes(header_bytes + payload)
    return output_path


__all__ = ["DEFAULT_STP_HEADER_SIZE", "save_wsxm_stp"]
