"""Parse spectroscopy-like pages from RHK SM4 files.

This is a lightweight bridge so the viewer can place markers / preview traces.
It does not attempt to reproduce `rhkpy`'s rich structuring; instead it emits
the same entry shape as `sxm_viewer.data.spectroscopy.parse_spectroscopy_file`.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import numpy as np

from ...utils.logging import log
from .adapter import _ensure_sm4_reader


def parse_rhksm4_spectroscopy(path: Path | str) -> List[Dict[str, object]]:
    """
    Return spectroscopy entries for `.sm4` files.

    We currently treat RHK "Line" pages (`RHK_PageDataType == 1`) as a set of
    spectra along the second dimension.
    """
    sm4_path = Path(path)
    reader = _ensure_sm4_reader()
    if reader is None:
        return []
    try:
        sm4 = reader.RHKsm4(str(sm4_path))
    except Exception as exc:
        log(f"[SM4] Failed to open {sm4_path.name}: {exc}")
        return []

    entries: List[Dict[str, object]] = []
    try:
        pages = list(sm4)
    except Exception:
        pages = []
    line_pages = [p for p in pages if getattr(p, "attrs", {}).get("RHK_PageDataType") == 1]
    if not line_pages:
        return []

    for page in line_pages:
        attrs = getattr(page, "attrs", {}) or {}
        label = getattr(page, "label", "") or "sm4_line"
        data = np.asarray(getattr(page, "data", None))
        coords = getattr(page, "coords", None)
        if data is None or data.size == 0:
            continue
        if coords and isinstance(coords, (list, tuple)) and len(coords) >= 1:
            axis_vals = np.asarray(coords[0][1], dtype=float).ravel()
            axis_label = str(attrs.get("RHK_Xlabel") or coords[0][0] or "X")
            axis_unit = str(attrs.get("RHK_Xunits") or "")
        else:
            axis_vals = np.arange(data.shape[0] if data.ndim else data.size, dtype=float)
            axis_label = "Index"
            axis_unit = ""

        # Normalize to (npoints, ntraces)
        if data.ndim == 1:
            traces = data.reshape((-1, 1))
        elif data.ndim >= 2:
            traces = data.reshape((data.shape[0], -1))
        else:
            continue

        time = _parse_time(attrs.get("RHK_DateTime")) or _mtime(sm4_path)
        x0 = _safe_float(attrs.get("RHK_Xoffset"))
        y0 = _safe_float(attrs.get("RHK_Yoffset"))
        x_nm = _coerce_pos_to_nm(x0)
        y_nm = _coerce_pos_to_nm(y0)

        for trace_idx in range(traces.shape[1]):
            y_key = f"trace{trace_idx:03d}"
            channel_values = np.asarray(traces[:, trace_idx], dtype=float).ravel()
            entry = {
                "path": str(sm4_path),
                "source": "rhksm4",
                "time": time,
                "x": x_nm,
                "y": y_nm,
                "channels": {label: channel_values},
                "channel_name": label,
                "channel_code": None,
                "V": axis_vals.copy(),
                "AxisLabel": axis_label,
                "AxisUnit": axis_unit,
                "AltAxis": None,
                "AltAxisLabel": None,
                "AltAxisUnit": None,
                "grid_rows": traces.shape[1],
                "grid_cols": 1,
                "grid_row": trace_idx,
                "grid_col": 0,
                "matrix_dataset": sm4_path.stem,
                "matrix_index": trace_idx,
                "matrix_file": True,
                "trace_key": y_key,
            }
            entries.append(entry)
    return entries


def _mtime(path: Path) -> Optional[datetime]:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return None


def _parse_time(value) -> Optional[datetime]:
    if isinstance(value, datetime):
        return value
    if not value:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%m/%d/%Y %H:%M:%S", "%d.%m.%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(text)
    except Exception:
        return None


def _safe_float(value) -> Optional[float]:
    try:
        return float(value)
    except Exception:
        return None


def _coerce_pos_to_nm(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    try:
        v = float(value)
    except Exception:
        return value
    if abs(v) < 1e-6:
        return v * 1e9
    return v


__all__ = ["parse_rhksm4_spectroscopy"]

