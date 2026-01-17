"""
Spectroscopy parsing and fitting utilities.

This module reintroduces the helpers that the GUI expects from the historical
project.  The aim is not to perfectly emulate every edge case from the lab's
old scripts; instead we provide forgiving parsers that understand common
Omicron/Anfatec exports (plain text ``.dat``/``.txt``) and reusable helpers for
matrix matching and parabola fits.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional
import math
import re

import numpy as np

from .matrix import parse_matrix_filename, matrix_dataset_key
from .channel_units import guess_channel_unit


SECTION_RE = re.compile(r"^\s*(?:\[|#|point\b|trace\b|spectrum\b)", re.IGNORECASE)


def parse_spectroscopy_file(path: Path | str) -> List[Dict[str, object]]:
    """
    Return a list of spectroscopy entries extracted from ``path``.

    Most files contain a single spectrum, but matrix acquisitions often embed
    multiple blocks separated by headers such as ``[Point 3]`` or ``# Trace 42``.
    Each entry includes:

    * ``path`` (str) – absolute path to the file.
    * ``V`` (np.ndarray) – bias axis, stored in volts when the header labels use
      ``mV``.
    * ``channels`` (dict[str, np.ndarray]) – remaining columns keyed by header
      labels or defaults like ``channel1``.
    * Metadata for downstream matching: ``time`` (datetime), ``x``/``y`` in nm
      when known, grid indices, and optional matrix index.
    """
    path = Path(path)
    matrix_specs = _parse_matrix_dat(path)
    if matrix_specs is not None:
        return matrix_specs
    text = _read_text(path)
    lines = text.replace("\r", "\n").split("\n")
    base_meta: Dict[str, object] = {}
    current_meta: Dict[str, object] = {}
    header_tokens: Optional[List[str]] = None
    rows: List[List[float]] = []
    specs: List[Dict[str, object]] = []
    block_index = 0

    def _flush():
        nonlocal rows, header_tokens, current_meta, block_index
        if not rows:
            header_tokens = None
            return
        entry = _rows_to_spec(
            rows,
            header_tokens,
            path,
            current_meta,
            block_index,
        )
        if entry:
            specs.append(entry)
            block_index += 1
        rows = []
        header_tokens = None
        current_meta = dict(base_meta)

    current_meta = dict(base_meta)
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if SECTION_RE.match(line):
            _flush()
            current_meta.update(_parse_section_metadata(line))
            continue
        key, value = _split_key_value(line)
        if key:
            norm = _normalize_meta_key(key)
            parsed = _coerce_value(value)
            base_meta[norm] = parsed
            current_meta[norm] = parsed
            continue
        tokens = _split_tokens(line)
        if not tokens:
            continue
        if _row_is_numeric(tokens):
            row = [float(tok) for tok in tokens]
            if rows:
                expected = len(rows[0])
                if len(row) > expected:
                    row = row[:expected]
                elif len(row) < expected:
                    # skip malformed lines
                    continue
            rows.append(row)
        else:
            header_tokens = _split_header_columns(raw_line)
    _flush()
    return specs


def fit_parabola_bias(V: Iterable[float], data: Iterable[float]) -> Dict[str, object]:
    """
    Fit ``data`` vs ``V`` to ``a*V^2 + b*V + c`` and return coefficients,
    uncertainties, RMSE, and a callable ``func(x)`` for plotting.
    """
    V = np.asarray(list(V), dtype=float).ravel()
    Y = np.asarray(list(data), dtype=float).ravel()
    mask = np.isfinite(V) & np.isfinite(Y)
    V = V[mask]
    Y = Y[mask]
    if V.size < 3 or Y.size < 3:
        raise ValueError("Need at least 3 finite points for a parabola fit.")
    A = np.column_stack([V ** 2, V, np.ones_like(V)])
    coeffs, residuals, rank, _ = np.linalg.lstsq(A, Y, rcond=None)
    a, b, c = coeffs
    if rank < 3:
        raise ValueError("Degenerate fit (input points are collinear).")
    if residuals.size:
        sse = float(residuals[0])
    else:
        pred = A @ coeffs
        sse = float(np.sum((Y - pred) ** 2))
    dof = max(1, V.size - 3)
    rmse = math.sqrt(max(sse / dof, 0.0))
    try:
        cov = np.linalg.inv(A.T @ A) * (sse / dof)
    except np.linalg.LinAlgError:
        cov = np.zeros((3, 3), dtype=float)
    errs = np.sqrt(np.clip(np.diag(cov), 0.0, np.inf))
    a_err, b_err, c_err = errs

    def _func(x):
        x = np.asarray(x, dtype=float)
        return a * x ** 2 + b * x + c

    return {
        "a": float(a),
        "b": float(b),
        "c": float(c),
        "a_err": float(a_err),
        "b_err": float(b_err),
        "c_err": float(c_err),
        "rmse": float(rmse),
        "func": _func,
    }


def find_last_image_for_spec(
    spec_time: Optional[datetime], images: Iterable[Dict[str, object]]
) -> Optional[Dict[str, object]]:
    """
    Return the latest image entry whose timestamp is <= ``spec_time``.
    """
    if spec_time is None:
        return None
    best = None
    best_time = None
    for img in images:
        t = img.get("time")
        if t is None:
            continue
        if t <= spec_time and (best_time is None or t > best_time):
            best = img
            best_time = t
    return best


def _matrix_base_name(stem: str) -> str:
    """
    Remove postfix tokens such as ``_matrix`` or ``-matrix`` so spectroscopy
    files map back to their parent SXM images.
    """
    stem = stem.lower().strip()
    stem = re.sub(r"(?:_matrix|-matrix).*", "", stem)
    stem = re.sub(r"(?:_spec|-spec).*", "", stem)
    return stem


# --------------------------------------------------------------------------- #
# Internal parsing helpers                                                    #
# --------------------------------------------------------------------------- #

META_KEY_MAP = {
    "x": "x",
    "x_nm": "x",
    "xn": "x",
    "xpos": "x",
    "positionx": "x",
    "y": "y",
    "y_nm": "y",
    "yn": "y",
    "ypos": "y",
    "positiony": "y",
    "row": "grid_row",
    "gridrow": "grid_row",
    "col": "grid_col",
    "column": "grid_col",
    "gridcol": "grid_col",
    "gridcols": "grid_cols",
    "gridcolumns": "grid_cols",
    "gridpointsx": "grid_cols",
    "gridrows": "grid_rows",
    "gridpointsy": "grid_rows",
    "matrixindex": "matrix_index",
    "index": "matrix_index",
    "pointindex": "matrix_index",
    "datetime": "datetime",
    "date": "date",
    "time": "time",
}


def _rows_to_spec(
    rows: List[List[float]],
    header_tokens: Optional[List[str]],
    path: Path,
    meta: Dict[str, object],
    block_idx: int,
) -> Optional[Dict[str, object]]:
    data = np.asarray(rows, dtype=float)
    if data.ndim == 1:
        data = data.reshape(-1, 1)
    n_rows, n_cols = data.shape
    if n_cols == 0 or n_rows == 0:
        return None
    # Primary axis defaults to first column (bias/time), optional Z in second.
    bias = data[:, 0].copy()
    z_axis = data[:, 1].copy() if n_cols >= 2 else None
    channel_labels = _channel_labels(header_tokens, n_cols)
    start_idx = 1 if n_cols == 1 else 2 if n_cols >= 3 else 1
    channels = {}
    for idx, label in enumerate(channel_labels, start=1):
        if idx < start_idx:
            continue
        col = data[:, idx].copy()
        channels[label] = col
    bias = _normalize_bias_axis(bias, header_tokens)

    axis_label = "Bias"
    axis_unit = "V"
    alt_axis = None
    alt_label = None
    alt_unit = None
    if header_tokens:
        first_label = header_tokens[0].strip().lower()
        if first_label == "mv" or first_label.endswith("(mv)"):
            axis_unit = "V"
        elif first_label.endswith("(v)"):
            axis_unit = "V"
        if len(header_tokens) > 1:
            alt_label = header_tokens[1].strip() or "Z"
    if z_axis is not None:
        alt_label = alt_label or "Z"
        alt_unit = "nm"
        alt_axis = z_axis.copy()
        try:
            if np.nanmax(np.abs(alt_axis)) < 1e-6:
                alt_axis = alt_axis * 1e9  # assume meters -> nm
        except Exception:
            pass

    entry = {
        "path": str(path),
        "V": bias,
        "channels": channels,
        "AxisLabel": axis_label,
        "AxisUnit": axis_unit,
        "AltAxis": alt_axis,
        "AltAxisLabel": alt_label,
        "AltAxisUnit": alt_unit,
    }
    entry.update(_extract_meta(meta, path, block_idx))
    return entry


def _channel_labels(header_tokens: Optional[List[str]], n_cols: int) -> List[str]:
    labels: List[str] = []
    tokens = header_tokens or []
    for idx in range(1, n_cols):
        label = tokens[idx] if idx < len(tokens) else ""
        label = _clean_channel_label(label) or f"channel{idx}"
        labels.append(label)
    return labels


def _clean_channel_label(label: str) -> str:
    label = str(label or "").strip()
    label = label.replace("/", "_").replace("(", "").replace(")", "")
    label = re.sub(r"[^a-zA-Z0-9_+-]", "_", label)
    label = re.sub(r"_{2,}", "_", label)
    return label.strip("_")


def _normalize_bias_axis(bias: np.ndarray, header_tokens: Optional[List[str]]) -> np.ndarray:
    if not bias.size:
        return bias
    scale = 1.0
    if header_tokens:
        unit_tokens = [
            str(tok).lower()
            for tok in header_tokens[: min(len(header_tokens), 3)]
            if tok
        ]
        if any("kv" in tok for tok in unit_tokens):
            scale = 1e3
        elif any("mv" in tok for tok in unit_tokens):
            scale = 1e-3
    return bias * scale


def _extract_meta(meta: Dict[str, object], path: Path, block_idx: int) -> Dict[str, object]:
    info: Dict[str, object] = {}
    info["time"] = (
        _parse_datetime(meta.get("datetime"))
        or _parse_date_and_time(meta.get("date"), meta.get("time"))
        or _mtime(path)
    )
    info["x"] = _maybe_float(meta.get("x"))
    info["y"] = _maybe_float(meta.get("y"))
    info["grid_cols"] = _maybe_int(meta.get("grid_cols"))
    info["grid_rows"] = _maybe_int(meta.get("grid_rows"))
    info["grid_col"] = _maybe_int(meta.get("grid_col"))
    info["grid_row"] = _maybe_int(meta.get("grid_row"))
    matrix_index = meta.get("matrix_index")
    if matrix_index is None:
        matrix_index = _guess_index_from_name(path, block_idx)
    info["matrix_index"] = _maybe_int(matrix_index)
    return info


def _guess_index_from_name(path: Path, block_idx: int) -> Optional[int]:
    stem = path.stem.lower()
    m = re.search(r"(?:matrix|spec|idx|point)[-_]?(\d+)", stem)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    m = re.search(r"(\d+)$", stem)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return block_idx if block_idx is not None else None


def _extract_section_value(pattern: str, line: str) -> Optional[int]:
    m = re.search(pattern, line, flags=re.IGNORECASE)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _parse_section_metadata(line: str) -> Dict[str, object]:
    meta: Dict[str, object] = {}
    idx = _extract_section_value(r"index[:= ]+(\d+)", line)
    if idx is not None:
        meta["matrix_index"] = idx
    row = _extract_section_value(r"row[:= ]+(\d+)", line)
    if row is not None:
        meta["grid_row"] = row
    col = _extract_section_value(r"col(?:umn)?[:= ]+(\d+)", line)
    if col is not None:
        meta["grid_col"] = col
    return meta


def _split_key_value(line: str):
    for sep in ("=", ":"):
        if sep in line:
            left, right = line.split(sep, 1)
            key = left.strip()
            value = right.strip()
            if key:
                return key, value
    # Tab-delimited metadata uses a single tab; skip data/header rows with multiple tabs.
    if "\t" in line and line.count("\t") == 1:
        left, right = line.split("\t", 1)
        key = left.strip()
        value = right.strip()
        if key:
            return key, value
    return None, None


def _split_tokens(line: str) -> List[str]:
    tokens = re.split(r"[;,\s]+", line)
    return [tok for tok in tokens if tok]


def _split_header_columns(line: str) -> List[str]:
    for sep in ("\t", ";", ","):
        if sep in line:
            return [part.strip() for part in line.split(sep) if part.strip()]
    parts = re.split(r"\s{2,}", line.strip())
    if len(parts) > 1:
        return [part.strip() for part in parts if part.strip()]
    return [line.strip()] if line.strip() else []


def _row_is_numeric(tokens: List[str]) -> bool:
    for tok in tokens:
        try:
            float(tok)
        except Exception:
            return False
    return True


def _normalize_meta_key(key: str) -> str:
    key = key.strip().lower()
    key = re.sub(r"[^a-z0-9]+", "_", key)
    key = key.strip("_")
    return META_KEY_MAP.get(key, key)


def _coerce_value(value: str):
    value = value.strip().strip('"')
    if not value:
        return ""
    try:
        if "." in value or "e" in value.lower():
            return float(value)
        return int(value)
    except Exception:
        return value


def _maybe_float(value) -> Optional[float]:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def _maybe_int(value) -> Optional[int]:
    try:
        if value is None:
            return None
        return int(value)
    except Exception:
        return None


def _parse_matrix_dat(path: Path) -> Optional[List[Dict[str, object]]]:
    """Return spectra for Omicron/Anfatec matrix ``.dat`` files."""
    name = path.name.lower()
    if not name.endswith(".dat") or "matrix" not in name:
        return None
    text = _read_text(path).replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    if len(lines) < 4:
        return None
    header_tokens = lines[0].split("\t")
    coord_tokens = lines[1].split("\t") if len(lines) > 1 else []
    if len(header_tokens) < 4 or len(coord_tokens) < 4:
        return None
    try:
        xs = [float(tok) for tok in header_tokens[3:]]
        ys = [float(tok) for tok in coord_tokens[3:]]
    except Exception:
        return None
    if not xs or not ys or len(xs) != len(ys):
        return None
    base, channel_code, channel_label = parse_matrix_filename(path.name)
    dataset_key, display_label = matrix_dataset_key(base, channel_code)
    channel_display = display_label or channel_label or Path(path).stem
    rows: List[tuple[float, List[float]]] = []
    for line in lines[2:]:
        tokens = line.split("\t")
        if len(tokens) < 3 + len(xs):
            continue
        try:
            bias = float(tokens[2])
            values = [float(tok) for tok in tokens[3 : 3 + len(xs)]]
        except Exception:
            continue
        rows.append((bias, values))
    if not rows:
        return None
    bias_axis = np.asarray([row[0] for row in rows], dtype=float)
    unit_tokens = coord_tokens[:3]
    if unit_tokens and len(unit_tokens) >= 3:
        unit = unit_tokens[2].lower()
        if "mv" in unit:
            bias_axis *= 1e-3
        elif "kv" in unit:
            bias_axis *= 1e3
    data = np.asarray([row[1] for row in rows], dtype=float)
    if data.shape[1] != len(xs):
        return None
    x_arr = np.asarray(xs, dtype=float)
    y_arr = np.asarray(ys, dtype=float)
    x_unique, x_inverse = np.unique(np.round(x_arr, 6), return_inverse=True)
    y_unique, y_inverse = np.unique(np.round(y_arr, 6), return_inverse=True)
    grid_cols = int(x_unique.size) or data.shape[1]
    grid_rows = int(y_unique.size) or 1
    specs: List[Dict[str, object]] = []
    for col in range(data.shape[1]):
        channel_series = data[:, col].copy()
        unit = guess_channel_unit(channel_display)
        unit_map = {channel_display: unit} if unit else {}
        entry = {
            "path": str(path),
            "V": bias_axis.copy(),
            "channels": {channel_display: channel_series},
            "data": (bias_axis.copy(), channel_series),
            "x": float(x_arr[col]),
            "y": float(y_arr[col]),
            "grid_cols": grid_cols,
            "grid_rows": grid_rows,
            "grid_col": int(x_inverse[col]),
            "grid_row": int(y_inverse[col]),
            "matrix_index": int(y_inverse[col] * grid_cols + x_inverse[col]),
            "time": _mtime(path),
            "matrix_dataset": dataset_key or base or Path(path).stem,
            "channel_name": channel_display,
            "channel_code": channel_code,
            "unit_map": unit_map,
        }
        if unit:
            entry["unit"] = unit
        specs.append(entry)
    return specs


def _parse_datetime(value) -> Optional[datetime]:
    if isinstance(value, datetime):
        return value
    if not value:
        return None
    text = str(value).strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%d.%m.%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(text)
    except Exception:
        return None


def _parse_date_and_time(date_val, time_val) -> Optional[datetime]:
    if not date_val and not time_val:
        return None
    date_str = str(date_val).strip() if date_val else ""
    time_str = str(time_val).strip() if time_val else ""
    combined = f"{date_str} {time_str}".strip()
    return _parse_datetime(combined) or _parse_datetime(date_str)


def _mtime(path: Path) -> Optional[datetime]:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return None


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return path.read_text(encoding="cp1252", errors="ignore")


__all__ = [
    "parse_spectroscopy_file",
    "fit_parabola_bias",
    "find_last_image_for_spec",
    "_matrix_base_name",
]



