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
from typing import Dict, Iterable, List, Optional, Tuple
import math
import re

import numpy as np

from .matrix import MatrixDataCube, parse_matrix_filename, matrix_dataset_key
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
    matrix_payload = _parse_matrix_dat(path)
    if matrix_payload is not None:
        matrix_specs, matrix_cube = matrix_payload
        if matrix_cube is not None:
            for entry in matrix_specs:
                entry['matrix_cube'] = matrix_cube
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
    if not header_tokens:
        raise SpectroscopyParseError(path, "Missing header row with column labels.")
    if not rows:
        raise SpectroscopyParseError(path, "No numeric data rows found.")
    expected_len = len(rows[0])
    if expected_len < 2:
        raise SpectroscopyParseError(path, "Expected at least two columns (bias + channel).")
    for idx, row in enumerate(rows, 1):
        if len(row) != expected_len:
            raise SpectroscopyParseError(
                path, f"Inconsistent column count: row {idx} has {len(row)}, expected {expected_len}."
            )
    data = np.asarray(rows, dtype=float)
    if data.ndim == 1:
        data = data.reshape(-1, 1)
    n_rows, n_cols = data.shape
    if n_cols == 0 or n_rows == 0:
        return None
    labels_raw = header_tokens or []
    cleaned_labels = [_clean_channel_label(tok) for tok in labels_raw]
    if len(cleaned_labels) < n_cols:
        cleaned_labels.extend([""] * (n_cols - len(cleaned_labels)))
    # Identify which column truly represents bias by header label.
    bias_col = None
    for idx, lbl in enumerate(cleaned_labels):
        low = (lbl or "").lower()
        if low == "bias" or low.startswith("bias"):
            bias_col = idx
            break
    if bias_col is None or bias_col >= n_cols:
        raise SpectroscopyParseError(path, "Unable to locate Bias column in header.")
    time_col = None
    for idx, lbl in enumerate(cleaned_labels):
        if idx == bias_col:
            continue
        low = (lbl or "").lower()
        if "time" in low or low == "t":
            time_col = idx
            break
    z_col = None
    topo_col = None
    for idx, lbl in enumerate(cleaned_labels):
        if idx in (bias_col, time_col):
            continue
        low = (lbl or "").lower()
        if low in ("dz", "z", "z_rel", "zrel", "height"):
            z_col = idx
            break
    for idx, lbl in enumerate(cleaned_labels):
        if idx in (bias_col, time_col, z_col):
            continue
        low = (lbl or "").lower()
        if "topo" in low or "topography" in low or "piezo" in low or low in ("zabs", "z_abs", "absz", "abs_z"):
            topo_col = idx
            break

    bias = data[:, bias_col].copy()
    z_axis = data[:, z_col].copy() if z_col is not None else None
    topo_axis = data[:, topo_col].copy() if topo_col is not None else None
    time_axis = data[:, time_col].copy() if time_col is not None else None
    channel_labels = cleaned_labels if cleaned_labels else ["" for _ in range(n_cols)]

    channels = {}
    for idx in range(n_cols):
        if idx in (bias_col, z_col, topo_col, time_col):
            continue
        label = channel_labels[idx] if idx < len(channel_labels) else ""
        label = label or f"channel{idx}"
        channels[label] = data[:, idx].copy()

    bias, axis_label, axis_unit = _normalize_bias_axis(bias, header_tokens)
    if header_tokens and bias_col < len(header_tokens):
        bias_label_raw = header_tokens[bias_col].strip() or None
        if bias_label_raw:
            axis_label = bias_label_raw

    axes_choices: List[Dict[str, object]] = [
        {"key": "bias", "label": axis_label, "unit": axis_unit, "values": bias.copy()}
    ]

    alt_axis = None
    alt_label = None
    alt_unit = None
    if z_axis is not None:
        if header_tokens and z_col is not None and z_col < len(header_tokens):
            alt_label = header_tokens[z_col].strip() or None
        alt_label = alt_label or "Z"
        alt_unit = alt_unit or "nm"
        alt_axis = z_axis.copy()
        try:
            if np.nanmax(np.abs(alt_axis)) < 1e-6:
                alt_axis = alt_axis * 1e9  # assume meters -> nm
        except Exception:
            pass
        axes_choices.append({"key": "z", "label": alt_label, "unit": alt_unit, "values": alt_axis})
    if topo_axis is not None:
        topo_label = None
        topo_unit = "nm"
        if header_tokens and topo_col is not None and topo_col < len(header_tokens):
            topo_label = header_tokens[topo_col].strip() or None
        topo_label = topo_label or "Topo"
        topo_axis = topo_axis.copy()
        try:
            max_abs = np.nanmax(np.abs(topo_axis))
            if max_abs < 1e-3:
                topo_axis = topo_axis * 1e9  # assume meters -> nm
        except Exception:
            pass
        axes_choices.append({"key": "topo", "label": topo_label, "unit": topo_unit, "values": topo_axis})

    # Expose a time axis choice only when an explicit time-like column exists.
    if time_axis is not None:
        time_label = None
        time_unit = "s"
        if header_tokens and time_col is not None and time_col < len(header_tokens):
            time_label = header_tokens[time_col].strip() or None
            low = time_label.lower() if time_label else ""
            if "ms" in low:
                time_unit = "ms"
            elif "us" in low:
                time_unit = "us"
            elif "ns" in low:
                time_unit = "ns"
        axes_choices.append({"key": "time", "label": time_label or "Time", "unit": time_unit, "values": time_axis.copy()})

    # Backward compatibility: expose the first alternate axis via AltAxis*
    if len(axes_choices) > 1 and alt_axis is None:
        alt = axes_choices[1]
        alt_axis = np.asarray(alt["values"], dtype=float)
        alt_label = alt.get("label")
        alt_unit = alt.get("unit")

    entry = {
        "path": str(path),
        "V": bias,
        "channels": channels,
        "AxisLabel": axis_label,
        "AxisUnit": axis_unit,
        "AxisChoices": axes_choices,
        "AltAxis": alt_axis,
        "AltAxisLabel": alt_label,
        "AltAxisUnit": alt_unit,
    }

    # Make axes plottable as Y channels too (e.g., bias vs time).
    for axis in axes_choices:
        ch_label = _clean_channel_label(axis.get("label") or axis.get("key") or "axis")
        if ch_label.lower() in ("", "bias", "v"):
            ch_label = "bias"
        channels.setdefault(ch_label, np.asarray(axis.get("values"), dtype=float))

    entry.update(_extract_meta(meta, path, block_idx))
    z_level = None
    z_label = None
    z_unit = None
    if topo_axis is not None:
        z_level = _constant_axis_value(topo_axis)
        if z_level is not None:
            z_label = topo_label or "Topo"
            z_unit = topo_unit or "nm"
    if z_level is None and alt_axis is not None:
        z_level = _constant_axis_value(alt_axis)
        if z_level is not None:
            z_label = alt_label or "Z"
            z_unit = alt_unit or "nm"
    if z_level is not None:
        entry["z_level_nm"] = float(z_level)
        entry["z_level_label"] = str(z_label or "Z")
        entry["z_level_unit"] = str(z_unit or "nm")
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


def _normalize_bias_axis(
    bias: np.ndarray, header_tokens: Optional[List[str]]
) -> tuple[np.ndarray, str, str]:
    """
    Normalize primary axis and preserve unit hints.
    For Omicron/Anfatec `.dat` inputs we treat the first column as Bias in mV
    unless the header explicitly states otherwise. Values are kept as-is.
    """
    if not bias.size:
        return bias, "Bias", "V"
    label = "Bias"
    unit = "mV"  # default for Omicron/Anfatec .dat spectra
    scale = 1.0
    if header_tokens:
        unit_tokens = [
            str(tok).lower()
            for tok in header_tokens[: min(len(header_tokens), 3)]
            if tok
        ]
        if any("kv" in tok for tok in unit_tokens):
            unit = "kV"
            scale = 1.0
        elif any("mv" in tok for tok in unit_tokens):
            unit = "mV"
            scale = 1.0
        elif any("v" in tok for tok in unit_tokens):
            unit = "V"
            scale = 1.0
    return bias * scale, label, unit


def _extract_meta(meta: Dict[str, object], path: Path, block_idx: int) -> Dict[str, object]:
    info: Dict[str, object] = {}
    file_mtime = _mtime(path)
    info["file_mtime"] = file_mtime
    info["time"] = (
        _parse_datetime(meta.get("datetime"))
        or _parse_date_and_time(meta.get("date"), meta.get("time"))
        or file_mtime
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


def _axis_values(values: np.ndarray, quantized: np.ndarray, uniques: np.ndarray) -> np.ndarray:
    """Return canonical axis values using the first occurrence of each quantized bin."""
    first_idx: Dict[float, int] = {}
    for idx, key in enumerate(quantized):
        fk = float(key)
        if fk not in first_idx:
            first_idx[fk] = idx
    return np.asarray([values[first_idx[float(key)]] for key in uniques], dtype=float)


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


def _constant_axis_value(values: np.ndarray, tol_nm: float = 1e-3) -> Optional[float]:
    try:
        arr = np.asarray(values, dtype=float).ravel()
    except Exception:
        return None
    if arr.size == 0:
        return None
    finite = arr[np.isfinite(arr)]
    if finite.size == 0:
        return None
    try:
        span = float(np.nanmax(finite) - np.nanmin(finite))
        center = float(np.nanmedian(finite))
    except Exception:
        return None
    limit = max(float(tol_nm), abs(center) * 1e-6)
    if span <= limit:
        return center
    return None


def _parse_matrix_dat(path: Path) -> Optional[Tuple[List[Dict[str, object]], MatrixDataCube]]:
    """Return spectra and structured dataset for Omicron/Anfatec matrix ``.dat`` files."""
    name = path.name.lower()
    if not name.endswith(".dat") or "matrix" not in name:
        return None
    text = _read_text(path).replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    if len(lines) < 4:
        raise MatrixDatError(path, "Matrix .dat file must include header, coordinate, and data rows.")
    header_tokens = lines[0].split("\t")
    coord_tokens = lines[1].split("\t") if len(lines) > 1 else []
    if len(header_tokens) < 4 or len(coord_tokens) < 4:
        raise MatrixDatError(path, "Matrix header rows must contain time/dz/bias and >=1 coordinate entry.")
    try:
        xs = [float(tok) for tok in header_tokens[3:]]
        ys = [float(tok) for tok in coord_tokens[3:]]
    except Exception:
        raise MatrixDatError(path, "Non-numeric X/Y coordinate entries.")
    if not xs or not ys:
        raise MatrixDatError(path, "Missing spatial coordinates.")
    if len(xs) != len(ys):
        raise MatrixDatError(path, "X/Y coordinate lengths differ.")
    base, channel_code, channel_label = parse_matrix_filename(path.name)
    dataset_key, display_label = matrix_dataset_key(base, channel_code)
    channel_display = display_label or channel_label or Path(path).stem
    rows: List[tuple[float, List[float]]] = []
    for idx, line in enumerate(lines[2:], start=3):
        tokens = line.split("\t")
        if len(tokens) < 3 + len(xs):
            raise MatrixDatError(path, f"Row {idx} has insufficient columns for {len(xs)} pixels.")
        try:
            bias = float(tokens[2])
            values = [float(tok) for tok in tokens[3 : 3 + len(xs)]]
        except Exception:
            raise MatrixDatError(path, f"Non-numeric data encountered in row {idx}.")
        rows.append((bias, values))
    if not rows:
        raise MatrixDatError(path, "Matrix file contains no bias rows.")
    bias_axis = np.asarray([row[0] for row in rows], dtype=float)
    bias_axis, axis_label, axis_unit = _normalize_bias_axis(bias_axis, header_tokens)
    data = np.asarray([row[1] for row in rows], dtype=float)
    if data.shape[1] != len(xs):
        raise MatrixDatError(path, "Data rows do not match declared coordinate count.")
    x_arr = np.asarray(xs, dtype=float)
    y_arr = np.asarray(ys, dtype=float)

    def _deduplicate_indices(x_vals, y_vals, decimals=9):
        seen = set()
        keep = []
        for idx, (xv, yv) in enumerate(zip(x_vals, y_vals)):
            key = (round(float(xv), decimals), round(float(yv), decimals))
            if key in seen:
                continue
            seen.add(key)
            keep.append(idx)
        if len(keep) == len(x_vals):
            return None
        return np.asarray(keep, dtype=int)

    dedup_idx = _deduplicate_indices(x_arr, y_arr)
    if dedup_idx is not None:
        x_arr = x_arr[dedup_idx]
        y_arr = y_arr[dedup_idx]
        data = data[:, dedup_idx]

    if data.shape[1] != len(x_arr):
        raise MatrixDatError(path, "Deduplicated coordinate count mismatches data columns.")

    x_quant = None
    y_quant = None
    x_unique = y_unique = None
    x_inverse = y_inverse = None
    grid_cols = grid_rows = 0
    for decimals in range(6, -3, -1):
        x_q = np.round(x_arr, decimals)
        y_q = np.round(y_arr, decimals)
        x_u, x_inv = np.unique(x_q, return_inverse=True)
        y_u, y_inv = np.unique(y_q, return_inverse=True)
        cols = int(x_u.size)
        rows = int(y_u.size)
        if cols <= 0 or rows <= 0:
            continue
        if cols * rows == x_arr.size:
            x_quant, y_quant = x_q, y_q
            x_unique, y_unique = x_u, y_u
            x_inverse, y_inverse = x_inv, y_inv
            grid_cols, grid_rows = cols, rows
            break
    # Fallback: treat 1×1 “matrices” as single-point traces instead of failing.
    if (grid_cols <= 0 or grid_rows <= 0 or x_unique is None or y_unique is None) and x_arr.size == 1:
        grid_cols = grid_rows = 1
        x_unique = np.array([x_arr[0]])
        y_unique = np.array([y_arr[0]])
        x_quant = np.array([x_arr[0]])
        y_quant = np.array([y_arr[0]])
        x_inverse = np.array([0], dtype=int)
        y_inverse = np.array([0], dtype=int)
    if grid_cols <= 0 or grid_rows <= 0 or x_unique is None or y_unique is None:
        raise MatrixDatError(path, "Unable to reconstruct grid dimensions from coordinates.")
    if grid_cols * grid_rows != x_arr.size:
        raise MatrixDatError(path, "X/Y coordinates do not form a rectangular grid.")

    x_axis = _axis_values(x_arr, x_quant, x_unique)
    y_axis = _axis_values(y_arr, y_quant, y_unique)

    cube = None
    is_single_point_matrix = grid_cols == 1 and grid_rows == 1
    if not is_single_point_matrix:
        cube = np.empty((bias_axis.size, grid_rows, grid_cols), dtype=float)
        cube.fill(np.nan)
        for col in range(data.shape[1]):
            cube[:, y_inverse[col], x_inverse[col]] = data[:, col]
        if np.isnan(cube).any():
            raise MatrixDatError(path, "Incomplete data grid (NaNs present after cube reconstruction).")

    specs: List[Dict[str, object]] = []
    for col in range(data.shape[1]):
        channel_series = data[:, col].copy()
        unit = guess_channel_unit(channel_display)
        unit_map = {channel_display: unit} if unit else {}
        channels = {channel_display: channel_series}
        # Expose bias as a pseudo-channel for plotting if desired.
        channels.setdefault("bias", bias_axis.copy())
        entry = {
            "path": str(path),
            "V": bias_axis.copy(),
            "channels": channels,
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
            "matrix_file": True,
            "unit_map": unit_map,
            "AxisLabel": axis_label,
            "AxisUnit": axis_unit,
            "AxisChoices": [{"key": "bias", "label": axis_label, "unit": axis_unit, "values": bias_axis.copy()}],
            "AltAxis": None,
            "AltAxisLabel": None,
            "AltAxisUnit": None,
        }
        if unit:
            entry["unit"] = unit
        if is_single_point_matrix:
            entry["matrix_file"] = False
            entry["matrix_dataset"] = None
            entry["matrix_index"] = None
            entry["grid_cols"] = 1
            entry["grid_rows"] = 1
        specs.append(entry)
    matrix_cube = None
    if not is_single_point_matrix:
        cube_metadata = {
            "channel_code": channel_code,
            "axis_label": axis_label,
            "axis_unit": axis_unit,
            "time": _mtime(path),
        }
        matrix_cube = MatrixDataCube(
            path=str(path),
            dataset_key=dataset_key or base or Path(path).stem,
            channel=channel_display,
            bias=bias_axis.copy(),
            x=x_axis,
            y=y_axis,
            data=cube,
            metadata=cube_metadata,
        )
    return specs, matrix_cube


def is_matrix_file_entry(entry: Optional[Dict[str, object]]) -> bool:
    """Return True when ``entry`` originates from an Omicron/Anfatec matrix .dat file."""
    if not entry:
        return False
    try:
        if entry.get("matrix_file"):
            return True
        path = entry.get("path")
        if not path:
            return False
        return "matrix" in Path(str(path)).name.lower()
    except Exception:
        return False


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
    "is_matrix_file_entry",
    "SpectroscopyParseError",
    "MatrixDatError",
]



class SpectroscopyParseError(Exception):
    """Raised when an Omicron/Anfatec spectroscopy file is malformed."""

    def __init__(self, path: Path | str, message: str):
        self.path = str(path)
        super().__init__(f"{self.path}: {message}")


class MatrixDatError(SpectroscopyParseError):
    """Raised when a Matrix .dat file violates the required structure."""

    def __init__(self, path: Path | str, message: str):
        super().__init__(path, message)
