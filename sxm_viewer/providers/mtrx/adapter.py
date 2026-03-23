"""Adapters for Scienta Omicron MATRIX ``.mtrx`` result files."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import hashlib
import math
import json
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np

from ...utils.logging import log
from ...data.channel_units import guess_channel_unit

try:
    from access2thematrix import MtrxData  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    MtrxData = None  # type: ignore


MTRX_CACHE_DIRNAME = ".sxmviewer_mtrx"
MTRX_CACHE_VERSION = 1
_IMPORT_ERROR = None


@dataclass
class ChannelExport:
    file_name: str
    caption: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


def is_mtrx_file(path: Path | str) -> bool:
    """Return True for the MATRIX result naming pattern used in the lab."""
    try:
        return Path(path).name.lower().endswith("_mtrx")
    except Exception:
        return False


def classify_mtrx_file(path: Path | str) -> str:
    """Best-effort split of MATRIX files into image-like or spectroscopy-like payloads."""
    path = Path(path)
    mdata = _try_open_mtrx(path)
    if mdata is not None:
        obj_type = str(getattr(mdata, "object_type", "") or "").lower()
        axis = getattr(mdata, "axis", None)
        if obj_type == "curve" and axis is None:
            return "spectroscopy"
        if "curve" in obj_type:
            return "spectroscopy"
        return "image"

    name = path.name.lower()
    if any(tok in name for tok in ("(v)", "(z)", "spectro", "sps", "curve")):
        return "spectroscopy"
    return "image"


def prepare_mtrx_folder(folder: Path | str) -> List[Path]:
    """Convert image-like MATRIX files within ``folder`` into viewer headers."""
    folder = Path(folder)
    if not folder.exists():
        return []
    cache_root = folder / MTRX_CACHE_DIRNAME
    cache_root.mkdir(exist_ok=True)
    generated: List[Path] = []
    for mtrx_path in sorted((p for p in folder.iterdir() if p.is_file() and is_mtrx_file(p)), key=lambda p: str(p).lower()):
        if classify_mtrx_file(mtrx_path) != "image":
            continue
        header_path = _convert_mtrx_image_file(mtrx_path, cache_root)
        if header_path is not None:
            generated.append(header_path)
    return generated


def prepare_mtrx_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit MATRIX image files into viewer headers."""
    generated: List[Path] = []
    seen = set()
    for raw_path in paths or []:
        mtrx_path = Path(raw_path)
        if not mtrx_path.is_file() or not is_mtrx_file(mtrx_path):
            continue
        try:
            key = str(mtrx_path.resolve()).lower()
        except Exception:
            key = str(mtrx_path).lower()
        if key in seen:
            continue
        seen.add(key)
        if classify_mtrx_file(mtrx_path) != "image":
            continue
        cache_root = mtrx_path.parent / MTRX_CACHE_DIRNAME
        cache_root.mkdir(exist_ok=True)
        header_path = _convert_mtrx_image_file(mtrx_path, cache_root)
        if header_path is not None:
            generated.append(header_path)
    return generated


def parse_mtrx_spectroscopy(path: Path | str) -> List[Dict[str, object]]:
    """Return spectroscopy entries extracted from a MATRIX curve file."""
    path = Path(path)
    if not is_mtrx_file(path):
        return []
    mdata = _try_open_mtrx(path) or _read_chainless_mtrx(path)
    if mdata is None:
        return []
    obj_type = str(getattr(mdata, "object_type", "") or "").lower()
    if "curve" not in obj_type:
        # Best-effort fallback: some standalone MATRIX exports only carry raw data.
        if "curve" not in _mtrx_file_kind(path):
            return []

    entries: List[Dict[str, object]] = []
    traces = list(getattr(mdata, "traces", []) or [])
    if traces and "curve" in obj_type:
        for trace_idx, trace in enumerate(traces):
            try:
                cu, _msg = mdata.select_curve(trace)
            except Exception:
                continue
            data = np.asarray(getattr(cu, "data", np.array([])), dtype=float)
            if data.ndim != 2 or data.shape[0] < 2 or data.shape[1] < 2:
                continue
            x_vals = np.asarray(data[0], dtype=float).ravel()
            y_vals = np.asarray(data[1], dtype=float).ravel()
            mask = np.isfinite(x_vals) & np.isfinite(y_vals)
            if mask.sum() < 2:
                continue
            x_vals = x_vals[mask]
            y_vals = y_vals[mask]
            x_name, x_unit = _pair_or_default(getattr(cu, "x_data_name_and_unit", None), default_name="Axis", default_unit="")
            y_name, y_unit = _pair_or_default(getattr(cu, "y_data_name_and_unit", None), default_name="Signal", default_unit="")
            ref_by = dict(getattr(cu, "referenced_by", {}) or {})
            entry: Dict[str, object] = {
                "path": str(path),
                "source": "omicron_mtrx",
                "trace": str(trace),
                "time": _extract_mtrx_time(mdata, path),
                "file_mtime": _mtime(path),
                "V": x_vals.copy(),
                "channels": {_clean_channel_label(y_name) or "channel": y_vals.copy()},
                "AxisLabel": x_name or "Axis",
                "AxisUnit": x_unit or "",
                "AxisChoices": [{"key": "bias", "label": x_name or "Axis", "unit": x_unit or "", "values": x_vals.copy()}],
                "AltAxis": None,
                "AltAxisLabel": None,
                "AltAxisUnit": None,
                "channel_name": y_name or str(trace),
                "channel_code": str(trace),
                "unit": y_unit or guess_channel_unit(y_name or str(trace)) or "",
                "order_idx": trace_idx + 1,
            }
            if ref_by:
                entry["referenced_by"] = ref_by
                loc_nm = ref_by.get("Location (m)")
                if isinstance(loc_nm, (list, tuple)) and len(loc_nm) >= 2:
                    entry["x"] = _meters_to_nm(loc_nm[0])
                    entry["y"] = _meters_to_nm(loc_nm[1])
                elif isinstance(ref_by.get("Location (px)"), (list, tuple)) and len(ref_by.get("Location (px)") or []) >= 2:
                    entry["x"] = _safe_float((ref_by.get("Location (px)") or [None, None])[0], default=None)
                    entry["y"] = _safe_float((ref_by.get("Location (px)") or [None, None])[1], default=None)
                ref_name = ref_by.get("Data File Name")
                if ref_name:
                    resolved = _resolve_related_path(path, str(ref_name))
                    if resolved is not None:
                        entry["image_key"] = str(resolved)
                        entry["image_path"] = str(resolved)
            entries.append(entry)
        entries.sort(key=lambda s: s.get("time") or datetime.min)
        return entries

    # Standalone fallback: create a single trace from the raw data block.
    data = np.asarray(getattr(mdata, "data", np.array([])), dtype=float).ravel()
    if data.size < 2:
        return []
    x_vals = np.arange(data.size, dtype=float)
    y_vals = np.asarray(data, dtype=float)
    y_name = _mtrx_channel_name(path)
    x_name = "Index"
    x_unit = "index"
    y_unit = guess_channel_unit(y_name) or ""
    entry = {
        "path": str(path),
        "source": "omicron_mtrx",
        "trace": "trace",
        "time": _extract_mtrx_time(mdata, path),
        "file_mtime": _mtime(path),
        "V": x_vals.copy(),
        "channels": {_clean_channel_label(y_name) or "channel": y_vals.copy()},
        "AxisLabel": x_name,
        "AxisUnit": x_unit,
        "AxisChoices": [{"key": "index", "label": x_name, "unit": x_unit, "values": x_vals.copy()}],
        "AltAxis": None,
        "AltAxisLabel": None,
        "AltAxisUnit": None,
        "channel_name": y_name or path.stem,
        "channel_code": y_name or path.stem,
        "unit": y_unit or "",
        "order_idx": 1,
        "fallback": True,
    }
    entries.append(entry)
    entries.sort(key=lambda s: s.get("time") or datetime.min)
    return entries


def _convert_mtrx_image_file(mtrx_path: Path, cache_root: Path) -> Optional[Path]:
    mdata = _try_open_mtrx(mtrx_path) or _read_chainless_mtrx(mtrx_path)
    if mdata is None:
        return None
    if classify_mtrx_file(mtrx_path) != "image":
        return None
    traces = list(getattr(mdata, "traces", []) or [])
    fallback_mode = not traces
    if fallback_mode:
        arr = np.asarray(getattr(mdata, "data", np.array([])), dtype=float).ravel()
        shape = _infer_mtrx_image_shape(arr.size)
        if shape is None:
            return None
        arr = _reshape_mtrx_image_data(arr, shape)
        traces = [mtrx_path.stem]

    cache_dir = _cache_dir_for(mtrx_path, cache_root)
    header_path = cache_dir / f"{mtrx_path.stem}_mtrx.txt"
    meta_path = cache_dir / "meta.json"
    src_stat = mtrx_path.stat()
    if header_path.exists() and meta_path.exists() and not _needs_rebuild(meta_path, src_stat.st_mtime, src_stat.st_size):
        return header_path

    if cache_dir.exists():
        for child in cache_dir.iterdir():
            try:
                if child.is_file():
                    child.unlink()
            except Exception:
                pass
    cache_dir.mkdir(parents=True, exist_ok=True)

    channels: List[ChannelExport] = []
    image_count = 0
    for idx, trace in enumerate(traces):
        if fallback_mode:
            im = None
            channel_name = _mtrx_channel_name(mtrx_path)
            channel_unit = "px"
            data_arr = arr
        else:
            try:
                im, _msg = mdata.select_image(trace)
            except Exception:
                continue
            data_arr = np.asarray(getattr(im, "data", np.array([])), dtype=float)
            if data_arr.ndim != 2 or data_arr.size == 0:
                continue
            channel_name, channel_unit = _pair_or_default(getattr(im, "channel_name_and_unit", None), default_name=str(trace), default_unit="")
        data_name = f"{_safe_token(mtrx_path.stem)}_{_safe_token(str(trace))}_{idx:02d}.dat"
        data_path = cache_dir / data_name
        np.savetxt(data_path, data_arr, fmt="%.9e")
        channels.append(
            ChannelExport(
                file_name=data_name,
                caption=f"{channel_name or Path(mtrx_path).stem} ({trace})",
                phys_unit=channel_unit or "",
                scale=1.0,
                offset=0.0,
            )
        )
        image_count += 1

    if not channels:
        return None

    dt = _extract_mtrx_datetime(mdata, mtrx_path)
    xpix = int(data_arr.shape[1]) if data_arr.ndim == 2 else int(shape[1])
    ypix = int(data_arr.shape[0]) if data_arr.ndim == 2 else int(shape[0])
    header: Dict[str, object] = {
        "xPixel": int(xpix) if image_count else 0,
        "yPixel": int(ypix) if image_count else 0,
        "XScanRange": float(xpix) if fallback_mode else float(getattr(im, "width", 0.0) or 0.0),
        "YScanRange": float(ypix) if fallback_mode else float(getattr(im, "height", 0.0) or 0.0),
        "XPhysUnit": "px" if fallback_mode else "nm",
        "YPhysUnit": "px" if fallback_mode else "nm",
        "xCenter": float(xpix) / 2.0 if fallback_mode else float(getattr(im, "x_offset", 0.0) or 0.0),
        "yCenter": float(ypix) / 2.0 if fallback_mode else float(getattr(im, "y_offset", 0.0) or 0.0),
        "Angle": 0.0 if fallback_mode else float(getattr(im, "angle", 0.0) or 0.0),
        "ScanAngle": 0.0 if fallback_mode else float(getattr(im, "angle", 0.0) or 0.0),
        "Date": dt.strftime("%Y-%m-%d") if dt else "",
        "Time": dt.strftime("%H:%M:%S") if dt else "",
        "Comment": getattr(mdata, "creation_comment", "") or "",
        "DataSetName": getattr(mdata, "data_set_name", "") or "",
        "SampleName": getattr(mdata, "sample_name", "") or "",
        "Session": getattr(mdata, "session", "") or "",
        "Cycle": getattr(mdata, "cycle", "") or "",
        "ChannelName": getattr(mdata, "channel_name", "") or "",
        "ObjectType": "image" if fallback_mode else (getattr(mdata, "object_type", "") or "image"),
        "ConvertedSource": str(mtrx_path),
        "ConvertedTimestamp": datetime.utcnow().isoformat(timespec="seconds"),
        "FallbackImport": bool(fallback_mode),
    }
    _write_sxm_style_header(header_path, header, channels, source=mtrx_path)
    meta = {
        "source": str(mtrx_path),
        "mtime": src_stat.st_mtime,
        "size": src_stat.st_size,
        "generated": datetime.utcnow().isoformat(timespec="seconds"),
        "header_name": header_path.name,
        "version": MTRX_CACHE_VERSION,
    }
    meta_path.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return header_path


def _write_sxm_style_header(
    header_path: Path,
    header: Dict[str, object],
    channels: Sequence[ChannelExport],
    *,
    source: Path,
):
    lines = [
        f"# Converted from {source.name} via MATRIX adapter",
        f"ConvertedSource = {source}",
        f"ConvertedTimestamp = {datetime.utcnow().isoformat(timespec='seconds')}",
    ]
    for key, value in header.items():
        formatted = _format_meta_value(value)
        if formatted is None:
            continue
        if isinstance(formatted, str) and formatted == "":
            continue
        lines.append(f"{key} = {formatted}")
    for ch in channels:
        lines.append("FileDescBegin")
        lines.append(f"FileName = {ch.file_name}")
        if ch.caption:
            lines.append(f"Caption = {ch.caption}")
        if ch.phys_unit:
            lines.append(f"PhysUnit = {ch.phys_unit}")
        lines.append(f"Scale = {ch.scale}")
        lines.append(f"Offset = {ch.offset}")
        lines.append("FileDescEnd")
    header_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _try_open_mtrx(path: Path) -> Optional[object]:
    if MtrxData is None:
        return None
    try:
        reader = MtrxData()
        traces, _msg = reader.open(str(path))
        if not traces and not getattr(reader, "traces", None):
            return None
        return reader
    except Exception:
        return None


def _read_chainless_mtrx(path: Path) -> Optional[object]:
    """Read the raw payload from a standalone MATRIX file.
    This is a fallback when the companion result-chain files are absent.
    """
    if MtrxData is None:
        return None
    try:
        raw = path.read_bytes()
    except Exception:
        return None
    if not raw.startswith(b"ONTMATRX0101"):
        return None
    try:
        reader = MtrxData()
        reader.raw_data = raw
        reader.raw_param = b""
        reader.param = {"BREF": ""}
        reader.channel_id = {}
        reader.bricklet_size = 0
        reader.data_item_count = 0
        reader.data = np.array([])
        reader.volume_scan = None
        reader.scan = None
        reader.axis = None
        reader._scan_raw_data(12, raw)
        if reader.data_item_count <= 0 or reader.data.size == 0:
            return None
        return reader
    except Exception:
        return None


def _extract_mtrx_datetime(mdata, path: Path) -> Optional[datetime]:
    for key in ("BTLK", "TLKB"):
        value = getattr(mdata, "param", {}).get(key)
        if not value:
            continue
        try:
            return datetime.strptime(str(value).strip(), "%A, %d %B %Y %H:%M:%S")
        except Exception:
            continue
    return _mtime(path)


def _extract_mtrx_time(mdata, path: Path) -> datetime:
    dt = _extract_mtrx_datetime(mdata, path)
    return dt or datetime.fromtimestamp(path.stat().st_mtime)


def _resolve_related_path(base_path: Path, ref_name: str) -> Optional[Path]:
    ref_name = str(ref_name or "").strip()
    if not ref_name:
        return None
    ref_path = Path(ref_name)
    candidates = []
    if ref_path.is_absolute() and ref_path.exists():
        return ref_path
    candidates.append(base_path.parent / ref_path.name)
    candidates.append(base_path.parent / ref_name)
    stem = ref_path.stem or ref_path.name
    if stem:
        for p in base_path.parent.iterdir():
            try:
                if p.is_file() and stem.lower() in p.name.lower():
                    candidates.append(p)
            except Exception:
                continue
    for candidate in candidates:
        try:
            if candidate.exists():
                return candidate
        except Exception:
            continue
    return None


def _pair_or_default(pair, default_name: str, default_unit: str) -> Tuple[str, str]:
    try:
        if isinstance(pair, (list, tuple)) and pair:
            name = str(pair[0] or default_name)
            unit = str(pair[1] or default_unit) if len(pair) > 1 else default_unit
            return name, unit
    except Exception:
        pass
    return default_name, default_unit


def _cache_dir_for(src: Path, cache_root: Path) -> Path:
    try:
        resolved = str(src.resolve())
    except Exception:
        resolved = str(src)
    digest = hashlib.sha1(resolved.encode("utf-8")).hexdigest()[:10]
    return cache_root / f"{src.stem}_{digest}"


def _needs_rebuild(meta_path: Path, mtime: float, size: int) -> bool:
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except Exception:
        return True
    if int(meta.get("version", -1)) != int(MTRX_CACHE_VERSION):
        return True
    if abs(meta.get("mtime", 0.0) - mtime) > 1e-6:
        return True
    if int(meta.get("size", -1)) != int(size):
        return True
    header_name = meta.get("header_name")
    if not header_name:
        return True
    header = meta_path.parent / header_name
    if not header.exists():
        return True
    return False


def _format_meta_value(value):
    if value is None:
        return None
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (list, tuple)):
        return ",".join(str(v) for v in value)
    if isinstance(value, np.ndarray):
        return ",".join(str(v) for v in value.tolist())
    return value


def _safe_token(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", str(text).strip())
    cleaned = cleaned.strip("_")
    return cleaned or "channel"


def _mtrx_file_kind(path: Path | str) -> str:
    """Heuristic file kind from a MATRIX filename when the chain is unavailable."""
    name = Path(path).name.lower()
    if any(tok in name for tok in ("(v)", "(z)", "curve", "spectro", "sps")):
        return "curve"
    return "image"


def _mtrx_channel_name(path: Path | str) -> str:
    """Return the filename token used as a display label for a standalone MATRIX file."""
    stem = Path(path).stem
    if "." in stem:
        stem = stem.split(".")[-1]
    return stem or Path(path).stem


def _infer_mtrx_image_shape(count: int) -> Optional[Tuple[int, int]]:
    if count <= 1:
        return None
    best = None
    root = int(math.sqrt(count))
    for rows in range(1, root + 1):
        if count % rows != 0:
            continue
        cols = count // rows
        cand = (rows, cols) if rows <= cols else (cols, rows)
        score = (abs(cand[1] - cand[0]), abs(cand[1] - 512), -cand[1])
        if best is None or score < best[0]:
            best = (score, cand)
    if best is not None:
        return best[1]
    rows = max(1, int(math.sqrt(count)))
    cols = int(math.ceil(count / rows))
    return (rows, cols)


def _reshape_mtrx_image_data(data: np.ndarray, shape: Tuple[int, int]) -> np.ndarray:
    rows, cols = shape
    need = rows * cols
    flat = np.asarray(data, dtype=float).ravel()
    if flat.size < need:
        flat = np.pad(flat, (0, need - flat.size), mode="edge")
    elif flat.size > need:
        flat = flat[:need]
    return flat.reshape((rows, cols))


def _safe_float(value, default: Optional[float] = 0.0) -> Optional[float]:
    try:
        if value is None:
            return default
        if isinstance(value, (float, int)):
            return float(value)
        txt = str(value).strip()
        if not txt:
            return default
        return float(txt)
    except Exception:
        return default


def _meters_to_nm(value) -> Optional[float]:
    try:
        if value is None:
            return None
        return float(value) * 1e9
    except Exception:
        try:
            return float(str(value).strip()) * 1e9
        except Exception:
            return None


def _clean_channel_label(label: str) -> str:
    label = str(label or "").strip()
    label = label.replace("/", "_").replace("(", "").replace(")", "")
    label = re.sub(r"[^a-zA-Z0-9_+-]", "_", label)
    label = re.sub(r"_{2,}", "_", label)
    return label.strip("_")


def _mtime(path: Path) -> Optional[datetime]:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return None
