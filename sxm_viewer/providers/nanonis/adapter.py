"""Adapters that convert Nanonis files into Omicron-style descriptors.

This module is isolated under the providers namespace to decouple parsing from
the GUI and the native (Omicron/Anfatec) pipeline.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import re
import shutil
import sys
import hashlib
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np

# NumPy 2.0 removed legacy scalar aliases; keep shims for vendored/third-party code.
if not hasattr(np, "float"):
    np.float = float  # type: ignore[attr-defined]
if not hasattr(np, "int"):
    np.int = int  # type: ignore[attr-defined]

from ...utils.logging import log
from ...data.channel_units import guess_channel_unit

try:
    from importlib import import_module
except ImportError:  # pragma: no cover - python <3.5 not supported, safeguard only
    import_module = None  # type: ignore


NANONIS_CACHE_DIRNAME = ".sxmviewer_nanonis"
NANONIS_CACHE_VERSION = 2
_NANONIS_READ = None
_IMPORT_ERROR = None


@dataclass
class ChannelExport:
    file_name: str
    caption: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


def prepare_nanonis_folder(folder: Path | str) -> List[Path]:
    """Convert Nanonis scans within ``folder`` and return generated header paths."""
    folder = Path(folder)
    reader = _ensure_nanonis_reader()
    if reader is None:
        # We already logged why the adapter is unavailable.
        return []
    scan_files = sorted({p for p in folder.glob("*.sxm") if p.is_file()})
    if not scan_files:
        return []
    cache_root = folder / NANONIS_CACHE_DIRNAME
    cache_root.mkdir(exist_ok=True)
    generated: List[Path] = []
    for scan_path in scan_files:
        try:
            header_path = _convert_scan_file(reader, scan_path, cache_root)
        except Exception as exc:
            log(f"[Nanonis] Failed to convert {scan_path.name}: {exc}")
            continue
        if header_path is not None:
            generated.append(header_path)
    return generated


def prepare_nanonis_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit Nanonis scan files and return generated header paths."""
    reader = _ensure_nanonis_reader()
    if reader is None:
        return []
    generated: List[Path] = []
    seen = set()
    for raw_path in paths or []:
        scan_path = Path(raw_path)
        if not scan_path.is_file():
            continue
        try:
            key = str(scan_path.resolve()).lower()
        except Exception:
            key = str(scan_path).lower()
        if key in seen:
            continue
        seen.add(key)
        cache_root = scan_path.parent / NANONIS_CACHE_DIRNAME
        cache_root.mkdir(exist_ok=True)
        try:
            header_path = _convert_scan_file(reader, scan_path, cache_root)
        except Exception as exc:
            log(f"[Nanonis] Failed to convert {scan_path.name}: {exc}")
            continue
        if header_path is not None:
            generated.append(header_path)
    return generated


# --------------------------------------------------------------------------- #
# Conversion helpers                                                         #
# --------------------------------------------------------------------------- #

def _convert_scan_file(reader, scan_path: Path, cache_root: Path) -> Optional[Path]:
    src_stat = scan_path.stat()
    cache_dir = _cache_dir_for(scan_path, cache_root)
    header_path = cache_dir / f"{scan_path.stem}_nanonis.txt"
    meta_path = cache_dir / "meta.json"
    if (
        header_path.exists()
        and meta_path.exists()
        and not _needs_rebuild(meta_path, src_stat.st_mtime, src_stat.st_size)
    ):
        return header_path

    if cache_dir.exists():
        shutil.rmtree(cache_dir, ignore_errors=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    scan = reader.Scan(str(scan_path))
    header = _extract_scan_header(scan)
    channels = _extract_scan_channels(scan, cache_dir)
    if not channels:
        log(f"[Nanonis] No usable channels found in {scan_path.name}")
        return None

    _write_sxm_style_header(header_path, header, channels, source=scan_path)
    meta = {
        "source": str(scan_path),
        "mtime": src_stat.st_mtime,
        "size": src_stat.st_size,
        "generated": datetime.utcnow().isoformat(timespec="seconds"),
        "channels": len(channels),
        "header_name": header_path.name,
        "version": NANONIS_CACHE_VERSION,
    }
    meta_path.write_text(json.dumps(meta, indent=2))
    return header_path


def _extract_scan_header(scan) -> Dict[str, object]:
    hdr = scan.header or {}
    xpix, ypix = _coerce_pixel_tuple(hdr.get("scan_pixels"))
    rng_x, rng_y = _meters_to_nm_pair(hdr.get("scan_range"))
    off_x, off_y = _meters_to_nm_pair(hdr.get("scan_offset"))
    angle = _safe_float(hdr.get("scan_angle"), default=0.0)
    bias = _safe_float(hdr.get("bias"), default=0.0)
    rec_date = _format_date_string(str(hdr.get("rec_date", "")).strip())
    rec_time = _format_time_string(str(hdr.get("rec_time", "")).strip())
    scan_dir = str(hdr.get("scan_dir", "")).strip()
    acq_time = hdr.get("acq_time")
    header = {
        "xPixel": xpix,
        "yPixel": ypix,
        "XScanRange": rng_x,
        "YScanRange": rng_y,
        "XPhysUnit": "nm",
        "YPhysUnit": "nm",
        "xCenter": off_x,
        "yCenter": off_y,
        "ScanAngle": angle,
        "Angle": angle,
        "ScanDir": scan_dir,
        "Bias": bias,
        "BiasPhysUnit": "V",
        "Date": rec_date,
        "Time": rec_time,
        "AcqTime[s]": _safe_float(acq_time) if acq_time is not None else "",
    }
    header["SessionPath"] = hdr.get("nanonismain>session path", "")
    header["Comment"] = hdr.get("comment", "")
    header["UserName"] = (
        hdr.get("user")
        or hdr.get("nanonismain>session user")
        or hdr.get("nanonismain>user")
        or ""
    )
    scan_time = hdr.get("scan_time")
    if isinstance(scan_time, (list, tuple)):
        if len(scan_time) >= 1:
            header["ScanTimeForward[s]"] = _safe_float(scan_time[0])
        if len(scan_time) >= 2:
            header["ScanTimeBackward[s]"] = _safe_float(scan_time[1])
    zctrl = hdr.get("z-controller")
    setp_val, setp_unit = _extract_zctrl_setpoint(zctrl)
    if setp_val is not None:
        header["SetPoint"] = setp_val
        if setp_unit:
            header["SetPointPhysUnit"] = setp_unit
    header["SampleTemp[K]"] = _safe_float(hdr.get("rec_temp"))
    header["ScanFile"] = hdr.get("scan_file")
    header["ScanType"] = hdr.get("scanit_type")
    header["BiasPolarity"] = hdr.get("bias")
    _flatten_nanonis_fields(header, hdr, prefix="Nanonis:")
    return header


def _extract_scan_channels(scan, cache_dir: Path) -> List[ChannelExport]:
    header_info = scan.header.get("data_info", {}) if scan.header else {}
    names = list(header_info.get("Name", []))
    units = list(header_info.get("Unit", []))
    directions = list(header_info.get("Direction", []))
    calibrations = list(header_info.get("Calibration", []))
    offsets = list(header_info.get("Offset", []))
    total = min(len(names), len(units), len(directions), len(calibrations), len(offsets))
    exports: List[ChannelExport] = []
    # Nanonis `.sxm` data is typically stored as float32 values that already
    # include calibration/offset. Integer formats require manual scaling.
    data_dtype = np.dtype(getattr(scan, "data_format", np.float32))
    needs_calibration = data_dtype.kind in ("i", "u")
    for idx in range(total):
        name = str(names[idx]).strip()
        unit = str(units[idx]).strip()
        direction = str(directions[idx]).strip().lower()
        scale = _safe_float(calibrations[idx], default=1.0)
        offset = _safe_float(offsets[idx], default=0.0)
        signal = scan.signals.get(name)
        if not signal:
            continue
        dir_keys = _direction_keys(direction, signal)
        for dir_key in dir_keys:
            arr = signal.get(dir_key)
            if arr is None:
                continue
            arr = np.asarray(arr, dtype=float)
            if np.isnan(arr).all():
                continue
            if needs_calibration:
                arr = arr * scale + offset
            safe_channel = _safe_token(name)
            suffix = "fwd" if dir_key == "forward" else "bwd"
            data_name = f"{scan.basename}_{safe_channel}_{suffix}.dat"
            data_path = cache_dir / data_name
            try:
                with open(data_path, "w", encoding="utf-8", newline="\n") as fh:
                    np.savetxt(fh, arr, fmt="%.9e")
            except UnicodeEncodeError:
                np.savetxt(data_path, arr, fmt="%.9e")
            caption_dir = "Forward" if dir_key == "forward" else "Backward"
            caption = _pretty_caption(name, caption_dir)
            exports.append(
                ChannelExport(
                    file_name=data_name,
                    caption=caption,
                    phys_unit=unit,
                    scale=1.0,
                    offset=0.0,
                )
            )
    return exports


def _write_sxm_style_header(
    header_path: Path,
    header: Dict[str, object],
    channels: Sequence[ChannelExport],
    *,
    source: Path,
):
    lines = [
        f"# Converted from {source.name} via Nanonis adapter",
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


# --------------------------------------------------------------------------- #
# Utility helpers                                                             #
# --------------------------------------------------------------------------- #

def _ensure_nanonis_reader():
    """Return the ``nanonispy.read`` module or ``None`` if unavailable."""
    global _NANONIS_READ, _IMPORT_ERROR
    if _NANONIS_READ is not None or _IMPORT_ERROR:
        return _NANONIS_READ
    module_names = ("nanonispy2.read", "nanonispy.read")
    for mod_name in module_names:
        try:
            _NANONIS_READ = import_module(mod_name) if import_module else None
            if _NANONIS_READ:
                return _NANONIS_READ
        except Exception:
            continue
    # Try adding the vendored copy that ships with the repository.
    vendor_path = Path(__file__).resolve().parent / "vendor" / "nanonispy2-1.2.0" / "nanonispy2-1.2.0"
    if vendor_path.exists():
        sys.path.append(str(vendor_path))
        try:
            _NANONIS_READ = import_module("nanonispy2.read") if import_module else None
            if _NANONIS_READ:
                return _NANONIS_READ
        except Exception as exc:
            _IMPORT_ERROR = exc
    else:
        _IMPORT_ERROR = RuntimeError("nanonispy package not found.")
    if _IMPORT_ERROR:
        log(f"[Nanonis] Adapter unavailable: {_IMPORT_ERROR}")
    return _NANONIS_READ


def _cache_dir_for(src: Path, cache_root: Path) -> Path:
    try:
        resolved = str(src.resolve())
    except Exception:
        resolved = str(src)
    digest = hashlib.sha1(resolved.encode("utf-8")).hexdigest()[:10]
    return cache_root / f"{src.stem}_{digest}"


def _needs_rebuild(meta_path: Path, mtime: float, size: int) -> bool:
    try:
        meta = json.loads(meta_path.read_text())
    except Exception:
        return True
    if int(meta.get("version", -1)) != int(NANONIS_CACHE_VERSION):
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


def _meters_to_nm_pair(values: Optional[Iterable[float]]) -> Tuple[float, float]:
    if values is None:
        return 0.0, 0.0
    vals = list(values)
    first = _safe_float(vals[0], default=0.0) if vals else 0.0
    second = _safe_float(vals[1], default=0.0) if len(vals) > 1 else 0.0
    return first * 1e9, second * 1e9


def _meters_to_nm_value(value) -> Optional[float]:
    try:
        if value is None:
            return None
        return float(value) * 1e9
    except Exception:
        parsed, unit = _split_value_and_unit(value)
        if parsed is None:
            try:
                return float(str(value).strip()) * 1e9
            except Exception:
                return None
        unit_key = str(unit or "").strip().lower().replace("µ", "u")
        scale_map = {
            "": 1e9,
            "m": 1e9,
            "meter": 1e9,
            "meters": 1e9,
            "nm": 1.0,
            "nanometer": 1.0,
            "nanometers": 1.0,
            "pm": 1e-3,
            "picometer": 1e-3,
            "picometers": 1e-3,
            "um": 1e3,
            "micrometer": 1e3,
            "micrometers": 1e3,
            "mm": 1e6,
            "a": 0.1,
            "angstrom": 0.1,
            "angstroms": 0.1,
            "å": 0.1,
        }
        factor = scale_map.get(unit_key)
        if factor is None:
            return None
        return float(parsed) * factor


def _coerce_pixel_tuple(values: Optional[Iterable[int]]) -> Tuple[int, int]:
    if values is None:
        return 0, 0
    vals = list(values)
    xpix = int(vals[0]) if vals else 0
    ypix = int(vals[1]) if len(vals) > 1 else xpix
    return xpix, ypix


def _format_date_string(text: str) -> str:
    if not text:
        return ""
    candidates = ("%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y")
    for fmt in candidates:
        try:
            return datetime.strptime(text.strip(), fmt).strftime("%Y-%m-%d")
        except Exception:
            continue
    return text.strip()


def _format_time_string(text: str) -> str:
    if not text:
        return ""
    candidates = ("%H:%M:%S", "%H.%M.%S")
    for fmt in candidates:
        try:
            return datetime.strptime(text.strip(), fmt).strftime("%H:%M:%S")
        except Exception:
            continue
    return text.strip()


def _safe_float(value, default: float = 0.0) -> float:
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


def _safe_token(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", text.strip())
    cleaned = cleaned.strip("_")
    return cleaned or "channel"


def _pretty_caption(name: str, direction: str) -> str:
    base = name.replace("_", " ").strip()
    title = base.title() if base else "Channel"
    return f"{title} ({direction})"


def _direction_keys(direction: str, signal: Dict[str, np.ndarray]) -> List[str]:
    available = []
    for candidate in ("forward", "backward"):
        if candidate in signal:
            available.append(candidate)
    if direction == "both":
        return available or list(signal.keys())
    if direction.startswith("forw"):
        return ["forward"] if "forward" in signal else available[:1]
    if direction.startswith("back"):
        return ["backward"] if "backward" in signal else available[-1:]
    if available:
        return available
    return list(signal.keys())


def _split_value_and_unit(text: str) -> Tuple[Optional[float], str]:
    if text is None:
        return None, ""
    s = str(text).strip()
    if not s:
        return None, ""
    m = re.match(r"^([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(.*)$", s)
    if m:
        try:
            value = float(m.group(1))
        except Exception:
            value = None
        unit = m.group(2).strip()
        return value, unit
    try:
        return float(s), ""
    except Exception:
        return None, ""


def _extract_zctrl_setpoint(zctrl) -> Tuple[Optional[float], str]:
    if not isinstance(zctrl, dict):
        return None, ""
    entries = zctrl.get("Setpoint") or zctrl.get("setpoint")
    if isinstance(entries, (list, tuple)) and entries:
        return _split_value_and_unit(entries[0])
    if isinstance(entries, str):
        return _split_value_and_unit(entries)
    return None, ""


def _extract_zctrl_absolute_z_nm(zctrl) -> Tuple[Optional[float], str]:
    if not isinstance(zctrl, dict):
        return None, ""
    for key in ("Z (m)", "Z", "z", "Z abs (m)", "Z abs"):
        value_nm = _meters_to_nm_value(zctrl.get(key))
        if value_nm is not None:
            return value_nm, "Z piezo absolute"
    return None, ""


def _extract_nanonis_z_level_nm(header: Dict[str, str]) -> Tuple[Optional[float], str]:
    for key in ("Z-Controller", "z-controller", "Z Controller", "z_controller", "Z_Controller"):
        value_nm, label = _extract_zctrl_absolute_z_nm(header.get(key))
        if value_nm is not None:
            return value_nm, label
    for key in (
        "Z-Controller>Z (m)",
        "Z-Controller>Z",
        "z-controller>Z (m)",
        "z-controller>Z",
        "Z Controller>Z (m)",
        "Z Controller>Z",
    ):
        value_nm = _meters_to_nm_value(header.get(key))
        if value_nm is not None:
            return value_nm, "Z piezo absolute"
    candidates = [
        ("Z piezo absolute (m)", "Z piezo absolute"),
        ("Z piezo absolute", "Z piezo absolute"),
        ("Z piezo abs (m)", "Z piezo absolute"),
        ("Z piezo abs", "Z piezo absolute"),
        ("Z piezo (m)", "Z piezo"),
        ("Z piezo", "Z piezo"),
        ("Absolute Z (m)", "Absolute Z"),
        ("Absolute Z", "Absolute Z"),
        ("Z absolute (m)", "Z absolute"),
        ("Z absolute", "Z absolute"),
        ("Z (m)", "Z"),
        ("Z", "Z"),
        ("Final Z (m)", "Final Z"),
        ("Final Z", "Final Z"),
        ("Z offset (m)", "Z offset"),
        ("Z offset", "Z offset"),
    ]
    for key, label in candidates:
        value_nm = _meters_to_nm_value(header.get(key))
        if value_nm is not None:
            return value_nm, label
    for key, value in (header or {}).items():
        key_low = str(key or "").strip().lower()
        if not key_low:
            continue
        if "z" not in key_low and "piezo" not in key_low:
            continue
        value_nm = _meters_to_nm_value(value)
        if value_nm is not None:
            return value_nm, re.sub(r"\s*\(.*?\)", "", str(key)).strip() or "Z"
    for key, value in (header or {}).items():
        key_txt = str(key or "").strip()
        if isinstance(value, dict):
            nested_nm, nested_label = _extract_nanonis_z_level_nm(value)
            if nested_nm is not None:
                return nested_nm, nested_label or (re.sub(r"\s*\(.*?\)", "", key_txt).strip() or "Z")
        elif isinstance(value, (list, tuple)):
            for item in value:
                if not isinstance(item, dict):
                    continue
                nested_nm, nested_label = _extract_nanonis_z_level_nm(item)
                if nested_nm is not None:
                    return nested_nm, nested_label or (re.sub(r"\s*\(.*?\)", "", key_txt).strip() or "Z")
    return None, ""


def _extract_nanonis_z_level_from_raw_header(path: Path) -> Tuple[Optional[float], str]:
    patterns = (
        (re.compile(r"^\s*Z-Controller>\s*Z\s*\(m\)\s*(?:\t+| {2,}|:\s*|=\s*)(\S+)\s*$", re.IGNORECASE), "Z piezo absolute"),
        (re.compile(r"^\s*Z\s*\(m\)\s*(?:\t+| {2,}|:\s*|=\s*)(\S+)\s*$", re.IGNORECASE), "Z"),
        (re.compile(r"^\s*Absolute\s+Z\s*\(m\)\s*(?:\t+| {2,}|:\s*|=\s*)(\S+)\s*$", re.IGNORECASE), "Absolute Z"),
    )
    try:
        with Path(path).open("r", encoding="utf-8", errors="ignore") as handle:
            for raw_line in handle:
                line = str(raw_line or "").strip()
                if not line:
                    continue
                if line.upper().startswith("[DATA]"):
                    break
                for pattern, label in patterns:
                    match = pattern.match(line)
                    if not match:
                        continue
                    value_nm = _meters_to_nm_value(match.group(1))
                    if value_nm is not None:
                        return value_nm, label
    except Exception:
        return None, ""
    return None, ""


def _signal_unit_to_nm(values, unit_hint: str) -> Optional[np.ndarray]:
    try:
        arr = np.asarray(values, dtype=float).ravel()
    except Exception:
        return None
    if arr.size == 0:
        return None
    unit_key = str(unit_hint or "").strip().lower().replace("µ", "u")
    if unit_key in ("m", "meter", "meters"):
        return arr * 1e9
    if unit_key in ("nm", "nanometer", "nanometers", ""):
        try:
            max_abs = float(np.nanmax(np.abs(arr[arr == arr]))) if arr.size else 0.0
        except Exception:
            max_abs = 0.0
        return arr * 1e9 if max_abs and max_abs < 1e-3 else arr
    if unit_key in ("pm", "picometer", "picometers"):
        return arr * 1e-3
    if unit_key in ("um", "micrometer", "micrometers"):
        return arr * 1e3
    if unit_key in ("a", "å", "angstrom", "angstroms"):
        return arr * 0.1
    return None


def _extract_constant_signal_z_level_nm(signals: Dict[str, np.ndarray]) -> Tuple[Optional[float], str]:
    if not isinstance(signals, dict):
        return None, ""
    for name, values in signals.items():
        low = str(name or "").strip().lower()
        if not low:
            continue
        if not any(token in low for token in ("topo", "topography", "z piezo", "absolute z", "z absolute", "z_abs", "z-abs")):
            continue
        unit_hint = ""
        if "(m)" in low:
            unit_hint = "m"
        elif "(nm)" in low:
            unit_hint = "nm"
        elif "(pm)" in low:
            unit_hint = "pm"
        elif "(um)" in low:
            unit_hint = "um"
        arr_nm = _signal_unit_to_nm(values, unit_hint)
        if arr_nm is None:
            continue
        finite = arr_nm[np.isfinite(arr_nm)]
        if finite.size == 0:
            continue
        try:
            span = float(np.nanmax(finite) - np.nanmin(finite))
            center = float(np.nanmedian(finite))
        except Exception:
            continue
        if span <= max(1e-3, abs(center) * 1e-6):
            return center, re.sub(r"\s*\(.*?\)", "", str(name)).strip() or "Topo"
    return None, ""


def _try_parse_datetime(text: str) -> Optional[datetime]:
    if not text:
        return None
    cleaned = str(text).strip()
    if not cleaned:
        return None
    fmts = [
        "%d.%m.%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%d.%m.%Y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%H:%M:%S",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(cleaned, fmt)
        except Exception:
            continue
    return None


def _nanonis_spec_metadata(header: Dict[str, str], path: Path) -> Dict[str, object]:
    meta: Dict[str, object] = {}
    date_txt = (
        header.get("Start date")
        or header.get("Start Date")
        or header.get("Date")
        or ""
    )
    time_txt = header.get("Start time") or header.get("Start Time") or ""
    dt = _try_parse_datetime(f"{date_txt} {time_txt}".strip())
    if dt is None:
        dt = _try_parse_datetime(date_txt) or _try_parse_datetime(time_txt)
    if dt is not None:
        meta["time"] = dt
    x_nm = _meters_to_nm_value(header.get("X (m)"))
    y_nm = _meters_to_nm_value(header.get("Y (m)"))
    if x_nm is not None:
        meta["x"] = x_nm
    if y_nm is not None:
        meta["y"] = y_nm
    z_nm, z_label = _extract_nanonis_z_level_nm(header)
    if z_nm is None:
        z_nm, z_label = _extract_nanonis_z_level_from_raw_header(path)
    if z_nm is not None:
        meta["z_level_nm"] = z_nm
        meta["z_level_label"] = z_label or "Z"
        meta["z_level_unit"] = "nm"
    # Ensure positions exist so thumbnails can render markers even when metadata is partial.
    if "x" not in meta:
        meta["x"] = 0.0
    if "y" not in meta:
        meta["y"] = 0.0
    if "time" not in meta:
        try:
            meta["time"] = datetime.fromtimestamp(Path(path).stat().st_mtime)
        except Exception:
            pass
    return meta


def _sanitize_channel_label(label: str) -> str:
    lbl = str(label or "").strip()
    lbl = lbl.replace("/", "_").replace("(", "").replace(")", "")
    lbl = re.sub(r"[^a-zA-Z0-9_+-]", "_", lbl)
    lbl = re.sub(r"_{2,}", "_", lbl)
    return lbl.strip("_")


def _select_z_axis(signals: Dict[str, np.ndarray]) -> Tuple[Optional[str], Optional[np.ndarray]]:
    """Best-effort selection of a Z axis for distance-based spectroscopies."""
    candidates = [
        "Z (m)",
        "Z",
        "Z rel (m)",
        "Z rel",
        "Delta Z (m)",
        "Z offset (m)",
        "Z offset",
        "Z piezo (m)",
        "Z piezo",
        "Distance (m)",
        "Distance",
    ]
    for name in candidates:
        if name in signals:
            return name, signals[name]
    for name, data in signals.items():
        low = name.lower()
        if low.startswith("z") or "z " in low or " z" in low or "distance" in low:
            return name, data
    return None, None


def _select_z_rel_axis(signals: Dict[str, np.ndarray]) -> Tuple[Optional[str], Optional[np.ndarray]]:
    """Select a relative Z axis if present (z_rel naming)."""
    for name, data in signals.items():
        low = name.lower()
        if "z_rel" in low or "rel z" in low:
            return name, data
    return None, None


def _select_bias_axis(signals: Dict[str, np.ndarray]) -> Tuple[Optional[str], Optional[np.ndarray]]:
    candidates = [
        "Bias calc (V)",
        "Sample bias (V)",
        "Bias (V)",
        "Tip bias (V)",
    ]
    for name in candidates:
        if name in signals:
            return name, signals[name]
    for name, data in signals.items():
        if "(V)" in name or name.lower().startswith("bias"):
            return name, data
    return None, None


def parse_nanonis_spectroscopy(path: Path | str) -> List[Dict[str, object]]:
    reader = _ensure_nanonis_reader()
    if reader is None:
        return []
    try:
        spec = reader.Spec(str(path))
    except Exception as exc:
        msg = str(exc)
        if "Could not find the [DATA] end tag" in msg:
            # Corrupt/incomplete file; skip quietly so Omicron parser can try.
            return []
        log(f"[Nanonis] Failed to parse spectroscopy {path}: {msg}")
        return []
    prefer_z = False
    try:
        name_l = str(path).lower()
        if "z-spectro" in name_l or "z_spectro" in name_l or "z spectro" in name_l or "z-spectroscopy" in name_l:
            prefer_z = True
    except Exception:
        pass
    axis_name = None
    axis_data = None
    if prefer_z:
        axis_name, axis_data = _select_z_axis(spec.signals)
    alt_axis_name = None
    alt_axis_data = None
    if prefer_z:
        alt_axis_name, alt_axis_data = _select_z_rel_axis(spec.signals)
    if axis_name is None or axis_data is None:
        axis_name, axis_data = _select_bias_axis(spec.signals)
    if axis_name is None or axis_data is None:
        return []
    axis = np.asarray(axis_data, dtype=float)
    axis_unit = "V"
    axis_label = axis_name or "Axis"
    if axis_name:
        low = axis_name.lower()
        axis_label = re.sub(r"\s*\(.*?\)", "", axis_name).strip() or axis_label
        if "(m)" in low or " distance" in low or "distance " in low:
            axis = axis * 1e9  # convert meters to nm for display consistency
            axis_unit = "nm"
            if "z" in axis_label.lower():
                axis_label = "Z"
    alt_axis_unit = None
    if alt_axis_name is not None and alt_axis_data is not None:
        alt_axis = np.asarray(alt_axis_data, dtype=float)
        alt_axis_unit = "nm"
        try:
            if np.nanmax(np.abs(alt_axis)) < 1e-6:
                alt_axis = alt_axis * 1e9
        except Exception:
            pass
    else:
        alt_axis = None
    channels: Dict[str, np.ndarray] = {}
    unit_map: Dict[str, str] = {}
    for name, values in spec.signals.items():
        if name == axis_name:
            continue
        arr = np.asarray(values, dtype=float)
        if arr.shape != axis.shape:
            continue
        clean = _sanitize_channel_label(name) or _safe_token(name)
        label = clean
        counter = 1
        while label in channels:
            label = f"{clean}_{counter}"
            counter += 1
        channels[label] = arr.copy()
        unit_guess = guess_channel_unit(name)
        if unit_guess:
            unit_map[label] = unit_guess
    if not channels:
        return []
    meta = _nanonis_spec_metadata(spec.header or {}, Path(path))
    if meta.get("z_level_nm") is None:
        z_nm, z_label = _extract_constant_signal_z_level_nm(spec.signals)
        if z_nm is not None:
            meta["z_level_nm"] = z_nm
            meta["z_level_label"] = z_label or "Topo"
            meta["z_level_unit"] = "nm"
    entry = {
        "path": str(path),
        "V": axis.copy(),
        "AxisLabel": axis_label,
        "AxisUnit": axis_unit,
        "AltAxis": alt_axis.copy() if alt_axis is not None else None,
        "AltAxisLabel": re.sub(r"\s*\(.*?\)", "", alt_axis_name).strip() if alt_axis_name else None,
        "AltAxisUnit": alt_axis_unit,
        "channels": channels,
        "unit_map": unit_map or None,
    }
    entry.update(meta)
    _flatten_nanonis_fields(entry, spec.header or {}, prefix="NanonisSpec:")
    return [entry]


def parse_nanonis_3ds(path: Path | str) -> List[Dict[str, object]]:
    """Parse Nanonis .3ds files (grid spectroscopy) into matrix-compatible entries."""
    reader = _ensure_nanonis_reader()
    if reader is None:
        log("[Nanonis] nanonispy2 not available; cannot parse .3ds")
        return []
    GridCls = None
    try:
        GridCls = getattr(reader, "Grid", None)
    except Exception:
        GridCls = None
    if GridCls is None:
        log("[Nanonis] Grid reader not found in nanonispy2; skipping .3ds")
        return []
    try:
        grid = GridCls(str(path))
    except Exception as exc:
        log(f"[Nanonis] Failed to read {path}: {exc}")
        return []
    try:
        return _parse_nanonis_3ds_grid(grid, path, chans=getattr(grid, "signals", {}) or {})
    except Exception as exc:
        log(f"[Nanonis] Unexpected failure parsing {path}: {exc}")
        return []


def _parse_nanonis_3ds_grid(grid, path: Path | str, chans: Dict[str, object]) -> List[Dict[str, object]]:
    entries: List[Dict[str, object]] = []
    if not chans:
        log(f"[Nanonis] No channels found in {path}")
        return entries
    def _parse_time(value):
        try:
            if isinstance(value, datetime):
                return value
            if isinstance(value, (int, float)):
                return datetime.fromtimestamp(float(value))
            if isinstance(value, str) and value.strip():
                # nanonispy2 returns datetime for start_time/end_time; still guard strings
                return datetime.fromisoformat(value)
        except Exception:
            return None
        return None

    def _first_non_null(*vals):
        for v in vals:
            if v is None:
                continue
            try:
                arr = np.asarray(v)
                # numpy arrays cannot be used in truth-testing; rely on size instead
                if arr.size == 0:
                    continue
            except Exception:
                pass
            return v
        return None

    # Diagnostic logs removed for normal runs (too noisy)
    bias_raw = chans.get("sweep_signal")
    bias = np.asarray(bias_raw, dtype=float) if bias_raw is not None else np.asarray([], dtype=float)
    if bias.size == 0:
        try:
            bias = np.asarray(grid._derive_sweep_signal(), dtype=float)
        except Exception:
            bias = np.asarray([], dtype=float)
    # grid dimensions
    dim_px = _first_non_null(grid.header.get("dim_px"), grid.header.get("Grid dim"))
    nx = ny = None
    if dim_px is not None:
        try:
            # dim_px is typically (nx, ny) but sometimes includes a params dimension
            if len(dim_px) >= 2:
                nx = int(dim_px[0])
                ny = int(dim_px[1])
        except Exception:
            nx = ny = None
    if nx is None or ny is None:
        # infer from first usable channel
        sample_arr = next(iter(chans.values()))
        shape = np.shape(sample_arr)
        if len(shape) >= 2:
            ny, nx = int(shape[0]), int(shape[1])
        else:
            nx = ny = 1
    nx = max(int(nx or 1), 1)
    ny = max(int(ny or 1), 1)
    # physical ranges if present (m -> nm)
    try:
        scan_range = _first_non_null(
            grid.header.get("size_xy"),
            grid.header.get("scan_range"),
            grid.header.get("ScanRange"),
        )
        center = _first_non_null(
            grid.header.get("pos_xy"),
            grid.header.get("center_xy"),
            (0.0, 0.0),
        )
        rx, ry = scan_range
        cx, cy = center
        rx_nm = float(rx) * 1e9 if abs(rx) < 1e-3 else float(rx)
        ry_nm = float(ry) * 1e9 if abs(ry) < 1e-3 else float(ry)
        cx_nm = float(cx) * 1e9 if abs(cx) < 1e-3 else float(cx)
        cy_nm = float(cy) * 1e9 if abs(cy) < 1e-3 else float(cy)
        x_offsets = np.linspace(cx_nm - rx_nm / 2, cx_nm + rx_nm / 2, nx)
        y_offsets = np.linspace(cy_nm - ry_nm / 2, cy_nm + ry_nm / 2, ny)
    except Exception:
        x_offsets = np.arange(nx, dtype=float)
        y_offsets = np.arange(ny, dtype=float)
    dataset_key = Path(path).stem
    # acquisition time from header if available
    spec_time = _parse_time(grid.header.get("start_time")) or _parse_time(grid.header.get("end_time"))
    if spec_time is None:
        try:
            spec_time = datetime.fromtimestamp(Path(path).stat().st_mtime)
        except Exception:
            spec_time = None
    channel_data: Dict[str, np.ndarray] = {}
    skip_keys = {"params", "sweep_signal", "topo"}
    for raw_key, raw_arr in chans.items():
        if str(raw_key) in skip_keys or raw_key in skip_keys:
            continue
        ch_key = _sanitize_channel_label(str(raw_key)) or str(raw_key)
        try:
            arr = np.asarray(raw_arr, dtype=float)
        except Exception as exc:
            log(f"[Nanonis] Failed to coerce channel {raw_key} in {path}: {exc}")
            continue
        if arr.ndim != 3 or arr.size == 0:
            log(f"[Nanonis] Channel {ch_key} has unsupported shape {arr.shape} in {path}")
            continue
        # Normalize layout to (ny, nx, pts)
        if arr.shape[0] == ny and arr.shape[1] == nx:
            data = arr
        elif arr.shape[0] == nx and arr.shape[1] == ny:
            data = np.transpose(arr, (1, 0, 2))
        elif arr.shape[0] == bias.size and arr.shape[1] == ny and arr.shape[2] == nx:
            data = np.transpose(arr, (1, 2, 0))
        elif arr.shape[0] == ny and arr.shape[2] == nx:
            data = np.transpose(arr, (0, 2, 1))
        else:
            data = arr
        channel_data[ch_key] = data
    if not channel_data:
        log(f"[Nanonis] Parsed 0 spectra from {path} (channels: {list(chans.keys())})")
        return entries

    rows, cols, pts = next(iter(channel_data.values())).shape
    x_coords = x_offsets if len(x_offsets) == cols else np.linspace(0, cols - 1, cols)
    y_coords = y_offsets if len(y_offsets) == rows else np.linspace(0, rows - 1, rows)
    channel_count = len(channel_data)
    idx = 0
    for y in range(rows):
        for x in range(cols):
            idx += 1
            chan_vals = {name: np.asarray(data[y, x, :], dtype=float) for name, data in channel_data.items()}
            first_vals = next(iter(chan_vals.values()))
            axis = bias.copy() if bias.size == first_vals.size else np.linspace(0, 1, first_vals.size, dtype=float)
            entry = {
                "path": str(path),
                "matrix_dataset": dataset_key,
                "matrix_index": idx - 1,
                "grid_rows": rows,
                "grid_cols": cols,
                "x": float(x_coords[x]),
                "y": float(y_coords[y]),
                "channels": chan_vals,
                "channel_name": None,
                "channel_code": None,
                "AxisLabel": "Bias",
                "AxisUnit": "V",
                "V": axis,
                "points_per_trace": int(first_vals.size),
                "source": "nanonis_3ds",
                "time": spec_time,
            }
            entries.append(entry)
    log(f"[Nanonis] Parsed {len(entries)} spectra from {path} ({rows}x{cols}, channels={channel_count})")
    return entries


def _flatten_nanonis_fields(target: Dict[str, object], source: Dict[str, object] | None, prefix: str):
    if not source:
        return
    for key, value in source.items():
        if key in target:
            continue
        formatted_key = f"{prefix}{str(key).strip()}"
        formatted_key = formatted_key.replace(">", "_").replace(":", "_").replace(" ", "_")
        if formatted_key in target:
            continue
        target[formatted_key] = _format_meta_value(value)


def _format_meta_value(value):
    if isinstance(value, np.ndarray):
        try:
            flat = value.ravel()
            return ", ".join(str(v) for v in flat)
        except Exception:
            try:
                return np.array2string(value)
            except Exception:
                return str(value)
    if isinstance(value, dict):
        try:
            return json.dumps(value)
        except Exception:
            return str(value)
    if isinstance(value, (list, tuple, set)):
        try:
            return ", ".join(str(_format_meta_value(v)) for v in value)
        except Exception:
            return ", ".join(str(v) for v in value)
    return value


__all__ = ["prepare_nanonis_folder", "parse_nanonis_spectroscopy", "parse_nanonis_3ds"]
