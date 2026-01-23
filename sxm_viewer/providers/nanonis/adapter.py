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
        try:
            return float(str(value).strip()) * 1e9
        except Exception:
            return None


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
    if axis_name:
        low = axis_name.lower()
        if "(m)" in low or " distance" in low or "distance " in low:
            axis = axis * 1e9  # convert meters to nm for display consistency
            axis_unit = "nm"
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
    if not channels:
        return []
    meta = _nanonis_spec_metadata(spec.header or {}, Path(path))
    entry = {
        "path": str(path),
        "V": axis.copy(),
        "AxisLabel": axis_name,
        "AxisUnit": axis_unit,
        "AltAxis": alt_axis.copy() if alt_axis is not None else None,
        "AltAxisLabel": alt_axis_name,
        "AltAxisUnit": alt_axis_unit,
        "channels": channels,
    }
    entry.update(meta)
    _flatten_nanonis_fields(entry, spec.header or {}, prefix="NanonisSpec:")
    return [entry]


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


__all__ = ["prepare_nanonis_folder", "parse_nanonis_spectroscopy"]
