"""Adapters that convert Nanonis files into Omicron-style descriptors.

This module is isolated under the providers namespace to decouple parsing from
the GUI and the native (Omicron/Anfatec) pipeline.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
import json
import math
import os
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
# Bumped: ScanTimeForward[s]/ScanTimeBackward[s] were previously never written
# to the cached header (isinstance(scan_time, (list, tuple)) silently missed
# nanonispy2's numpy-array scan_time) - existing caches need rebuilding to
# pick up the fix.
NANONIS_CACHE_VERSION = 4
_NANONIS_READ = None
_IMPORT_ERROR = None


@dataclass
class ChannelExport:
    file_name: str
    caption: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


_NANONIS_CONVERT_MAX_WORKERS = min(8, max(1, (os.cpu_count() or 4)))


def _convert_one_scan_safe(reader, scan_path: Path, cache_root: Path) -> Optional[Path]:
    try:
        return _convert_scan_file(reader, scan_path, cache_root)
    except Exception as exc:
        log(f"[Nanonis] Failed to convert {scan_path.name}: {exc}")
        return None


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
    # Each scan file converts into its own hashed cache subdirectory (see
    # _cache_dir_for), so conversions are fully independent of each other —
    # a first-time load of a folder full of never-before-seen scans (no
    # cache hits at all) was previously converting them one at a time at
    # ~75-85 ms/file, which adds up to many seconds for a few hundred files.
    # ThreadPoolExecutor.map still returns results in `scan_files` order.
    with ThreadPoolExecutor(max_workers=_NANONIS_CONVERT_MAX_WORKERS) as executor:
        results = executor.map(
            lambda scan_path: _convert_one_scan_safe(reader, scan_path, cache_root),
            scan_files,
        )
        generated = [header_path for header_path in results if header_path is not None]
    return generated


def prepare_nanonis_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit Nanonis scan files and return generated header paths."""
    reader = _ensure_nanonis_reader()
    if reader is None:
        return []
    jobs: List[Tuple[Path, Path]] = []
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
        jobs.append((scan_path, cache_root))
    if not jobs:
        return []
    with ThreadPoolExecutor(max_workers=_NANONIS_CONVERT_MAX_WORKERS) as executor:
        results = executor.map(
            lambda job: _convert_one_scan_safe(reader, job[0], job[1]),
            jobs,
        )
        generated = [header_path for header_path in results if header_path is not None]
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
    # nanonispy2 parses this as a numpy array (its own entries_to_be_floated
    # coercion), never a plain list/tuple - a plain isinstance(list, tuple)
    # check silently misses it for every real file, so ScanTimeForward[s]/
    # ScanTimeBackward[s] never got populated at all.
    if isinstance(scan_time, (list, tuple, np.ndarray)):
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
    # A scan's slow-axis "Direction" header (up/down) records which way the
    # tip physically swept while acquiring lines, and rows are stored in
    # acquisition order - for direction='up' the tip starts at the bottom of
    # the frame, so row 0 is the *south*-most line, not the north-most one
    # our origin='upper' display (and _map_spec_to_pixels's row-0-is-north
    # convention) expects. Confirmed on real data (K1030, Direction=up):
    # spectroscopy points independently verified against Nanonis's own
    # viewer landed on the wrong blobs without this flip. direction='down'
    # already starts at the top, so it needs no flip.
    scan_dir = str((scan.header or {}).get("scan_dir", "")).strip().lower()
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
            if scan_dir == "up":
                arr = np.flipud(arr)
            # Store converted channels as native float32 arrays to avoid the
            # expensive ASCII round-trip on subsequent viewer loads.
            arr = np.asarray(arr, dtype=np.float32)
            safe_channel = _safe_token(name)
            suffix = "fwd" if dir_key == "forward" else "bwd"
            data_name = f"{scan.basename}_{safe_channel}_{suffix}.npy"
            data_path = cache_dir / data_name
            np.save(data_path, arr, allow_pickle=False)
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

def _fast_spec_load_data(self):
    """Drop-in replacement for nanonispy2/nanonispy's Spec._load_data.

    The original parses a .dat file's ASCII data block with
    numpy.genfromtxt, whose per-value type-inference dominates cold-scan
    time on large folders (measured: ~28s of a ~29s cold scan across 1655
    real .dat files, ~40M calls into genfromtxt's internal _loose_call).
    It also opens the file about three times per call, including a full
    separate readlines() just to count header lines for genfromtxt's
    skip_header - redundant, since self.byte_offset (already computed by
    NanonisFile.__init__'s header parser) already points exactly at the
    start of the column-names line.

    This reuses that byte_offset to seek once and parses the data with
    np.loadtxt on the same open, already-seeked handle. Verified
    byte-for-byte identical output against the original genfromtxt path
    across all 1655 .dat files in a real reference dataset (0 mismatches),
    ~2.9x faster. Applied as a monkeypatch (see _ensure_nanonis_reader)
    rather than edited into the vendored copy - the vendor/ directory
    mirrors upstream nanonispy2 and must stay untouched (see CLAUDE.md).
    """
    with open(self.fname, 'r') as f:
        f.seek(self.byte_offset)
        column_names = f.readline().strip('\n').split('\t')
        specdata = np.loadtxt(f, delimiter='\t', ndmin=2)
    data_dict = {}
    for i, name in enumerate(column_names):
        data_dict[name] = specdata[:, i]
    return data_dict


def _patch_fast_spec_loader(reader_module) -> None:
    try:
        spec_cls = getattr(reader_module, "Spec", None)
        if spec_cls is not None and hasattr(spec_cls, "_load_data"):
            spec_cls._load_data = _fast_spec_load_data
    except Exception as exc:
        log(f"[Nanonis] Could not install fast .dat parser, falling back to default: {exc}")


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
                _patch_fast_spec_loader(_NANONIS_READ)
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
                _patch_fast_spec_loader(_NANONIS_READ)
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


# Threshold above which an embedded spectroscopy "Start time" is treated as
# implausible and the file's own disk mtime is trusted instead. Confirmed on
# real data: every single-spectrum .dat file in one dataset had its embedded
# Start time running a consistent ~58-59 minutes *ahead* of its own disk
# mtime (spanning files from different days, so not a one-off DST fluke),
# while every image in the same folder had an embedded time matching its
# own file's mtime exactly (0.0 min difference) - proving mtime is reliable
# in that environment and the embedded field is specifically wrong for
# spectroscopy. A real acquisition's file write happens seconds after the
# measurement finishes, so any large gap is a red flag, not normal jitter.
_SPEC_TIME_SANITY_THRESHOLD_S = 600


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
        try:
            mtime = datetime.fromtimestamp(Path(path).stat().st_mtime)
            if abs((dt - mtime).total_seconds()) > _SPEC_TIME_SANITY_THRESHOLD_S:
                meta["time"] = mtime
                meta["time_source_fallback"] = "mtime_sanity_check"
        except Exception:
            pass
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
    """Select a relative Z axis if present (z_rel naming, e.g. "Z rel (m)")."""
    for name, data in signals.items():
        low = name.lower()
        if "z_rel" in low or "zrel" in low or "rel z" in low or "z rel" in low:
            return name, data
    return None, None


def _select_topo_axis(signals: Dict[str, np.ndarray]) -> Tuple[Optional[str], Optional[np.ndarray]]:
    """Best-effort selection of an absolute Z/piezo axis, distinct from a
    relative-Z channel (see `_select_z_rel_axis`) - i.e. the same candidate
    list as `_select_z_axis` but excluding anything "rel"-named."""
    candidates = [
        "Z (m)",
        "Z",
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
        if "rel" in low:
            continue
        if low.startswith("z") or "z " in low or " z" in low or "distance" in low:
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


def _select_true_bias_axis(signals: Dict[str, np.ndarray]) -> Tuple[Optional[str], Optional[np.ndarray]]:
    """Like `_select_bias_axis`, but without the broad "any (V)-named
    column" fallback - used only for the Axis dropdown's "Bias" choice,
    where mislabeling an unrelated voltage channel (e.g. an oscillation
    excitation amplitude, which also happens to be measured in volts) as
    "Bias" would be actively misleading. `_select_bias_axis`'s broader match
    stays as-is for picking the file's *primary* sweep axis, where a
    best-effort guess is better than failing to parse the file at all."""
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
        if name.lower().startswith("bias"):
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

    # Build a richer, always-available AxisChoices list (bias / relative-Z /
    # absolute-Z), independent of the `prefer_z` filename heuristic above -
    # that heuristic only picks the *default* axis; a relative-Z channel
    # (e.g. "Z rel (m)") should be selectable in the Axis dropdown whenever
    # it's present, not just for files whose name flags them as Z-spectroscopy.
    axes_choices: List[Dict[str, object]] = []
    bias_choice = None
    bias_name, bias_data = _select_true_bias_axis(spec.signals)
    if bias_name is not None and bias_data is not None:
        bias_choice = {
            "key": "bias",
            "label": re.sub(r"\s*\(.*?\)", "", bias_name).strip() or "Bias",
            "unit": "V",
            "values": np.asarray(bias_data, dtype=float).copy(),
        }
    else:
        # No per-point Bias column - common for a pure Z-sweep, where bias
        # is held fixed for the whole spectrum rather than swept. Fall back
        # to the header's own fixed scalar bias value as a constant-valued
        # choice, instead of silently omitting "Bias" from the Axis list.
        fixed_bias = None
        for hdr_key in ("Bias>Bias (V)", "Bias (V)"):
            try:
                if hdr_key in (spec.header or {}):
                    fixed_bias = float(spec.header[hdr_key])
                    break
            except Exception:
                continue
        if fixed_bias is not None and axis.size:
            bias_choice = {
                "key": "bias",
                "label": "Bias",
                "unit": "V",
                "values": np.full(axis.shape, fixed_bias, dtype=float),
            }
    zrel_choice = None
    zrel_values = None
    zrel_name, zrel_data = _select_z_rel_axis(spec.signals)
    if zrel_name is not None and zrel_data is not None:
        zrel_values = np.asarray(zrel_data, dtype=float).copy()
        try:
            if np.nanmax(np.abs(zrel_values)) < 1e-6:
                zrel_values = zrel_values * 1e9  # assume meters -> nm
        except Exception:
            pass
        zrel_choice = {
            "key": "z",
            "label": re.sub(r"\s*\(.*?\)", "", zrel_name).strip() or "Z rel",
            "unit": "nm",
            "values": zrel_values,
        }
    topo_choice = None
    topo_values = None
    topo_name, topo_data = _select_topo_axis(spec.signals)
    if topo_name is not None and topo_data is not None and topo_name != zrel_name:
        topo_values = np.asarray(topo_data, dtype=float).copy()
        try:
            if np.nanmax(np.abs(topo_values)) < 1e-6:
                topo_values = topo_values * 1e9  # assume meters -> nm
        except Exception:
            pass
        topo_choice = {
            "key": "topo",
            "label": re.sub(r"\s*\(.*?\)", "", topo_name).strip() or "Topo",
            "unit": "nm",
            "values": topo_values,
        }
    # When both a relative-Z and an absolute-Z axis exist, record the
    # constant offset between them so combining multiple spectra can add it
    # back to each one's own relative-Z values, keeping their true relative
    # height differences meaningful instead of every trace starting at zero.
    if zrel_choice is not None and topo_choice is not None:
        try:
            diff = topo_values - zrel_values
            finite = diff[np.isfinite(diff)]
            if finite.size:
                zrel_choice["origin_abs"] = float(np.nanmedian(finite))
        except Exception:
            pass
    # For Z-spectroscopy files the meaningful sweep axis is Z itself, so list
    # absolute Z first, then relative Z, then Bias last (Bias is normally
    # held fixed during a Z sweep - still offered, just the least useful
    # default here). Bias-spectroscopy files keep the conventional opposite
    # order (Bias first).
    ordered_choices = (
        [topo_choice, zrel_choice, bias_choice]
        if prefer_z
        else [bias_choice, zrel_choice, topo_choice]
    )
    axes_choices.extend(choice for choice in ordered_choices if choice is not None)

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
        "AxisChoices": axes_choices or None,
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
        # A grid can be acquired at any rotation, independent of whatever
        # image it later gets anchored to (nanonispy2 exposes this as
        # grid.header['angle'], parsed from the raw file's "Grid settings"
        # field - see nanonispy2/read.py). Building x_offsets/y_offsets as
        # a naive axis-aligned linspace (the previous behavior) silently
        # discarded that angle, fabricating x/y as if the grid were
        # unrotated - confirmed wrong against real rotated data (a grid
        # sharing its anchor image's exact center/size/angle round-tripped
        # to exactly the image's own pixel bounds only once this rotation
        # was applied; the unrotated version did not). Local per-pixel
        # offsets are rotated into true absolute (x, y) here, using the
        # same rotation convention _map_spec_to_pixels (main_window.py)
        # uses to go the other way (absolute -> local for display).
        angle_deg = _safe_float(grid.header.get("angle"), default=0.0)
        local_x = np.linspace(-rx_nm / 2, rx_nm / 2, nx)
        local_y = np.linspace(-ry_nm / 2, ry_nm / 2, ny)
        lx_grid, ly_grid = np.meshgrid(local_x, local_y)  # shape (ny, nx)
        if angle_deg:
            theta = math.radians(angle_deg)
            cos_t, sin_t = math.cos(theta), math.sin(theta)
            abs_dx = lx_grid * cos_t + ly_grid * sin_t
            abs_dy = -lx_grid * sin_t + ly_grid * cos_t
        else:
            abs_dx, abs_dy = lx_grid, ly_grid
        x_abs = cx_nm + abs_dx
        y_abs = cy_nm + abs_dy
    except Exception:
        x_abs = np.tile(np.arange(nx, dtype=float), (ny, 1))
        y_abs = np.tile(np.arange(ny, dtype=float).reshape(-1, 1), (1, nx))
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
    if x_abs.shape == (rows, cols) and y_abs.shape == (rows, cols):
        x_coords_2d, y_coords_2d = x_abs, y_abs
    else:
        x_coords_2d, y_coords_2d = np.meshgrid(
            np.arange(cols, dtype=float), np.arange(rows, dtype=float)
        )
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
                "x": float(x_coords_2d[y, x]),
                "y": float(y_coords_2d[y, x]),
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
