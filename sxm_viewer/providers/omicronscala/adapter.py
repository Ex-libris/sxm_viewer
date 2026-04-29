"""Adapter that converts Omicron SCALA exports into SXMViewer descriptors.

The viewer's native image pipeline expects an Omicron/Anfatec-style `.txt`
header with `FileDescBegin` blocks that reference per-channel grid files.

SCALA acquisitions are distributed as a `.par` file (metadata) plus one binary
file per image channel.  This adapter creates a small cache next to the source
file containing:
  - `<stem>_scala.txt` : synthesized header
  - `chNN_<label>.int` : raw channel grids in little-endian int16
  - `meta.json`        : cache coherency info
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import hashlib
import json
import re
import shutil
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np

from ...utils.logging import log


SCALA_CACHE_DIRNAME = ".sxmviewer_scala"
SCALA_CACHE_VERSION = 1


@dataclass
class ChannelExport:
    file_name: str
    caption: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


def prepare_scala_folder(folder: Path | str) -> List[Path]:
    folder = Path(folder)
    scan_files = sorted({p for p in folder.glob("*.par") if p.is_file()})
    if not scan_files:
        return []
    generated: List[Path] = []
    cache_root = folder / SCALA_CACHE_DIRNAME
    cache_root.mkdir(exist_ok=True)
    for par_path in scan_files:
        try:
            header_path = _convert_par_file(par_path, cache_root)
        except Exception as exc:
            log(f"[SCALA] Failed to convert {par_path.name}: {exc}")
            continue
        if header_path is not None:
            generated.append(header_path)
    return generated


def prepare_scala_files(paths: Iterable[Path | str]) -> List[Path]:
    generated: List[Path] = []
    seen = set()
    for raw_path in paths or []:
        par_path = Path(raw_path)
        if not par_path.is_file():
            continue
        if par_path.suffix.lower() != ".par":
            continue
        try:
            key = str(par_path.resolve()).lower()
        except Exception:
            key = str(par_path).lower()
        if key in seen:
            continue
        seen.add(key)
        cache_root = par_path.parent / SCALA_CACHE_DIRNAME
        cache_root.mkdir(exist_ok=True)
        try:
            header_path = _convert_par_file(par_path, cache_root)
        except Exception as exc:
            log(f"[SCALA] Failed to convert {par_path.name}: {exc}")
            continue
        if header_path is not None:
            generated.append(header_path)
    return generated


def _convert_par_file(par_path: Path, cache_root: Path) -> Optional[Path]:
    src_stat = par_path.stat()
    cache_dir = _cache_dir_for(par_path, cache_root)
    header_path = cache_dir / f"{par_path.stem}_scala.txt"
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

    meta = _parse_par_metadata(par_path)
    xpix = _safe_int(meta.get("ImageSizeinX"), default=None)
    ypix = _safe_int(meta.get("ImageSizeinY"), default=None)
    if not xpix or not ypix:
        log(f"[SCALA] Missing ImageSizeinX/ImageSizeinY in {par_path.name}")
        return None

    header = _make_header(meta, xpix=xpix, ypix=ypix, source=par_path)
    channels = _extract_channels(meta, par_path.parent, cache_dir, xpix=xpix, ypix=ypix)
    if not channels:
        log(f"[SCALA] No usable channels found in {par_path.name}")
        return None

    _write_sxm_style_header(header_path, header, channels, source=par_path)
    meta_out = {
        "source": str(par_path),
        "mtime": src_stat.st_mtime,
        "size": src_stat.st_size,
        "generated": datetime.utcnow().isoformat(timespec="seconds"),
        "channels": len(channels),
        "header_name": header_path.name,
        "version": SCALA_CACHE_VERSION,
    }
    meta_path.write_text(json.dumps(meta_out, indent=2), encoding="utf-8")
    return header_path


def _parse_par_metadata(par_path: Path) -> Dict[str, str]:
    raw = par_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    cleaned: List[str] = []
    for line in raw:
        line = line.strip()
        if not line:
            continue
        # strip comments
        if ";" in line:
            line = line.split(";", 1)[0].strip()
        if not line:
            continue
        cleaned.append(line.replace(" ", ""))

    # Normalize SCALA's channel blocks by prefixing keys with a short channel id.
    meta = list(cleaned)
    chlist: List[str] = []
    imglist: List[str] = []
    spec_parameter: str = ""

    for i, entry in enumerate(list(meta)):
        if "TopographicChannel" in entry:
            # SCALA files encode a channel identifier in the filename, typically the last 3 chars.
            try:
                ch_name = meta[i + 8][-3:].upper() + "_"
            except Exception:
                ch_name = f"CH{i:02d}_"
            for off, key in enumerate(
                [
                    "TopographicChannel",
                    "Direction",
                    "MinimumRawValue",
                    "MaximumRawValue",
                    "MinimumPhysValue",
                    "MaximumPhysValue",
                    "Resolution",
                    "PhysicalUnit",
                    "Filename",
                    "DisplayName",
                ]
            ):
                idx = i + off
                if 0 <= idx < len(meta):
                    meta[idx] = ch_name + key + ":" + meta[idx].split(":", 1)[-1]
            chlist.append(ch_name)
            imglist.append(ch_name)
        elif "SpectroscopyChannel" in entry:
            # Skip spectroscopy channels for now; the viewer's image pipeline consumes grids.
            try:
                ch_name = meta[i + 16][-3:].upper() + "_"
            except Exception:
                ch_name = f"SP{i:02d}_"
            spec_parameter = meta[i + 1]
            for off, key in enumerate(
                [
                    "SpectroscopyChannel",
                    "Parameter",
                    "Direction",
                    "MinimumRawValue",
                    "MaximumRawValue",
                    "MinimumPhysValue",
                    "MaximumPhysValue",
                    "Resolution",
                    "PhysicalUnit",
                    "NumberSpecPoints",
                    "StartPoint",
                    "EndPoint",
                    "Increment",
                    "AcqTimePerPoint",
                    "DelayTimePerPoint",
                    "Feedback",
                    "Filename",
                    "DisplayName",
                ]
            ):
                idx = i + off
                if 0 <= idx < len(meta):
                    meta[idx] = ch_name + key + ":" + meta[idx].split(":", 1)[-1]
            chlist.append(ch_name)
        elif spec_parameter and (spec_parameter + "Parameter") in entry:
            # Record the chosen sweep parameter (e.g. Bias/Current); keep it globally.
            meta[i] = "SpecParam:" + spec_parameter

    pairs = []
    for entry in meta:
        if ":" not in entry:
            continue
        k, v = entry.split(":", 1)
        if not k:
            continue
        pairs.append((k, v))

    meta_dict: Dict[str, str] = {k: v for k, v in pairs}
    # Keep the normalized channel prefix lists for channel extraction
    meta_dict["_scala_chlist"] = ",".join(chlist)
    meta_dict["_scala_imglist"] = ",".join(imglist)
    return meta_dict


def _make_header(meta: Dict[str, str], *, xpix: int, ypix: int, source: Path) -> Dict[str, object]:
    header: Dict[str, object] = {
        "xPixel": int(xpix),
        "yPixel": int(ypix),
    }
    xr_nm = _safe_float(meta.get("FieldXSizeinnm"))
    yr_nm = _safe_float(meta.get("FieldYSizeinnm"))
    if xr_nm is not None:
        header["XScanRange"] = float(xr_nm)
    if yr_nm is not None:
        header["YScanRange"] = float(yr_nm)
    # best-effort offsets are also in nm in SCALA
    xo_nm = _safe_float(meta.get("XOffset"))
    yo_nm = _safe_float(meta.get("YOffset"))
    if xo_nm is not None:
        header["XOffset"] = float(xo_nm)
    if yo_nm is not None:
        header["YOffset"] = float(yo_nm)
    header["ConvertedSourceType"] = "omicronscala"
    header["ConvertedSourceName"] = source.name
    ts = meta.get("Timestamp") or ""
    if ts:
        header["Timestamp"] = ts
    return header


def _extract_channels(
    meta: Dict[str, str],
    src_dir: Path,
    cache_dir: Path,
    *,
    xpix: int,
    ypix: int,
) -> List[ChannelExport]:
    imglist = [s for s in (meta.get("_scala_imglist") or "").split(",") if s]
    exports: List[ChannelExport] = []
    for idx, prefix in enumerate(imglist):
        filename = meta.get(prefix + "Filename") or ""
        if not filename:
            continue
        src_path = src_dir / filename
        if not src_path.exists():
            log(f"[SCALA] Missing channel file {filename} referenced by {cache_dir.parent.name}")
            continue
        caption = meta.get(prefix + "DisplayName") or meta.get(prefix + "TopographicChannel") or prefix.rstrip("_")
        unit = meta.get(prefix + "PhysicalUnit") or ""
        scale, offset = _scala_scale_offset(meta, prefix)
        out_name = _safe_filename(f"ch{idx:02d}_{caption}.int")
        out_path = cache_dir / out_name

        raw = np.fromfile(src_path, dtype=np.dtype(">i2"))
        if raw.size < xpix * ypix:
            log(f"[SCALA] Channel {filename} too small ({raw.size} samples)")
            continue
        raw = raw[: xpix * ypix]
        # Viewer expects row-major (y, x) grids.
        grid = raw.reshape((ypix, xpix)).astype(np.dtype("<i2"), copy=False)
        grid.tofile(out_path)

        exports.append(
            ChannelExport(
                file_name=out_name,
                caption=str(caption),
                phys_unit=str(unit),
                scale=float(scale),
                offset=float(offset),
            )
        )
    return exports


def _scala_scale_offset(meta: Dict[str, str], prefix: str) -> Tuple[float, float]:
    min_raw = _safe_float(meta.get(prefix + "MinimumRawValue"))
    max_raw = _safe_float(meta.get(prefix + "MaximumRawValue"))
    min_phys = _safe_float(meta.get(prefix + "MinimumPhysValue"))
    max_phys = _safe_float(meta.get(prefix + "MaximumPhysValue"))
    if None in (min_raw, max_raw, min_phys, max_phys):
        return 1.0, 0.0
    if float(max_raw) == float(min_raw):
        return 1.0, float(min_phys)
    scale = (float(max_phys) - float(min_phys)) / (float(max_raw) - float(min_raw))
    offset = float(min_phys) - scale * float(min_raw)
    return scale, offset


def _write_sxm_style_header(
    header_path: Path,
    header: Dict[str, object],
    channels: Sequence[ChannelExport],
    *,
    source: Path,
):
    lines = [
        f"# Converted from {source.name} via Omicron SCALA adapter",
        f"ConvertedSource = {source}",
        f"ConvertedTimestamp = {datetime.utcnow().isoformat(timespec='seconds')}",
    ]
    for key, value in header.items():
        if value is None:
            continue
        lines.append(f"{key} = {value}")
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
    if int(meta.get("version", -1)) != int(SCALA_CACHE_VERSION):
        return True
    if abs(float(meta.get("mtime", 0.0)) - float(mtime)) > 1e-6:
        return True
    if int(meta.get("size", -1)) != int(size):
        return True
    header_name = meta.get("header_name")
    if not header_name:
        return True
    if not (meta_path.parent / header_name).exists():
        return True
    return False


def _safe_int(value: object, *, default: Optional[int] = 0) -> Optional[int]:
    try:
        return int(str(value))
    except Exception:
        return default


def _safe_float(value: object) -> Optional[float]:
    try:
        return float(str(value))
    except Exception:
        return None


_BAD_CHARS = re.compile(r"[^a-zA-Z0-9_.-]+")


def _safe_filename(value: str, *, fallback: str = "channel") -> str:
    value = (value or "").strip()
    if not value:
        return fallback
    value = value.replace(" ", "_")
    value = _BAD_CHARS.sub("_", value)
    value = value.strip("._-")
    return value or fallback


__all__ = ["prepare_scala_folder", "prepare_scala_files"]

