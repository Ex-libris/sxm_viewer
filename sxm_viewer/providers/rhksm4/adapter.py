"""Adapter that converts RHK SM4 (.sm4) files into SXMViewer descriptors."""

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

try:
    from importlib import import_module
except Exception:  # pragma: no cover
    import_module = None  # type: ignore


SM4_CACHE_DIRNAME = ".sxmviewer_sm4"
SM4_CACHE_VERSION = 1
_SM4_MODULE = None
_SM4_IMPORT_ERROR = None


@dataclass
class ChannelExport:
    file_name: str
    caption: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


def prepare_sm4_folder(folder: Path | str) -> List[Path]:
    folder = Path(folder)
    scan_files = sorted({p for p in folder.glob("*.sm4") if p.is_file()} | {p for p in folder.glob("*.SM4") if p.is_file()})
    if not scan_files:
        return []
    cache_root = folder / SM4_CACHE_DIRNAME
    cache_root.mkdir(exist_ok=True)
    generated: List[Path] = []
    for sm4_path in scan_files:
        try:
            generated.extend(_convert_sm4_file(sm4_path, cache_root))
        except Exception as exc:
            log(f"[SM4] Failed to convert {sm4_path.name}: {exc}")
    return generated


def prepare_sm4_files(paths: Iterable[Path | str]) -> List[Path]:
    generated: List[Path] = []
    seen = set()
    for raw_path in paths or []:
        sm4_path = Path(raw_path)
        if not sm4_path.is_file():
            continue
        if sm4_path.suffix.lower() != ".sm4":
            continue
        try:
            key = str(sm4_path.resolve()).lower()
        except Exception:
            key = str(sm4_path).lower()
        if key in seen:
            continue
        seen.add(key)
        cache_root = sm4_path.parent / SM4_CACHE_DIRNAME
        cache_root.mkdir(exist_ok=True)
        try:
            generated.extend(_convert_sm4_file(sm4_path, cache_root))
        except Exception as exc:
            log(f"[SM4] Failed to convert {sm4_path.name}: {exc}")
    return generated


def _convert_sm4_file(sm4_path: Path, cache_root: Path) -> List[Path]:
    reader = _ensure_sm4_reader()
    if reader is None:
        return []
    src_stat = sm4_path.stat()
    cache_dir = _cache_dir_for(sm4_path, cache_root)
    meta_path = cache_dir / "meta.json"
    # we may generate multiple headers per file; keep a single meta.json for rebuild decisions
    if cache_dir.exists() and meta_path.exists() and not _needs_rebuild(meta_path, src_stat.st_mtime, src_stat.st_size):
        return sorted(cache_dir.glob("*_sm4.txt"))

    if cache_dir.exists():
        shutil.rmtree(cache_dir, ignore_errors=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    sm4 = reader.RHKsm4(str(sm4_path))
    pages = [p for p in sm4 if getattr(p, "attrs", {}).get("RHK_PageDataType") == 0]
    if not pages:
        log(f"[SM4] No image pages found in {sm4_path.name}")
        return []

    # Group pages by raster shape so each header has consistent x/y pixels.
    groups: Dict[Tuple[int, int], List[object]] = {}
    for p in pages:
        try:
            xpix = int(p.attrs.get("RHK_Xsize"))
            ypix = int(p.attrs.get("RHK_Ysize"))
        except Exception:
            continue
        groups.setdefault((xpix, ypix), []).append(p)

    out_headers: List[Path] = []
    group_idx = 0
    for (xpix, ypix), group_pages in sorted(groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        group_idx += 1
        header = _make_header(sm4_path, group_pages[0], xpix=xpix, ypix=ypix, group_idx=group_idx)
        channels = _export_pages(group_pages, cache_dir, xpix=xpix, ypix=ypix, prefix=f"g{group_idx:02d}_")
        if not channels:
            continue
        header_path = cache_dir / f"{sm4_path.stem}_sm4_{group_idx:02d}.txt"
        _write_sxm_style_header(header_path, header, channels, source=sm4_path)
        out_headers.append(header_path)

    meta_out = {
        "source": str(sm4_path),
        "mtime": src_stat.st_mtime,
        "size": src_stat.st_size,
        "generated": datetime.utcnow().isoformat(timespec="seconds"),
        "headers": [p.name for p in out_headers],
        "version": SM4_CACHE_VERSION,
    }
    meta_path.write_text(json.dumps(meta_out, indent=2), encoding="utf-8")
    return out_headers


def _make_header(source: Path, page: object, *, xpix: int, ypix: int, group_idx: int) -> Dict[str, object]:
    attrs = getattr(page, "attrs", {}) or {}
    header: Dict[str, object] = {
        "xPixel": int(xpix),
        "yPixel": int(ypix),
        "ConvertedSourceType": "rhksm4",
        "ConvertedSourceName": source.name,
        "ConvertedPageGroup": int(group_idx),
    }
    # Scan size from coordinate scales (if meters, convert to nm).
    xr, yr = _scan_ranges_nm(attrs, xpix=xpix, ypix=ypix)
    if xr is not None:
        header["XScanRange"] = float(xr)
    if yr is not None:
        header["YScanRange"] = float(yr)
    # Preserve a few high-value fields if present.
    for key in ("RHK_DateTime", "RHK_Bias", "RHK_Angle", "RHK_UserText"):
        if key in attrs:
            header[key] = attrs.get(key)
    return header


def _export_pages(
    pages: Sequence[object],
    cache_dir: Path,
    *,
    xpix: int,
    ypix: int,
    prefix: str,
) -> List[ChannelExport]:
    exports: List[ChannelExport] = []
    for idx, page in enumerate(pages):
        attrs = getattr(page, "attrs", {}) or {}
        label = getattr(page, "label", "") or attrs.get("RHK_Label") or f"page{idx}"
        caption = str(label).strip() or f"page{idx}"
        unit = str(attrs.get("RHK_Zunits") or "")
        scale = _safe_float(attrs.get("RHK_Zscale"), default=1.0)
        offset = _safe_float(attrs.get("RHK_Zoffset"), default=0.0)

        data = np.asarray(getattr(page, "data", None))
        if data is None or data.size == 0:
            continue
        if data.shape != (xpix, ypix) and data.shape != (ypix, xpix):
            # try to coerce flattened arrays
            if data.ndim == 1 and data.size >= xpix * ypix:
                data = data[: xpix * ypix].reshape((ypix, xpix))
            else:
                log(f"[SM4] Skipping page {caption}: unexpected shape {data.shape}")
                continue
        if data.shape == (xpix, ypix):
            data = data.reshape((xpix, ypix)).T
        # Store raw int32 when possible; float32 otherwise.
        out_name = _safe_filename(f"{prefix}ch{idx:02d}_{caption}.dat")
        out_path = cache_dir / out_name
        if np.issubdtype(data.dtype, np.integer):
            np.asarray(data, dtype=np.dtype("<i4")).tofile(out_path)
        else:
            np.asarray(data, dtype=np.dtype("<f4")).tofile(out_path)
            # float exports are already in numeric units; keep scaling if present anyway.
        exports.append(ChannelExport(file_name=out_name, caption=caption, phys_unit=unit, scale=float(scale), offset=float(offset)))
    return exports


def _scan_ranges_nm(attrs: Dict[str, object], *, xpix: int, ypix: int) -> Tuple[Optional[float], Optional[float]]:
    xscale = _safe_float(attrs.get("RHK_Xscale"), default=None)
    yscale = _safe_float(attrs.get("RHK_Yscale"), default=None)
    xunit = str(attrs.get("RHK_Xunits") or "").strip().lower()
    yunit = str(attrs.get("RHK_Yunits") or "").strip().lower()
    if xscale is None or yscale is None:
        return None, None
    xr = abs(float(xscale)) * float(xpix)
    yr = abs(float(yscale)) * float(ypix)
    xr_nm = _to_nm(xr, xunit)
    yr_nm = _to_nm(yr, yunit)
    return xr_nm, yr_nm


def _to_nm(value: float, unit: str) -> Optional[float]:
    unit = (unit or "").strip().lower()
    if not unit:
        return None
    if unit in {"nm"}:
        return float(value)
    if unit in {"m", "meter", "meters"}:
        return float(value) * 1e9
    if unit in {"um", "µm"}:
        return float(value) * 1e3
    if unit in {"mm"}:
        return float(value) * 1e6
    if unit in {"a", "å", "angstrom"}:
        return float(value) * 0.1
    return None


def _write_sxm_style_header(
    header_path: Path,
    header: Dict[str, object],
    channels: Sequence[ChannelExport],
    *,
    source: Path,
):
    lines = [
        f"# Converted from {source.name} via RHK SM4 adapter",
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


def _ensure_sm4_reader():
    global _SM4_MODULE, _SM4_IMPORT_ERROR
    if _SM4_MODULE is not None or _SM4_IMPORT_ERROR:
        return _SM4_MODULE
    # Prefer an installed spym, but fall back to the vendored reader.
    candidates = ("spym.io.rhksm4._sm4", "sxm_viewer.providers.rhksm4.vendor.spym_sm4")
    for name in candidates:
        try:
            _SM4_MODULE = import_module(name) if import_module else None
            if _SM4_MODULE:
                return _SM4_MODULE
        except Exception as exc:
            _SM4_IMPORT_ERROR = exc
            continue
    if _SM4_IMPORT_ERROR:
        log(f"[SM4] Adapter unavailable: {_SM4_IMPORT_ERROR}")
    return None


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
    if int(meta.get("version", -1)) != int(SM4_CACHE_VERSION):
        return True
    if abs(float(meta.get("mtime", 0.0)) - float(mtime)) > 1e-6:
        return True
    if int(meta.get("size", -1)) != int(size):
        return True
    headers = meta.get("headers") or []
    if not headers:
        # Fall back to a glob if older meta.json lacks explicit names.
        return not bool(list(meta_path.parent.glob("*_sm4_*.txt")))
    for name in headers:
        if not (meta_path.parent / str(name)).exists():
            return True
    return False


def _safe_float(value: object, *, default: Optional[float] = 0.0) -> Optional[float]:
    try:
        return float(value)
    except Exception:
        return default


_BAD_CHARS = re.compile(r"[^a-zA-Z0-9_.-]+")


def _safe_filename(value: str, *, fallback: str = "channel") -> str:
    value = (value or "").strip()
    if not value:
        return fallback
    value = value.replace(" ", "_")
    value = _BAD_CHARS.sub("_", value)
    value = value.strip("._-")
    return value or fallback


__all__ = ["prepare_sm4_folder", "prepare_sm4_files"]
