"""Adapter that converts WSxM session files into Omicron-style descriptors.

A WSxM session is a `<name>.wsxm` manifest (INI-style text) plus a sibling
data folder (named by the manifest's `Directory for files`, conventionally
`<name>_files`) holding one `.stp` file per channel/direction. Each `.stp`
file is itself a small ASCII header (size given by a `Image header size: N`
marker) followed by a raw little-endian float32/float64 raster.

This module is isolated under the providers namespace, like
`providers/nanonis/adapter.py`, to decouple parsing from the GUI and the
native (Omicron/Anfatec) pipeline. It follows the same conversion pattern:
turn foreign per-channel files into a synthetic Omicron-style header (one per
physical scan) plus cached `.npy` channel payloads, so the rest of the app
never needs to know WSxM exists.
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

WSXM_CACHE_DIRNAME = ".sxmviewer_wsxm"
WSXM_CACHE_VERSION = 1

_HEADER_SIZE_RE = re.compile(rb"Image header size:\s*(\d+)")
_VIEW_NAME_RE = re.compile(r"View (\d+) name:\s*(.+)")
_DIRECTORY_RE = re.compile(r"Directory for files:\s*(.+)")
# "<series> Image <channel> (<unit>)<variant>.stp", e.g.
# "NDST2-K1001 Image LI_Demod_1_X (A)1.stp" -> series="NDST2-K1001",
# channel="LI_Demod_1_X", unit="A", variant="1".
_STP_STEM_RE = re.compile(r"^(?P<series>.+?)\s+Image\s+(?P<channel>.+?)\s*\((?P<unit>[^)]*)\)(?P<variant>\d*)$")
_VALUE_UNIT_RE = re.compile(r"^([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*(.*)$")
_NANONIS_MARKER = "NANONIS ASCII HEADER:"

_LENGTH_TO_M = {
    "m": 1.0,
    "mm": 1e-3,
    "um": 1e-6,
    "µm": 1e-6,
    "nm": 1e-9,
    "pm": 1e-12,
    "a": 1e-10,
    "Å": 1e-10,
}


@dataclass
class _StpMember:
    path: Path
    series: str
    channel: str
    unit: str
    variant: str
    sections: Dict[str, Dict[str, str]]
    rows: int
    cols: int
    payload_offset: int
    dtype: str


@dataclass
class _ChannelExport:
    file_name: str
    caption: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


def prepare_wsxm_folder(folder: Path | str) -> List[Path]:
    """Convert every WSxM session (``*.wsxm``) in ``folder``; return generated header paths."""
    folder = Path(folder)
    session_files = sorted({p for p in folder.glob("*.wsxm") if p.is_file()})
    generated: List[Path] = []
    for session_path in session_files:
        generated.extend(_convert_session_safe(session_path))
    return generated


def prepare_wsxm_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit ``.wsxm`` session files; return generated header paths."""
    jobs: List[Path] = []
    seen = set()
    for raw_path in paths or []:
        session_path = Path(raw_path)
        if not session_path.is_file() or session_path.suffix.lower() != ".wsxm":
            continue
        try:
            key = str(session_path.resolve()).lower()
        except Exception:
            key = str(session_path).lower()
        if key in seen:
            continue
        seen.add(key)
        jobs.append(session_path)
    generated: List[Path] = []
    for session_path in jobs:
        generated.extend(_convert_session_safe(session_path))
    return generated


def _convert_session_safe(session_path: Path) -> List[Path]:
    try:
        return _convert_session_file(session_path)
    except Exception as exc:
        log(f"[WSxM] Failed to convert {session_path.name}: {exc}")
        return []


def _convert_session_file(session_path: Path) -> List[Path]:
    folder = session_path.parent
    directory_name, view_names = _parse_wsxm_manifest(session_path)
    files_dir = folder / directory_name
    if not files_dir.is_dir():
        fallback = folder / f"{session_path.stem}_files"
        if fallback.is_dir():
            files_dir = fallback
        else:
            log(f"[WSxM] {session_path.name}: data folder not found ({directory_name})")
            return []
    if not view_names:
        log(f"[WSxM] {session_path.name}: no views listed in session manifest")
        return []

    src_stat = session_path.stat()
    cache_root = folder / WSXM_CACHE_DIRNAME
    cache_root.mkdir(exist_ok=True)
    cache_dir = _cache_dir_for(session_path, cache_root)
    meta_path = cache_dir / "meta.json"

    files_signature = _files_dir_signature(files_dir, view_names)
    if meta_path.exists() and not _needs_rebuild(
        meta_path, src_stat.st_mtime, src_stat.st_size, files_signature
    ):
        cached = _reuse_cached_headers(meta_path, cache_dir)
        if cached is not None:
            return cached

    if cache_dir.exists():
        shutil.rmtree(cache_dir, ignore_errors=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    groups: Dict[str, List[_StpMember]] = {}
    for name in view_names:
        classified = _classify_stp_name(name)
        if classified is None:
            log(f"[WSxM] Skipping unrecognized session entry: {name}")
            continue
        series, channel, unit, variant = classified
        stp_path = files_dir / name
        if not stp_path.is_file():
            log(f"[WSxM] Referenced file missing: {stp_path}")
            continue
        try:
            sections, rows, cols, header_size, dtype = _parse_stp_header(stp_path)
        except Exception as exc:
            log(f"[WSxM] Failed to parse header for {name}: {exc}")
            continue
        groups.setdefault(series, []).append(
            _StpMember(
                path=stp_path,
                series=series,
                channel=channel,
                unit=unit,
                variant=variant,
                sections=sections,
                rows=rows,
                cols=cols,
                payload_offset=header_size,
                dtype=dtype,
            )
        )

    generated: List[Path] = []
    for series, members in groups.items():
        header_path = _write_series(cache_dir, session_path, series, members)
        if header_path is not None:
            generated.append(header_path)

    meta = {
        "source": str(session_path),
        "mtime": src_stat.st_mtime,
        "size": src_stat.st_size,
        "files_signature": files_signature,
        "generated": datetime.utcnow().isoformat(timespec="seconds"),
        "headers": [p.name for p in generated],
        "version": WSXM_CACHE_VERSION,
    }
    meta_path.write_text(json.dumps(meta, indent=2))
    return generated


def _reuse_cached_headers(meta_path: Path, cache_dir: Path) -> Optional[List[Path]]:
    try:
        meta = json.loads(meta_path.read_text())
    except Exception:
        return None
    names = meta.get("headers")
    if not isinstance(names, list):
        return None
    headers = [cache_dir / name for name in names]
    if headers and all(p.exists() for p in headers):
        return headers
    return None


# --------------------------------------------------------------------------- #
# Session manifest parsing                                                    #
# --------------------------------------------------------------------------- #

def _parse_wsxm_manifest(path: Path) -> Tuple[str, List[str]]:
    raw = path.read_text(encoding="latin-1", errors="replace")
    directory = f"{path.stem}_files"
    m = _DIRECTORY_RE.search(raw)
    if m:
        candidate = m.group(1).strip().rstrip("\\/")
        if candidate:
            directory = candidate
    entries: Dict[int, str] = {}
    for m in _VIEW_NAME_RE.finditer(raw):
        idx = int(m.group(1))
        entries[idx] = m.group(2).strip()
    names = [entries[idx] for idx in sorted(entries)]
    return directory, names


def _classify_stp_name(name: str) -> Optional[Tuple[str, str, str, str]]:
    stem = name[:-4] if name.lower().endswith(".stp") else name
    m = _STP_STEM_RE.match(stem.strip())
    if not m:
        return None
    return (
        m.group("series").strip(),
        m.group("channel").strip(),
        m.group("unit").strip(),
        m.group("variant"),
    )


# --------------------------------------------------------------------------- #
# .stp header/payload parsing                                                 #
# --------------------------------------------------------------------------- #

def _parse_stp_header(path: Path) -> Tuple[Dict[str, Dict[str, str]], int, int, int, str]:
    raw = path.read_bytes()
    m = _HEADER_SIZE_RE.search(raw)
    if not m:
        raise ValueError("missing 'Image header size' marker")
    header_size = int(m.group(1))
    text = raw[:header_size].decode("latin-1", errors="replace")
    sections: Dict[str, Dict[str, str]] = {}
    current: Optional[str] = None
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("[") and stripped.endswith("]"):
            current = stripped[1:-1].strip()
            sections.setdefault(current, {})
            continue
        if current is None or ":" not in stripped:
            continue
        key, _, value = stripped.partition(":")
        sections[current][key.strip()] = value.strip()

    general = sections.get("General Info", {})
    cols = _safe_int(general.get("Number of columns"))
    rows = _safe_int(general.get("Number of rows"))
    if cols <= 0 or rows <= 0:
        raise ValueError(f"could not parse grid dimensions ({cols}x{rows})")

    payload_len = len(raw) - header_size
    n_pixels = rows * cols
    if payload_len >= n_pixels * 8:
        dtype = "<f8"
    elif payload_len >= n_pixels * 4:
        dtype = "<f4"
    else:
        raise ValueError(f"payload too small for {cols}x{rows} grid ({payload_len} bytes)")
    return sections, rows, cols, header_size, dtype


def _read_stp_payload(path: Path, rows: int, cols: int, header_size: int, dtype: str) -> np.ndarray:
    itemsize = 8 if dtype == "<f8" else 4
    n_pixels = rows * cols
    with open(path, "rb") as fh:
        fh.seek(header_size)
        buf = fh.read(n_pixels * itemsize)
    # Raw payload only - deliberately no orientation flip here. WSxM's own
    # column-mirror convention (relative to Nanonis) has only ever been
    # validated by hand against a single sample file; baking an unverified
    # flip into every import risks the exact silent-orientation bug class
    # documented at length for the Nanonis adapter and the Grid Map Explorer
    # (see CLAUDE.md "Spectroscopy position mapping" / "Grid Map Explorer").
    # `ScanDir`/`Y scanning direction` are still recorded on the header for
    # whoever validates this properly against an independent reference.
    arr = np.frombuffer(buf, dtype=dtype).reshape(rows, cols)
    return np.asarray(arr, dtype=np.float32)


# --------------------------------------------------------------------------- #
# Embedded (Nanonis-origin) header recovery                                   #
# --------------------------------------------------------------------------- #

def _parse_embedded_nanonis_header(comments: str) -> Dict[str, str]:
    r"""Recover the original instrument header WSxM embeds for converted scans.

    When a `.stp` file was produced by WSxM importing a Nanonis `.sxm` scan,
    `[Miscellaneous] Comments` contains the *entire* original Nanonis ASCII
    header, double-escaped: the header was first serialized with real
    newlines turned into literal `\n` two-character sequences, and *that*
    result was embedded as-is, with every backslash (including the ones
    from the `\n` markers themselves, and the ones in Windows paths like
    `SCAN_FILE`) doubled again. Confirmed byte-for-byte on a real sample:
    the raw bytes between fields are `\`, `\`, `n` (two literal backslashes
    then 'n'), and `SCAN_FILE`'s path separators are doubled too. Undoing
    only the outer `\n`-to-newline pass first (matching just the trailing
    backslash+n of that triple) leaves a stray leading backslash on every
    line - undoubling backslashes first, then decoding `\n`, recovers the
    original single-backslash text cleanly. Absent entirely for native
    (non-Nanonis-origin) WSxM acquisitions.
    """
    if not comments or _NANONIS_MARKER not in comments:
        return {}
    _, _, blob = comments.partition(_NANONIS_MARKER)
    blob = blob.replace("\\\\", "\\")
    blob = blob.replace("\\n", "\n")
    fields: Dict[str, str] = {}
    for line in blob.splitlines():
        stripped = line.strip()
        if not stripped.startswith(":"):
            continue
        body = stripped[1:]
        key, sep, value = body.partition(":")
        if not sep:
            continue
        fields[key.strip()] = value.strip()
    return fields


# --------------------------------------------------------------------------- #
# Synthetic Omicron-style header emission                                     #
# --------------------------------------------------------------------------- #

def _write_series(
    cache_dir: Path, session_path: Path, series: str, members: List[_StpMember]
) -> Optional[Path]:
    if not members:
        return None
    def _sort_key(m: _StpMember):
        acq_channel = m.sections.get("General Info", {}).get("Acquisition channel", "").strip()
        return (acq_channel or m.channel, m.variant, m.path.name)

    members = sorted(members, key=_sort_key)
    rows, cols = members[0].rows, members[0].cols
    kept = []
    for m in members:
        if (m.rows, m.cols) != (rows, cols):
            log(
                f"[WSxM] {series}: '{m.path.name}' grid size {m.cols}x{m.rows} "
                f"differs from series {cols}x{rows}, skipping"
            )
            continue
        kept.append(m)
    members = kept
    if not members:
        return None

    rep = members[0]
    control = rep.sections.get("Control", {})
    general = rep.sections.get("General Info", {})
    head_settings = rep.sections.get("Head Settings", {})
    misc = rep.sections.get("Miscellaneous", {})

    x_amp_val, x_amp_unit = _split_value_unit(control.get("X Amplitude"))
    y_amp_val, y_amp_unit = _split_value_unit(control.get("Y Amplitude"))
    angle_val, _ = _split_value_unit(control.get("Angle"))
    bias_val, bias_unit = _split_value_unit(control.get("Topography Bias"))
    setpoint_val, setpoint_unit = _split_value_unit(control.get("Set Point"))

    comments = misc.get("Comments", "")
    nanonis_fields = _parse_embedded_nanonis_header(comments)
    original_source = nanonis_fields.get("SCAN_FILE", "").strip()
    rec_date = _format_date_string(nanonis_fields.get("REC_DATE", ""))
    rec_time = _format_time_string(nanonis_fields.get("REC_TIME", ""))
    if not rec_date and not rec_time:
        try:
            dt = datetime.fromtimestamp(rep.path.stat().st_mtime)
            rec_date, rec_time = dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M:%S")
        except Exception:
            pass

    scan_dir = nanonis_fields.get("SCAN_DIR", "").strip().lower() or general.get(
        "Y scanning direction", ""
    ).strip().lower()

    header: Dict[str, object] = {
        "xPixel": cols,
        "yPixel": rows,
        "XScanRange": _to_nm(x_amp_val, x_amp_unit),
        "YScanRange": _to_nm(y_amp_val, y_amp_unit),
        "XPhysUnit": "nm",
        "YPhysUnit": "nm",
        "Angle": angle_val if angle_val is not None else 0.0,
        "ScanAngle": angle_val if angle_val is not None else 0.0,
        "ScanDir": scan_dir,
        "Bias": bias_val if bias_val is not None else "",
        "BiasPhysUnit": bias_unit or "V",
        "SetPoint": setpoint_val if setpoint_val is not None else "",
        "SetPointPhysUnit": setpoint_unit,
        "HeadType": general.get("Head type", ""),
        "Date": rec_date,
        "Time": rec_time,
        "OriginalSourcePath": original_source,
        "WsxmSeries": series,
        "WsxmSessionFile": str(session_path),
        "DocumentTitle": misc.get("Document title", ""),
        "Comment": comments,
    }
    _flatten_fields(header, control, prefix="WsxmControl:")
    _flatten_fields(header, general, prefix="WsxmGeneralInfo:")
    _flatten_fields(header, head_settings, prefix="WsxmHeadSettings:")
    if nanonis_fields:
        _flatten_fields(header, nanonis_fields, prefix="Nanonis:")

    exports: List[_ChannelExport] = []
    for m in members:
        try:
            arr = _read_stp_payload(m.path, m.rows, m.cols, m.payload_offset, m.dtype)
        except Exception as exc:
            log(f"[WSxM] Failed reading payload for {m.path.name}: {exc}")
            continue
        m_general = m.sections.get("General Info", {})
        # `Acquisition channel` is the authoritative channel name - the
        # filename-parsed one can be corrupted by WSxM's own uniquification
        # (confirmed on a real sample: the backward copy of "OC_M1_Freq.
        # _Shift" was written to disk as "...Freq1._Shift...", which would
        # otherwise look like a distinct channel from its forward pair).
        channel_name = m_general.get("Acquisition channel", "").strip() or m.channel
        direction = m_general.get("X scanning direction", "").strip().lower()
        if direction.startswith("back"):
            dir_tag, caption = "Bwd", f"{channel_name} (Bwd)"
        elif direction.startswith("forw"):
            dir_tag, caption = "Fwd", channel_name
        else:
            dir_tag = m.variant or "0"
            caption = f"{channel_name} (#{dir_tag})"
        m_misc = m.sections.get("Miscellaneous", {})
        scale = _safe_float(m_misc.get("Z Scale Factor"), default=1.0)
        offset = _safe_float(m_misc.get("Z Scale Offset"), default=0.0)
        safe_channel = _safe_token(f"{channel_name}_{dir_tag}")
        data_name = f"{_safe_token(series)}_{safe_channel}.npy"
        np.save(cache_dir / data_name, arr, allow_pickle=False)
        exports.append(
            _ChannelExport(
                file_name=data_name,
                caption=caption,
                phys_unit=m.unit,
                scale=scale,
                offset=offset,
            )
        )

    if not exports:
        return None
    _dedupe_captions(exports)

    header_path = cache_dir / f"{_safe_token(series)}_wsxm.txt"
    _write_wsxm_style_header(header_path, header, exports, source=session_path)
    return header_path


def _dedupe_captions(exports: Sequence[_ChannelExport]) -> None:
    seen: Dict[str, int] = {}
    for ch in exports:
        count = seen.get(ch.caption, 0)
        if count:
            ch.caption = f"{ch.caption} [{count + 1}]"
        seen[ch.caption] = count + 1


def _write_wsxm_style_header(
    header_path: Path,
    header: Dict[str, object],
    channels: Sequence[_ChannelExport],
    *,
    source: Path,
) -> None:
    lines = [
        f"# Converted from {source.name} via WSxM adapter",
        f"ConvertedSource = {source}",
        f"ConvertedTimestamp = {datetime.utcnow().isoformat(timespec='seconds')}",
    ]
    for key, value in header.items():
        if value is None or value == "":
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


def _flatten_fields(target: Dict[str, object], source: Dict[str, str], prefix: str) -> None:
    if not source:
        return
    for key, value in source.items():
        formatted_key = f"{prefix}{str(key).strip()}"
        formatted_key = formatted_key.replace(">", "_").replace(":", "_").replace(" ", "_")
        if formatted_key in target:
            continue
        target[formatted_key] = value


# --------------------------------------------------------------------------- #
# Cache bookkeeping                                                           #
# --------------------------------------------------------------------------- #

def _cache_dir_for(src: Path, cache_root: Path) -> Path:
    try:
        resolved = str(src.resolve())
    except Exception:
        resolved = str(src)
    digest = hashlib.sha1(resolved.encode("utf-8")).hexdigest()[:10]
    return cache_root / f"{src.stem}_{digest}"


def _files_dir_signature(files_dir: Path, names: Iterable[str]) -> float:
    total = 0.0
    for name in names:
        try:
            st = (files_dir / name).stat()
        except OSError:
            continue
        total += st.st_mtime + st.st_size
    return total


def _needs_rebuild(meta_path: Path, mtime: float, size: int, files_signature: float) -> bool:
    try:
        meta = json.loads(meta_path.read_text())
    except Exception:
        return True
    if int(meta.get("version", -1)) != int(WSXM_CACHE_VERSION):
        return True
    if abs(float(meta.get("mtime", 0.0)) - mtime) > 1e-6:
        return True
    if int(meta.get("size", -1)) != int(size):
        return True
    if abs(float(meta.get("files_signature", -1.0)) - files_signature) > 1e-6:
        return True
    return False


# --------------------------------------------------------------------------- #
# Small parsing utilities                                                     #
# --------------------------------------------------------------------------- #

def _split_value_unit(text) -> Tuple[Optional[float], str]:
    if text is None:
        return None, ""
    s = str(text).strip()
    if not s:
        return None, ""
    m = _VALUE_UNIT_RE.match(s)
    if not m:
        return None, ""
    try:
        value = float(m.group(1))
    except Exception:
        return None, m.group(2).strip()
    return value, m.group(2).strip()


def _to_nm(value: Optional[float], unit: str) -> float:
    if value is None:
        return 0.0
    factor = _LENGTH_TO_M.get(unit.strip().lower())
    if factor is None:
        return value
    return value * factor * 1e9


def _safe_float(value, default: Optional[float] = 0.0):
    try:
        if value is None:
            return default
        if isinstance(value, (float, int)):
            return float(value)
        text = str(value).strip()
        if not text:
            return default
        return float(text)
    except Exception:
        return default


def _safe_int(value, default: int = 0) -> int:
    try:
        return int(float(value))
    except Exception:
        return default


def _format_date_string(text: str) -> str:
    if not text:
        return ""
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text.strip(), fmt).strftime("%Y-%m-%d")
        except Exception:
            continue
    return text.strip()


def _format_time_string(text: str) -> str:
    if not text:
        return ""
    for fmt in ("%H:%M:%S", "%H.%M.%S"):
        try:
            return datetime.strptime(text.strip(), fmt).strftime("%H:%M:%S")
        except Exception:
            continue
    return text.strip()


def _safe_token(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", str(text).strip()).strip("_") or "x"
