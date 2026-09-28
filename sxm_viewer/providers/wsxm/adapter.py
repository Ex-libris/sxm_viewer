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
from ...geometry import spec_mapping
from ...data.channel_units import guess_channel_unit
from ...processing import filters as native_filters

WSXM_CACHE_DIRNAME = ".sxmviewer_wsxm"
# v2: files whose name doesn't match _STP_STEM_RE (e.g. a WSxM-renamed
# crop/filter export like "Crop.stp") used to be silently skipped entirely -
# confirmed on real data, they never appeared as thumbnails at all. Fixed by
# _classify_from_header falling back to the file's own header when the name
# doesn't parse. v2 also adds WsxmProcessingHistory and the recovered-profile
# fields (WsxmProfileChannel/WsxmProfilePoints) - none of this is a data-
# orientation change, but existing caches predate these fields entirely, so
# bump to force a rebuild rather than serve stale headers missing them.
# v3: header filenames are now prefixed with the session's own name
# unconditionally (see _write_series) - two different .wsxm sessions in the
# same folder could otherwise produce identically-named, indistinguishable
# thumbnails for a same-named scan. Bump so existing caches regenerate under
# the new naming instead of keeping their pre-fix filenames forever.
# v4: adds filter-match detection (WsxmFilterMatch* fields) - a derived
# image (e.g. "Simple Flatten") gets paired with its raw sibling in the same
# session (matched via embedded Nanonis provenance) and tested against
# sxm_viewer's own native filters; a match confirmed to near machine
# precision is recorded so the image can eventually be reproduced live from
# the raw data instead of only importing WSxM's baked pixels. Bump so
# existing caches pick up the new fields.
# v5: DATA-ORIENTATION CHANGE. `_read_stp_payload` now flips the raster 180
# degrees (both axes) on import - confirmed by direct comparison against
# WSxM's own on-screen display of a real session; the previous version
# deliberately imported unflipped pending exactly this confirmation (see the
# v1-era comment this replaced). Every cache built before this version has
# every WSxM-derived image mirrored on both axes relative to correct - this
# bump is load-bearing, not cosmetic (see CLAUDE.md's Nanonis
# Direction=up precedent for why a data-orientation fix that skips the
# version bump silently ships only to never-before-converted files).
WSXM_CACHE_VERSION = 5

_HEADER_SIZE_RE = re.compile(rb"Image header size:\s*(\d+)")
_VIEW_NAME_RE = re.compile(r"View (\d+) name:\s*(.+)")
_DIRECTORY_RE = re.compile(r"Directory for files:\s*(.+)")
# "<series> Image <channel> (<unit>)<variant>.stp", e.g.
# "NDST2-K1001 Image LI_Demod_1_X (A)1.stp" -> series="NDST2-K1001",
# channel="LI_Demod_1_X", unit="A", variant="1".
_STP_STEM_RE = re.compile(r"^(?P<series>.+?)\s+Image\s+(?P<channel>.+?)\s*\((?P<unit>[^)]*)\)(?P<variant>\d*)$")
_VALUE_UNIT_RE = re.compile(r"^([+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*(.*)$")
_NANONIS_MARKER = "NANONIS ASCII HEADER:"
# "Layer 3: Name: Profiles; Active: Yes; Number of Points: 2;
#  Points 0: ^81,589.; Points 1: ^443,575.; ..." - WSxM's own [Graphic
# Layers] entry for a drawn profile line, confirmed against a real saved
# session (a bare crop/profile/filter test session with no public format
# docs available). Points are pixel coordinates in the image's own raster.
_PROFILE_POINT_RE = re.compile(r"Points\s+\d+:\s*\^([+-]?[\d.eE+-]+),([+-]?[\d.eE+-]+)\.")

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
    if not session_files:
        # Confirmed a real point of confusion: a WSxM session's own data
        # folder is conventionally named "<session>_files", which reads as
        # "this is where the images live" - a natural place to point a
        # folder loader, even though the actual session lives one level up
        # as a sibling ".wsxm" file. Redirect rather than silently returning
        # nothing whenever a parent session explicitly claims this exact
        # folder as its data directory.
        owner = _find_owning_session(folder)
        if owner is not None:
            session_files = [owner]
    generated: List[Path] = []
    for session_path in session_files:
        generated.extend(_convert_session_safe(session_path))
    return generated


def _find_owning_session(folder: Path) -> Optional[Path]:
    try:
        if not any(folder.glob("*.stp")):
            return None
    except OSError:
        return None
    parent = folder.parent
    try:
        resolved_folder = folder.resolve()
    except OSError:
        resolved_folder = folder
    for candidate in sorted(parent.glob("*.wsxm")):
        try:
            directory_name, _ = _parse_wsxm_manifest(candidate)
        except Exception:
            continue
        files_dir = parent / directory_name
        if not files_dir.is_dir():
            files_dir = parent / f"{candidate.stem}_files"
        try:
            if files_dir.resolve() == resolved_folder:
                return candidate
        except OSError:
            if files_dir == folder:
                return candidate
    return None


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
        stp_path = files_dir / name
        if not stp_path.is_file():
            log(f"[WSxM] Referenced file missing: {stp_path}")
            continue
        try:
            sections, rows, cols, header_size, dtype = _parse_stp_header(stp_path)
        except Exception as exc:
            log(f"[WSxM] Failed to parse header for {name}: {exc}")
            continue
        classified = _classify_stp_name(name)
        if classified is None:
            # Doesn't match the "<series> Image <channel> (<unit>)N.stp"
            # convention - e.g. a WSxM-renamed crop/filter export like
            # "Crop.stp". Confirmed on real data: these used to be silently
            # dropped entirely (never became a thumbnail at all) rather than
            # imported as their own standalone scan. Fall back to whatever
            # the file's own header says about itself.
            classified = _classify_from_header(name, sections)
        series, channel, unit, variant = classified
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

    all_members = [m for members in groups.values() for m in members]
    filter_matches = _find_filter_matches(all_members)

    generated: List[Path] = []
    for series, members in groups.items():
        header_path = _write_series(cache_dir, session_path, series, members, filter_matches)
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


def _classify_from_header(name: str, sections: Dict[str, Dict[str, str]]) -> Tuple[str, str, str, str]:
    """Fallback classification for a `.stp` file whose name doesn't match
    the series/channel/unit/variant convention (e.g. a WSxM-renamed crop or
    filter export like "Crop.stp" or "Filtered.stp"). Since nothing else in
    the session shares its filename pattern, it has no siblings to group
    with anyway - treat it as its own standalone one-channel scan, keyed by
    its own filename stem."""
    stem = name[:-4] if name.lower().endswith(".stp") else name
    general = sections.get("General Info", {})
    channel = general.get("Acquisition channel", "").strip() or stem.strip()
    unit = guess_channel_unit(channel) or ""
    series = _safe_token(stem)
    return series, channel, unit, ""


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
    arr = np.frombuffer(buf, dtype=dtype).reshape(rows, cols)
    # WSxM stores the raster mirrored on both axes relative to the
    # orientation the rest of the app (and WSxM's own on-screen display)
    # treat as "real" - confirmed by direct comparison against WSxM's own
    # display of a real session, not inferred. A prior version of this
    # function deliberately left the payload unflipped pending that
    # confirmation (see git history / CLAUDE.md's rotation-bug precedents
    # for why an unverified flip is worse than none); flip160_pixel below
    # keeps the drawn-profile mapping in lock-step with this exact
    # transform, since flipping one without the other would silently point
    # a recovered profile at the wrong feature.
    arr = arr[::-1, ::-1]
    return np.asarray(arr, dtype=np.float32)


def _flip_pixel_180(px: float, py: float, cols: int, rows: int) -> Tuple[float, float]:
    """Map a pixel coordinate through the same 180-degree flip `_read_stp_payload`
    applies to the raster, so a WSxM-recorded position (drawn against the
    *unflipped* raw payload) still lands on the same physical feature in the
    flipped array this adapter actually imports."""
    return (cols - 1 - px), (rows - 1 - py)


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


def _parse_profile_points(sections: Dict[str, Dict[str, str]]) -> Optional[List[Tuple[float, float]]]:
    """Recover a drawn profile line's endpoints from `[Graphic Layers]`.

    Confirmed on a real saved session (no public WSxM format docs exist for
    this): a drawn profile shows up as one of the numbered `Layer N` entries,
    e.g. ``Name: Profiles; Active: Yes; Number of Points: 2; Points 0:
    ^81,589.; Points 1: ^443,575.; ...`` - NOT in the separate `[Profiles]`
    section, which only records a window-placement rectangle for whatever
    profile-analysis window was open, not the line itself. Points are pixel
    coordinates in the image's own raster (row 0 = same orientation as the
    imported array - verified by rendering both interpretations against a
    real step-edge feature; the un-flipped one lands a sensible "measure
    step height" line, the flipped one runs parallel to the edge instead).
    `_add_saved_profile_from_pts` only takes a 2-point (start, end) line, so
    a WSxM polyline with more than 2 points is truncated to its first two.
    """
    layers = sections.get("Graphic Layers", {})
    for value in layers.values():
        if "Name: Profiles" not in value and "Name:Profiles" not in value:
            continue
        pts = [(float(x), float(y)) for x, y in _PROFILE_POINT_RE.findall(value)]
        if len(pts) >= 2:
            return pts[:2]
    return None


def _member_caption(m: "_StpMember") -> Tuple[str, str, str]:
    """Return (channel_name, dir_tag, caption) for one `.stp` member.

    Shared by `_write_series` (per-channel export naming) and
    `_find_filter_matches` (pairing a filtered image back to its raw
    sibling) so the two can never disagree about which physical channel a
    member represents.
    """
    m_general = m.sections.get("General Info", {})
    # `Acquisition channel` is the authoritative channel name - the
    # filename-parsed one can be corrupted by WSxM's own uniquification
    # (confirmed on a real sample: the backward copy of "OC_M1_Freq._Shift"
    # was written to disk as "...Freq1._Shift...", which would otherwise
    # look like a distinct channel from its forward pair).
    channel_name = m_general.get("Acquisition channel", "").strip() or m.channel
    direction = m_general.get("X scanning direction", "").strip().lower()
    if direction.startswith("back"):
        return channel_name, "Bwd", f"{channel_name} (Bwd)"
    if direction.startswith("forw"):
        return channel_name, "Fwd", channel_name
    dir_tag = m.variant or "0"
    return channel_name, dir_tag, f"{channel_name} (#{dir_tag})"


# --------------------------------------------------------------------------- #
# Filter-match detection: pair a derived image back to its raw sibling and    #
# confirm which native filter (if any) reproduces it exactly.                 #
# --------------------------------------------------------------------------- #

# Each candidate is (name, callable). Verified against real data: WSxM's
# "Simple Flatten" matched `line_flatten:row:poly2` to ~1e-15 relative
# residual (machine precision) against its raw sibling in a real test
# session - not merely "similar", an exact reproduction. The others are
# untested against real WSxM exports so far but are included on the same
# basis (whole-frame or per-line background removal is the universe of
# things a "flatten"-style SPM filter does) so a newly-seen WSxM filter name
# has a chance of matching one without code changes; `_match_filter_candidate`
# only ever reports a match it has itself verified numerically, so an absent
# or wrong candidate here just means "no match found", never a false one.
_FILTER_CANDIDATES: List[Tuple[str, "callable"]] = [
    ("line_flatten:row:median", lambda a: native_filters.line_flatten_image(a, axis="row", method="median")),
    ("line_flatten:row:mean", lambda a: native_filters.line_flatten_image(a, axis="row", method="mean")),
    ("line_flatten:row:poly1", lambda a: native_filters.line_flatten_image(a, axis="row", method="poly1")),
    ("line_flatten:row:poly2", lambda a: native_filters.line_flatten_image(a, axis="row", method="poly2")),
    ("line_flatten:col:median", lambda a: native_filters.line_flatten_image(a, axis="col", method="median")),
    ("line_flatten:col:mean", lambda a: native_filters.line_flatten_image(a, axis="col", method="mean")),
    ("line_flatten:col:poly1", lambda a: native_filters.line_flatten_image(a, axis="col", method="poly1")),
    ("line_flatten:col:poly2", lambda a: native_filters.line_flatten_image(a, axis="col", method="poly2")),
    ("flatten_median:row", lambda a: native_filters.flatten_remove_median(a, axis="row")),
    ("flatten_median:col", lambda a: native_filters.flatten_remove_median(a, axis="col")),
    ("flatten_median:both", lambda a: native_filters.flatten_remove_median(a, axis="both")),
    ("plane_fit:1st_order", lambda a: native_filters.subtract_best_fit_plane(a)),
    ("plane_fit:2nd_order", lambda a: native_filters.subtract_2nd_order_plane(a)),
]

# Relative residual threshold for declaring a match. Real verified matches
# land near 1e-15 (float64 machine precision); genuinely wrong candidates in
# the same test landed at 0.19-0.26 - six orders of magnitude apart, so this
# has wide margin on both sides rather than being a knife-edge tuning.
_FILTER_MATCH_RELTOL = 1e-6

# Processing-history tokens that describe something other than a pixel-value
# transform (geometry, format) - never worth testing against pixel-filter
# candidates, and "zoom" specifically can never be reconstructed anyway
# since WSxM records no crop origin (see the adapter's module docstring
# history / CLAUDE.md).
_NON_FILTER_TOKENS = {"converted", "zoom"}


def _match_filter_candidate(raw_arr: np.ndarray, derived_arr: np.ndarray) -> Optional[Tuple[str, float]]:
    """Return (candidate_name, relative_residual) for the first native
    filter that reproduces `derived_arr` from `raw_arr` to near machine
    precision, or None. Allows a single global additive constant between
    the two (WSxM may re-center the display range), which is why the
    residual is de-medianed before scoring rather than compared directly.
    """
    if raw_arr.shape != derived_arr.shape:
        return None
    raw64 = raw_arr.astype(np.float64)
    derived64 = derived_arr.astype(np.float64)
    scale = float(np.nanstd(derived64))
    if not np.isfinite(scale) or scale <= 0:
        return None
    for name, fn in _FILTER_CANDIDATES:
        try:
            candidate = np.asarray(fn(raw64), dtype=np.float64)
        except Exception:
            continue
        if candidate.shape != derived64.shape:
            continue
        residual = derived64 - candidate
        residual = residual - np.nanmedian(residual)
        rel_err = float(np.nanstd(residual)) / scale
        if np.isfinite(rel_err) and rel_err < _FILTER_MATCH_RELTOL:
            return name, rel_err
    return None


def _member_provenance_key(m: "_StpMember") -> Optional[Tuple[str, str, str, str, str]]:
    """Identity of the underlying acquisition a member's pixel data came
    from - same source .sxm file, same timestamp, same channel, same scan
    direction - so a raw image and its filtered/derived sibling can be
    recognized as "the same acquisition" even though WSxM saved them as two
    independent, differently-named `.stp` files. Returns None when the
    embedded Nanonis header (see `_parse_embedded_nanonis_header`) is
    absent, since that provenance is the only reliable link available -
    native (non-Nanonis-origin) WSxM files can't be paired this way.
    """
    misc = m.sections.get("Miscellaneous", {})
    nanonis_fields = _parse_embedded_nanonis_header(misc.get("Comments", ""))
    scan_file = nanonis_fields.get("SCAN_FILE", "").strip()
    if not scan_file:
        return None
    rec_date = nanonis_fields.get("REC_DATE", "").strip()
    rec_time = nanonis_fields.get("REC_TIME", "").strip()
    channel_name, dir_tag, _ = _member_caption(m)
    return (scan_file, rec_date, rec_time, channel_name, dir_tag)


def _find_filter_matches(members: List["_StpMember"]) -> Dict[str, dict]:
    """Pair every derived image in this session back to its raw sibling and
    confirm the exact native filter that reproduces it, keyed by the
    derived member's own `str(path)`.

    Pairing key: same underlying acquisition (`_member_provenance_key`).
    Within a group sharing that key, the member with the *fewest*
    `Image processes` tokens is treated as the raw reference; any other
    member whose own token list equals the reference's plus exactly one
    trailing token is a one-step derivative of it, and that trailing token
    is the WSxM filter name responsible. Only WSxM-name tokens that aren't
    known non-filter markers (`_NON_FILTER_TOKENS`) get tested against the
    candidate registry - and even then, only a numerically-confirmed match
    is ever recorded (see `_match_filter_candidate`).
    """
    groups: Dict[Tuple[str, str, str, str, str], List["_StpMember"]] = {}
    for m in members:
        key = _member_provenance_key(m)
        if key is None:
            continue
        groups.setdefault(key, []).append(m)

    matches: Dict[str, dict] = {}
    for group in groups.values():
        if len(group) < 2:
            continue
        histories = []
        for m in group:
            hist_raw = m.sections.get("General Info", {}).get("Image processes", "")
            tokens = tuple(t.strip() for t in hist_raw.split(",") if t.strip())
            histories.append((m, tokens))
        histories.sort(key=lambda pair: len(pair[1]))
        raw_member, raw_tokens = histories[0]
        raw_arr = None
        for m, tokens in histories[1:]:
            if len(tokens) != len(raw_tokens) + 1 or tokens[: len(raw_tokens)] != raw_tokens:
                continue
            wsxm_name = tokens[-1]
            if wsxm_name.lower() in _NON_FILTER_TOKENS:
                continue
            try:
                if raw_arr is None:
                    raw_arr = _read_stp_payload(
                        raw_member.path, raw_member.rows, raw_member.cols,
                        raw_member.payload_offset, raw_member.dtype,
                    )
                derived_arr = _read_stp_payload(m.path, m.rows, m.cols, m.payload_offset, m.dtype)
            except Exception as exc:
                log(f"[WSxM] Failed reading payload while matching filter for {m.path.name}: {exc}")
                continue
            match = _match_filter_candidate(raw_arr, derived_arr)
            if match is None:
                continue
            native_name, rel_err = match
            _, _, raw_caption = _member_caption(raw_member)
            matches[str(m.path)] = {
                "wsxm_name": wsxm_name,
                "native_match": native_name,
                "rel_err": rel_err,
                "source_series": raw_member.series,
                "source_caption": raw_caption,
            }
            log(
                f"[WSxM] Confirmed '{wsxm_name}' on {m.path.name} == native "
                f"{native_name} (rel. residual {rel_err:.2e}), source: "
                f"{raw_member.series}/{raw_caption}"
            )
    return matches


# --------------------------------------------------------------------------- #
# Synthetic Omicron-style header emission                                     #
# --------------------------------------------------------------------------- #

def _write_series(
    cache_dir: Path,
    session_path: Path,
    series: str,
    members: List[_StpMember],
    filter_matches: Optional[Dict[str, dict]] = None,
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
        # Audit trail (e.g. "converted, plane, Simple Flatten" or "converted,
        # plane, zoom") - a history of *names*, not a parameter recipe, so it
        # is not by itself something to replay through sxm_viewer's own
        # filter pipeline. When a raw sibling of this image also exists in
        # the session, `_find_filter_matches` independently confirms (by
        # comparing actual pixel data, not by trusting this name) whether a
        # native filter reproduces the step exactly - see the
        # WsxmFilterMatch* fields below when that succeeds.
        "WsxmProcessingHistory": general.get("Image processes", ""),
    }
    _flatten_fields(header, control, prefix="WsxmControl:")
    _flatten_fields(header, general, prefix="WsxmGeneralInfo:")
    _flatten_fields(header, head_settings, prefix="WsxmHeadSettings:")
    if nanonis_fields:
        _flatten_fields(header, nanonis_fields, prefix="Nanonis:")

    exports: List[_ChannelExport] = []
    profile_channel_idx: Optional[int] = None
    profile_points_nm: Optional[List[Tuple[float, float]]] = None
    filter_match_idx: Optional[int] = None
    filter_match_info: Optional[dict] = None
    filter_matches = filter_matches or {}
    for m in members:
        try:
            arr = _read_stp_payload(m.path, m.rows, m.cols, m.payload_offset, m.dtype)
        except Exception as exc:
            log(f"[WSxM] Failed reading payload for {m.path.name}: {exc}")
            continue
        channel_name, dir_tag, caption = _member_caption(m)
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
        if profile_channel_idx is None:
            pixel_pts = _parse_profile_points(m.sections)
            if pixel_pts is not None:
                extent = spec_mapping.header_extent(header)
                angle = spec_mapping.header_scan_angle(header)
                # WSxM recorded these points against its own (unflipped)
                # raster; _read_stp_payload flips the array 180 degrees on
                # import, so the points need the identical flip to keep
                # pointing at the same physical feature.
                profile_points_nm = [
                    spec_mapping.pixel_to_nm(*_flip_pixel_180(px, py, cols, rows), extent, angle, cols, rows)
                    for px, py in pixel_pts
                ]
                profile_channel_idx = len(exports) - 1
        if filter_match_idx is None:
            info = filter_matches.get(str(m.path))
            if info is not None:
                filter_match_idx = len(exports) - 1
                filter_match_info = info

    if not exports:
        return None
    _dedupe_captions(exports)

    if profile_points_nm is not None:
        header["WsxmProfileChannel"] = profile_channel_idx
        header["WsxmProfilePoints"] = ";".join(
            f"{x:.6g},{y:.6g}" for x, y in profile_points_nm
        )

    if filter_match_info is not None:
        # Confirmed numerically at import time (see _find_filter_matches) -
        # not a guess from the WSxM filter's name alone. NativeFilter/Params
        # are parseable as "<function-key>:<axis>:<method>" so a future
        # consumer can call straight into processing/filters.py against the
        # source channel named here, reproducing this image live from raw
        # data instead of only showing WSxM's baked pixels.
        header["WsxmFilterMatchChannel"] = filter_match_idx
        header["WsxmFilterMatchWsxmName"] = filter_match_info["wsxm_name"]
        header["WsxmFilterMatchNative"] = filter_match_info["native_match"]
        header["WsxmFilterMatchRelError"] = f"{filter_match_info['rel_err']:.3e}"
        header["WsxmFilterMatchSourceSeries"] = filter_match_info["source_series"]
        header["WsxmFilterMatchSourceCaption"] = filter_match_info["source_caption"]
        header["WsxmFilterMatchSourceHeader"] = (
            f"{_safe_token(session_path.stem)}_{_safe_token(filter_match_info['source_series'])}_wsxm.txt"
        )

    # Prefixed with the *session's own* name unconditionally, not just on a
    # detected collision - two different .wsxm sessions in the same root
    # folder can each contain a scan they both happen to call e.g.
    # "NDST2-K1005" (confirmed: WSxM sessions are independent snapshots of
    # whatever images were open, so nothing stops two sessions from
    # overlapping). The header's *filename* is what the GUI actually shows
    # as the thumbnail caption/tooltip (`Path(header_path).name` in
    # gui/viewer/thumbnail_ui.py) and preview/popup titles, so collision-
    # proofing it here is what keeps two sessions' images from looking
    # identical in the grid - a cache-directory-only distinction (already in
    # place via _cache_dir_for) prevents data clobbering but is invisible to
    # the user.
    header_path = cache_dir / f"{_safe_token(session_path.stem)}_{_safe_token(series)}_wsxm.txt"
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
