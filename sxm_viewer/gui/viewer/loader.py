"""Loader helpers for SXMGridViewer."""
from __future__ import annotations

import time
from ..._shared import (
    QtCore,
    QtGui,
    QtWidgets,
    QIcon,
    QPixmap,
    QImage,
    QPainter,
    QPen,
    QBrush,
    FigureCanvas,
    Figure,
    Line2D,
    colormaps,
    np,
    Path,
    defaultdict,
    OrderedDict,
    datetime,
    hashlib,
    itertools,
    io,
    json,
    math,
    os,
    sys,
    threading,
    _scipy_ndimage,
    log_status,
    matplotlib,
)
from ...data.io import (
    parse_header,
    read_channel_file,
    normalize_unit_and_data,
    _split_key_value,
    _coerce_value,
    _canonical_header_key,
    _parse_inline_channels,
    _trailing_digits,
    _load_ascii_grid,
    _load_binary_grid,
    _load_tokenized_grid,
    _load_binary_with_inference,
    _binary_dtype_candidates,
)
from ...config import save_config
from ...data.spectroscopy import (
    parse_spectroscopy_file,
    SpectroscopyParseError,
    fit_parabola_bias,
    find_last_image_for_spec,
    _matrix_base_name,
    _rows_to_spec,
    _channel_labels,
    _clean_channel_label,
    _normalize_bias_axis,
    _extract_meta,
    _guess_index_from_name,
    _extract_section_value,
    _parse_section_metadata,
    _split_key_value,
    _split_tokens,
    _split_header_columns,
    _row_is_numeric,
    _normalize_meta_key,
    _coerce_value,
    _maybe_float,
    _maybe_int,
    _parse_datetime,
    _parse_date_and_time,
    _mtime,
    _read_text,
)
from ...data.matrix import MatrixDataset, parse_matrix_filename, matrix_dataset_key
from ...processing.detection import (
    _detect_dtype_for_file,
    _sample_channel_values_for_tagging,
    header_indicates_constant,
    _find_topography_channel,
    filedesc_indicates_current_or_topo,
)
from ...providers import convert_nanonis, convert_nanonis_files, parse_nanonis_spectroscopy, parse_nanonis_3ds
from ..detail_panels import SpectroscopyPopup, SpectroscopyCompareDialog


_SPECTRO_MANIFEST_VERSION = 3
_SPECTRO_CACHE_MTIME_TOLERANCE = 2.0
_SPECTRO_MANIFEST_FILE = "manifest_v2.json"
_SPECTRO_METADATA_ARRAY_KEYS = {
    "channels",
    "V",
    "data",
    "AltAxis",
}
_SPECTRO_METADATA_KEYS = {
    "path",
    "time",
    "x",
    "y",
    "source",
    "display_time",
    "available_channels",
    "trace_length",
    "unit_map",
    "AxisLabel",
    "AxisUnit",
    "AltAxisLabel",
    "AltAxisUnit",
    "order_idx",
    "matrix_dataset",
    "matrix_index",
    "grid_rows",
    "grid_cols",
    "grid_row",
    "grid_col",
    "channel_name",
    "channel_code",
    "image_key",
    "image_path",
    "primary_image_key",
    "shared_image_keys",
    "shared_repeat_assignment",
    "xy_stack_count",
    "xy_stack_display",
    "xy_stack_key",
    "xy_stack_summary",
    "xy_stack_z_varies",
    "xy_stack_z_level_nm",
    "xy_stack_z_label",
    "xy_stack_z_min_nm",
    "xy_stack_z_max_nm",
    "z_level_nm",
    "z_level_label",
    "z_level_unit",
}


def _fast_abs_path(path: Path | str | None) -> str:
    try:
        raw = os.fspath(path)
    except Exception:
        raw = str(path or "")
    try:
        return os.path.abspath(raw)
    except Exception:
        return raw


def _normalize_spectro_path_key(path: Path | str | None) -> str:
    key = _fast_abs_path(path)
    return key.lower() if os.name == "nt" else key


def _spectro_relative_key(base_folder: Path | None, filepath: Path) -> str:
    if base_folder is not None:
        try:
            return os.path.relpath(_fast_abs_path(filepath), _fast_abs_path(base_folder)).replace("\\", "/")
        except Exception:
            pass
    return filepath.name


def _discover_spectro_file_records(folder_path: Path | None, files) -> list[dict]:
    records = []
    seen = set()
    valid_exts = {".dat", ".3ds"}
    if files:
        for f in files:
            if not f:
                continue
            try:
                p = Path(f)
            except Exception:
                continue
            suffix = p.suffix.lower()
            if suffix not in valid_exts:
                continue
            norm_key = _normalize_spectro_path_key(p)
            if norm_key in seen:
                continue
            seen.add(norm_key)
            try:
                st = p.stat()
                mtime = float(st.st_mtime)
                size = int(st.st_size)
            except Exception:
                mtime = 0.0
                size = -1
            records.append(
                {
                    "path": p,
                    "norm_key": norm_key,
                    "rel_key": _spectro_relative_key(folder_path, p),
                    "suffix": suffix,
                    "mtime": mtime,
                    "size": size,
                }
            )
        records.sort(key=lambda rec: str(rec["path"]).lower())
        return records
    if folder_path is None:
        return records
    try:
        base_abs = _fast_abs_path(folder_path)
        with os.scandir(base_abs) as it:
            for entry in it:
                try:
                    if not entry.is_file():
                        continue
                except Exception:
                    continue
                suffix = Path(entry.name).suffix.lower()
                if suffix not in valid_exts:
                    continue
                full_path = Path(entry.path)
                norm_key = _normalize_spectro_path_key(full_path)
                if norm_key in seen:
                    continue
                seen.add(norm_key)
                try:
                    st = entry.stat()
                    mtime = float(st.st_mtime)
                    size = int(st.st_size)
                except Exception:
                    mtime = 0.0
                    size = -1
                records.append(
                    {
                        "path": full_path,
                        "norm_key": norm_key,
                        "rel_key": entry.name.replace("\\", "/"),
                        "suffix": suffix,
                        "mtime": mtime,
                        "size": size,
                    }
                )
    except Exception:
        try:
            fallback_files = []
            for pat in ("*.dat", "*.DAT", "*.3ds", "*.3DS"):
                fallback_files.extend(folder_path.glob(pat))
            return _discover_spectro_file_records(folder_path, fallback_files)
        except Exception:
            return []
    records.sort(key=lambda rec: str(rec["path"]).lower())
    return records


def _spectro_payload_cache_key(base_folder: Path | None, filepath: Path, *, legacy: bool = False) -> str:
    source = filepath.name if legacy else _spectro_relative_key(base_folder, filepath)
    return hashlib.md5(source.encode("utf-8")).hexdigest()[:16]


def _spectro_payload_paths(cache_dir: Path, base_folder: Path | None, filepath: Path, *, legacy: bool = False):
    cache_key = _spectro_payload_cache_key(base_folder, filepath, legacy=legacy)
    return (
        cache_dir / f"{cache_key}_meta.json",
        cache_dir / f"{cache_key}_data.npy",
        cache_key,
    )


def _serialize_cache_value(value):
    if isinstance(value, datetime):
        return {"__datetime__": value.isoformat()}
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, np.ndarray):
        return None
    if isinstance(value, dict):
        out = {}
        for key, item in value.items():
            serialized = _serialize_cache_value(item)
            if serialized is not None:
                out[str(key)] = serialized
        return out
    if isinstance(value, (list, tuple)):
        out = []
        for item in value:
            serialized = _serialize_cache_value(item)
            if serialized is not None:
                out.append(serialized)
        return out
    return value


def _deserialize_cache_value(value):
    if isinstance(value, dict):
        if set(value.keys()) == {"__datetime__"}:
            try:
                return datetime.fromisoformat(str(value["__datetime__"]))
            except Exception:
                return None
        return {key: _deserialize_cache_value(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_deserialize_cache_value(item) for item in value]
    return value


def _spec_channel_names(spec) -> list[str]:
    names = []
    channels = spec.get("channels") or {}
    if isinstance(channels, dict):
        names.extend(str(name) for name in channels.keys() if str(name).strip())
    if not names:
        names.extend(str(name) for name in (spec.get("available_channels") or []) if str(name).strip())
    deduped = []
    seen = set()
    for name in names:
        if name in seen:
            continue
        seen.add(name)
        deduped.append(name)
    return deduped


def _spec_points_per_trace(spec) -> int | None:
    trace_length = spec.get("trace_length")
    if trace_length is not None:
        try:
            return int(trace_length)
        except Exception:
            pass
    vals = spec.get("V")
    if vals is not None:
        try:
            arr = np.asarray(vals)
            if arr.size:
                return int(arr.size)
        except Exception:
            pass
    channels = spec.get("channels") or {}
    if isinstance(channels, dict):
        for values in channels.values():
            try:
                arr = np.asarray(values)
                if arr.size:
                    return int(arr.size)
            except Exception:
                continue
    return None


def _spec_has_payload_data(spec) -> bool:
    try:
        vals = spec.get("V")
        if vals is not None and np.asarray(vals).size:
            return True
    except Exception:
        pass
    try:
        alt_vals = spec.get("AltAxis")
        if alt_vals is not None and np.asarray(alt_vals).size:
            return True
    except Exception:
        pass
    channels = spec.get("channels") or {}
    if isinstance(channels, dict):
        for values in channels.values():
            try:
                if values is not None and np.asarray(values).size:
                    return True
            except Exception:
                continue
    for axis in spec.get("AxisChoices") or []:
        if not isinstance(axis, dict):
            continue
        try:
            if np.asarray(axis.get("values", [])).size:
                return True
        except Exception:
            continue
    return False


def _payload_specs_usable(specs) -> bool:
    specs = list(specs or [])
    if not specs:
        return False
    return any(_spec_has_payload_data(spec) for spec in specs)


def _spec_metadata_entry(spec):
    entry = {}
    spec_dict = dict(spec or {})
    for key in _SPECTRO_METADATA_KEYS:
        if key not in spec_dict:
            continue
        value = spec_dict.get(key)
        if key == "AxisChoices":
            axes = []
            for axis in value or []:
                if not isinstance(axis, dict):
                    continue
                axis_meta = {}
                for axis_key in ("key", "label", "unit"):
                    axis_val = axis.get(axis_key)
                    if axis_val not in (None, ""):
                        axis_meta[axis_key] = axis_val
                if axis_meta:
                    axes.append(axis_meta)
            if axes:
                entry["AxisChoices"] = axes
            continue
        serialized = _serialize_cache_value(value)
        if serialized is not None:
            entry[key] = _deserialize_cache_value(serialized)
    axis_choices = spec_dict.get("AxisChoices")
    if axis_choices:
        axes = []
        for axis in axis_choices or []:
            if not isinstance(axis, dict):
                continue
            axis_meta = {}
            for axis_key in ("key", "label", "unit"):
                axis_val = axis.get(axis_key)
                if axis_val not in (None, ""):
                    axis_meta[axis_key] = axis_val
            if axis_meta:
                axes.append(axis_meta)
        if axes:
            entry["AxisChoices"] = axes
    entry["available_channels"] = _spec_channel_names(spec)
    trace_length = _spec_points_per_trace(spec)
    if trace_length is not None:
        entry["trace_length"] = int(trace_length)
    entry["_payload_hydrated"] = False
    return entry


def _spec_metadata_list(specs):
    return [_spec_metadata_entry(spec) for spec in list(specs or [])]


def _restore_spec_metadata(entry):
    spec = dict(_deserialize_cache_value(entry or {}))
    spec["_payload_hydrated"] = bool(spec.get("_payload_hydrated", False))
    return spec


def _restore_spectro_payload_specs(entries):
    specs = []
    for entry_dict in entries or []:
        spec = dict(entry_dict)
        if "channels" in spec and isinstance(spec["channels"], dict):
            spec["channels"] = {k: np.array(v) for k, v in spec["channels"].items()}
        if "V" in spec:
            spec["V"] = np.array(spec["V"])
        spec["available_channels"] = _spec_channel_names(spec)
        trace_length = _spec_points_per_trace(spec)
        if trace_length is not None:
            spec["trace_length"] = int(trace_length)
        spec["_payload_hydrated"] = True
        specs.append(spec)
    return specs


def _clone_payload_spec_entry(spec):
    clone = dict(spec or {})
    channels = spec.get("channels")
    if isinstance(channels, dict):
        clone["channels"] = {
            key: (np.array(value) if value is not None else value)
            for key, value in channels.items()
        }
    if "V" in spec:
        try:
            clone["V"] = np.array(spec["V"])
        except Exception:
            clone["V"] = spec.get("V")
    axis_choices = spec.get("AxisChoices")
    if isinstance(axis_choices, (list, tuple)):
        rebuilt = []
        for axis in axis_choices:
            if not isinstance(axis, dict):
                continue
            axis_clone = dict(axis)
            if "values" in axis:
                try:
                    axis_clone["values"] = np.array(axis.get("values"))
                except Exception:
                    axis_clone["values"] = axis.get("values")
            rebuilt.append(axis_clone)
        clone["AxisChoices"] = rebuilt
    clone["available_channels"] = _spec_channel_names(clone)
    clone["_payload_hydrated"] = True
    return clone


def _spec_identity_token(spec):
    base = _normalize_spectro_path_key(spec.get("path") or "")
    matrix_index = spec.get("matrix_index")
    if matrix_index is not None:
        return f"{base}#idx:{matrix_index}"
    x_val = spec.get("x")
    y_val = spec.get("y")
    if x_val is not None or y_val is not None:
        try:
            return f"{base}#pos:{round(float(x_val), 6)}:{round(float(y_val), 6)}"
        except Exception:
            return f"{base}#pos:{x_val}:{y_val}"
    order_idx = spec.get("order_idx")
    if order_idx is not None:
        return f"{base}#order:{order_idx}"
    return base


def _merge_payload_into_spec(target, payload):
    if target is payload:
        return target
    preserved = {
        "image_key": target.get("image_key"),
        "image_path": target.get("image_path"),
        "primary_image_key": target.get("primary_image_key"),
        "shared_image_keys": target.get("shared_image_keys"),
        "shared_repeat_assignment": target.get("shared_repeat_assignment"),
        "display_time": target.get("display_time"),
        "xy_stack_summary": target.get("xy_stack_summary"),
        "xy_stack_display": target.get("xy_stack_display"),
        "xy_stack_count": target.get("xy_stack_count"),
        "xy_stack_key": target.get("xy_stack_key"),
        "xy_stack_z_level_nm": target.get("xy_stack_z_level_nm"),
    }
    target.clear()
    target.update(payload)
    for key, value in preserved.items():
        if value not in (None, "", []):
            target[key] = value
    target["available_channels"] = _spec_channel_names(target)
    trace_length = _spec_points_per_trace(target)
    if trace_length is not None:
        target["trace_length"] = int(trace_length)
    target["_payload_hydrated"] = True
    return target


def _load_spectro_manifest(cache_dir: Path | None):
    if cache_dir is None:
        return {}
    manifest_path = cache_dir / _SPECTRO_MANIFEST_FILE
    if not manifest_path.exists():
        return {}
    try:
        with open(manifest_path, "r", encoding="utf-8") as handle:
            payload = json.load(handle)
        if int(payload.get("version", 0)) != _SPECTRO_MANIFEST_VERSION:
            return {}
        entries = payload.get("entries")
        if not isinstance(entries, dict):
            return {}
        return entries
    except Exception:
        return {}


def _save_spectro_manifest(cache_dir: Path | None, manifest_entries: dict):
    if cache_dir is None:
        return
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        manifest_path = cache_dir / _SPECTRO_MANIFEST_FILE
        payload = {
            "version": _SPECTRO_MANIFEST_VERSION,
            "saved_at": datetime.utcnow().isoformat(),
            "entries": manifest_entries,
        }
        with open(manifest_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
    except Exception:
        pass


def save_spectro_manifest_snapshot(folder: Path | None, manifest_entries: dict):
    try:
        folder_path = Path(folder) if folder else None
    except Exception:
        folder_path = None
    if folder_path is None:
        return
    cache_dir = folder_path / ".sxmviewer_spectro_cache"
    _save_spectro_manifest(cache_dir, manifest_entries or {})


def _manifest_entry_valid(entry, *, mtime: float, fsize: int) -> bool:
    if not isinstance(entry, dict):
        return False
    meta_mtime = entry.get("mtime")
    meta_size = entry.get("size")
    if meta_mtime is None or meta_size is None:
        return False
    try:
        drift = abs(float(meta_mtime) - float(mtime))
    except Exception:
        return False
    try:
        size_ok = int(meta_size) == int(fsize)
    except Exception:
        return False
    return drift <= _SPECTRO_CACHE_MTIME_TOLERANCE and size_ok


def collect_folder_image_paths(viewer, folder: Path) -> list[Path]:
    """Return the image header files implied by a folder load."""
    folder = Path(folder)
    txts = sorted(folder.glob("*.txt"))
    if getattr(viewer, "convert_nanonis_enabled", True):
        converted = convert_nanonis(folder)
        if converted:
            txts = sorted(list(txts) + list(converted), key=lambda p: str(p).lower())
            log_status(f"Converted {len(converted)} Nanonis scan(s)")
    else:
            log_status("Skipping Nanonis .sxm conversion (disabled in config)")
    return txts


def classify_dropped_paths(viewer, paths):
    """Split explicit drops into image headers and spectroscopy files."""
    image_paths = []
    spectro_paths = []
    seen = set()
    for raw in paths or []:
        path = Path(raw)
        if not path.exists() or path.is_dir():
            continue
        try:
            key = str(path.resolve()).lower()
        except Exception:
            key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        suffix = path.suffix.lower()
        if suffix in {".dat", ".3ds"}:
            spectro_paths.append(path)
            continue
        if suffix == ".sxm":
            image_paths.append(path)
            continue
        if suffix == ".txt":
            try:
                _header, fds = parse_header(path)
            except Exception:
                fds = []
            header_like = bool(fds) and any(
                str(fd.get("FileName") or "").strip() for fd in fds
            )
            if header_like:
                image_paths.append(path)
            else:
                spectro_paths.append(path)
    return image_paths, spectro_paths


def _collect_explicit_image_paths(viewer, paths) -> list[Path]:
    """Return header files for an explicit file drop without scanning the folder."""
    collected: list[Path] = []
    sxm_paths = []
    seen = set()
    for raw in paths or []:
        path = Path(raw)
        if not path.exists() or path.is_dir():
            continue
        try:
            key = str(path.resolve()).lower()
        except Exception:
            key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        if path.suffix.lower() == ".txt":
            collected.append(path)
        elif path.suffix.lower() == ".sxm":
            sxm_paths.append(path)
    if sxm_paths:
        if getattr(viewer, "convert_nanonis_enabled", True):
            converted = convert_nanonis_files(sxm_paths)
            if converted:
                collected.extend(converted)
        else:
            log_status("Skipping Nanonis .sxm conversion (disabled in config)")
    return sorted(collected, key=lambda p: str(p).lower())


def load_files(
    viewer,
    files,
    folder_hint: Path | None = None,
    source_label: str = "files",
    *,
    append: bool = False,
    refresh_spectros: bool = True,
):
    files = [Path(p) for p in (files or []) if p]
    if not files:
        return
    log_status(f"Loading {source_label}: {len(files)} file(s)")
    t0 = time.perf_counter()
    viewer._update_toolbar_actions(False)
    prev_last_dir = getattr(viewer, 'last_dir', None)
    existing_files = list(getattr(viewer, "files", []) or []) if append else []
    existing_headers = dict(getattr(viewer, "headers", {}) or {}) if append else None
    folder = Path(folder_hint) if folder_hint is not None else None
    if folder is None:
        parents = {p.parent for p in files if p.parent}
        if len(parents) == 1:
            folder = next(iter(parents))
        # Mixed-folder drops should not force a synthetic root path into the UI.
    if folder is not None:
        viewer.last_dir = folder
        try:
            viewer.path_le.setText(str(folder))
        except Exception:
            pass
        viewer.config['last_dir'] = str(folder)
        viewer._record_recent_dir(folder)

    txts = _collect_explicit_image_paths(viewer, files)
    log_status(f"Found {len(txts)} header file(s)")
    if append:
        merged_files = []
        seen = set()
        for path in existing_files + txts:
            key = str(path)
            if key in seen:
                continue
            seen.add(key)
            merged_files.append(path)
        viewer.files = merged_files
        viewer.headers = existing_headers if existing_headers is not None else viewer.headers
    else:
        viewer.files = txts
        viewer.headers.clear()
        viewer._invalidate_thumbnail_cache()
        viewer._invalidate_channel_cache()
        viewer.thumb_multi_select = set()
    cache_hits = 0
    cache_miss = 0
    for t in txts:
        key = str(t)
        if append and key in (existing_headers or {}):
            continue
        cached = viewer._get_cached_header(t)
        if cached:
            hdr, fds = cached
            cache_hits += 1
        else:
            try:
                hdr, fds = parse_header(t)
                cache_miss += 1
                viewer._store_header_cache(t, hdr, fds)
            except Exception:
                continue
        viewer.headers[key] = (hdr, fds)
    if cache_miss:
        viewer._save_header_cache()
    log_status(f"Headers loaded (hits={cache_hits}, miss={cache_miss})")
    if not viewer.headers:
        viewer.meta_box.setPlainText("No valid .txt headers found")
        viewer.clear_thumbs(); return
    viewer._build_image_timestamp_index()
    viewer._rebuild_frame_map_entries()
    t_headers = time.perf_counter()

    # build channel dropdown from first header
    first_key = next(iter(viewer.headers))
    _, first_fds = viewer.headers[first_key]
    labels = []
    for idx, fd in enumerate(first_fds):
        cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
        labels.append(f"{idx}: {cap}")
    max_channels = max(len(v[1]) for v in viewer.headers.values())
    if max_channels > len(labels):
        for idx in range(len(labels), max_channels):
            labels.append(f"{idx}: chan{idx}")

    viewer.channel_dropdown.blockSignals(True)
    viewer.channel_dropdown.clear()
    for lab in labels:
        viewer.channel_dropdown.addItem(lab)
        viewer.channel_dropdown.setItemData(viewer.channel_dropdown.count()-1, lab, QtCore.Qt.ToolTipRole)
    viewer.channel_dropdown.setMinimumWidth(240)
    if 0 <= viewer.last_channel_index < viewer.channel_dropdown.count():
        viewer.channel_dropdown.setCurrentIndex(viewer.last_channel_index)
    else:
        viewer.last_channel_index = 0; viewer.channel_dropdown.setCurrentIndex(0)
    viewer.channel_dropdown.setEnabled(True)
    viewer.channel_dropdown.blockSignals(False)
    try:
        viewer._sync_channel_nav_buttons()
    except Exception:
        pass

    # set cmaps
    try: viewer.thumb_cmap_combo.setCurrentText(viewer.thumb_cmap)
    except: pass
    try: viewer.preview_cmap_combo.setCurrentText(viewer.preview_cmap)
    except: pass
    # set icon sizes for cmap combos
    try:
        viewer.thumb_cmap_combo.setIconSize(QtCore.QSize(96, 14))
        viewer.preview_cmap_combo.setIconSize(QtCore.QSize(96, 14))
    except Exception:
        pass

    # auto-detect tags for files not already tagged (can be disabled via config)
    should_tag = getattr(viewer, "auto_detect_tags", True)
    if should_tag:
        log_status("Auto-detecting tags...")
        viewer._auto_detect_tags_for_folder()
    else:
        log_status("Skipping auto-detect tags (disabled in config)")
    t_tags = time.perf_counter()

    if refresh_spectros:
        # Keep spectroscopy folder aligned only when opening a real folder.
        if folder is not None:
            try:
                spec_path = Path(getattr(viewer, 'spec_folder_path', folder))
            except Exception:
                spec_path = folder
            auto_follow = False
            if not spec_path.exists():
                auto_follow = True
            elif prev_last_dir and spec_path.resolve() == Path(prev_last_dir).resolve():
                auto_follow = True
            if auto_follow:
                viewer.spec_folder_path = folder
                viewer.config['spectra_folder'] = str(folder)
                save_config(viewer.config)
                try:
                    viewer.spec_folder_le.setText(str(folder))
                except Exception:
                    pass

        defer_spectros = bool(
            getattr(viewer, "lazy_spectros_enabled", False)
            and source_label == "folder"
        )
        if defer_spectros:
            log_status("Deferring spectroscopy references until after initial folder load...")
            viewer._spectros_loaded = False
            viewer._spectros_pending = True
            viewer._spectros_loading = False
            viewer.spectros = []
            viewer.matrix_spectros = []
            viewer.matrix_datasets = {}
            viewer.spectros_by_image = defaultdict(list)
            try:
                viewer._clear_multi_spec_selection()
            except Exception:
                pass
            try:
                viewer._update_spectro_stats_label()
            except Exception:
                pass
            if hasattr(viewer, "_schedule_pending_spectro_load"):
                try:
                    viewer._schedule_pending_spectro_load(delay_ms=1200)
                except Exception:
                    pass
        else:
            log_status("Loading spectroscopy references...")
            viewer._spectros_pending = False
            viewer._reload_spectros(refresh=False)
        t_specs = time.perf_counter()
    else:
        log_status("Skipping spectroscopy reload for explicit file drop")
        t_specs = time.perf_counter()

    QtCore.QTimer.singleShot(0, lambda: viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex()))
    log_status(f"{source_label.capitalize()} load complete.")
    log_status(
        f"[Perf] Load stages: headers { (t_headers - t0)*1000:.0f} ms | "
        f"tags { (t_tags - t_headers)*1000:.0f} ms | "
        f"spectros { (t_specs - t_tags)*1000:.0f} ms | "
        f"total { (t_specs - t0)*1000:.0f} ms"
    )


def load_folder(viewer, folder:Path):
    folder = Path(folder)
    log_status(f"Loading folder: {folder}")
    files = collect_folder_image_paths(viewer, folder)
    return load_files(viewer, files, folder_hint=folder, source_label="folder")


def load_spectroscopy_files(viewer, files, folder_hint: Path | None = None, *, append: bool = True, refresh: bool = True):
    files = [Path(p) for p in (files or []) if p]
    if not files:
        return
    log_status(f"Loading spectroscopy files: {len(files)} file(s)")
    prev_specs = list(getattr(viewer, "spectros", []) or []) if append else []
    prev_matrix = dict(getattr(viewer, "matrix_datasets", {}) or {}) if append else {}
    t0 = time.perf_counter()
    viewer._update_toolbar_actions(False)
    if folder_hint is not None:
        try:
            viewer.last_dir = Path(folder_hint)
            viewer.config["last_dir"] = str(viewer.last_dir)
            viewer._record_recent_dir(viewer.last_dir)
        except Exception:
            pass
    scan_folder = Path(folder_hint) if folder_hint is not None else None
    if scan_folder is not None and not scan_folder.exists():
        scan_folder = None
    new_specs, spec_stats = _scan_spectros(
        viewer,
        scan_folder,
        files=files,
        image_paths=[str(p) for p in getattr(viewer, "files", []) or []],
        image_meta=getattr(viewer, "image_meta", None),
        use_disk_cache=False,
    )
    if append:
        merged_specs = prev_specs + list(new_specs or [])
    else:
        merged_specs = list(new_specs or [])
    # Rebuild the session cache from the merged spectroscopy set.
    viewer.spectros = merged_specs
    merged_matrix = dict(prev_matrix)
    for key, ds in (getattr(viewer, "matrix_datasets", {}) or {}).items():
        if key in merged_matrix:
            try:
                merged_matrix[key].channels.extend(list(getattr(ds, "channels", []) or []))
            except Exception:
                merged_matrix[key] = ds
        else:
            merged_matrix[key] = ds
    viewer.matrix_datasets = merged_matrix
    viewer._spectros_loaded = False
    viewer._spectros_pending = False
    viewer._assign_spectros_to_images()
    try:
        refresh_spectro_manifest_from_viewer(viewer)
    except Exception:
        pass
    viewer.matrix_spectros = [spec for spec in viewer.spectros if spec.get('matrix_index') is not None]
    viewer._update_spectro_stats_label(spec_stats)
    viewer._spectros_loaded = True
    viewer._update_matrix_summary_banner()
    if refresh:
        viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        if viewer.last_preview:
            try:
                viewer.show_file_channel(viewer.last_preview[0], viewer.last_preview[1])
            except Exception:
                pass
    scan_ms = (time.perf_counter() - t0) * 1000.0
    log_status(f"[Perf] Spectroscopy files: {scan_ms:.0f} ms | appended={append}")
    return merged_specs


def _parse_header_datetime(viewer, header, path: Path | str | None = None):
    """Return a sortable key (float timestamp) parsed from header Date/Time if possible; otherwise 0.0.
    Accepts common formats and uses file mtime as a tie-breaker for ambiguous day/month formats."""
    try:
        if path and getattr(viewer, "image_time_source", None) == "mtime":
            try:
                return Path(path).stat().st_mtime
            except Exception:
                pass
        date = str(header.get('Date', '') or '').strip()
        time = str(header.get('Time', '') or '').strip()
        if not date and not time:
            return 0.0
        candidates = []
        if date and time:
            candidates.append(f"{date} {time}")
        if date:
            candidates.append(date)
        fmts = [
            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M:%S',
            '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %H:%M', '%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M',
            '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M',
            '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'
        ]
        parsed = []
        for s in candidates:
            for fmt in fmts:
                try:
                    dt = datetime.strptime(s, fmt)
                    parsed.append(dt)
                except Exception:
                    continue
        if not parsed:
            return 0.0
        file_dt = None
        if path:
            try:
                file_dt = datetime.fromtimestamp(Path(path).stat().st_mtime)
            except Exception:
                file_dt = None
        if file_dt:
            parsed.sort(key=lambda dt: abs((dt - file_dt).total_seconds()))
        return parsed[0].timestamp()
    except Exception:
        return 0.0


def _load_spectro_disk_payload(cache_dir: Path | None, base_folder: Path | None, filepath: Path, mtime: float, fsize: int):
    if not cache_dir or not cache_dir.exists():
        return None, None
    candidate_paths = [
        _spectro_payload_paths(cache_dir, base_folder, filepath, legacy=False),
        _spectro_payload_paths(cache_dir, base_folder, filepath, legacy=True),
    ]
    seen = set()
    for meta_file, data_file, cache_key in candidate_paths:
        token = (str(meta_file), str(data_file))
        if token in seen:
            continue
        seen.add(token)
        if not meta_file.exists() or not data_file.exists():
            continue
        try:
            with open(meta_file, "r", encoding="utf-8") as handle:
                cache_meta = json.load(handle)
            if not _manifest_entry_valid(cache_meta, mtime=mtime, fsize=fsize):
                continue
            cache_data = np.load(data_file, allow_pickle=True)
            specs = _restore_spectro_payload_specs(cache_data)
            return specs, {
                "payload_meta": meta_file.name,
                "payload_data": data_file.name,
                "cache_key": cache_key,
            }
        except Exception:
            continue
    return None, None


def _store_spectro_disk_payload(cache_dir: Path | None, base_folder: Path | None, filepath: Path, mtime: float, fsize: int, specs):
    if not cache_dir:
        return None
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        meta_file, data_file, cache_key = _spectro_payload_paths(cache_dir, base_folder, filepath, legacy=False)
        cache_meta = {
            "filename": filepath.name,
            "mtime": float(mtime),
            "size": int(fsize),
            "cached_at": datetime.utcnow().isoformat(),
            "spec_count": len(specs or []),
            "relative_path": _spectro_relative_key(base_folder, filepath),
        }
        with open(meta_file, "w", encoding="utf-8") as handle:
            json.dump(cache_meta, handle, indent=2)
        serializable_specs = []
        for spec in specs or []:
            entry = dict(spec)
            if "channels" in entry and isinstance(entry["channels"], dict):
                entry["channels"] = {
                    key: (value.tolist() if hasattr(value, "tolist") else value)
                    for key, value in entry["channels"].items()
                }
            if "V" in entry and hasattr(entry["V"], "tolist"):
                entry["V"] = entry["V"].tolist()
            serializable_specs.append(entry)
        np.save(data_file, np.array(serializable_specs, dtype=object), allow_pickle=True)
        return {
            "payload_meta": meta_file.name,
            "payload_data": data_file.name,
            "cache_key": cache_key,
        }
    except Exception as exc:
        try:
            log_status(f"  - spectroscopy cache store failed: {exc}")
        except Exception:
            pass
        return None


def _build_manifest_entry(base_folder: Path | None, filepath: Path, mtime: float, fsize: int, specs, payload_info=None):
    info = payload_info or {}
    return {
        "relpath": _spectro_relative_key(base_folder, filepath),
        "filename": filepath.name,
        "extension": filepath.suffix.lower(),
        "mtime": float(mtime),
        "size": int(fsize),
        "payload_meta": info.get("payload_meta"),
        "payload_data": info.get("payload_data"),
        "cache_key": info.get("cache_key"),
        "spec_count": len(specs or []),
        "source_type": str((specs or [{}])[0].get("source") or ""),
        "specs": [_serialize_cache_value(spec) for spec in _spec_metadata_list(specs or [])],
    }


def _parse_spectro_file_payload(filepath: Path, mtime: float):
    spec_list = None
    parse_error = None
    ext = filepath.suffix.lower()
    if ext == ".dat":
        try:
            spec_list = parse_nanonis_spectroscopy(filepath)
        except Exception:
            spec_list = None
    elif ext == ".3ds":
        try:
            spec_list = parse_nanonis_3ds(filepath)
        except Exception:
            spec_list = None
    if spec_list is None and ext not in (".3ds",):
        try:
            spec_list = parse_spectroscopy_file(filepath)
        except SpectroscopyParseError as exc:
            parse_error = exc
            spec_list = None
        except Exception:
            spec_list = None
    if not spec_list:
        return spec_list, parse_error
    if ext == ".dat" and any(s.get("x") is None or s.get("y") is None for s in spec_list):
        try:
            with open(filepath, "r", encoding="latin-1") as handle:
                for _ in range(30):
                    line = handle.readline()
                    if not line:
                        break
                    if "x/y-Pos:" not in line:
                        continue
                    parts = line.split(":", 1)[1].strip().split("/")
                    if len(parts) != 2:
                        break
                    try:
                        x_nm = _coerce_pos_to_nm(float(parts[0]))
                        y_nm = _coerce_pos_to_nm(float(parts[1]))
                        for spec in spec_list:
                            if spec.get("x") is None:
                                spec["x"] = x_nm
                            if spec.get("y") is None:
                                spec["y"] = y_nm
                    except ValueError:
                        pass
                    break
        except Exception:
            pass
    for spec in spec_list or []:
        if "path" not in spec or not spec.get("path"):
            spec["path"] = str(filepath)
        time_value = spec.get("time")
        use_mtime = False
        if time_value is None:
            use_mtime = True
        elif isinstance(time_value, datetime):
            if time_value.year < 1990:
                use_mtime = True
        elif isinstance(time_value, (int, float)):
            try:
                spec["time"] = datetime.fromtimestamp(float(time_value))
            except Exception:
                use_mtime = True
        elif isinstance(time_value, str):
            if not time_value.strip():
                use_mtime = True
            else:
                try:
                    spec["time"] = datetime.fromisoformat(time_value)
                except Exception:
                    use_mtime = True
        if use_mtime:
            try:
                spec["time"] = datetime.fromtimestamp(mtime)
            except Exception:
                pass
        spec["available_channels"] = _spec_channel_names(spec)
        trace_length = _spec_points_per_trace(spec)
        if trace_length is not None:
            spec["trace_length"] = int(trace_length)
        spec["_payload_hydrated"] = True
    return spec_list, parse_error


def hydrate_spectro_file(viewer, spec_or_path, *, log_perf: bool = True, return_stage: bool = False):
    if not spec_or_path:
        return None
    try:
        filepath = Path(spec_or_path.get("path")) if isinstance(spec_or_path, dict) else Path(spec_or_path)
    except Exception:
        return None
    if not filepath:
        return None
    norm_key = _normalize_spectro_path_key(filepath)
    try:
        st = filepath.stat()
        mtime = st.st_mtime
        fsize = st.st_size
    except Exception:
        mtime = 0.0
        fsize = -1
    cached = (getattr(viewer, "_spectro_cache", {}) or {}).get(norm_key)
    full_specs = None
    hydrate_stage = "memory"
    t_hydrate = time.perf_counter()
    if cached and not cached.get("deferred"):
        try:
            if abs(float(cached.get("mtime", 0.0)) - float(mtime)) <= _SPECTRO_CACHE_MTIME_TOLERANCE:
                candidate_specs = [_clone_payload_spec_entry(spec) for spec in (cached.get("data") or [])]
                if _payload_specs_usable(candidate_specs):
                    full_specs = candidate_specs
                else:
                    viewer._spectro_cache.pop(norm_key, None)
            else:
                viewer._spectro_cache.pop(norm_key, None)
        except Exception:
            full_specs = None
    folder = getattr(viewer, "spec_folder_path", None) or getattr(viewer, "last_dir", None)
    try:
        folder = Path(folder) if folder else filepath.parent
    except Exception:
        folder = filepath.parent
    disk_cache_dir = None
    if getattr(viewer, "spectro_disk_cache_enabled", True):
        try:
            disk_cache_dir = folder / ".sxmviewer_spectro_cache"
        except Exception:
            disk_cache_dir = None
    if full_specs is None:
        hydrate_stage = "disk"
        full_specs, payload_info = _load_spectro_disk_payload(disk_cache_dir, folder, filepath, mtime, fsize)
        if full_specs is not None and not _payload_specs_usable(full_specs):
            full_specs = None
        if full_specs is None:
            hydrate_stage = "parse"
            full_specs, parse_error = _parse_spectro_file_payload(filepath, mtime)
            if parse_error is not None:
                raise parse_error
            if not full_specs:
                return (None, hydrate_stage) if return_stage else None
            payload_info = _store_spectro_disk_payload(disk_cache_dir, folder, filepath, mtime, fsize, full_specs)
        if full_specs:
            viewer._spectro_cache[norm_key] = {
                "mtime": float(mtime),
                "data": [_clone_payload_spec_entry(spec) for spec in full_specs],
            }
            if getattr(viewer, "spectro_manifest_cache_enabled", True):
                manifest = getattr(viewer, "_spectro_manifest_entries", {}) or {}
                rel_key = _spectro_relative_key(folder, filepath)
                manifest[rel_key] = _build_manifest_entry(folder, filepath, mtime, fsize, full_specs, payload_info=payload_info)
                viewer._spectro_manifest_entries = manifest
                if hasattr(viewer, "_schedule_spectro_manifest_save"):
                    viewer._schedule_spectro_manifest_save()
                else:
                    _save_spectro_manifest(disk_cache_dir, manifest)
    elapsed_ms = (time.perf_counter() - t_hydrate) * 1000.0
    if log_perf and (hydrate_stage != "memory" or elapsed_ms >= 5.0):
        try:
            log_status(f"[Perf] Spectro hydrate {filepath.name}: {hydrate_stage} {elapsed_ms:.0f} ms")
        except Exception:
            pass
    full_specs = list(full_specs or [])
    target_specs = [spec for spec in getattr(viewer, "spectros", []) or [] if _normalize_spectro_path_key(spec.get("path")) == norm_key]
    payload_by_id = {_spec_identity_token(spec): spec for spec in full_specs}
    for target in target_specs:
        payload = payload_by_id.get(_spec_identity_token(target))
        if payload is None and full_specs:
            payload = full_specs[0]
        if payload is not None:
            _merge_payload_into_spec(target, dict(payload))
    if isinstance(spec_or_path, dict):
        payload = payload_by_id.get(_spec_identity_token(spec_or_path))
        if payload is None and full_specs:
            payload = full_specs[0]
        if payload is not None:
            _merge_payload_into_spec(spec_or_path, dict(payload))
            return (spec_or_path, hydrate_stage) if return_stage else spec_or_path
    return (full_specs, hydrate_stage) if return_stage else full_specs


def hydrate_spectro_entries(viewer, specs):
    hydrated = []
    by_path = OrderedDict()
    stages = {"memory": 0, "disk": 0, "parse": 0}
    for spec in list(specs or []):
        path = spec.get("path")
        if not path:
            continue
        norm_key = _normalize_spectro_path_key(path)
        by_path.setdefault(norm_key, []).append(spec)
    for grouped_specs in by_path.values():
        seed = grouped_specs[0]
        full_specs, stage = hydrate_spectro_file(viewer, seed.get("path"), log_perf=False, return_stage=True)
        if full_specs is None:
            continue
        stages[str(stage or "memory")] = stages.get(str(stage or "memory"), 0) + 1
        payload_by_id = {
            _spec_identity_token(entry): entry for entry in (full_specs or [])
        }
        for spec in grouped_specs:
            payload = payload_by_id.get(_spec_identity_token(spec))
            if payload is None and full_specs:
                payload = full_specs[0]
            if payload is not None:
                _merge_payload_into_spec(spec, _clone_payload_spec_entry(payload))
                hydrated.append(spec)
    total = sum(stages.values())
    if total > 1:
        try:
            log_status(
                f"[Perf] Spectro hydrate batch: files={total} | memory={stages['memory']} | disk={stages['disk']} | parsed={stages['parse']}"
            )
        except Exception:
            pass
    return hydrated


def refresh_spectro_manifest_from_viewer(viewer):
    if not getattr(viewer, "spectro_manifest_cache_enabled", True):
        return
    folder = getattr(viewer, "spec_folder_path", None) or getattr(viewer, "last_dir", None)
    try:
        folder = Path(folder) if folder else None
    except Exception:
        folder = None
    if folder is None:
        return
    cache_dir = folder / ".sxmviewer_spectro_cache"
    manifest = dict(getattr(viewer, "_spectro_manifest_entries", {}) or {})
    grouped = defaultdict(list)
    for spec in getattr(viewer, "spectros", []) or []:
        path = spec.get("path")
        if not path:
            continue
        grouped[_normalize_spectro_path_key(path)].append(spec)
    updated = False
    for norm_key, spec_list in grouped.items():
        try:
            filepath = Path(spec_list[0].get("path"))
        except Exception:
            continue
        rel_key = _spectro_relative_key(folder, filepath)
        previous = manifest.get(rel_key, {})
        mtime = previous.get("mtime")
        fsize = previous.get("size")
        if mtime is None or fsize is None:
            try:
                st = filepath.stat()
                mtime = st.st_mtime
                fsize = st.st_size
            except Exception:
                mtime = 0.0 if mtime is None else mtime
                fsize = -1 if fsize is None else fsize
        payload_info = {
            "payload_meta": previous.get("payload_meta"),
            "payload_data": previous.get("payload_data"),
            "cache_key": previous.get("cache_key"),
        }
        manifest[rel_key] = _build_manifest_entry(folder, filepath, float(mtime), int(fsize), spec_list, payload_info=payload_info)
        updated = True
    if updated:
        viewer._spectro_manifest_entries = manifest
        if hasattr(viewer, "_schedule_spectro_manifest_save"):
            viewer._schedule_spectro_manifest_save()
        else:
            _save_spectro_manifest(cache_dir, manifest)


def _scan_spectros(
    viewer,
    folder: Path | None,
    files: list[Path] | None = None,
    *,
    image_paths=None,
    image_meta=None,
    use_disk_cache: bool = True,
):
    specs = []
    stats = {
        'display_count': 0,
        'matrix_files': 0,   # matrix-format files (true grids)
        'matrix_specs': 0,   # spectra originating from matrix files
        'total_specs': 0,
        'matrix_samples': [],
        'dat_files': 0,
        'txt_files': 0,
        'matrix_dat_files': 0,
        'single_dat_files': 0,
        'empty_files': 0,
        'single_entries': 0,
        'deferred_files': 0,
        'invalid_files': 0,
        'manifest_ms': 0.0,
        'disk_cache_ms': 0.0,
        'parse_ms': 0.0,
    }
    prefer_grid_as_matrix = bool(getattr(viewer, "spectro_single_grid_as_matrix", False))
    force_single_mode = bool(getattr(viewer, "spectro_force_single_mode", False))

    def _points_per_trace_for_list(spec_list):
        for spec in spec_list:
            points = _spec_points_per_trace(spec)
            if points:
                return int(points)
        return None

    def _derive_grid_from_specs(spec_list):
        col_candidates = [spec.get('grid_col') for spec in spec_list if spec.get('grid_col') is not None]
        row_candidates = [spec.get('grid_row') for spec in spec_list if spec.get('grid_row') is not None]
        matrix_indices = [spec.get('matrix_index') for spec in spec_list if spec.get('matrix_index') is not None]
        grid_cols = grid_rows = None
        zero_based = True
        if col_candidates and row_candidates:
            grid_cols = max(col_candidates) + 1
            grid_rows = max(row_candidates) + 1
        elif matrix_indices:
            min_idx = min(matrix_indices)
            max_idx = max(matrix_indices)
            # detect 1-based indexing
            if min_idx >= 1:
                zero_based = False
                max_idx -= 1
            side = int(round(math.sqrt(max_idx + 1)))
            if side > 0:
                grid_cols = grid_rows = side
        if not grid_cols or not grid_rows:
            total = len(spec_list)
            grid_cols = int(round(math.sqrt(total))) or 1
            grid_rows = int(math.ceil(total / grid_cols)) or 1
        return grid_rows, grid_cols, zero_based

    def _ensure_grid_indices(spec_list, grid_rows, grid_cols, zero_based=True):
        for idx, spec in enumerate(spec_list):
            row = spec.get('grid_row')
            col = spec.get('grid_col')
            if row is None or col is None:
                matrix_index = spec.get('matrix_index')
                if matrix_index is not None:
                    try:
                        val = int(matrix_index)
                        if not zero_based:
                            val -= 1
                        row = val // grid_cols
                        col = val % grid_cols
                    except Exception:
                        row = col = None
                if row is None or col is None:
                    row = idx // grid_cols
                    col = idx % grid_cols
            spec['grid_row'] = int(row)
            spec['grid_col'] = int(col)
            spec['matrix_index'] = int(row * grid_cols + col)

    def _clone_spec_entry(spec):
        clone = dict(spec)
        channels = spec.get("channels")
        if isinstance(channels, dict):
            clone["channels"] = dict(channels)
        axis_choices = spec.get("AxisChoices")
        if isinstance(axis_choices, (list, tuple)):
            clone["AxisChoices"] = [dict(ax) for ax in axis_choices]
        return clone

    def _reset_spec_classification(spec):
        # Preserve nanonis .3ds matrix metadata so grids are classified correctly.
        if spec.get("source") == "nanonis_3ds":
            return
        for key in ("matrix_dataset", "grid_rows", "grid_cols", "matrix_index", "grid_row", "grid_col", "channel_name", "channel_code"):
            if key in spec:
                spec.pop(key, None)

    def _rebuild_matrix_datasets_from_specs(spec_list):
        datasets = {}
        grouped = defaultdict(list)
        for spec in list(spec_list or []):
            if spec.get("matrix_index") is None:
                continue
            dataset_key = spec.get("matrix_dataset") or Path(spec.get("path") or "").stem
            grouped[(str(dataset_key), _normalize_spectro_path_key(spec.get("path") or ""))].append(spec)
        for (dataset_key, _path_key), members in grouped.items():
            if not members:
                continue
            first = members[0]
            rows = int(first.get("grid_rows") or 1)
            cols = int(first.get("grid_cols") or 1)
            ds = datasets.get(dataset_key)
            if ds is None:
                ds = MatrixDataset(dataset_key, rows, cols)
                datasets[dataset_key] = ds
            filename = Path(first.get("path") or "").name or str(first.get("path") or dataset_key)
            ds.add_channel(
                filename,
                channel_code=first.get("channel_code"),
                label=first.get("channel_name") or first.get("channel_code"),
                spectra_count=len(members),
                path=first.get("path"),
                points_per_trace=_points_per_trace_for_list(members),
            )
        viewer.matrix_datasets = datasets
        return datasets

    def _classify_file(spec_list, path_obj: Path):
        info = {
            "is_matrix": False,
            "dataset_key": None,
            "channel_code": None,
            "channel_label": None,
            "grid_rows": None,
            "grid_cols": None,
            "zero_based": True,
            "points_per_trace": None,
        }
        if not spec_list:
            return info
        # Force nanonis .3ds to be treated as matrix datasets
        if any(s.get("source") == "nanonis_3ds" for s in spec_list):
            grid_rows = spec_list[0].get("grid_rows") or spec_list[0].get("grid_row") or 0
            grid_cols = spec_list[0].get("grid_cols") or spec_list[0].get("grid_col") or 0
            if not grid_rows or not grid_cols:
                grid_rows, grid_cols, zero_based = _derive_grid_from_specs(spec_list)
            else:
                zero_based = True
            points_per_trace = _points_per_trace_for_list(spec_list)
            dataset_key = path_obj.stem
            info.update(
                {
                    "is_matrix": True,
                    "dataset_key": dataset_key,
                    "channel_code": None,
                    "channel_label": None,
                    "grid_rows": grid_rows,
                    "grid_cols": grid_cols,
                    "zero_based": zero_based,
                    "points_per_trace": points_per_trace,
                }
            )
            return info
        grid_rows, grid_cols, zero_based = _derive_grid_from_specs(spec_list)
        points_per_trace = _points_per_trace_for_list(spec_list)
        base, channel_code, ch_label = parse_matrix_filename(path_obj.name)
        dataset_key, display_label = matrix_dataset_key(base, channel_code)
        stem_base = _matrix_base_name(path_obj.stem)
        has_grid = grid_rows and grid_cols and (grid_rows * grid_cols == len(spec_list))
        single_point_matrix = bool(has_grid and grid_rows == 1 and grid_cols == 1 and len(spec_list) == 1)
        has_matrix_meta = any(
            (s.get('matrix_dataset') or (s.get('grid_cols') and s.get('grid_rows')))
            for s in spec_list
        )
        is_named_matrix = "matrix" in path_obj.name.lower()
        if force_single_mode:
            is_matrix = False
        elif single_point_matrix:
            is_matrix = False
        elif has_matrix_meta:
            is_matrix = True
        elif prefer_grid_as_matrix and has_grid and len(spec_list) > 1:
            is_matrix = True
        elif is_named_matrix and (has_grid or len(spec_list) > 1):
            is_matrix = True
        else:
            is_matrix = False
        ds_key = None
        if is_matrix:
            ds_key = (
                spec_list[0].get('matrix_dataset')
                or dataset_key
                or (f"{base}_{channel_code}" if base and channel_code else None)
                or stem_base
                or path_obj.stem
            )
        info.update(
            {
                "is_matrix": is_matrix,
                "dataset_key": ds_key,
                "channel_code": channel_code,
                "channel_label": display_label or ch_label or channel_code,
                "grid_rows": grid_rows,
                "grid_cols": grid_cols,
                "zero_based": zero_based,
                "points_per_trace": points_per_trace,
            }
        )
        return info

    def _assign_matrix_reference(spec_list, headers, ref_mtime: float, image_paths=None, image_meta=None):
        """Attach a single reference image_key to all spectra in a matrix dataset.
        We only anchor to actual loaded images (thumbnails). If none exist, we skip anchoring.
        """
        if not spec_list:
            return None
        candidates = []
        try:
            if image_meta:
                for img in image_meta:
                    p = str(img.get("path"))
                    if not p or "commands" in p.lower():
                        continue
                    candidates.append((p, img.get("time")))
        except Exception:
            candidates = []
        if not candidates:
            try:
                if image_paths:
                    for p in image_paths:
                        if p and "commands" not in str(p).lower():
                            candidates.append((str(p), None))
            except Exception:
                candidates = []
        if not candidates:
            return None
        valid_paths_lower = {p.lower(): p for p, _ in candidates}
        ref_epoch = None
        try:
            st = spec_list[0].get("time")
            if isinstance(st, datetime):
                ref_epoch = st.timestamp()
        except Exception:
            ref_epoch = None
        if ref_epoch is None:
            ref_epoch = ref_mtime if ref_mtime is not None else None
        best_key = None
        best_delta = None
        # Prefer loaded images (using image_meta time if available, else file mtime)
        for pth, tval in candidates:
            delta = None
            try:
                if isinstance(tval, datetime):
                    delta = abs(tval.timestamp() - ref_epoch) if ref_epoch is not None else None
                if delta is None:
                    ht = Path(pth).stat().st_mtime
                    delta = abs(ht - ref_epoch) if ref_epoch is not None else None
            except Exception:
                delta = None
            if delta is None:
                continue
            if best_delta is None or delta < best_delta:
                best_delta = delta
                best_key = pth
        # Final fallback: first candidate
        if not best_key:
            best_key = candidates[0][0]
        if best_key:
            key_norm = best_key
            if key_norm.lower() in valid_paths_lower:
                key_norm = valid_paths_lower[key_norm.lower()]
            for s in spec_list:
                s["image_key"] = key_norm
                s["image_path"] = key_norm
        return best_key

    viewer.matrix_datasets = {}
    if not folder and not files:
        return specs, stats
    folder_path = Path(folder) if folder else None
    if folder_path is not None and not folder_path.exists():
        folder_path = None
    # persistent spectroscopy cache directory (per folder)
    disk_cache_dir = None
    if use_disk_cache and folder_path is not None and getattr(viewer, "spectro_disk_cache_enabled", True):
        disk_cache_dir = folder_path / ".sxmviewer_spectro_cache"
        try:
            disk_cache_dir.mkdir(exist_ok=True)
        except Exception:
            disk_cache_dir = None
    manifest_enabled = bool(getattr(viewer, "spectro_manifest_cache_enabled", True))
    lazy_payload = bool(getattr(viewer, "spectro_lazy_payload_enabled", True))
    manifest_entries = _load_spectro_manifest(disk_cache_dir) if manifest_enabled else {}
    manifest_changed = False
    cache = viewer._spectro_cache
    seen_keys = set()
    cache_miss_logged = 0
    discovery_start = time.perf_counter()
    file_records = _discover_spectro_file_records(folder_path, files)
    files = [record["path"] for record in file_records]
    stats["discovery_ms"] = (time.perf_counter() - discovery_start) * 1000.0
    total = len(files)
    if total:
        log_status(f"Scanning {total} spectroscopy file(s)...")
    progress_step = max(1, total // 20) if total else 1
    # Use the known image files (thumbnails) as anchoring targets.
    # Prefer thumbnail metadata paths (image_meta), then files, then headers (skipping commands).
    image_paths = []
    try:
        image_paths = [str(img.get("path")) for img in (getattr(viewer, "image_meta", []) or []) if img.get("path")]
    except Exception:
        image_paths = []
    if image_paths is None:
        image_paths = []
        try:
            image_paths = [str(img.get("path")) for img in (getattr(viewer, "image_meta", []) or []) if img.get("path")]
        except Exception:
            image_paths = []
        if not image_paths:
            try:
                image_paths = [str(p) for p in getattr(viewer, "files", []) or []]
            except Exception:
                image_paths = []
        # Debug log removed for normal operation
        if not image_paths:
            try:
                image_paths = [
                    str(Path(k))
                    for k in (viewer.headers.keys() if getattr(viewer, "headers", None) else [])
                    if "commands" not in str(k).lower()
                ]
            except Exception:
                image_paths = []
    if not image_paths:
        try:
            image_paths = [str(p) for p in getattr(viewer, "files", []) or []]
        except Exception:
            image_paths = []

    bulk_manifest_ready = False
    if manifest_enabled and lazy_payload and file_records:
        try:
            bulk_manifest_ready = all(
                _manifest_entry_valid(manifest_entries.get(record["rel_key"]), mtime=record["mtime"], fsize=record["size"])
                for record in file_records
            )
        except Exception:
            bulk_manifest_ready = False

    if bulk_manifest_ready:
        t_manifest = time.perf_counter()
        for record in file_records:
            p = record["path"]
            norm_key = record["norm_key"]
            ext = record["suffix"]
            if ext == ".dat":
                stats['dat_files'] += 1
                stats['single_dat_files'] += 1
            elif ext == ".txt":
                stats['txt_files'] += 1
            elif ext == ".3ds":
                stats['dat_files'] += 1
            seen_keys.add(norm_key)
            entry = manifest_entries.get(record["rel_key"]) or {}
            spec_list = [_restore_spec_metadata(spec_entry) for spec_entry in (entry.get("specs") or [])]
            specs.extend(spec_list)
            stats["manifest_hits"] = stats.get("manifest_hits", 0) + 1
        stats["manifest_ms"] += (time.perf_counter() - t_manifest) * 1000.0
        stats['display_count'] = len(specs)
        stats['single_entries'] = sum(1 for spec in specs if spec.get("matrix_index") is None)
        stats['matrix_specs'] = sum(1 for spec in specs if spec.get("matrix_index") is not None)
        matrix_datasets = _rebuild_matrix_datasets_from_specs(specs)
        stats['matrix_dat_files'] = sum(len(ds.channels) for ds in matrix_datasets.values())
        stats['matrix_files'] = len(matrix_datasets)
        try:
            log_status("  - spectroscopy finalize metadata...")
            log_status("  - spectroscopy finalize ordering...")
        except Exception:
            pass
        stale = [k for k in list(cache.keys()) if k not in seen_keys]
        for k in stale:
            cache.pop(k, None)
        current_rel_keys = {record["rel_key"] for record in file_records}
        stale_manifest = [key for key in list(manifest_entries.keys()) if key not in current_rel_keys]
        if stale_manifest:
            for key in stale_manifest:
                manifest_entries.pop(key, None)
            manifest_changed = True
        viewer._spectro_manifest_entries = manifest_entries
        if manifest_changed:
            if hasattr(viewer, "_schedule_spectro_manifest_save"):
                viewer._schedule_spectro_manifest_save()
            else:
                _save_spectro_manifest(disk_cache_dir, manifest_entries)
        specs.sort(key=lambda s: s.get('time') or datetime.min)
        stats['total_specs'] = len(specs)
        log_status("Spectroscopy scan summary:")
        log_status(
            f"  Files: {len(file_records)} total  |  singles: {stats['single_dat_files']}  |  matrices: {stats['matrix_dat_files']}  |  empty/deferred: {stats['empty_files']}/{stats['deferred_files']}  |  invalid: {stats['invalid_files']}"
        )
        log_status(
            f"  Spectra: {stats['total_specs']} total  |  from singles: {stats['single_entries']} traces  |  from matrices: {stats['matrix_specs']} traces"
        )
        log_status(
            f"  Cache: {len(file_records)}/{len(file_records)} files (100% hit rate)  |  memory: 0  |  manifest: {stats.get('manifest_hits', 0)}  |  disk: 0  |  parsed: 0"
        )
        log_status(
            f"  Timings: discovery {stats.get('discovery_ms', 0.0):.0f} ms  |  manifest {stats.get('manifest_ms', 0.0):.0f} ms  |  disk payload 0 ms  |  raw parse 0 ms"
        )
        try:
            json_line = {
                "folder": str(folder_path) if folder_path is not None else "",
                "files_scanned": len(file_records),
                "spectra_total": stats.get("total_specs", 0),
                "single_files": stats.get("single_dat_files", 0),
                "single_entries": stats.get("single_entries", 0),
                "matrix_datasets": len(matrix_datasets),
                "matrix_spectra": stats.get("matrix_specs", 0),
                "empty_files": stats.get("empty_files", 0),
                "invalid_files": stats.get("invalid_files", 0),
            }
            log_status(f"[SXMViewer-JSON] {json.dumps(json_line)}")
        except Exception:
            pass
        return specs, stats

    for idx, record in enumerate(file_records, 1):
        spec_list = None
        parse_error = None
        disk_cached = None
        p = record["path"]
        norm_key = record["norm_key"]
        ext = record["suffix"]
        if ext == ".dat":
            stats['dat_files'] += 1
        elif ext == ".txt":
            stats['txt_files'] += 1
        elif ext == ".3ds":
            stats['dat_files'] += 1
        if norm_key in seen_keys:
            continue
        seen_keys.add(norm_key)
        mtime = float(record.get("mtime", 0.0))
        fsize = int(record.get("size", -1))
        cached = cache.get(norm_key)
        rel_key = record["rel_key"]
        manifest_entry = manifest_entries.get(rel_key) if manifest_enabled else None
        payload_info = None
        # eager parse limit (0 means no deferral)
        if viewer.spectro_eager_limit and idx > viewer.spectro_eager_limit:
            if manifest_entry and _manifest_entry_valid(manifest_entry, mtime=mtime, fsize=fsize):
                t_manifest = time.perf_counter()
                spec_list = [_restore_spec_metadata(entry) for entry in (manifest_entry.get("specs") or [])]
                stats["manifest_ms"] += (time.perf_counter() - t_manifest) * 1000.0
                stats["manifest_hits"] = stats.get("manifest_hits", 0) + 1
            elif disk_cache_dir:
                t_disk = time.perf_counter()
                disk_cached, payload_info = _load_spectro_disk_payload(disk_cache_dir, folder_path, p, mtime, fsize)
                stats["disk_cache_ms"] += (time.perf_counter() - t_disk) * 1000.0
                if disk_cached is not None:
                    manifest_entries[rel_key] = _build_manifest_entry(folder_path, p, mtime, fsize, disk_cached, payload_info=payload_info)
                    manifest_changed = True
                    spec_list = _spec_metadata_list(disk_cached) if lazy_payload else [_clone_spec_entry(entry) for entry in disk_cached]
                    stats["disk_cache_hits"] = stats.get("disk_cache_hits", 0) + 1
            if not spec_list:
                stats['deferred_files'] += 1
                cache[norm_key] = {'mtime': mtime, 'deferred': True, 'path': str(p)}
                viewer._spectro_deferred.add(norm_key)
                continue

        if cached and abs(cached.get('mtime', 0.0) - mtime) <= _SPECTRO_CACHE_MTIME_TOLERANCE and not cached.get('deferred'):
            raw_list = cached.get('data') or []
            restored = [_clone_payload_spec_entry(entry) for entry in raw_list]
            spec_list = _spec_metadata_list(restored) if lazy_payload else [_clone_spec_entry(entry) for entry in restored]
            stats["cache_hits"] = stats.get("cache_hits", 0) + 1
        else:
            if manifest_entry and _manifest_entry_valid(manifest_entry, mtime=mtime, fsize=fsize):
                t_manifest = time.perf_counter()
                spec_list = [_restore_spec_metadata(entry) for entry in (manifest_entry.get("specs") or [])]
                stats["manifest_ms"] += (time.perf_counter() - t_manifest) * 1000.0
                stats["manifest_hits"] = stats.get("manifest_hits", 0) + 1
            elif disk_cache_dir:
                t_disk = time.perf_counter()
                disk_cached, payload_info = _load_spectro_disk_payload(disk_cache_dir, folder_path, p, mtime, fsize)
                stats["disk_cache_ms"] += (time.perf_counter() - t_disk) * 1000.0
                if disk_cached is not None:
                    manifest_entries[rel_key] = _build_manifest_entry(folder_path, p, mtime, fsize, disk_cached, payload_info=payload_info)
                    manifest_changed = True
                    spec_list = _spec_metadata_list(disk_cached) if lazy_payload else [_clone_spec_entry(entry) for entry in disk_cached]
                    stats["disk_cache_hits"] = stats.get("disk_cache_hits", 0) + 1
            if spec_list is None and disk_cache_dir and cache_miss_logged < 10:
                cache_miss_logged += 1
                try:
                    log_status(f"  - spectroscopy cache miss: {p.name}")
                except Exception:
                    pass
            if spec_list is None:
                stats["cache_miss"] = stats.get("cache_miss", 0) + 1
                t_parse = time.perf_counter()
                parsed_specs, parse_error = _parse_spectro_file_payload(p, mtime)
                stats["parse_ms"] += (time.perf_counter() - t_parse) * 1000.0
                spec_list = parsed_specs
            if parse_error is not None:
                stats['invalid_files'] += 1
                try:
                    log_status(f"Spectroscopy parse rejected: {parse_error}")
                except Exception:
                    pass
                continue
            if not spec_list:
                stats['empty_files'] += 1
                continue
            cache[norm_key] = {'mtime': mtime, 'data': [_clone_payload_spec_entry(spec) for spec in spec_list]}
            if disk_cache_dir:
                payload_info = _store_spectro_disk_payload(disk_cache_dir, folder_path, p, mtime, fsize, spec_list)
            if manifest_enabled:
                manifest_entries[rel_key] = _build_manifest_entry(folder_path, p, mtime, fsize, spec_list, payload_info=payload_info)
                manifest_changed = True
            if lazy_payload:
                spec_list = _spec_metadata_list(spec_list)
        for spec in spec_list or []:
            _reset_spec_classification(spec)
        specs.extend(spec_list or [])
        info = _classify_file(spec_list, p)
        if info.get("is_matrix"):
            grid_rows = info.get("grid_rows") or 1
            grid_cols = info.get("grid_cols") or 1
            _ensure_grid_indices(spec_list, grid_rows, grid_cols, zero_based=info.get("zero_based", True))
            # Anchor all points of the matrix to a single reference image to avoid scatter
            chosen_key = _assign_matrix_reference(spec_list, viewer.headers, mtime, image_paths=image_paths, image_meta=getattr(viewer, "image_meta", None))
            if not chosen_key and image_paths:
                fallback_key = image_paths[0]
                for s in spec_list:
                    s["image_key"] = fallback_key
                    s["image_path"] = fallback_key
                chosen_key = fallback_key
            if chosen_key:
                try:
                    log_status(f"  - matrix anchored to: {Path(chosen_key).name} ({len(spec_list)} pts, {grid_cols}x{grid_rows})")
                except Exception:
                    pass
                # Warn if anchor is not among known image paths (may prevent markers from drawing)
                # No extra debug here; anchor info is logged above
            stats['matrix_files'] += 1
            stats['matrix_specs'] += len(spec_list)
            stats['display_count'] += 1
            if ext == ".dat":
                stats['matrix_dat_files'] += 1
            ds_key = info.get("dataset_key") or Path(p).stem
            ds = viewer.matrix_datasets.get(ds_key)
            if ds is None:
                ds = MatrixDataset(ds_key, grid_rows, grid_cols)
                viewer.matrix_datasets[ds_key] = ds
            label = info.get("channel_label") or info.get("channel_code") or Path(p).stem
            ds.add_channel(
                p.name,
                channel_code=info.get("channel_code"),
                label=label,
                spectra_count=len(spec_list),
                path=p,
                points_per_trace=info.get("points_per_trace"),
            )
            for spec in spec_list or []:
                spec.setdefault('matrix_dataset', ds_key)
                if label:
                    spec.setdefault('channel_name', label)
                if info.get("channel_code"):
                    spec.setdefault('channel_code', info.get("channel_code"))
                spec.setdefault('grid_rows', grid_rows)
                spec.setdefault('grid_cols', grid_cols)
            if len(stats['matrix_samples']) < 3:
                grid_desc = f"{grid_cols}x{grid_rows}"
                pts = info.get("points_per_trace")
                pts_txt = f", {pts} pts/trace" if pts else ""
                stats['matrix_samples'].append(f"{p.name}: {grid_desc} ({len(spec_list)} spectra{pts_txt})")
        else:
            if spec_list:
                stats['single_dat_files'] += 1
                stats['single_entries'] += len(spec_list)
                stats['display_count'] += len(spec_list)
            else:
                stats['empty_files'] += 1
        if total and (idx % progress_step == 0 or idx == total):
            pct = idx / total * 100.0
            log_status(f"  - spectroscopy load {idx}/{total} ({pct:4.0f}%)")
    try:
        log_status("  - spectroscopy finalize metadata...")
    except Exception:
        pass
    stale = [k for k in list(cache.keys()) if k not in seen_keys]
    for k in stale:
        cache.pop(k, None)
    if manifest_enabled:
        current_rel_keys = {record["rel_key"] for record in file_records}
        stale_manifest = [key for key in list(manifest_entries.keys()) if key not in current_rel_keys]
        for key in stale_manifest:
            manifest_entries.pop(key, None)
            manifest_changed = True
        viewer._spectro_manifest_entries = manifest_entries
        if manifest_changed:
            if hasattr(viewer, "_schedule_spectro_manifest_save"):
                viewer._schedule_spectro_manifest_save()
            else:
                _save_spectro_manifest(disk_cache_dir, manifest_entries)
    try:
        log_status("  - spectroscopy finalize ordering...")
    except Exception:
        pass
    specs.sort(key=lambda s: s.get('time') or datetime.min)
    stats['total_specs'] = len(specs)
    # logging summary
    single_files = stats.get('single_dat_files', 0)
    empty_files = stats.get('empty_files', 0)
    invalid_files = stats.get('invalid_files', 0)
    matrix_count = len(viewer.matrix_datasets)
    matrix_specs = stats.get('matrix_specs', 0)
    single_entries = stats.get('single_entries', single_files)
    log_status("Spectroscopy scan summary:")
    log_status(
        f"  Files: {total} total  |  singles: {single_files}  |  matrices: {stats.get('matrix_files', matrix_count)}  |  empty/deferred: {empty_files}/{stats.get('deferred_files',0)}  |  invalid: {invalid_files}"
    )
    log_status(
        f"  Spectra: {stats['total_specs']} total  |  from singles: {single_entries} traces  |  from matrices: {matrix_specs} traces"
    )
    cache_hits = stats.get("cache_hits", 0)
    manifest_hits = stats.get("manifest_hits", 0)
    disk_hits = stats.get("disk_cache_hits", 0)
    cache_miss = stats.get("cache_miss", 0)
    total_cached = cache_hits + manifest_hits + disk_hits
    if total:
        cache_pct = (total_cached / max(total, 1)) * 100
        log_status(
            f"  Cache: {total_cached}/{total} files ({cache_pct:.0f}% hit rate)  |  memory: {cache_hits}  |  manifest: {manifest_hits}  |  disk: {disk_hits}  |  parsed: {cache_miss}"
        )
        log_status(
            f"  Timings: discovery {stats.get('discovery_ms', 0.0):.0f} ms  |  manifest {stats.get('manifest_ms', 0.0):.0f} ms  |  disk payload {stats.get('disk_cache_ms', 0.0):.0f} ms  |  raw parse {stats.get('parse_ms', 0.0):.0f} ms"
        )
    if viewer.matrix_datasets:
        log_status("  Matrix datasets:")
        for key, ds in sorted(viewer.matrix_datasets.items(), key=lambda kv: kv[0]):
            chans = []
            spectra_per_ch = []
            points_per_trace = []
            mtimes = []
            for ch in ds.channels:
                label = ch.get('label') or ch.get('channel_code') or Path(ch.get('filename', '')).stem
                chans.append(label)
                try:
                    spectra_per_ch.append(int(ch.get('spectra_count', 0)))
                except Exception:
                    pass
                pts = ch.get('points_per_trace')
                if pts:
                    try:
                        points_per_trace.append(int(pts))
                    except Exception:
                        pass
                try:
                    mtimes.append(Path(ch.get('path')).stat().st_mtime)
                except Exception:
                    continue
            chan_txt = ", ".join(chans) if chans else "1 channel"
            spectra_txt = ""
            if spectra_per_ch:
                spectra_txt = f" | spectra/ch: {max(spectra_per_ch)}"
            points_txt = ""
            if points_per_trace:
                points_txt = f" | points/trace: {max(points_per_trace)}"
            acq_txt = ""
            if mtimes:
                try:
                    acq_txt = f" | acquired: {datetime.fromtimestamp(min(mtimes)).strftime('%Y-%m-%d %H:%M')}"
                except Exception:
                    pass
            label = key or ds.base or "matrix"
            if ds.base and key and key.startswith(f"{ds.base}_"):
                suffix = key[len(ds.base) + 1 :]
                label = f"{ds.base}_{suffix}"
            log_status(
                f"    - {label}: {ds.cols}x{ds.rows} px | channels: {chan_txt}{spectra_txt}{points_txt}{acq_txt}"
            )
    try:
        verbose = os.environ.get("SXM_VERBOSE")
        json_line = {
            "folder": str(folder),
            "files_scanned": total,
            "spectra_total": stats['total_specs'],
            "single_files": single_files,
            "single_entries": single_entries,
            "matrix_datasets": matrix_count,
            "matrix_spectra": matrix_specs,
            "empty_files": empty_files,
            "invalid_files": invalid_files,
        }
        log_status(f"[SXMViewer-JSON] {json.dumps(json_line)}")
        if verbose:
            log_status("Matrix datasets:")
            for key, ds in viewer.matrix_datasets.items():
                log_status(
                    f"  - {ds.base}: {len(ds.channels)} channel(s)  {ds.rows}x{ds.cols} -> "
                    f"{sum(c.get('spectra_count',0) for c in ds.channels)} spectra"
                )
                for ch in ds.channels:
                    log_status(
                        f"      * {Path(ch['path']).name} ({ch.get('channel_code')}) {ch.get('label','')} "
                        f"-> {ch.get('spectra_count')} spectra"
                    )
    except Exception:
        pass
    return specs, stats


def _coerce_pos_to_nm(value: float) -> float:
    """
    Best-effort unit coercion for spectroscopy positions.

    Heuristic:
    - |v| < 1e-6 -> assume meters, convert to nm.
    - Otherwise assume already in nm.
    """
    try:
        v = float(value)
    except Exception:
        return value
    if abs(v) < 1e-6:
        return v * 1e9
    return v
__all__ = [
    "load_folder",
    "_parse_header_datetime",
    "_scan_spectros",
]
