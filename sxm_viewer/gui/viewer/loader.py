"""Loader helpers for SXMGridViewer."""
from __future__ import annotations

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
from ...providers import convert_nanonis, parse_nanonis_spectroscopy
from ..detail_panels import SpectroscopyPopup, SpectroscopyCompareDialog

def load_folder(viewer, folder:Path):
    folder = Path(folder)
    log_status(f"Loading folder: {folder}")
    viewer._update_toolbar_actions(False)
    prev_last_dir = getattr(viewer, 'last_dir', None)
    viewer.last_dir = folder
    viewer.path_le.setText(str(folder))
    # persist last dir early
    viewer.config['last_dir'] = str(folder)
    viewer._record_recent_dir(folder)

    txts = sorted(folder.glob("*.txt"))
    converted = convert_nanonis(folder)
    if converted:
        txts = sorted(list(txts) + list(converted), key=lambda p: str(p).lower())
        log_status(f"Converted {len(converted)} Nanonis scan(s)")
    log_status(f"Found {len(txts)} header file(s)")
    viewer.files = txts
    viewer.headers.clear()
    viewer._invalidate_thumbnail_cache()
    viewer._invalidate_channel_cache()
    viewer.thumb_multi_select = set()
    cache_hits = 0
    cache_miss = 0
    for t in txts:
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
        viewer.headers[str(t)] = (hdr, fds)
    if cache_miss:
        viewer._save_header_cache()
    log_status(f"Headers loaded (hits={cache_hits}, miss={cache_miss})")
    if not viewer.headers:
        viewer.meta_box.setPlainText("No valid .txt headers found")
        viewer.clear_thumbs(); return
    viewer._build_image_timestamp_index()
    viewer._rebuild_frame_map_entries()

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
    viewer.channel_dropdown.setMinimumWidth(380)
    if 0 <= viewer.last_channel_index < viewer.channel_dropdown.count():
        viewer.channel_dropdown.setCurrentIndex(viewer.last_channel_index)
    else:
        viewer.last_channel_index = 0; viewer.channel_dropdown.setCurrentIndex(0)
    viewer.channel_dropdown.blockSignals(False)

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

    # auto-detect tags for files not already tagged
    log_status("Auto-detecting tags...")
    viewer._auto_detect_tags_for_folder()

    # keep spectroscopy folder aligned with the SXM folder unless the user picked a custom path
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

    # load spectroscopy markers referencing this folder
    log_status("Loading spectroscopy references...")
    viewer._reload_spectros(refresh=False)

    QtCore.QTimer.singleShot(0, lambda: viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex()))
    log_status("Folder load complete.")


def _parse_header_datetime(viewer, header):
    """Return a sortable key (float timestamp) parsed from header Date/Time if possible; otherwise 0.0.
    Accepts common formats, falls back to 0.0 on failure."""
    try:
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
            '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M:%S', '%d/%m/%Y %H:%M:%S',
            '%d-%m-%Y %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y'
        ]
        for s in candidates:
            for fmt in fmts:
                try:
                    dt = datetime.strptime(s, fmt)
                    return dt.timestamp()
                except Exception:
                    continue
        return 0.0
    except Exception:
        return 0.0


def _scan_spectros(viewer, folder:Path):
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
    }
    def _is_matrix_file(spec_list, path_obj: Path) -> bool:
        if not spec_list:
            return False
        if any(s.get('matrix_dataset') for s in spec_list):
            return True
        if any((s.get('grid_cols') or 0) > 1 or (s.get('grid_rows') or 0) > 1 for s in spec_list):
            return True
        name = path_obj.name.lower()
        return "matrix" in name and len(spec_list) > 1

    viewer.matrix_datasets = {}
    if not folder or not Path(folder).exists():
        return specs, stats
    patterns = ("*.dat","*.DAT")
    cache = viewer._spectro_cache
    seen_keys = set()
    file_map = {}
    for pat in patterns:
        for f in folder.glob(pat):
            # normalize path for dedup (case-insensitive on Windows)
            try:
                key = str(f.resolve())
            except Exception:
                key = str(f)
            norm_key = key.lower() if os.name == "nt" else key
            if norm_key not in file_map:
                file_map[norm_key] = f
    files = sorted(file_map.values(), key=lambda p: str(p).lower())
    total = len(files)
    if total:
        log_status(f"Scanning {total} spectroscopy file(s)...")
    progress_step = max(1, total // 20) if total else 1
    for idx, f in enumerate(files, 1):
        p = Path(f)
        if p.is_dir():
            continue
        try:
            key = str(p.resolve())
        except Exception:
            key = str(p)
        norm_key = key.lower() if os.name == "nt" else key
        ext = p.suffix.lower()
        if ext == ".dat":
            stats['dat_files'] += 1
        elif ext == ".txt":
            stats['txt_files'] += 1
        if norm_key in seen_keys:
            continue
        seen_keys.add(norm_key)
        try:
            mtime = p.stat().st_mtime
        except Exception:
            mtime = 0.0
        cached = cache.get(norm_key)
        # eager parse limit (0 means no deferral)
        if viewer.spectro_eager_limit and idx > viewer.spectro_eager_limit:
            stats['deferred_files'] += 1
            cache[norm_key] = {'mtime': mtime, 'deferred': True, 'path': str(p)}
            viewer._spectro_deferred.add(norm_key)
            continue

        if cached and abs(cached.get('mtime', 0.0) - mtime) <= 1e-6 and not cached.get('deferred'):
            spec_list = cached.get('data') or []
        else:
            spec_list = None
            if ext == ".dat":
                # Prefer Nanonis parsing first for .dat; fallback to legacy/Omicron parser if empty.
                try:
                    spec_list = parse_nanonis_spectroscopy(p)
                except Exception:
                    spec_list = None
            if not spec_list:
                try:
                    spec_list = parse_spectroscopy_file(p)
                except Exception:
                    spec_list = None
            if not spec_list:
                stats['empty_files'] += 1
                continue

            # --- Fallback: Parse coordinates from header comments if missing (e.g. Nanonis .dat) ---
            if spec_list and ext == ".dat":
                if any(s.get('x') is None or s.get('y') is None for s in spec_list):
                    try:
                        with open(p, 'r', encoding='latin-1') as f:
                            # Scan first 30 lines for metadata
                            for _ in range(30):
                                line = f.readline()
                                if not line: break
                                # Look for line like: ;x/y-Pos: -104.75/-266
                                if "x/y-Pos:" in line:
                                    parts = line.split(":", 1)[1].strip().split("/")
                                    if len(parts) == 2:
                                        try:
                                            x_val, y_val = float(parts[0]), float(parts[1])
                                            x_nm = _coerce_pos_to_nm(x_val)
                                            y_nm = _coerce_pos_to_nm(y_val)
                                            for s in spec_list:
                                                if s.get('x') is None: s['x'] = x_nm
                                                if s.get('y') is None: s['y'] = y_nm
                                        except ValueError: pass
                                    break
                    except Exception: pass

            # ensure basic metadata is present for assignment
            for s in spec_list or []:
                if 'path' not in s or not s.get('path'):
                    s['path'] = str(p)
                # normalize/ensure time for ordering; fallback to file mtime
                t = s.get('time')
                use_mtime = False
                if t is None:
                    use_mtime = True
                elif isinstance(t, datetime):
                    if t.year < 1990:
                        use_mtime = True
                elif isinstance(t, (int, float)):
                    try:
                        s['time'] = datetime.fromtimestamp(float(t))
                    except Exception:
                        use_mtime = True
                elif isinstance(t, str):
                    if not t.strip():
                        use_mtime = True
                    else:
                        try:
                            s['time'] = datetime.fromisoformat(t)
                        except Exception:
                            use_mtime = True
                
                if use_mtime:
                    try:
                        s['time'] = datetime.fromtimestamp(mtime)
                    except Exception:
                        pass
            cache[norm_key] = {'mtime': mtime, 'data': spec_list}
        specs.extend(spec_list or [])
        # counting logic: only treat files as matrix when flagged as such (dataset/grid or name hint)
        is_matrix = _is_matrix_file(spec_list, p)
        if is_matrix:
            stats['matrix_files'] += 1
            stats['matrix_specs'] += len(spec_list)
            stats['display_count'] += 1
            if ext == ".dat":
                stats['matrix_dat_files'] += 1
            # build/augment MatrixDataset
            base, channel_code, ch_label = parse_matrix_filename(p.name)
            dataset_key, display_label = matrix_dataset_key(base, channel_code)
            label = display_label or ch_label or channel_code or Path(p).stem
            # infer grid
            grid_cols = None
            grid_rows = None
            for s in spec_list:
                grid_cols = grid_cols or s.get('grid_cols')
                grid_rows = grid_rows or s.get('grid_rows')
                if grid_cols and grid_rows:
                    break
            if not grid_cols or not grid_rows:
                n = max(1, len(spec_list))
                side = int(round(n ** 0.5))
                grid_cols = grid_cols or side
                grid_rows = grid_rows or side
            ds_key = dataset_key or base or Path(p).stem
            ds = viewer.matrix_datasets.get(ds_key)
            if ds is None:
                ds = MatrixDataset(ds_key, grid_rows, grid_cols)
                viewer.matrix_datasets[ds_key] = ds
            ds.add_channel(p.name, channel_code=channel_code, label=label, spectra_count=len(spec_list), path=p)
            for spec in spec_list or []:
                spec.setdefault('matrix_dataset', ds_key)
                if label:
                    spec.setdefault('channel_name', label)
                if channel_code:
                    spec.setdefault('channel_code', channel_code)
            # describe grid if available
            grid_cols = None
            grid_rows = None
            for s in spec_list:
                grid_cols = grid_cols or s.get('grid_cols')
                grid_rows = grid_rows or s.get('grid_rows')
                if grid_cols and grid_rows:
                    break
            if grid_cols and grid_rows:
                grid_desc = f"{grid_cols}x{grid_rows}"
            else:
                # fallback: infer from count if square
                n = len(spec_list)
                side = int(round(n ** 0.5))
                grid_desc = f"~{side}x{side}" if side * side == n else f"{n} spectra"
            if len(stats['matrix_samples']) < 3:
                stats['matrix_samples'].append(
                    f"{p.name}: {grid_desc} ({len(spec_list)} spectra)"
                )
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
    stale = [k for k in list(cache.keys()) if k not in seen_keys]
    for k in stale:
        cache.pop(k, None)
    specs.sort(key=lambda s: s.get('time') or datetime.min)
    stats['total_specs'] = len(specs)
    # logging summary
    single_files = stats.get('single_dat_files', 0)
    empty_files = stats.get('empty_files', 0)
    matrix_count = len(viewer.matrix_datasets)
    matrix_specs = stats.get('matrix_specs', 0)
    single_entries = stats.get('single_entries', single_files)
    log_status("Spectroscopy scan summary:")
    log_status(
        f"  Files: {total} total  |  singles: {single_files}  |  matrices: {stats.get('matrix_files', matrix_count)}  |  empty/deferred: {empty_files}/{stats.get('deferred_files',0)}"
    )
    log_status(
        f"  Spectra: {stats['total_specs']} total  |  from singles: {single_entries} traces  |  from matrices: {matrix_specs} traces"
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
                try:
                    # assume each channel in the dataset has the same trace length; store first nonzero
                    cache_entry = cache.get(str(Path(ch.get('path')).resolve()).lower(), {})
                    data = cache_entry.get('data') if isinstance(cache_entry, dict) else None
                    if isinstance(data, list):
                        trace = data[0] if data else None
                    else:
                        trace = None
                    vals = None
                    if isinstance(trace, dict):
                        vals = trace.get('V')
                        if vals is None and trace.get('channels'):
                            vals = next(iter(trace['channels'].values()))
                    if vals is not None:
                        arr = np.asarray(vals)
                        points_per_trace.append(int(arr.size))
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
        import json
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
        }
        log_status(f"[SXMViewer-JSON] {json.dumps(json_line)}")
        if verbose:
            log_status("Matrix datasets:")
            for key, ds in viewer.matrix_datasets.items():
                log_status(f"  - {ds.base}: {len(ds.channels)} channel(s)  {ds.rows}x{ds.cols} -> "
                           f"{sum(c.get('spectra_count',0) for c in ds.channels)} spectra")
                for ch in ds.channels:
                    log_status(f"      * {Path(ch['path']).name} ({ch.get('channel_code')}) {ch.get('label','')} "
                               f"-> {ch.get('spectra_count')} spectra")
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
