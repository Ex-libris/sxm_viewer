"""Export helpers for SXMGridViewer."""
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
from ..workers.batch_export import BatchExportSignals, BatchExportWorker
from ...utils.units import _safe_float
from ..wsxm_stp import save_wsxm_stp

try:  # Late imports to avoid circular references in stripped-down builds
    from ..thumbnail_render import convert_to_si, _value_in_nm
except Exception:  # pragma: no cover - fallback for testing
    convert_to_si = None

    def _value_in_nm(val, unit):
        return None

def _collect_channel_exports(viewer, header_path_str, main_channel_idx=None):
    header_path = Path(header_path_str)
    file_key = str(header_path)
    header, fds = viewer.headers.get(file_key, (None, None))
    if header is None or not fds:
        return header, []
    base_extent = viewer._header_extent(header)
    exports = []
    channel_idx = main_channel_idx
    if channel_idx is None:
        channel_idx = 0
    if channel_idx < 0 or channel_idx >= len(fds):
        channel_idx = 0
    def _append(idx, cmap=None):
        if idx is None or idx < 0 or idx >= len(fds):
            return
        fd = fds[idx]
        try:
            unit_final, arr_conv = viewer._get_filtered_channel_array(file_key, idx, header, fd)
        except Exception:
            return
        cap = fd.get('Caption', fd.get('FileName', f"chan{idx}"))
        adj_arr, adj_extent = viewer._apply_adjustments_for_channel(file_key, idx, arr_conv, base_extent)
        disp_extent = viewer._display_extent(adj_extent, header)
        disp_unit, disp_arr, _ = viewer._scale_unit_for_display(unit_final, adj_arr, file_key)
        cbar_label = cap
        if disp_unit:
            cbar_label = f"{cap} [{disp_unit}]"
        date = str(header.get('Date', '') or '').strip()
        time_txt = str(header.get('Time', '') or '').strip()
        datetime_txt = " ".join([t for t in (date, time_txt) if t]).strip()
        if datetime_txt:
            title_txt = f"{header_path.name}  {cap}  {datetime_txt}"
        else:
            title_txt = f"{header_path.name}  {cap}"
        exports.append({
            'arr': disp_arr,
            'extent': disp_extent,
            'unit': disp_unit,
            'caption': cap,
            'idx': idx,
            'cmap': cmap,
            'fd': fd,
            'relative_axes': bool(viewer.relative_axes),
            'colorbar_label': cbar_label,
            'title': title_txt,
        })
    cmap_main = viewer.per_file_channel_cmap.get((file_key, channel_idx), viewer.preview_cmap_combo.currentText() or viewer.preview_cmap)
    _append(channel_idx, cmap_main)
    for spec in getattr(viewer, 'extra_view_specs', []):
        try:
            idx2 = viewer._find_channel_index_for_spec(fds, spec)
        except Exception:
            idx2 = None
        if idx2 is None:
            continue
        cmap2 = viewer._resolve_extra_spec_cmap(spec, file_key)
        _append(idx2, cmap2)
    return header, exports


def export_targets(viewer, paths=None):
    """``paths`` (or the thumbnail selection) in thumbnail-grid order, else
    the previewed image."""
    if paths:
        order = {str(fp): i for i, fp in enumerate(getattr(viewer, "current_thumb_files", []) or [])}
        uniq = list(dict.fromkeys(str(p) for p in paths if p))
        return sorted(uniq, key=lambda k: (order.get(k, len(order)), k))
    targets = []
    try:
        targets = list(viewer._ordered_thumbnail_selection() or [])
    except Exception:
        targets = []
    if not targets and viewer.last_preview:
        targets = [str(viewer.last_preview[0])]
    return [str(t) for t in targets if t]


def _unique_path(path):
    path = Path(path)
    if not path.exists():
        return path
    for n in itertools.count(2):
        cand = path.with_name(f"{path.stem}_{n}{path.suffix}")
        if not cand.exists():
            return cand


def png_name_for(viewer, file_key, channel_idx=None):
    """``<file>__<channel>__<date>__<time>.png`` for one displayed image."""
    header, fds = (getattr(viewer, "headers", {}) or {}).get(str(file_key), (None, None))
    header = header or {}
    chan = ""
    try:
        if fds and channel_idx is not None and 0 <= int(channel_idx) < len(fds):
            fd = fds[int(channel_idx)]
            chan = fd.get("Caption", fd.get("FileName", "")) or ""
    except Exception:
        chan = ""
    clean = viewer._sanitize_filename_component
    parts = [clean(Path(str(file_key)).stem)]
    if chan:
        parts.append(clean(chan))
    for key in ("Date", "Time"):
        val = str(header.get(key, "") or "").strip()
        if val:
            parts.append(clean(val))
    return "__".join(parts) + ".png"


def render_targets_as_displayed(viewer, targets, out_dir, *, dpi=300, title="Export"):
    """Save each target as a PNG exactly as the main preview shows it.

    Every target is loaded into the preview canvas in turn (same route as
    "Send selected to PowerPoint"), so the active display preset, overlays,
    per-file colormap / contrast / Zero / quarter-turn and extra channel
    views all apply; the hidden-overlay (H) state and the shortcut hint are
    honoured by ``_save_current_figure_without_shortcut_hint``. The
    previously shown image is restored afterwards.

    Returns ``(results, errors, cancelled)``; ``results`` holds dicts with
    ``path``, ``file_key``, ``channel_idx``.
    """
    canvas = getattr(viewer, "preview_canvas", None)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    results, errors = [], []
    if canvas is None:
        return results, ["Preview canvas is unavailable."], False
    original_last = getattr(viewer, "last_preview", None)
    progress = None
    if len(targets) > 1:
        progress = QtWidgets.QProgressDialog(f"{title}...", "Cancel", 0, len(targets), viewer)
        progress.setWindowTitle(title)
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.setMinimumDuration(300)
    cancelled = False
    try:
        for i, file_key in enumerate(targets):
            if progress is not None:
                progress.setValue(i)
                progress.setLabelText(f"{title}: {Path(file_key).name} ({i + 1}/{len(targets)})")
                QtWidgets.QApplication.processEvents()
                if progress.wasCanceled():
                    cancelled = True
                    break
            try:
                channel_idx = viewer._thumbnail_channel_idx_for_file(file_key)
                viewer.show_file_channel(file_key, channel_idx, use_local_cmap=True)
                QtWidgets.QApplication.processEvents(QtCore.QEventLoop.AllEvents, 50)
                if not list(getattr(canvas, "views", None) or []):
                    errors.append(f"{Path(file_key).name}: nothing to render")
                    continue
                out_path = _unique_path(out_dir / png_name_for(viewer, file_key, channel_idx))

                def _save(path=out_path):
                    canvas.fig.savefig(str(path), format="png", dpi=dpi, bbox_inches="tight")

                canvas._save_current_figure_without_shortcut_hint(_save)
                results.append({"path": out_path, "file_key": str(file_key), "channel_idx": channel_idx})
            except Exception as exc:
                errors.append(f"{Path(file_key).name}: {exc}")
    finally:
        if progress is not None:
            progress.setValue(len(targets))
            progress.close()
        try:
            if original_last:
                viewer.show_file_channel(original_last[0], original_last[1], use_local_cmap=True)
        except Exception:
            pass
    return results, errors, cancelled


def on_export_pngs(viewer, paths=None):
    """Export ``paths`` / the selected thumbnails (or the previewed image) as
    PNGs, rendered as displayed in the main preview (display preset included)."""
    targets = export_targets(viewer, paths)
    if not targets:
        QtWidgets.QMessageBox.information(viewer, "Export", "Select one or more images first.")
        return
    default_dir = str(getattr(viewer, 'last_export_dir', None) or getattr(viewer, 'last_dir', Path(targets[0]).parent))
    out_dir = QtWidgets.QFileDialog.getExistingDirectory(viewer, "Select export folder", default_dir)
    if not out_dir:
        return
    viewer.last_export_dir = out_dir
    try:
        results, errors, cancelled = render_targets_as_displayed(viewer, targets, out_dir, title="Export PNGs")
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, "Export", f"Export failed: {exc}")
        return
    preset = getattr(getattr(viewer, "preview_canvas", None), "_active_display_preset", None)
    lines = [f"Exported {len(results)} of {len(targets)} PNG(s) to\n{out_dir}"]
    if preset:
        lines.append(f"Display preset: {str(preset).capitalize()}")
    if cancelled:
        lines.append("Cancelled.")
    if errors:
        lines.append("Errors:\n" + "\n".join(errors[:8]))
    QtWidgets.QMessageBox.information(viewer, "Export", "\n".join(lines))


def on_export_xyz_files(viewer):
    targets = list(getattr(viewer, 'thumb_multi_select', set()))
    if not targets:
        if getattr(viewer, 'selected_file_for_thumbs', None):
            targets = [viewer.selected_file_for_thumbs]
        elif viewer.last_preview:
            targets = [viewer.last_preview[0]]
    if not targets:
        QtWidgets.QMessageBox.information(viewer, "Export", "No thumbnails selected.")
        return
    out_dir = QtWidgets.QFileDialog.getExistingDirectory(viewer, "Select folder for XYZ export", str(viewer.last_dir))
    if not out_dir:
        return
    out_dir = Path(out_dir)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, "Export", f"Cannot create folder: {exc}")
        return
    exported = []
    channel_idx = viewer.channel_dropdown.currentIndex()
    for file_key in targets:
        header, exports = viewer._collect_channel_exports(file_key, channel_idx)
        if header is None or not exports:
            log_status(f"[XYZ Export] No channels for {file_key}")
            continue
        header_path = Path(file_key)
        for item in exports:
            arr_si, z_unit = convert_to_si(item['arr'], item.get('unit'))
            if not z_unit:
                z_unit = item.get('unit') or 'arb.'
            extent = item.get('extent')
            x_vals, y_vals, x_unit, y_unit = viewer._axes_from_extent(header, arr_si.shape, extent)
            date_token = viewer._sanitize_filename_component(header.get('Date', ''))
            time_token = viewer._sanitize_filename_component(header.get('Time', ''))
            base_name = viewer._sanitize_filename_component(header_path.stem)
            chan_token = viewer._sanitize_filename_component(item.get('caption') or f"chan{item.get('idx')}")
            parts = [p for p in (chan_token, base_name, date_token, time_token) if p]
            fname = "__".join(parts) + ".xyz"
            full_path = out_dir / fname
            meta_lines = [
                f"Source file: {header_path.name}",
                f"Channel: {item.get('caption') or ''} (index {item.get('idx')})",
                f"Date: {header.get('Date', '')} Time: {header.get('Time', '')}",
                f"Bias: {header.get('Bias', '')} {header.get('BiasPhysUnit', '')}",
                f"Dimensions: {header.get('xPixel','?')} x {header.get('yPixel','?')} pixels",
                f"X range: {header.get('XScanRange', header.get('ScanRange','?'))} {header.get('XPhysUnit','')}",
                f"Y range: {header.get('YScanRange', header.get('ScanRange','?'))} {header.get('YPhysUnit','')}",
            ]
            try:
                viewer._write_xyz_file(full_path, x_vals, y_vals, arr_si, x_unit, y_unit, z_unit, meta_lines)
                exported.append(str(full_path))
            except Exception as exc:
                QtWidgets.QMessageBox.warning(viewer, "Export", f"Failed to export {fname}: {exc}")
                log_status(f"[XYZ Export] Failed {full_path}: {exc}")
    if not exported:
        QtWidgets.QMessageBox.information(viewer, "Export", "No XYZ files were created.")
    else:
        preview = "\n".join(exported[:5])
        if len(exported) > 5:
            preview += "\n..."
        QtWidgets.QMessageBox.information(viewer, "Export", f"Exported {len(exported)} XYZ file(s) to {out_dir}:\n{preview}")

def on_export_wsxm_stp_files(viewer):
    targets = list(getattr(viewer, 'thumb_multi_select', set()))
    if not targets:
        focus = getattr(viewer, 'selected_file_for_thumbs', None)
        if focus:
            targets = [focus]
        elif viewer.last_preview:
            targets = [viewer.last_preview[0]]
    if not targets:
        QtWidgets.QMessageBox.information(viewer, "Export", "No thumbnails selected.")
        return
    out_dir = QtWidgets.QFileDialog.getExistingDirectory(
        viewer, "Select folder for WSxM STP export", str(viewer.last_dir)
    )
    if not out_dir:
        return
    out_path = Path(out_dir)
    try:
        out_path.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, "Export", f"Cannot create folder: {exc}")
        return
    channel_idx = viewer.channel_dropdown.currentIndex()
    exported = []
    errors = []
    for file_key in targets:
        header, exports = viewer._collect_channel_exports(file_key, channel_idx)
        if header is None or not exports:
            errors.append(f"{Path(file_key).name}: no channel data available")
            continue
        for item in exports:
            try:
                arr, meta = _prepare_stp_payload(viewer, file_key, item, header)
                fname = _build_stp_filename(viewer, file_key, item)
                save_wsxm_stp(out_path / fname, arr, **meta)
                exported.append(str(out_path / fname))
            except Exception as exc:
                errors.append(f"{Path(file_key).name}: {exc}")
    if not exported and errors:
        QtWidgets.QMessageBox.warning(viewer, "Export", "\n".join(errors[:5]))
        return
    summary = f"Exported {len(exported)} STP file(s) to {out_path}"
    if errors:
        summary += f"\nErrors ({len(errors)}):\n" + "\n".join(errors[:5])
    QtWidgets.QMessageBox.information(viewer, "Export", summary)

def export_view_as_stp(viewer, view: dict | None):
    if not view:
        QtWidgets.QMessageBox.information(viewer, "Export", "No view selected.")
        return
    meta = view.get('meta') or {}
    file_path = meta.get('file_path')
    chan_idx = meta.get('channel_index')
    if not file_path or chan_idx is None:
        QtWidgets.QMessageBox.warning(viewer, "Export", "View metadata missing file reference.")
        return
    header, _fds = viewer.headers.get(str(file_path), (None, None))
    if header is None:
        QtWidgets.QMessageBox.warning(viewer, "Export", "Header information unavailable.")
        return
    item = {
        'arr': np.asarray(view.get('arr')),
        'unit': view.get('unit'),
        'caption': meta.get('channel') or view.get('colorbar_label'),
        'idx': chan_idx,
        'extent': view.get('extent') or view.get('extent_raw'),
    }
    try:
        arr, meta_params = _prepare_stp_payload(viewer, file_path, item, header)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, "Export", f"Cannot prepare STP export: {exc}")
        return
    default_name = _build_stp_filename(viewer, file_path, item)
    out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
        viewer,
        "Export view as WSxM STP",
        str(Path(viewer.last_dir) / default_name),
        "WSxM STP (*.stp)",
    )
    if not out_path:
        return
    try:
        save_wsxm_stp(out_path, arr, **meta_params)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, "Export", f"Failed to save STP file:\n{exc}")
        return
    QtWidgets.QMessageBox.information(viewer, "Export", f"Saved STP file to\n{out_path}")


def on_export_selected_same_view(viewer):
    targets = list(getattr(viewer, 'thumb_multi_select', set()))
    if not targets:
        if getattr(viewer, 'selected_file_for_thumbs', None):
            targets = [viewer.selected_file_for_thumbs]
        elif viewer.last_preview:
            targets = [viewer.last_preview[0]]
    if not targets:
        QtWidgets.QMessageBox.information(viewer, "Export", "No thumbnails selected.")
        return
    config = viewer.get_current_detail_config()
    if not config.get('channels'):
        QtWidgets.QMessageBox.information(viewer, "Export", "No channels configured to export.")
        return
    out_dir = QtWidgets.QFileDialog.getExistingDirectory(viewer, "Select export folder", str(viewer.last_dir))
    if not out_dir:
        return
    worker = BatchExportWorker(viewer, targets, config, out_dir)
    worker.signals.progress.connect(viewer._on_batch_export_progress)
    worker.signals.finished.connect(viewer._on_batch_export_finished)
    viewer._batch_export_worker = worker
    progress = QtWidgets.QProgressDialog("Exporting...", "Cancel", 0, len(targets), viewer)
    progress.setWindowTitle("Batch export")
    progress.setWindowModality(QtCore.Qt.WindowModal)
    progress.canceled.connect(worker.cancel)
    progress.show()
    viewer._batch_export_progress = progress
    QtCore.QThreadPool.globalInstance().start(worker)


def _on_batch_export_progress(viewer, current, total, path):
    dlg = getattr(viewer, '_batch_export_progress', None)
    if dlg is None:
        return
    dlg.setMaximum(total)
    dlg.setValue(current)
    dlg.setLabelText(f"Exporting {Path(path).name} ({current}/{total})")


def _on_batch_export_finished(viewer, saved_paths, errors, cancelled):
    dlg = getattr(viewer, '_batch_export_progress', None)
    if dlg is not None:
        dlg.close()
        viewer._batch_export_progress = None
    viewer._batch_export_worker = None
    msg_lines = [f"Saved {len(saved_paths)} file(s)."]
    if saved_paths:
        preview_paths = "\n".join(saved_paths[:5])
        msg_lines.append(preview_paths + ("\n..." if len(saved_paths) > 5 else ""))
    if cancelled:
        msg_lines.append("Operation cancelled.")
    if errors:
        msg_lines.append("Errors:\n" + "\n".join(errors[:10]))
    QtWidgets.QMessageBox.information(viewer, "Batch export", "\n".join(msg_lines))


def _normalize_stp_array(arr, unit):
    data = np.asarray(arr, dtype=np.float64)
    converter = convert_to_si
    if callable(converter):
        try:
            data_si, base_unit = converter(data, unit)
        except Exception:
            data_si, base_unit = data, unit
    else:
        data_si, base_unit = data, unit
    base_unit = base_unit or (unit or "arb.")
    data_si = np.asarray(data_si, dtype=np.float64)
    if base_unit == "m":
        return data_si / 1e-9, "nm"
    return data_si, base_unit


def _convert_scalar_to_unit(value, unit, target_unit):
    if value in (None, ""):
        return None
    try:
        arr = np.array([float(value)], dtype=np.float64)
    except Exception:
        return None
    converter = convert_to_si
    if callable(converter):
        try:
            arr_si, base_unit = converter(arr, unit)
        except Exception:
            arr_si, base_unit = arr, unit
    else:
        arr_si, base_unit = arr, unit
    val = float(arr_si[0])
    base = (base_unit or unit or target_unit or "").lower()
    if target_unit and target_unit.lower() == "pa":
        if base == "a":
            return val * 1e12
        if base == "pa":
            return val
    if target_unit and target_unit.lower() == "v":
        return val
    return val


def _range_nm_from_header(header, axis, extent):
    key = f"{axis}ScanRange"
    unit_key = f"{axis}PhysUnit"
    rng = _value_in_nm(header.get(key), header.get(unit_key) or header.get("PhysUnit"))
    if rng is None:
        rng = _value_in_nm(header.get("ScanRange"), header.get("PhysUnit"))
    if rng is None and extent is not None:
        idx = (0, 1) if axis == "X" else (2, 3)
        try:
            rng = abs(float(extent[idx[1]]) - float(extent[idx[0]]))
        except Exception:
            rng = None
    return rng


def _build_stp_comment(file_key, caption, header, x_nm, y_nm):
    now_txt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"Exported from SXM Viewer on {now_txt}",
        f"Source: {Path(file_key).name}",
        f"Channel: {caption}",
    ]
    date = str(header.get("Date", "") or "").strip()
    time_txt = str(header.get("Time", "") or "").strip()
    acquired = " ".join(t for t in (date, time_txt) if t).strip()
    if acquired:
        lines.append(f"Acquired: {acquired}")
    if x_nm is not None and y_nm is not None:
        lines.append(f"Scan size: {x_nm:.3f} nm × {y_nm:.3f} nm")
    bias = header.get("Bias")
    bias_unit = header.get("BiasPhysUnit", "")
    if bias not in (None, ""):
        lines.append(f"Bias (raw): {bias} {bias_unit}".strip())
    setp = header.get("SetPoint")
    setp_unit = header.get("SetPointPhysUnit", "")
    if setp not in (None, ""):
        lines.append(f"Setpoint (raw): {setp} {setp_unit}".strip())
    user = header.get("UserName")
    if user:
        lines.append(f"Operator: {user}")
    return "\n".join(lines)[:1800]


def _build_wsxm_comment_block(header_path):
    """Return a WSxM-style comment block (escaped newlines) using the source header."""
    try:
        raw = Path(header_path).read_text(encoding="utf-8", errors="ignore")
    except Exception:
        try:
            raw = Path(header_path).read_text(encoding="cp1252", errors="ignore")
        except Exception:
            raw = ""
    raw = raw.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not raw:
        return None
    prefix = "Converted from SXM Viewer files\n\nHEADER DUMP:\n"
    return prefix + raw


def _prepare_stp_payload(viewer, file_key, item, header):
    arr = item.get('arr')
    if arr is None:
        raise ValueError("Missing image data")
    arr_norm, z_unit = _normalize_stp_array(arr, item.get('unit'))
    if arr_norm.ndim != 2:
        raise ValueError("WSxM export expects 2-D data")
    extent = item.get('extent') or item.get('extent_raw')
    x_nm = _range_nm_from_header(header, "X", extent)
    y_nm = _range_nm_from_header(header, "Y", extent)
    if x_nm is None:
        x_nm = float(np.shape(arr_norm)[1])
    if y_nm is None:
        y_nm = float(np.shape(arr_norm)[0])
    setpoint_pa = _convert_scalar_to_unit(header.get('SetPoint'), header.get('SetPointPhysUnit'), 'pA')
    bias_v = _convert_scalar_to_unit(header.get('Bias'), header.get('BiasPhysUnit'), 'V')
    angle = _safe_float(header.get('Angle'), 0.0)
    timestamp_txt = " ".join(
        t for t in (
            str(header.get('Date', '') or '').strip(),
            str(header.get('Time', '') or '').strip(),
        ) if t
    ).strip()
    if not timestamp_txt:
        timestamp_txt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    head_type = header.get('HeadType') or header.get('Head type') or header.get('Head') or 'STM'
    caption = item.get('caption') or Path(file_key).name
    z_gain_val = _safe_float(header.get('ZGain'), None)
    z_gain = z_gain_val if z_gain_val is not None else 1.0
    comment_block = _build_wsxm_comment_block(file_key)
    comment = _build_stp_comment(file_key, caption, header, x_nm, y_nm)
    meta = {
        'channel': caption,
        'x_nm': x_nm,
        'y_nm': y_nm,
        'z_unit': z_unit or 'arb.',
        'setpoint_pa': setpoint_pa,
        'bias_v': bias_v,
        'angle_deg': angle,
        'comment': comment,
        'comment_block': comment_block,
        'head_type': head_type,
        'timestamp': timestamp_txt,
        'z_gain': z_gain,
        'name': Path(file_key).name,
    }
    return arr_norm, meta


def _build_stp_filename(viewer, file_key, item):
    chan = viewer._sanitize_filename_component(item.get('caption') or f"chan{item.get('idx', 0)}")
    base = viewer._sanitize_filename_component(Path(file_key).stem)
    parts = [p for p in (chan, base) if p]
    if not parts:
        parts = ["export"]
    return "__".join(parts) + ".stp"
__all__ = [
    "_collect_channel_exports",
    "on_export_pngs",
    "on_export_xyz_files",
    "on_export_wsxm_stp_files",
    "on_export_selected_same_view",
    "export_view_as_stp",
    "_on_batch_export_progress",
    "_on_batch_export_finished",
    "export_targets",
    "render_targets_as_displayed",
    "export_to_obsidian",
    "populate_obsidian_menu",
]


# ---------------------------------------------------------------------------
# Obsidian export: PNGs rendered as displayed, embedded at the end of a note
# ---------------------------------------------------------------------------
def _obsidian_settings(viewer):
    cfg = getattr(viewer, "config", None)
    return dict((cfg or {}).get("obsidian_export") or {}) if isinstance(cfg, dict) else {}


def _store_obsidian_settings(viewer, values):
    from ...config import save_config
    settings = _obsidian_settings(viewer)
    vault = str(values["vault"])
    attachments = dict(settings.get("attachments") or {})
    attachments[vault] = str(values.get("attachment") or "")
    settings.update({
        "vault": vault,
        "note": str(values["note"]),
        "attachments": attachments,
        "width": int(values.get("width") or 0),
        "heading": bool(values.get("heading", True)),
        "captions": bool(values.get("captions", True)),
    })
    try:
        viewer.config["obsidian_export"] = settings
        save_config(viewer.config)
    except Exception:
        pass


def _obsidian_caption(viewer, file_key, channel_idx):
    from .preview import _fmt_overlay_number
    header, fds = (getattr(viewer, "headers", {}) or {}).get(str(file_key), (None, None))
    header = header or {}
    parts = [f"`{Path(str(file_key)).name}`"]
    try:
        fd = fds[int(channel_idx)] if fds and channel_idx is not None else None
        chan = (fd or {}).get("Caption") or ""
    except Exception:
        chan = ""
    if chan:
        parts.append(str(chan))
    try:
        bias = _fmt_overlay_number(header.get("Bias"), header.get("BiasPhysUnit") or "V")
        setp = _fmt_overlay_number(header.get("SetPoint"), header.get("SetPointPhysUnit") or "")
    except Exception:
        bias = setp = ""
    if bias:
        parts.append(bias)
    if setp:
        parts.append(setp)
    when = " ".join(t for t in (str(header.get("Date", "") or "").strip(),
                                 str(header.get("Time", "") or "").strip()) if t)
    if when:
        parts.append(when)
    return " · ".join(parts)


def export_to_obsidian(viewer, note=None, vault=None, targets=None):
    """Render ``targets`` (default: thumbnail selection / previewed image) as
    displayed and append them to an Obsidian note. With no ``note`` a dialog
    asks for vault/note/options; otherwise the remembered options are used."""
    from ...integrations import obsidian
    from ..dialogs.obsidian_export import ObsidianExportDialog, suggested_attachment_setting

    title = "Export to Obsidian"
    targets = export_targets(viewer, targets)
    if not targets:
        QtWidgets.QMessageBox.information(viewer, title, "Select one or more images first.")
        return
    settings = _obsidian_settings(viewer)
    if note is None:
        dlg = ObsidianExportDialog(viewer, n_images=len(targets), settings=settings)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return
        try:
            values = dlg.values()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(viewer, title, str(exc))
            return
    else:
        vault = Path(vault) if vault else obsidian.vault_for_path(Path(note), obsidian.list_vaults())
        if vault is None:
            QtWidgets.QMessageBox.warning(viewer, title, f"{note} is not inside an Obsidian vault.")
            return
        attach = (settings.get("attachments") or {}).get(str(vault))
        values = {
            "vault": Path(vault),
            "note": Path(note),
            "attachment": attach if attach is not None else suggested_attachment_setting(vault),
            "width": int(settings.get("width", 500) or 0),
            "heading": bool(settings.get("heading", True)),
            "captions": bool(settings.get("captions", True)),
        }
    vault, note_path = Path(values["vault"]), Path(values["note"])
    img_dir = obsidian.attachment_dir(vault, note_path, values.get("attachment") or None)
    for path, what in ((note_path, "note"), (img_dir, "image folder")):
        try:
            Path(path).resolve().relative_to(vault.resolve())
        except ValueError:
            QtWidgets.QMessageBox.warning(viewer, title, f"The {what} must be inside the vault.")
            return
    try:
        results, errors, cancelled = render_targets_as_displayed(viewer, targets, img_dir, title=title)
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, title, f"Rendering failed: {exc}")
        return
    if not results:
        detail = ("\n" + "\n".join(errors[:8])) if errors else ""
        QtWidgets.QMessageBox.warning(viewer, title, "Nothing was exported." + detail)
        return
    lines = []
    if values.get("heading"):
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        folders = sorted({Path(r["file_key"]).parent.name for r in results})
        src = f" · `{folders[0]}`" if len(folders) == 1 else ""
        lines += [f"#### SXM Viewer export · {stamp}{src}", ""]
    for res in results:
        lines.append(obsidian.embed_line(vault, note_path, res["path"], width=values.get("width") or 0))
        if values.get("captions"):
            lines.append(_obsidian_caption(viewer, res["file_key"], res["channel_idx"]))
        lines.append("")
    try:
        obsidian.append_to_note(note_path, "\n".join(lines))
    except Exception as exc:
        QtWidgets.QMessageBox.warning(viewer, title, f"Could not write the note: {exc}")
        return
    _store_obsidian_settings(viewer, values)
    msg = f"Added {len(results)} image(s) to {note_path.relative_to(vault).as_posix()} ({vault.name})"
    if cancelled:
        msg += "\nCancelled before all images were rendered."
    if errors:
        msg += "\nErrors:\n" + "\n".join(errors[:8])
    log_status(f"[Obsidian] {msg}")
    QtWidgets.QMessageBox.information(viewer, title, msg)


def populate_obsidian_menu(viewer, menu, targets=None):
    """Fill ``menu`` with one-click "append to the active note" entries for
    the open vault(s), the last-used note, and "Choose note..." (dialog)."""
    from ...integrations import obsidian

    menu.clear()
    targets = list(targets) if targets else None
    vaults = obsidian.list_vaults()
    shown = [v for v in vaults if v.is_open] or vaults[:1]
    settings = _obsidian_settings(viewer)
    added = set()
    for vault in shown:
        note = obsidian.active_note(vault.path)
        if note is None:
            continue
        rel = note.relative_to(vault.path).as_posix()
        act = menu.addAction(f"Active note: {rel}  ({vault.name})")
        act.setToolTip(str(note))
        act.triggered.connect(lambda _=False, n=note, v=vault.path: export_to_obsidian(viewer, n, v, targets))
        added.add(str(note))
    last = settings.get("note")
    if last and last not in added and Path(last).is_file():
        last_vault = settings.get("vault")
        act = menu.addAction(f"Last used: {Path(last).name}  ({Path(last_vault).name if last_vault else ''})")
        act.setToolTip(str(last))
        act.triggered.connect(lambda _=False, n=Path(last), v=last_vault: export_to_obsidian(viewer, n, v, targets))
    if not menu.isEmpty():
        menu.addSeparator()
    choose = menu.addAction("Choose note / options...")
    choose.triggered.connect(lambda _=False: export_to_obsidian(viewer, None, None, targets))




