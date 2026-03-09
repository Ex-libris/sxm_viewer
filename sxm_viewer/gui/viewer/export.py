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
from ..detail_panels import BatchExportSignals, BatchExportWorker
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
        disp_unit, disp_arr, _ = viewer._scale_unit_for_display(unit_final, adj_arr)
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


def on_export_pngs(viewer):
    # Export high-quality PNGs for the currently selected file's visible channels (main + extras)
    if not viewer.last_preview:
        QtWidgets.QMessageBox.information(viewer, "No selection", "Select a file/channel first.")
        return
    header_path_str, channel_idx = viewer.last_preview
    header_path = Path(header_path_str)
    header, exports = viewer._collect_channel_exports(header_path_str, channel_idx)
    if header is None or not exports:
        QtWidgets.QMessageBox.information(viewer, "Export", "No channels to export.")
        return

    default_dir = str(getattr(viewer, 'last_dir', header_path.parent))
    out_dir = QtWidgets.QFileDialog.getExistingDirectory(viewer, "Select export folder", default_dir)
    if not out_dir:
        return

    # Metadata for naming
    date = viewer._sanitize_filename_component(header.get('Date', ''))
    time = viewer._sanitize_filename_component(header.get('Time', ''))
    file_base = viewer._sanitize_filename_component(Path(header_path_str).stem)

    # Save each channel as a separate high-DPI PNG
    from matplotlib.figure import Figure
    for item in exports:
        try:
            fig = Figure(figsize=(6, 5), dpi=300)
            ax = fig.add_subplot(1,1,1)
            arr = np.asarray(item['arr'])
            flip = bool(item.get('relative_axes'))
            if flip:
                arr_plot = np.flipud(arr)
            else:
                arr_plot = arr
            origin = 'lower' if flip else 'upper'
            cmapname = item.get('cmap', 'viridis')
            extent = item.get('extent')
            if extent is None:
                im = ax.imshow(arr_plot, origin=origin, interpolation='nearest', cmap=cmapname)
            else:
                im = ax.imshow(arr_plot, extent=extent, origin=origin, interpolation='nearest', aspect='equal', cmap=cmapname)
            if item.get('relative_axes') and extent is not None:
                pass
            cbar_label = item.get('colorbar_label') or item.get('unit') or ''
            if cbar_label:
                cbar = fig.colorbar(im, ax=ax, fraction=0.08, pad=0.02)
                cbar.set_label(cbar_label)
            ax.set_title(item.get('title') or item.get('caption') or '')
            try:
                fig.tight_layout()
            except Exception:
                pass

            chan_name = viewer._sanitize_filename_component(item.get('caption') or f"chan{item.get('idx',0)}")
            parts = [p for p in (chan_name, file_base, date, time) if p]
            fname = "__".join(parts) + ".png"
            out_path = str(Path(out_dir) / fname)
            fig.savefig(out_path, dpi=300, bbox_inches='tight')
        except Exception as e:
            # keep going for other channels
            print('Export failed for a channel:', e)

    QtWidgets.QMessageBox.information(viewer, "Export", f"Exported {len(exports)} PNG(s) to\n{out_dir}")


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
]




