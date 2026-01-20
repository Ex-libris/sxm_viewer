"""Preview helpers for SXMGridViewer."""
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
from ...config import save_config
from ...data.spectroscopy import is_matrix_file_entry

def _build_metadata_html(viewer, header_path:Path, header:dict, fd:dict, channel_idx:int,
                         unit_normalized:str, unit_display:str, arr_display:np.ndarray, zero_offset:float|None) -> str:
    """Return HTML for the metadata pane with clearer styling and sections."""
    def esc(s):
        try:
            return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        except Exception:
            return ''
    dark = bool(getattr(viewer, 'dark_mode', False))
    show_preview_specs = bool(getattr(viewer, 'show_preview_spectra', getattr(viewer, 'show_spectra', True)))
    text_color = '#e0e0e0' if dark else '#222'
    label_color = '#a0a0a0' if dark else '#555'
    accent_border = '#6fa8ff' if dark else '#4a7edb'
    accent_bg = 'rgba(111,168,255,0.16)' if dark else 'rgba(74,126,219,0.10)'
    # Respect the user's configured metadata font size (in px). Default to 10px if not present.
    try:
        font_size = int(getattr(viewer, 'config', {}).get('meta_font_size', 10))
    except Exception:
        font_size = 10
    filename = header_path.name
    date = header.get('Date', '')
    time = header.get('Time', '')
    bias = header.get('Bias', None); bias_unit = header.get('BiasPhysUnit', '')
    setp = header.get('SetPoint', None); setp_unit = header.get('SetPointPhysUnit', '')
    user = header.get('UserName', '')
    cap = fd.get('Caption','')
    phys_orig = fd.get('PhysUnit','')
    scale = fd.get('Scale','')
    offset = fd.get('Offset','')
    def fmt_number(val, precision=3):
        try:
            num = float(val)
            return f"{num:.{precision}f}".rstrip('0').rstrip('.')
        except Exception:
            if val is None:
                return ''
            return esc(val)

    # stats
    try:
        flat = np.asarray(arr_display).ravel()
        vmin = np.nanmin(flat); vmax = np.nanmax(flat); vmed = np.nanmedian(flat)
        stats = f"min={vmin:.6g} | max={vmax:.6g} | median={vmed:.6g}"
    except Exception:
        stats = "min/max/median: N/A"
    # tags
    taginfo = viewer.tags.get(str(header_path), {})
    tag_label = taginfo.get('tag', None)
    tag_chip = ''
    if tag_label == 'constant-height':
        chip_color = '#2e7d32'; chip_text = 'CH'
    elif tag_label == 'constant-current':
        chip_color = '#1565c0'; chip_text = 'CC'
    else:
        chip_color = None; chip_text = ''
    if chip_color:
        tag_chip = f"<span style='background:{chip_color};color:#fff;border-radius:10px;padding:2px 8px;font-weight:600'>" \
                   f"{chip_text}</span> <span style='color:#555'>({esc(tag_label)})</span>"
    # abs z + dzs
    ch_lines = ''
    abs_nm = None
    if tag_label == 'constant-height':
        abs_pm = taginfo.get('abs_z_pm', None)
        if abs_pm is not None:
            abs_nm = abs_pm/1000.0
            ch_lines += f"<div>Const-height (abs z): <b>{abs_nm:.3f} nm</b></div>"
        dz_prev_nonch, prevname = viewer._dz_vs_last_before_ch(header_path)
        if dz_prev_nonch is not None:
            ch_lines += f"<div>dz vs prev non-CH (<i>{esc(prevname)}</i>): <b>{dz_prev_nonch:+.0f} pm</b> ({dz_prev_nonch/1000.0:+.3f} nm)</div>"
        dz_prev_ch, prevch_name = viewer._dz_vs_previous_ch(header_path)
        if dz_prev_ch is not None:
            ch_lines += f"<div>dz vs prev CH (<i>{esc(prevch_name)}</i>): <b>{dz_prev_ch:+.0f} pm</b> ({dz_prev_ch/1000.0:+.3f} nm)</div>"

    # control params
    params = {}
    def collect_params(d):
        for k,v in (d or {}).items():
            kl = str(k).lower()
            if any(tok in kl for tok in ('ki','kp','pll','ampl','amplitude','amp','setpoint','natural','natfreq','freq','f0','kpl','kipl','lockin')):
                try:
                    params[k] = float(v)
                except Exception:
                    params[k] = v
    collect_params(header); collect_params(fd)
    params_rows = ''.join([f"<tr><td>{esc(k)}</td><td style='text-align:right'>{esc(v)}</td></tr>" for k,v in params.items()])

    spec_section = ''
    spec_entries = viewer.spectros_by_image.get(str(header_path), [])
    if show_preview_specs and spec_entries:
        rows = []
        for idx, spec in enumerate(spec_entries[:6], 1):
            name = Path(spec['path']).name
            matrix_idx = spec.get('matrix_index')
            if matrix_idx is not None:
                name = f"{name} [{matrix_idx}]"
            xs = spec.get('x')
            ys = spec.get('y')
            pos_txt = f"{xs:.1f}/{ys:.1f} nm" if xs is not None and ys is not None else "n/a"
            rows.append(f"<tr><td>S{idx}</td><td>{esc(name)}</td><td style='text-align:right'>{esc(pos_txt)}</td></tr>")
        if len(spec_entries) > 6:
            rows.append(f"<tr><td colspan='3' style='text-align:center;color:{label_color}'>+ {len(spec_entries)-6} more?</td></tr>")
        spec_section = f"""
        <div style='height:6px'></div>
        <div style='font-weight:600; color:{label_color}; margin-bottom:2px'>Spectroscopies ({len(spec_entries)})</div>
        <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
          {''.join(rows)}
        </table>
        """

    scan_entries = [
        ('XScanRange', 'X scan', header.get('XScanRange'), header.get('XPhysUnit', header.get('PhysUnit',''))),
        ('YScanRange', 'Y scan', header.get('YScanRange'), header.get('YPhysUnit', header.get('PhysUnit',''))),
        ('Speed', 'Speed', header.get('Speed'), ''),
        ('LineRate', 'Line rate', header.get('LineRate'), ''),
        ('Angle', 'Angle', header.get('Angle'), 'deg'),
        ('xPixel', 'x pixels', header.get('xPixel'), ''),
        ('yPixel', 'y pixels', header.get('yPixel'), ''),
        ('xCenter', 'x center', header.get('xCenter'), header.get('XPhysUnit', '')),
        ('yCenter', 'y center', header.get('yCenter'), header.get('YPhysUnit', '')),
        ('dzdx', 'dz/dx', header.get('dzdx') or header.get('dz/dx'), ''),
        ('dzdy', 'dz/dy', header.get('dzdy') or header.get('dz/dy'), ''),
        ('overscan[%]', 'Overscan (%)', header.get('overscan[%]'), '%'),
    ]
    scan_rows = []
    for key, label, val, extra_unit in scan_entries:
        if val is None or val == '':
            continue
        if isinstance(val, float):
            val_txt = f"{val:.3f}"
        else:
            val_txt = esc(val)
        unit_txt = extra_unit or ''
        scan_rows.append(f"<tr><td>{esc(label)}</td><td style='text-align:right'>{val_txt} {esc(unit_txt)}</td></tr>")
    scan_section = ""
    if scan_rows:
        scan_section = f"""
        <div style='height:6px'></div>
        <div style='font-weight:600; color:{label_color}; margin-bottom:2px'>Scan metadata</div>
        <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
          {''.join(scan_rows)}
        </table>
        """

    # key metadata highlight
    x_range = header.get('XScanRange'); y_range = header.get('YScanRange')
    x_unit = header.get('XPhysUnit', header.get('PhysUnit','nm'))
    y_unit = header.get('YPhysUnit', header.get('PhysUnit','nm'))
    xpix = header.get('xPixel') or header.get('XPixel')
    ypix = header.get('yPixel') or header.get('YPixel')
    x_center = header.get('xCenter'); y_center = header.get('yCenter')
    piezo_txt = f"{abs_nm:.3f} nm" if abs_nm is not None else ""
    date_display = " ".join(t for t in (date, time) if t).strip() or ""
    size_txt = ""
    if x_range is not None and y_range is not None:
        size_txt = f"{fmt_number(x_range)} {esc(x_unit)}  {fmt_number(y_range)} {esc(y_unit)}"
    pixel_txt = ""
    if xpix is not None and ypix is not None:
        pixel_txt = f"{fmt_number(xpix,0)}  {fmt_number(ypix,0)}"
    center_txt = ""
    if x_center is not None and y_center is not None:
        center_txt = f"{fmt_number(x_center)} / {fmt_number(y_center)} {esc(x_unit)}"
    bias_txt = f"{fmt_number(bias)} {esc(bias_unit)}" if bias is not None else ""
    setp_txt = f"{fmt_number(setp)} {esc(setp_unit)}" if setp is not None else ""
    key_rows = [
        ("Acquired", date_display),
        ("Bias", bias_txt),
        ("Setpoint", setp_txt),
        ("Image size", size_txt),
        ("Pixels", pixel_txt),
        ("X/Y center", center_txt),
        ("Piezo Z", piezo_txt),
    ]
    key_section_rows = "".join(
        f"<tr><td style='padding:2px 6px;color:{label_color};font-weight:600'>{esc(lbl)}</td>"
        f"<td style='padding:2px 6px;text-align:right'><span style='color:{text_color};font-weight:600'>{val or ''}</span></td></tr>"
        for lbl, val in key_rows if val
    )
    key_section = f"""
    <div style='border:1px solid {accent_border}; border-radius:12px; background:{accent_bg}; padding:8px; margin-bottom:8px;'>
      <table style='width:100%; border-collapse:collapse'>{key_section_rows}</table>
    </div>
    """

    relative_row = ""
    if zero_offset is not None:
        relative_row = f"<tr><td style='color:{label_color}'>Relative zero</td><td style='text-align:right'>{zero_offset:.6g} {esc(unit_display)}</td></tr>"

    html = f"""
    <div style='font-family:Segoe UI, Arial; font-size:{font_size}px; color:{text_color}; background: transparent;'>
      <div style='font-weight:600; font-size:1.15em; margin-bottom:4px'>{esc(filename)} {tag_chip}</div>
      {key_section}
      <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
        <tr><td style='color:{label_color}'>Date</td><td style='text-align:right'>{esc(date) or '&nbsp;'}</td></tr>
        <tr><td style='color:{label_color}'>Time</td><td style='text-align:right'>{esc(time) or '&nbsp;'}</td></tr>
        <tr><td style='color:{label_color}'>Bias</td><td style='text-align:right'>{'' if bias is None else esc(bias)} {esc(bias_unit)}</td></tr>
        <tr><td style='color:{label_color}'>SetPoint</td><td style='text-align:right'>{'' if setp is None else esc(setp)} {esc(setp_unit)}</td></tr>
        <tr><td style='color:{label_color}'>User</td><td style='text-align:right'>{esc(user)}</td></tr>
      </table>
      <div style='height:6px'></div>
      {spec_section}
      <div style='height:6px'></div>
      <div style='font-weight:600; color:%s; margin-bottom:2px'>Channel</div>
      <table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>
        <tr><td style='color:{label_color}'>Index</td><td style='text-align:right'>{channel_idx}</td></tr>
        <tr><td style='color:{label_color}'>Caption</td><td style='text-align:right'>{esc(cap)}</td></tr>
        <tr><td style='color:{label_color}'>Unit (orig)</td><td style='text-align:right'>{esc(phys_orig)}</td></tr>
        <tr><td style='color:{label_color}'>Normalized (SI)</td><td style='text-align:right'><b>{esc(unit_normalized)}</b></td></tr>
        <tr><td style='color:{label_color}'>Shown unit</td><td style='text-align:right'><b>{esc(unit_display)}</b></td></tr>
        {relative_row}
        <tr><td style='color:{label_color}'>Scale</td><td style='text-align:right'>{esc(scale)}</td></tr>
        <tr><td style='color:{label_color}'>Offset</td><td style='text-align:right'>{esc(offset)}</td></tr>
        <tr><td style='color:{label_color}'>Stats</td><td style='text-align:right'>{esc(stats)}</td></tr>
      </table>
      <div style='height:6px'></div>
      {ch_lines}
      {("<div style='height:6px'></div><div style='font-weight:600; color:#333; margin-bottom:2px'>Control params</div>" if params_rows else '')}
      {("<table style='width:100%; border-collapse:collapse' cellspacing='0' cellpadding='2'>" + params_rows + "</table>") if params_rows else ''}
      {scan_section}
    </div>
    """
    return html


def show_file_channel(viewer, header_path_str, channel_idx:int, use_local_cmap=False):
    viewer.last_preview = (str(header_path_str), int(channel_idx))
    if hasattr(viewer, 'adjust_image_btn'):
        viewer.adjust_image_btn.setEnabled(True)
    viewer._update_toolbar_actions(True)
    header_path = Path(header_path_str)
    # track selected file for thumbnail highlighting
    try:
        viewer.selected_file_for_thumbs = str(header_path)
        viewer._refresh_thumb_selection_styles()
    except Exception:
        pass
    viewer._update_frame_map_active(str(header_path))
    file_key = str(header_path)
    header, fds = viewer.headers.get(file_key, (None,None))
    if header is None or channel_idx < 0 or channel_idx >= len(fds): return
    fd = fds[channel_idx]; fname = fd.get("FileName")
    axis_unit = 'px'
    try:
        xpix = int(header.get('xPixel', 128)); ypix = int(header.get('yPixel', xpix))
        base_extent = viewer._header_extent(header)
        unit_normalized, arr_base = viewer._get_filtered_channel_array(file_key, channel_idx, header, fd)
        viewer._last_base_array = np.asarray(arr_base)
        viewer._last_base_extent = base_extent
        viewer._last_base_unit = unit_normalized
        arr_adj, adjusted_extent = viewer._apply_adjustments_for_channel(file_key, channel_idx, viewer._last_base_array, base_extent)
        display_extent = viewer._display_extent(adjusted_extent, header)
        display_unit, display_arr, zero_offset = viewer._scale_unit_for_display(unit_normalized, arr_adj)
        viewer._last_display_array = np.asarray(display_arr)
        viewer._last_display_unit = display_unit
        viewer._last_display_extent = display_extent
        viewer._last_colorbar_label = None
        axis_unit = header.get('XPhysUnit') or header.get('YPhysUnit') or header.get('ScanUnit') or ''
        if not axis_unit:
            axis_unit = 'px' if display_extent is None else 'nm'
        viewer._last_axis_unit = axis_unit
    except Exception as e:
        viewer.meta_box.setPlainText("Error reading channel: %s" % str(e)); return

    cmap_to_use = viewer.preview_cmap_combo.currentText() or viewer.preview_cmap
    if use_local_cmap:
        cmap_to_use = viewer.per_file_channel_cmap.get((file_key, channel_idx), cmap_to_use)

    # Spectroscopy entries for this file (singles only for overlay)
    show_preview_specs = bool(getattr(viewer, 'show_preview_spectra', getattr(viewer, 'show_spectra', True)))
    spec_entries = viewer.spectros_by_image.get(str(header_path), []) if show_preview_specs else []
    overlay_specs = []
    if spec_entries and show_preview_specs:
        if viewer.show_single_markers:
            overlay_specs.extend([
                s for s in spec_entries
                if s.get('matrix_index') is None or not is_matrix_file_entry(s)
            ])
        if viewer.show_matrix_markers:
            overlay_specs.extend([
                s for s in spec_entries
                if s.get('matrix_index') is not None and is_matrix_file_entry(s)
            ])

    # build views (main + dynamic extras based on current file)
    views = []
    caption = fd.get('Caption', fd.get('FileName', ''))
    date = str(header.get('Date', '') or '').strip()
    time_txt = str(header.get('Time', '') or '').strip()
    datetime_txt = " ".join([t for t in (date, time_txt) if t]).strip()
    base_title = Path(header_path).name
    if datetime_txt:
        title_text = f"{base_title}  {caption}  {datetime_txt}"
    else:
        title_text = f"{base_title}  {caption}"
    colorbar_label = caption
    if display_unit:
        colorbar_label = f"{caption} [{display_unit}]"
    viewer._last_colorbar_label = colorbar_label
    meta = {
        'file_path': str(header_path),
        'file_name': header_path.name,
        'date': date,
        'time': time_txt,
        'datetime': datetime_txt,
        'channel': caption,
        'channel_index': int(channel_idx),
    }
    main = {
        'arr': display_arr,
        'extent': display_extent,
        'extent_raw': base_extent,
        'cmap': cmap_to_use,
        'unit': display_unit,
        'title': title_text,
        'colorbar_label': colorbar_label,
        'axis_unit': axis_unit,
        'relative_axes': bool(viewer.relative_axes),
        'meta': meta,
        'spectra': overlay_specs,
    }
    views.append(main)

    # Rebuild extra views for the currently selected file using stored specifications
    for spec in getattr(viewer, 'extra_view_specs', []):
        try:
            # Find matching channel in this file (by caption first, then by index)
            idx2 = viewer._find_channel_index_for_spec(fds, spec)
            if idx2 is None:
                continue
            fd2 = fds[idx2]
            unit2_final, arr2_conv = viewer._get_filtered_channel_array(file_key, idx2, header, fd2)
            cmap2 = viewer._resolve_extra_spec_cmap(spec, file_key)
            arr2_adj, adj2_extent = viewer._apply_adjustments_for_channel(file_key, idx2, arr2_conv, base_extent)
            extent2 = viewer._display_extent(adj2_extent, header)
            unit2_display, arr2_display, _ = viewer._scale_unit_for_display(unit2_final, arr2_adj)
            caption2 = fd2.get('Caption', fd2.get('FileName', ''))
            if datetime_txt:
                title2 = f"{Path(header_path).name}  {caption2}  {datetime_txt}"
            else:
                title2 = f"{Path(header_path).name}  {caption2}"
            cbar_label2 = caption2
            if unit2_display:
                cbar_label2 = f"{caption2} [{unit2_display}]"
            meta2 = dict(meta)
            meta2['channel'] = caption2
            meta2['channel_index'] = int(idx2)
            views.append({'arr': arr2_display, 'extent': extent2, 'cmap': cmap2, 'unit': unit2_display,
                          'title': title2, 'colorbar_label': cbar_label2, 'axis_unit': axis_unit,
                          'relative_axes': bool(viewer.relative_axes), 'meta': meta2})
        except Exception:
            # Skip extra view if anything fails for this file
            continue

    preserve = False
    try:
        last = viewer.last_preview[0] if viewer.last_preview else None
        preserve = (
            bool(getattr(viewer, 'preserve_profiles_on_channel_change', False))
            and last == str(header_path)
            and getattr(viewer, 'current_mode', viewer.MODE_BROWSE) == viewer.MODE_MEASURE
        )
    except Exception:
        preserve = False
    viewer.preview_canvas.set_views(views, preserve_profiles=preserve)
    suppress_profile_restart = getattr(viewer, '_suppress_profile_restart', False)
    if suppress_profile_restart:
        viewer._suppress_profile_restart = False
    if getattr(viewer, 'current_mode', viewer.MODE_BROWSE) == viewer.MODE_MEASURE:
        try:
            canvas = getattr(viewer, 'preview_canvas', None)
            angle_active = bool(canvas and getattr(canvas, 'angle_enabled', False))
            profile_active = bool(canvas and getattr(canvas, 'profile_enabled', False))
            if not suppress_profile_restart and not angle_active and not profile_active:
                viewer._on_start_profile(force_enable=True)
        except Exception:
            pass
    elif getattr(viewer, '_pending_profile_enable', False):
        viewer._pending_profile_enable = False
    else:
        # Ensure profiles are cleared when not in Measure mode.
        try:
            viewer.preview_canvas.enable_profile(False)
            if hasattr(viewer.preview_canvas, 'clear_saved_profiles'):
                viewer.preview_canvas.clear_saved_profiles()
        except Exception:
            pass

    # Styled HTML metadata
    try:
        html = viewer._build_metadata_html(header_path, header, fd, channel_idx, unit_normalized, display_unit, display_arr, zero_offset)
        viewer.meta_box.setHtml(html)
    except Exception:
        viewer.meta_box.setPlainText(f"File: {header_path.name}")


def _on_preview_value(viewer, value, x, y, view):
    if value is None or view is None:
        viewer.preview_value_label.setText("Value: --")
        return
    unit = view.get('unit') or ''
    title = view.get('title') or ''
    text = f"{title}: {value:.4g}"
    if unit:
        text += f" {unit}"
    viewer.preview_value_label.setText(text)

# ---------- manual tagging (still available) ----------

def on_preview_cmap_changed(viewer, idx):
    viewer._suppress_profile_restart = False
    viewer.preview_cmap = viewer.preview_cmap_combo.currentText(); viewer.config['preview_cmap'] = viewer.preview_cmap; save_config(viewer.config)
    if viewer.last_preview:
        viewer._suppress_profile_restart = True
        viewer.show_file_channel(viewer.last_preview[0], viewer.last_preview[1])
__all__ = [
    "_build_metadata_html",
    "show_file_channel",
    "_on_preview_value",
    "on_preview_cmap_changed",
]




