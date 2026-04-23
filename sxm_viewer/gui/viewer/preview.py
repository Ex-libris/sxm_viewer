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
from ...config import save_config, CH_EQUALITY_TOL_NM, CH_SAMPLE_POINTS
from ...data.io import parse_header
from ...processing.detection import _find_topography_channel, _sample_channel_values_for_tagging, header_indicates_constant
from ...data.io import normalize_unit_and_data
from ...data.spectroscopy import is_matrix_file_entry
from ...utils.units import _auto_display_unit, _safe_float
from ..spectroscopy import overlays as spectro_overlays
from ..thumbnail_render import detect_valid_scan_region

# Tolerance floor (~20 pm) for deciding constant-height by percentile spread
CH_RANGE_TOL_NM = max(CH_EQUALITY_TOL_NM, 0.02)


def _resolve_view_clim(viewer, file_key, channel_idx, arr, *, relative_zero: bool = False):
    helper = getattr(viewer, "_resolve_preview_clim", None)
    if callable(helper):
        try:
            return helper(str(file_key), int(channel_idx), arr, relative_zero=relative_zero)
        except Exception:
            pass
    return _auto_preview_clim(arr, relative_zero=relative_zero)

def _build_spec_transform(header, xpix, ypix):
    if not header:
        return None
    try:
        xc = float(header.get('xCenter'))
        yc = float(header.get('yCenter'))
        xr = float(header.get('XScanRange'))
        yr = float(header.get('YScanRange'))
    except Exception:
        return None
    if xr == 0.0 or yr == 0.0:
        return None
    try:
        xp = int(xpix)
        yp = int(ypix)
    except Exception:
        xp = yp = None
    try:
        angle = float(header.get('Angle', 0.0))
    except Exception:
        angle = 0.0
    return {
        'x_center': xc,
        'y_center': yc,
        'x_range': xr,
        'y_range': yr,
        'angle': angle,
        'xpix': xp,
        'ypix': yp,
    }


def _coerce_float(value):
    try:
        return float(value)
    except Exception:
        return None


def _fmt_overlay_number(value, unit="", precision=4):
    if value is None:
        return ""
    unit_txt = str(unit or "").strip()
    numeric = _safe_float(value, default=None)
    if numeric is None:
        return f"{str(value).strip()} {unit_txt}".strip()
    display_value = float(numeric)
    display_unit = unit_txt
    if unit_txt:
        try:
            display_unit, factor = _auto_display_unit(unit_txt, np.asarray([display_value], dtype=float))
            display_value *= float(factor)
        except Exception:
            display_unit = unit_txt
    text = f"{display_value:.{precision}g}"
    return f"{text} {display_unit}".strip()


def _candidate_abs_z_keys(header):
    if not header:
        return []
    keys = []
    for key in header.keys():
        kl = str(key).strip().lower().replace(" ", "")
        if kl in {"topography", "z", "zabs", "z_abs", "zpiezo", "zposition"}:
            keys.append(str(key))
    priority = {"topography": 0, "z_abs": 1, "zabs": 1, "z": 2, "zpiezo": 3, "zposition": 4}
    keys.sort(key=lambda k: priority.get(str(k).strip().lower(), 99))
    return keys


def _extract_abs_z_nm_from_header(header):
    if not header:
        return None
    for key in _candidate_abs_z_keys(header):
        raw = header.get(key)
        num = _coerce_float(raw)
        if num is None:
            continue
        unit = (
            header.get(f"{key}PhysUnit")
            or header.get(f"{key}Unit")
            or header.get("ZPhysUnit")
            or header.get("PhysUnit")
            or "nm"
        )
        try:
            _, arr_nm = normalize_unit_and_data(np.asarray([num], dtype=float), unit)
            if arr_nm.size:
                return float(arr_nm[0])
        except Exception:
            continue
    return None


def _infer_abs_z_pm_from_topography(viewer, file_key: str, header: dict, fds: list, channel_idx: int):
    if not fds:
        return None
    try:
        topo_idx = _find_topography_channel(fds)
        if topo_idx is None:
            topo_idx = channel_idx if (0 <= channel_idx < len(fds)) else None
        if topo_idx is None:
            return None
        fd_topo = fds[topo_idx]
        samples = _sample_channel_values_for_tagging(file_key, header, fd_topo, CH_SAMPLE_POINTS)
        if samples is None or not np.asarray(samples).size:
            return None
        _, arr_nm = normalize_unit_and_data(samples, fd_topo.get("PhysUnit", ""))
        vals = np.asarray(arr_nm, dtype=float).ravel()
        vals = vals[np.isfinite(vals)]
        if vals.size == 0:
            return None
        return int(round(float(np.nanmedian(vals)) * 1000.0))
    except Exception:
        return None


def _build_acquisition_overlay_info(viewer, header_path: Path, header: dict, fds: list, channel_idx: int):
    file_key = str(header_path)
    taginfo = viewer.tags.get(file_key, {}) or {}
    tag_label = str(taginfo.get("tag") or "").strip().lower()
    header_hint = header_indicates_constant(header)
    if tag_label in {"constant-height", "constant-current"}:
        effective_tag = tag_label
    elif header_hint == "CH":
        effective_tag = "constant-height"
    elif header_hint == "CC":
        effective_tag = "constant-current"
    else:
        effective_tag = None

    bias_val = header.get("Bias")
    bias_unit = header.get("BiasPhysUnit") or "V"
    setp_val = header.get("SetPoint")
    setp_unit = header.get("SetPointPhysUnit") or ""
    bias_txt = _fmt_overlay_number(bias_val, bias_unit)
    setp_txt = _fmt_overlay_number(setp_val, setp_unit)
    header_abs_nm = _extract_abs_z_nm_from_header(header)
    if effective_tag is None:
        if taginfo.get("abs_z_pm") is not None or (header_abs_nm is not None and not (bias_txt or setp_txt)):
            effective_tag = "constant-height"
        elif bias_txt or setp_txt:
            effective_tag = "constant-current"
        else:
            return {
                "mode": "",
                "text": "",
                "z_abs_nm": None,
                "bias_text": bias_txt,
                "setpoint_text": setp_txt,
            }

    if effective_tag == "constant-height":
        abs_pm = taginfo.get("abs_z_pm", None)
        if abs_pm is None:
            if header_abs_nm is not None:
                abs_pm = int(round(float(header_abs_nm) * 1000.0))
        if abs_pm is None:
            abs_pm = _infer_abs_z_pm_from_topography(viewer, file_key, header, fds, channel_idx)
        z_txt = ""
        z_nm = None
        if abs_pm is not None:
            z_nm = float(abs_pm) / 1000.0
            z_txt = f"{z_nm:.3f} nm"
            if tag_label == "constant-height" and taginfo.get("abs_z_pm") != abs_pm:
                try:
                    updated = dict(taginfo)
                    updated["abs_z_pm"] = int(abs_pm)
                    viewer.tags[file_key] = updated
                    viewer.config["tags"] = viewer.tags
                    save_config(viewer.config)
                except Exception:
                    pass
        text = f"CH  z_abs {z_txt}" if z_txt else "CH"
        return {
            "mode": "CH",
            "text": text,
            "z_abs_nm": z_nm,
            "bias_text": bias_txt,
            "setpoint_text": setp_txt,
        }

    parts = []
    if bias_txt:
        parts.append(f"Bias {bias_txt}")
    if setp_txt:
        parts.append(f"Iset {setp_txt}")
    text = "CC"
    if parts:
        text = "CC  " + " | ".join(parts)
    return {
        "mode": "CC",
        "text": text,
        "z_abs_nm": None,
        "bias_text": bias_txt,
        "setpoint_text": setp_txt,
    }

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
    tag_key = str(header_path)
    taginfo = viewer.tags.get(tag_key, {})
    try:
        _hdr, fds_all = viewer.headers.get(tag_key, (None, None))
    except Exception:
        fds_all = None
    tag_label = taginfo.get('tag', None)
    header_hint = header_indicates_constant(header)
    hinted_label = None
    if header_hint == 'CH':
        hinted_label = 'constant-height'
    elif header_hint == 'CC':
        hinted_label = 'constant-current'
    effective_tag = tag_label or hinted_label
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
    if effective_tag == 'constant-height':
        abs_pm = taginfo.get('abs_z_pm', None)
        inferred = False
        if abs_pm is None:
            # derive from topography channel values (median) when not stored
            try:
                topo_idx = _find_topography_channel(fds_all)
                if topo_idx is None:
                    topo_idx = channel_idx if (fds_all and 0 <= channel_idx < len(fds_all)) else None
                if topo_idx is not None:
                    fd_topo = fds_all[topo_idx] if fds_all else None
                    if fd_topo is None:
                        raise ValueError("Topography channel not available")
                    samples = _sample_channel_values_for_tagging(tag_key, header, fd_topo, CH_SAMPLE_POINTS)
                    if samples is not None and np.asarray(samples).size:
                        _, arr_nm = normalize_unit_and_data(samples, fd_topo.get('PhysUnit', ''))
                        vals = np.asarray(arr_nm, dtype=float).ravel()
                        vals = vals[np.isfinite(vals)]
                        if vals.size:
                            median = float(np.nanmedian(vals))
                            abs_pm = int(round(median * 1000.0))
                            inferred = True
            except Exception:
                abs_pm = None
        if abs_pm is not None:
            abs_nm = abs_pm / 1000.0
            suffix = " (inferred)" if inferred else ""
            ch_lines += f"<div>Piezo Z (abs){suffix}: <b>{abs_nm:.3f} nm</b></div>"
            # persist when we already tagged as CH
            if tag_label == 'constant-height' and taginfo.get('abs_z_pm') != abs_pm:
                try:
                    taginfo = dict(taginfo)
                    taginfo['abs_z_pm'] = abs_pm
                    viewer.tags[tag_key] = taginfo
                    viewer.config['tags'] = viewer.tags
                    save_config(viewer.config)
                except Exception:
                    pass
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
    bias_txt = esc(_fmt_overlay_number(bias, bias_unit)) if bias is not None else ""
    setp_txt = esc(_fmt_overlay_number(setp, setp_unit)) if setp is not None else ""
    key_rows = [
        ("Acquired", date_display),
        ("Bias", bias_txt),
        ("Setpoint", setp_txt),
        ("Image size", size_txt),
        ("Pixels", pixel_txt),
        ("X/Y center", center_txt),
        ("Piezo Z (abs)", piezo_txt),
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


def build_single_channel_view(viewer, header_path_str, channel_idx: int, *, cmap_override=None, use_local_cmap=False):
    header_path = Path(header_path_str)
    file_key = str(header_path)
    header, fds = viewer.headers.get(file_key, (None, None))
    if header is None or fds is None:
        try:
            header, fds = parse_header(header_path)
            viewer.headers[file_key] = (header, fds)
        except Exception:
            header, fds = None, None
    if header is None or fds is None or channel_idx < 0 or channel_idx >= len(fds):
        return None
    fd = fds[channel_idx]
    axis_unit = "px"
    xpix = int(header.get("xPixel", 128))
    ypix = int(header.get("yPixel", xpix))
    base_extent = None
    if getattr(viewer, "_is_processed_key", lambda _k: False)(file_key):
        try:
            processed_view = getattr(viewer, "_processed_views", {}).get(file_key) or {}
            stored_extent = processed_view.get("extent_raw")
            if stored_extent is not None and len(stored_extent) == 4:
                base_extent = tuple(float(v) for v in stored_extent)
        except Exception:
            base_extent = None
    if base_extent is None:
        base_extent = viewer._header_extent(header)
    unit_normalized, arr_base = viewer._get_filtered_channel_array(file_key, channel_idx, header, fd)
    arr_base = np.asarray(arr_base)
    arr_adj, adjusted_extent = viewer._apply_adjustments_for_channel(file_key, channel_idx, arr_base, base_extent)
    display_extent = viewer._display_extent(adjusted_extent, header)
    display_unit, display_arr, zero_offset = viewer._scale_unit_for_display(unit_normalized, arr_adj)
    display_arr = np.asarray(display_arr)
    axis_unit = header.get("XPhysUnit") or header.get("YPhysUnit") or header.get("ScanUnit") or ""
    if not axis_unit:
        axis_unit = "px" if display_extent is None else "nm"

    cmap_to_use = cmap_override or viewer.preview_cmap_combo.currentText() or viewer.preview_cmap
    if use_local_cmap and cmap_override is None:
        cmap_to_use = viewer.per_file_channel_cmap.get((file_key, channel_idx), cmap_to_use)

    show_preview_specs = bool(getattr(viewer, "show_preview_spectra", getattr(viewer, "show_spectra", True)))
    spec_entries = viewer.spectros_by_image.get(str(header_path), []) if show_preview_specs else []
    overlay_specs = []
    highlight_spec = None
    highlight_candidate = getattr(viewer, "_highlighted_spec", None)
    if highlight_candidate and getattr(viewer, "spectro_highlight_glow", True):
        try:
            highlight_path = str(highlight_candidate.get("image_key") or highlight_candidate.get("path") or "")
            shared_keys = [str(key) for key in (highlight_candidate.get("shared_image_keys") or []) if key]
        except Exception:
            highlight_path = ""
            shared_keys = []
        if (highlight_path and highlight_path == str(header_path)) or str(header_path) in shared_keys:
            highlight_spec = highlight_candidate
    if spec_entries and show_preview_specs:
        if viewer.show_single_markers:
            overlay_specs.extend([
                s for s in spec_entries
                if s.get("matrix_index") is None or not is_matrix_file_entry(s)
            ])
        if viewer.show_matrix_markers:
            overlay_specs.extend([
                s for s in spec_entries
                if s.get("matrix_index") is not None and is_matrix_file_entry(s)
            ])

    caption = fd.get("Caption", fd.get("FileName", ""))
    date = str(header.get("Date", "") or "").strip()
    time_txt = str(header.get("Time", "") or "").strip()
    datetime_txt = " ".join([t for t in (date, time_txt) if t]).strip()
    base_title = header_path.name
    title_text = f"{base_title}  {caption}  {datetime_txt}" if datetime_txt else f"{base_title}  {caption}"
    colorbar_label = f"{caption} [{display_unit}]" if display_unit else caption
    acq_overlay = _build_acquisition_overlay_info(viewer, header_path, header, fds, channel_idx)
    meta = {
        "file_path": str(header_path),
        "file_name": header_path.name,
        "date": date,
        "time": time_txt,
        "datetime": datetime_txt,
        "channel": caption,
        "channel_index": int(channel_idx),
        "acquisition_mode": acq_overlay.get("mode"),
        "acquisition_overlay_text": acq_overlay.get("text", ""),
        "acquisition_z_abs_nm": acq_overlay.get("z_abs_nm"),
        "acquisition_bias_text": acq_overlay.get("bias_text", ""),
        "acquisition_setpoint_text": acq_overlay.get("setpoint_text", ""),
    }
    spec_pixels = []
    for spec in overlay_specs:
        coords = None
        try:
            coords = viewer._map_spec_to_pixels(spec, header, xpix, ypix, file_key=str(header_path))
        except Exception:
            coords = None
        if coords is not None:
            spec_pixels.append((spec, float(coords[0]), float(coords[1])))
    stack_badges = spectro_overlays._stack_badges_from_coords(spec_pixels)
    spec_pixels = spectro_overlays._spread_overlapping_marker_coords(
        spec_pixels,
        marker_size=float(getattr(viewer, "spectro_marker_size", 5.0) or 5.0),
    )

    view = {
        "arr": display_arr,
        "extent": display_extent,
        "extent_raw": adjusted_extent,
        "cmap": cmap_to_use,
        "unit_normalized": unit_normalized,
        "unit": display_unit,
        "display_relative_zero": bool(getattr(viewer, "display_units_relative", False)),
        "zero_offset": zero_offset,
        "title": title_text,
        "colorbar_label": colorbar_label,
        "axis_unit": axis_unit,
        "relative_axes": bool(viewer.relative_axes),
        "meta": meta,
        "path": str(header_path),
        "channel_idx": int(channel_idx),
        "acquisition_overlay_text": acq_overlay.get("text", ""),
        "spectra": overlay_specs,
        "highlight_spec": highlight_spec,
        "spec_pixels": list(spec_pixels),
        "stack_badges": list(stack_badges),
    }
    clim = _resolve_view_clim(
        viewer,
        header_path,
        channel_idx,
        display_arr,
        relative_zero=bool(getattr(viewer, "display_units_relative", False)),
    )
    if clim:
        view["clim"] = clim
    return {
        "header_path": header_path,
        "header": header,
        "fds": fds,
        "fd": fd,
        "channel_idx": int(channel_idx),
        "unit_normalized": unit_normalized,
        "display_unit": display_unit,
        "display_arr": display_arr,
        "zero_offset": zero_offset,
        "axis_unit": axis_unit,
        "base_extent": base_extent,
        "view": view,
    }


def show_file_channel(viewer, header_path_str, channel_idx:int, use_local_cmap=False):
    current_path_str = str(header_path_str)
    prev_preview = getattr(viewer, "last_preview", None)
    viewer.last_preview = (current_path_str, int(channel_idx))
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

    local_cmap = viewer.per_file_channel_cmap.get((file_key, channel_idx))
    cmap_to_use = local_cmap or getattr(viewer, "preview_cmap", None) or (viewer.preview_cmap_combo.currentText() or viewer.preview_cmap)
    try:
        viewer._sync_cmap_controls_for_selection(
            file_key,
            channel_idx,
            thumb_cmap=getattr(viewer, "thumb_cmap", None),
            preview_cmap=getattr(viewer, "preview_cmap", None),
        )
    except Exception:
        pass

    # Spectroscopy entries for this file (singles only for overlay)
    show_preview_specs = bool(getattr(viewer, 'show_preview_spectra', getattr(viewer, 'show_spectra', True)))
    spec_entries = viewer.spectros_by_image.get(str(header_path), []) if show_preview_specs else []
    overlay_specs = []
    highlight_spec = None
    highlight_candidate = getattr(viewer, '_highlighted_spec', None)
    if highlight_candidate and getattr(viewer, "spectro_highlight_glow", True):
        try:
            highlight_path = str(highlight_candidate.get('image_key') or highlight_candidate.get('path') or '')
            shared_keys = [str(key) for key in (highlight_candidate.get("shared_image_keys") or []) if key]
        except Exception:
            highlight_path = ''
            shared_keys = []
        if (highlight_path and highlight_path == str(header_path)) or str(header_path) in shared_keys:
            highlight_spec = highlight_candidate
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
    acq_overlay = _build_acquisition_overlay_info(viewer, header_path, header, fds, channel_idx)
    meta = {
        'file_path': str(header_path),
        'file_name': header_path.name,
        'date': date,
        'time': time_txt,
        'datetime': datetime_txt,
        'channel': caption,
        'channel_index': int(channel_idx),
        'acquisition_mode': acq_overlay.get("mode"),
        'acquisition_overlay_text': acq_overlay.get("text", ""),
        'acquisition_z_abs_nm': acq_overlay.get("z_abs_nm"),
        'acquisition_bias_text': acq_overlay.get("bias_text", ""),
        'acquisition_setpoint_text': acq_overlay.get("setpoint_text", ""),
    }
    clim_main = _resolve_view_clim(
        viewer,
        file_key,
        channel_idx,
        display_arr,
        relative_zero=bool(getattr(viewer, "display_units_relative", False)),
    )
    spec_pixels = []
    for spec in overlay_specs:
        coords = None
        try:
            coords = viewer._map_spec_to_pixels(spec, header, xpix, ypix, file_key=str(header_path))
        except Exception:
            coords = None
        if coords is not None:
            spec_pixels.append((spec, float(coords[0]), float(coords[1])))
    stack_badges = spectro_overlays._stack_badges_from_coords(spec_pixels)
    spec_pixels = spectro_overlays._spread_overlapping_marker_coords(
        spec_pixels,
        marker_size=float(getattr(viewer, "spectro_marker_size", 5.0) or 5.0),
    )
    main = {
        'arr': display_arr,
        'extent': display_extent,
        'extent_raw': adjusted_extent,
        'cmap': cmap_to_use,
        'unit_normalized': unit_normalized,
        'unit': display_unit,
        'display_relative_zero': bool(getattr(viewer, 'display_units_relative', False)),
        'zero_offset': zero_offset,
        'title': title_text,
        'colorbar_label': colorbar_label,
        'axis_unit': axis_unit,
        'relative_axes': bool(viewer.relative_axes),
        'meta': meta,
        'path': str(header_path),
        'channel_idx': int(channel_idx),
        'acquisition_overlay_text': acq_overlay.get("text", ""),
        'spectra': overlay_specs,
        'highlight_spec': highlight_spec,
        'spec_pixels': list(spec_pixels),
        'stack_badges': list(stack_badges),
    }
    if clim_main:
        main['clim'] = clim_main
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
            unit2_display, arr2_display, zero2_offset = viewer._scale_unit_for_display(unit2_final, arr2_adj)
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
            clim2 = _resolve_view_clim(
                viewer,
                file_key,
                idx2,
                arr2_display,
                relative_zero=bool(getattr(viewer, "display_units_relative", False)),
            )
            vdict = {'arr': arr2_display, 'extent': extent2, 'extent_raw': adj2_extent,
                     'cmap': cmap2, 'unit_normalized': unit2_final, 'unit': unit2_display,
                     'display_relative_zero': bool(getattr(viewer, 'display_units_relative', False)),
                     'zero_offset': zero2_offset,
                     'title': title2,
                     'colorbar_label': cbar_label2, 'axis_unit': axis_unit,
                     'relative_axes': bool(viewer.relative_axes), 'meta': meta2,
                     'path': str(header_path), 'channel_idx': int(idx2),
                     'acquisition_overlay_text': acq_overlay.get("text", ""),
                     'spec_pixels': list(spec_pixels)}
            if clim2:
                vdict['clim'] = clim2
            views.append(vdict)
        except Exception:
            # Skip extra view if anything fails for this file
            continue

    preserve = False
    try:
        last = prev_preview[0] if prev_preview else None
        preserve = (
            not getattr(viewer, '_suppress_profile_restore', False) and
            bool(getattr(viewer, 'preserve_profiles_on_channel_change', False))
            and last == current_path_str
            and getattr(viewer, 'current_mode', viewer.MODE_BROWSE) == viewer.MODE_MEASURE
        )
    except Exception:
        preserve = False
    viewer.preview_canvas.set_views(views, preserve_profiles=preserve)
    if getattr(viewer, '_suppress_profile_restore', False):
        viewer._suppress_profile_restore = False
    suppress_profile_restart = getattr(viewer, '_suppress_profile_restart', False)
    if suppress_profile_restart:
        viewer._suppress_profile_restart = False
    # Do NOT auto-start profiles; the user must click the Profile tool explicitly.
    if getattr(viewer, 'current_mode', viewer.MODE_BROWSE) == viewer.MODE_MEASURE:
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

    # Styled HTML metadata (preserve scroll position while browsing)
    try:
        sb = viewer.meta_box.verticalScrollBar()
        prev_pos = sb.value()
        html = viewer._build_metadata_html(header_path, header, fd, channel_idx, unit_normalized, display_unit, display_arr, zero_offset)
        viewer.meta_box.setHtml(html)
        QtCore.QTimer.singleShot(0, lambda pos=prev_pos: sb.setValue(pos))
    except Exception:
        viewer.meta_box.setPlainText(f"File: {header_path.name}")
    # Optional auto-tagging (constant-height/current) based on topography variance.
    try:
        _maybe_auto_tag_file(viewer, header_path, header, fds, channel_idx)
    except Exception:
        pass


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
    cmap_name = viewer.preview_cmap_combo.currentText()
    if viewer.last_preview:
        try:
            viewer._set_thumbnail_entry_cmap([viewer.last_preview[0]], cmap_name)
        except Exception:
            pass
        try:
            viewer._set_combo_text_silent(getattr(viewer, "preview_cmap_combo", None), cmap_name)
        except Exception:
            pass
        return
    viewer.preview_cmap = cmap_name
    viewer.config['preview_cmap'] = viewer.preview_cmap
    save_config(viewer.config)
__all__ = [
    "_build_metadata_html",
    "show_file_channel",
    "_on_preview_value",
    "on_preview_cmap_changed",
]

def _auto_preview_clim(arr, *, relative_zero: bool = False):
    """
    Compute color limits with automatic aborted scan detection and optional flat suppression.
    """
    try:
        a = np.asarray(arr, dtype=float)
        if a.ndim == 2:
            region = detect_valid_scan_region(a)
            if region:
                r0, r1 = region
                a = a[r0:r1 + 1, :]
        finite = a[np.isfinite(a)]
        if finite.size == 0:
            return None
        hist, edges = np.histogram(finite, bins=256)
        idx_max = int(np.argmax(hist))
        frac = hist[idx_max] / float(finite.size)
        if frac > 0.7:
            lo_edge, hi_edge = edges[idx_max], edges[idx_max + 1]
            trimmed = finite[(finite < lo_edge) | (finite > hi_edge)]
            if trimmed.size >= max(10, int(0.001 * finite.size)):
                if trimmed.size > 100:
                    if np.std(trimmed) > 1e-12 and np.ptp(trimmed) > 1e-12:
                        finite = trimmed
                else:
                    finite = trimmed
        vmin = float(np.nanpercentile(finite, 1.0))
        vmax = float(np.nanpercentile(finite, 99.0))
        if relative_zero:
            vmin = 0.0
        if vmin == vmax:
            return None
        return (vmin, vmax)
    except Exception:
        return None


def _classify_topography_values(vals, tolerance_nm: float | None = None):
    """
    Simple CH/CC classifier: if any row in a 2D topography image is exactly flat
    (all finite values identical), mark as constant-height; otherwise constant-current.
    For 1D data, mark CH only if the entire vector is flat.
    """
    try:
        arr = np.asarray(vals, dtype=float)
    except Exception:
        return None

    if arr.ndim == 0:
        return None

    if arr.ndim == 2:
        if arr.shape[0] == 0:
            return None
        def _is_flat(row):
            row_fin = row[np.isfinite(row)]
            return row_fin.size > 0 and np.ptp(row_fin) == 0.0
        top_flat = _is_flat(arr[0])
        bottom_flat = _is_flat(arr[-1])
        median = float(np.nanmedian(arr))
        prange = float(np.nanmax(arr) - np.nanmin(arr))
        if top_flat and bottom_flat:
            return {
                'tag': 'constant-height',
                'abs_pm': int(round(median * 1000.0)),
                'rng_nm': prange,
                'median_nm': median,
            }
        return {
            'tag': 'constant-current',
            'abs_pm': None,
            'rng_nm': prange,
            'median_nm': median,
        }

    # 1D fallback
    arr = arr[np.isfinite(arr)]
    if arr.size == 0:
        return None
    if np.ptp(arr) == 0.0:
        median = float(np.nanmedian(arr))
        return {
            'tag': 'constant-height',
            'abs_pm': int(round(median * 1000.0)),
            'rng_nm': 0.0,
            'median_nm': median,
        }
    return {'tag': 'constant-current', 'abs_pm': None, 'rng_nm': float(np.nanmax(arr) - np.nanmin(arr)),
            'median_nm': float(np.nanmedian(arr))}


def _maybe_auto_tag_file(viewer, header_path:Path, header:dict, fds:list, channel_idx:int):
    """Auto-tag constant-height/current using topography variance; respects manual tags."""
    if not getattr(viewer, "auto_detect_tags", False):
        return
    key = str(header_path)
    existing = viewer.tags.get(key, {})
    if existing.get("manual"):
        return
    if not fds:
        return
    topo_idx = _find_topography_channel(fds)
    if topo_idx is None:
        topo_idx = channel_idx if 0 <= channel_idx < len(fds) else 0
    if topo_idx is None or topo_idx >= len(fds):
        return
    fd_topo = fds[topo_idx]
    try:
        raw_arr = viewer._get_channel_array(key, topo_idx, header, fd_topo)
        _, arr_nm = normalize_unit_and_data(raw_arr, fd_topo.get('PhysUnit',''))
    except Exception:
        return
    classifier = getattr(viewer, "_classify_topography_values", _classify_topography_values)
    tag_info = classifier(arr_nm)
    if not tag_info:
        return
    if tag_info['tag'] == 'constant-height':
        viewer.tags[key] = {'tag': 'constant-height', 'abs_z_pm': tag_info.get('abs_pm'), 'auto': True,
                            'rng_nm': tag_info.get('rng_nm')}
    else:
        viewer.tags[key] = {'tag': 'constant-current', 'auto': True, 'rng_nm': tag_info.get('rng_nm')}
    viewer.config['tags'] = viewer.tags
    save_config(viewer.config)




