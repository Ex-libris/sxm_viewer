"""Spectroscopy overlay helpers for SXMGridViewer."""
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
from ...data.spectroscopy import is_matrix_file_entry

_MARKER_SYMBOLS = {"circle", "square", "triangle", "diamond"}


def _normalized_symbol(viewer):
    symbol = getattr(viewer, "spectro_marker_symbol", "circle") or "circle"
    symbol = symbol.lower()
    return symbol if symbol in _MARKER_SYMBOLS else "circle"


def _effective_marker_size(viewer, crowded: bool, reveal_points: bool) -> float:
    size = float(getattr(viewer, "spectro_marker_size", 5.0) or 5.0)
    compact = bool(getattr(viewer, "compact_markers", False))
    if compact:
        size *= 0.7
    if crowded:
        size *= 0.85
    if reveal_points and size < 3.0:
        size = 3.0
    return max(1.5, min(size, 12.0))


def _marker_path(symbol: str, center: QtCore.QPointF, size: float) -> QtGui.QPainterPath:
    path = QtGui.QPainterPath()
    x = center.x()
    y = center.y()
    if symbol == "square":
        rect = QtCore.QRectF(x - size, y - size, size * 2.0, size * 2.0)
        radius = max(1.5, size * 0.3)
        path.addRoundedRect(rect, radius, radius)
    elif symbol == "triangle":
        path.moveTo(x, y - size)
        path.lineTo(x + size, y + size)
        path.lineTo(x - size, y + size)
        path.closeSubpath()
    elif symbol == "diamond":
        path.moveTo(x, y - size)
        path.lineTo(x + size, y)
        path.lineTo(x, y + size)
        path.lineTo(x - size, y)
        path.closeSubpath()
    else:
        rect = QtCore.QRectF(x - size, y - size, size * 2.0, size * 2.0)
        path.addEllipse(rect)
    return path


def _draw_marker_symbol(painter, x, y, symbol, size, base_color, highlight=False, pulse=1.0):
    center = QtCore.QPointF(x, y)
    path = _marker_path(symbol, center, size)
    stroke_color = QtGui.QColor(base_color)
    if stroke_color.alpha() == 0:
        stroke_color.setAlpha(255)
    fill_color = QtGui.QColor(stroke_color)
    fill_color.setAlpha(min(255, max(90, stroke_color.alpha() - 40)))
    pen = QtGui.QPen(stroke_color)
    pen.setJoinStyle(QtCore.Qt.RoundJoin)
    pen.setWidthF(max(1.0, size * 0.35))
    painter.setBrush(QtGui.QBrush(fill_color))
    painter.setPen(pen)
    painter.drawPath(path)
    if highlight:
        glow_scale = 2.15 + 0.65 * pulse
        halo_size = size * glow_scale
        gradient = QtGui.QRadialGradient(center, halo_size)
        peak_alpha = min(255, int(170 * (0.8 + 0.4 * pulse)))
        gradient.setColorAt(0.0, QtGui.QColor(255, 248, 255, peak_alpha))
        gradient.setColorAt(0.4, QtGui.QColor(255, 190, 230, int(110 * pulse)))
        gradient.setColorAt(1.0, QtGui.QColor(255, 140, 210, 0))
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(QtGui.QBrush(gradient))
        painter.drawEllipse(center, halo_size, halo_size)
        halo = QtGui.QPen(QtGui.QColor(255, 90, 180, 200), max(2.0, size * 0.4 * pulse))
        halo.setJoinStyle(QtCore.Qt.RoundJoin)
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.setPen(halo)
        painter.drawEllipse(center, halo_size * 0.85, halo_size * 0.85)
        hi_pen = QtGui.QPen(QtGui.QColor(255, 245, 255, 190), max(1.6, size * 0.35))
        hi_pen.setJoinStyle(QtCore.Qt.RoundJoin)
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.setPen(hi_pen)
        painter.drawPath(path)
    return path.boundingRect().adjusted(-1.5, -1.5, 1.5, 1.5)


def _spread_overlapping_marker_coords(coords, *, marker_size=5.0, cluster_tol=0.5):
    """Fan out spectra that land on the same pixel so coincident points stay visible."""
    if not coords or len(coords) < 2:
        return list(coords or [])
    tol = max(0.25, float(cluster_tol or 0.5))
    buckets = OrderedDict()
    for spec, col, row in coords:
        try:
            key = (
                int(round(float(col) / tol)),
                int(round(float(row) / tol)),
            )
            buckets.setdefault(key, []).append((spec, float(col), float(row)))
        except Exception:
            buckets.setdefault(("raw", id(spec)), []).append((spec, col, row))

    spread = []
    for group in buckets.values():
        count = len(group)
        if count <= 1:
            spread.extend(group)
            continue
        cx = sum(col for _, col, _ in group) / float(count)
        cy = sum(row for _, _, row in group) / float(count)
        radius = max(0.9, min(3.6, float(marker_size or 5.0) * 0.38))
        if count > 8:
            radius *= 1.0 + min(1.2, (count - 8) * 0.08)
        for idx, (spec, _col, _row) in enumerate(group):
            angle = (-0.5 * math.pi) + (2.0 * math.pi * idx / float(count))
            spread.append((
                spec,
                cx + radius * math.cos(angle),
                cy + radius * math.sin(angle),
            ))
    return spread


def _stack_badge_text(spec):
    try:
        count = int(spec.get("xy_stack_count") or 0)
    except Exception:
        count = 0
    if count <= 1:
        return ""
    label = str(spec.get("xy_stack_display") or "").strip()
    return label or f"x{count}"


def _stack_badge_tooltip(spec):
    text = str(spec.get("xy_stack_summary") or "").strip()
    if text:
        return text
    label = _stack_badge_text(spec)
    return f"Coincident spectra: {label}" if label else ""


def _stack_badges_from_coords(coords):
    groups = OrderedDict()
    for spec, col, row in coords or []:
        label = _stack_badge_text(spec)
        if not label:
            continue
        key = str(spec.get("xy_stack_key") or f"{round(float(col), 3)}:{round(float(row), 3)}")
        groups.setdefault(key, {"spec": spec, "coords": [], "label": label})
        groups[key]["coords"].append((float(col), float(row)))
    badges = []
    for group in groups.values():
        pts = group["coords"]
        if not pts:
            continue
        cx = sum(col for col, _ in pts) / float(len(pts))
        cy = sum(row for _, row in pts) / float(len(pts))
        badges.append({
            "spec": group["spec"],
            "col": cx,
            "row": cy,
            "label": group["label"],
            "tooltip": _stack_badge_tooltip(group["spec"]),
        })
    return badges


def _draw_stack_badge(painter, x, y, label, *, tooltip=None):
    font = QtGui.QFont("Segoe UI", 7, QtGui.QFont.Bold)
    painter.save()
    painter.setFont(font)
    metrics = painter.fontMetrics()
    badge_w = max(18, metrics.horizontalAdvance(label) + 10)
    badge_h = max(14, metrics.height() + 2)
    rect = QtCore.QRectF(x + 7.0, y - badge_h - 2.0, badge_w, badge_h)
    painter.setPen(QtGui.QPen(QtGui.QColor(255, 240, 180), 1.0))
    painter.setBrush(QtGui.QColor(40, 30, 18, 220))
    painter.drawRoundedRect(rect, 5, 5)
    painter.setPen(QtGui.QPen(QtGui.QColor(255, 228, 120), 1.0))
    painter.drawText(rect, QtCore.Qt.AlignCenter, label)
    painter.restore()
    return rect


def _spectros_near_thumb_pos(viewer, file_key: str, header: dict, thumb_pos_px: QtCore.QPoint, thumb_dims):
    """
    Map a click in thumbnail pixel coordinates to spectroscopy list ordered by distance.
    Returns list of spectro dicts (nearest first).
    """
    entries = viewer.spectros_by_image.get(str(file_key), []) or []
    if not entries:
        return []
    w, h = thumb_dims if thumb_dims else viewer._thumb_dimensions()
    px, py = int(thumb_pos_px.x()), int(thumb_pos_px.y())
    px = min(max(px, 0), max(w - 1, 0))
    py = min(max(py, 0), max(h - 1, 0))
    extent = viewer._header_extent(header) if header is not None else [0.0, 1.0, 1.0, 0.0]
    x0, x1, y1, y0 = extent
    xspan = x1 - x0 if x1 != x0 else 1.0
    yspan = y1 - y0 if y1 != y0 else 1.0
    cols = max(1, w - 1)
    rows = max(1, h - 1)
    frac_x = px / float(cols)
    frac_y = py / float(rows)
    u = frac_x - 0.5
    v = 0.5 - frac_y
    angle_deg = viewer._header_scan_angle(header) if header is not None and hasattr(viewer, "_header_scan_angle") else 0.0
    if angle_deg:
        theta = math.radians(angle_deg)
        cos_t = math.cos(theta)
        sin_t = math.sin(theta)
        u_rot = u * cos_t - v * sin_t
        v_rot = u * sin_t + v * cos_t
    else:
        u_rot, v_rot = u, v
    cx = 0.5 * (x0 + x1)
    cy = 0.5 * (y0 + y1)
    sx = cx + u_rot * xspan
    sy = cy + v_rot * yspan
    hits = []
    for s in entries:
        sx_e = s.get('x'); sy_e = s.get('y')
        if sx_e is None or sy_e is None:
            continue
        dx = sx - sx_e; dy = sy - sy_e
        d2 = dx*dx + dy*dy
        hits.append((d2, s))
    hits.sort(key=lambda t: t[0])
    return [h[1] for h in hits]


def _render_spectroscopy_overlays(
    viewer,
    pixmap,
    header,
    file_key,
    xpix,
    ypix,
    reveal_points_override=None,
    selected_spec=None,
    entries_override=None,
    matrix_as_points=False,
    thumb_crop=None,
):
    """Render spectroscopy overlays with configurable marker symbols and matrix footprints."""
    if not viewer.show_spectra:
        return []
    specs = entries_override if entries_override is not None else viewer.spectros_by_image.get(file_key, [])
    if not specs:
        return []
    markers = []
    painter = QtGui.QPainter(pixmap)
    painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
    w_scale = pixmap.width() / max(1, xpix - 1)
    crop_rows = None
    if thumb_crop:
        try:
            r0 = int(thumb_crop.get("r0"))
            r1 = int(thumb_crop.get("r1"))
            if r1 > r0:
                crop_rows = r1 - r0 + 1
        except Exception:
            crop_rows = None
    y_denom = max(1, (crop_rows - 1)) if crop_rows else max(1, ypix - 1)
    h_scale = pixmap.height() / y_denom
    if reveal_points_override is None:
        reveal_points = hasattr(viewer, '_temp_reveal') and file_key in getattr(viewer, '_temp_reveal', set())
    else:
        reveal_points = bool(reveal_points_override)

    singles = []
    matrices = defaultdict(list)
    for s in specs:
        midx = s.get('matrix_index')
        is_matrix_file = is_matrix_file_entry(s)
        force_points = matrix_as_points or not is_matrix_file
        if midx is None or force_points:
            singles.append(s)
        else:
            key = s.get('matrix_dataset') or str(s.get('path'))
            matrices[key].append(s)

    # When requested (e.g., matrix preview dialog), render matrix entries as points too.
    if matrix_as_points and matrices:
        for ms in matrices.values():
            singles.extend(ms)

    # Matrix footprints (skip when explicitly rendering matrix entries as individual points)
    if viewer.show_matrix_markers and matrices and not matrix_as_points:
        matrix_color = QtGui.QColor(getattr(viewer, 'spectro_marker_color_matrix', QtGui.QColor(64, 200, 255, 200)))
        for m_specs in matrices.values():
            rect = viewer._matrix_bbox_pixels(
                m_specs, header, xpix, ypix, w_scale, h_scale, file_key, thumb_crop=thumb_crop
            )
            if rect is None:
                continue
            painter.save()
            border = QtGui.QColor(matrix_color)
            border.setAlpha(235)
            shadow = QtGui.QColor(10, 10, 20, 80)
            painter.setPen(QtGui.QPen(shadow, 1.5))
            painter.drawRoundedRect(rect.translated(2, 2), 8, 8)
            painter.setPen(QtGui.QPen(border, 2.4))
            fill = QtGui.QBrush(QtGui.QColor(matrix_color.red(), matrix_color.green(), matrix_color.blue(), 70))
            painter.setBrush(fill)
            painter.drawRoundedRect(rect, 8, 8)
            try:
                grid_cols = m_specs[0].get('grid_cols')
                grid_rows = m_specs[0].get('grid_rows')
                dims = f"{grid_cols}x{grid_rows}" if grid_cols and grid_rows else None
            except Exception:
                dims = None
            chip_text = dims or "MATRIX"
            chip_font = QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold)
            painter.setFont(chip_font)
            metrics = painter.fontMetrics()
            chip_w = max(min(metrics.horizontalAdvance(chip_text) + 12, rect.width() - 8), 28)
            chip_h = max(metrics.height() + 6, 16)
            chip_rect = QtCore.QRectF(rect.left() + 6, rect.top() + 6, chip_w, chip_h)
            painter.setBrush(QtGui.QColor(border.red(), border.green(), border.blue(), 220))
            painter.setPen(QtGui.QPen(QtCore.Qt.white, 1.2))
            painter.drawRoundedRect(chip_rect, 6, 6)
            painter.drawText(chip_rect, QtCore.Qt.AlignCenter, chip_text)
            painter.restore()
            tooltip = Path(m_specs[0].get('path', '')).name
            if dims:
                tooltip = f"{tooltip}\nGrid: {dims}"
            markers.append({'rect': rect, 'spec': m_specs[0], 'label': chip_text, 'kind': 'matrix', 'tooltip': tooltip})

    color_single = getattr(viewer, 'spectro_marker_color_single', QtGui.QColor(255, 160, 0, 200))
    color_matrix = getattr(viewer, 'spectro_marker_color_matrix', QtGui.QColor(64, 200, 255, 200))

    # Single spectroscopies (customizable markers)
    if (viewer.show_single_markers or reveal_points or matrix_as_points) and singles:
        coords = []
        for idx, spec in enumerate(singles, 1):
            c = viewer._map_spec_to_pixels(spec, header, xpix, ypix, file_key, thumb_crop=thumb_crop)
            if c is None:
                c = viewer._fallback_spec_coords(idx, xpix, ypix)
            col, row = c
            coords.append((spec, float(col), float(row)))

        count = len(coords)
        crowded = count > 200 or bool(getattr(viewer, "compact_markers", False))
        marker_symbol = _normalized_symbol(viewer)
        marker_size = _effective_marker_size(viewer, crowded, reveal_points)
        badge_defs = _stack_badges_from_coords(coords)
        coords = _spread_overlapping_marker_coords(coords, marker_size=marker_size)
        pulse = float(getattr(viewer, "_highlight_pulse_strength", 1.0) or 1.0)
        for spec, col, row in coords:
            x = col * w_scale
            y = row * h_scale
            highlight = False
            try:
                if selected_spec and viewer._spec_identity_key(spec) == viewer._spec_identity_key(selected_spec):
                    highlight = True
            except Exception:
                highlight = False
            is_matrix_spec = is_matrix_file_entry(spec)
            base_color = color_matrix if is_matrix_spec else color_single
            rect = _draw_marker_symbol(
                painter,
                x,
                y,
                marker_symbol,
                marker_size,
                base_color,
                highlight=highlight,
                pulse=pulse if highlight else 1.0,
            )
            markers.append({'rect': rect, 'spec': spec, 'label': ''})
        for badge in badge_defs:
            bx = float(badge["col"]) * w_scale
            by = float(badge["row"]) * h_scale
            rect = _draw_stack_badge(painter, bx, by, badge["label"], tooltip=badge.get("tooltip"))
            markers.append({
                'rect': rect,
                'spec': badge.get("spec"),
                'label': 'stack-badge',
                'tooltip': badge.get("tooltip"),
            })
    # summary badge (S/M counts and matrix grid if available)
    try:
        total_s = len(singles)
        total_m = sum(len(v) for v in matrices.values()) if matrices else 0
        badge_w = 64
        badge_h = 18
        bx = pixmap.width() - badge_w - 6
        by = 6
        painter.setPen(QtGui.QPen(QtCore.Qt.NoPen))
        painter.setBrush(QtGui.QColor(35, 35, 40, 200))
        painter.drawRoundedRect(bx, by, badge_w, badge_h, 7, 7)
        painter.setFont(QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold))
        painter.setPen(QtGui.QColor(240, 240, 240))
        # include matrix grid hint if available
        dims_hint = ""
        if matrices:
            try:
                any_list = next(iter(matrices.values()))
                gc = any_list[0].get('grid_cols'); gr = any_list[0].get('grid_rows')
                if gc and gr:
                    dims_hint = f" ({gc}x{gr})"
            except Exception:
                dims_hint = ""
        painter.drawText(bx + 6, by + 12, f"S:{total_s} M:{len(matrices)}{dims_hint}")
        # optional matrix dims hint
        if matrices:
            dims = None
            try:
                m_any = next(iter(matrices.values()))
                gc = m_any[0].get('grid_cols'); gr = m_any[0].get('grid_rows')
                if gc and gr:
                    dims = f"{gc}x{gr}"
            except Exception:
                dims = None
            if dims:
                painter.setFont(QtGui.QFont("Segoe UI", 7))
                painter.drawText(bx + badge_w - 32, by + 12, dims)
        markers.append({'rect': QtCore.QRectF(bx, by, badge_w, badge_h), 'spec': None, 'label': 'badge'})
    except Exception:
        pass

    painter.end()
    return markers
__all__ = [
    "_spectros_near_thumb_pos",
    "_render_spectroscopy_overlays",
    "_spread_overlapping_marker_coords",
    "_stack_badges_from_coords",
]



