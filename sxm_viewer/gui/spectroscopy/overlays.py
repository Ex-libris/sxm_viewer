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


def _draw_marker_symbol(painter, x, y, symbol, size, base_color, highlight=False, pulse=1.0, low_conf=False):
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
    if low_conf:
        warn_pen = QtGui.QPen(QtGui.QColor(255, 210, 96, 220), max(1.5, size * 0.28))
        warn_pen.setStyle(QtCore.Qt.DashLine)
        warn_pen.setJoinStyle(QtCore.Qt.RoundJoin)
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.setPen(warn_pen)
        painter.drawEllipse(center, size * 1.55, size * 1.55)
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


# --- Type icons -------------------------------------------------------------
# One glyph per spectroscopy "kind" (single / z-stack / grid / loose
# cluster), drawn the same way wherever that kind shows up - corner badge,
# footprint chip, or floating icon - so the shape itself is identifiable
# without reading an abbreviation ("S:"/"M:"/"L:"/"C:").

def _draw_grid_icon(painter, cx, cy, *, cell=2.2, gap=1.3, color=None):
    """3x3 dot grid marking a matrix/grid dataset."""
    color = QtGui.QColor(color) if color is not None else QtGui.QColor(20, 20, 24)
    painter.save()
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QBrush(color))
    step = cell + gap
    start = -step
    for row in range(3):
        for col in range(3):
            x = cx + start + col * step
            y = cy + start + row * step
            painter.drawEllipse(QtCore.QPointF(x, y), cell / 2.0, cell / 2.0)
    painter.restore()


def _draw_cluster_icon(painter, cx, cy, kind, *, size=7.0, color=None):
    """Three dots in a row for a 'line' cluster, three scattered dots for a
    'cloud' - the loose (non-grid-file) point groupings from
    _detect_spectro_groups."""
    color = QtGui.QColor(color) if color is not None else QtGui.QColor(20, 20, 24)
    painter.save()
    painter.setPen(QtCore.Qt.NoPen)
    painter.setBrush(QtGui.QBrush(color))
    r = size * 0.15
    if kind == "line":
        pts = [(-size * 0.36, 0.0), (0.0, 0.0), (size * 0.36, 0.0)]
    else:
        pts = [(-size * 0.32, size * 0.22), (size * 0.06, -size * 0.28), (size * 0.32, size * 0.16)]
    for dx, dy in pts:
        painter.drawEllipse(QtCore.QPointF(cx + dx, cy + dy), r, r)
    painter.restore()


def _draw_stack_icon(painter, x, y, count, *, color=None):
    """Icon for a Z-stack / repeated-measurement site: a few stacked bars
    with the trace count on a small numeral tag. Replaces the old floating
    'Zx44' text badge - the shape reads as 'a stack of traces' on its own,
    the number just says how many."""
    color = QtGui.QColor(color) if color is not None else QtGui.QColor(165, 141, 242, 235)
    bar_w = 15.0
    bar_h = 2.8
    gap = 1.6
    total_h = bar_h * 3 + gap * 2
    top = y - total_h - 6.0
    widths = (1.0, 0.76, 0.55)
    painter.save()
    painter.setPen(QtCore.Qt.NoPen)
    for i, w in enumerate(widths):
        c = QtGui.QColor(color)
        c.setAlpha(max(110, color.alpha() - i * 40))
        painter.setBrush(QtGui.QBrush(c))
        bw = bar_w * w
        rect = QtCore.QRectF(x - bw / 2.0, top + i * (bar_h + gap), bw, bar_h)
        painter.drawRoundedRect(rect, bar_h / 2.0, bar_h / 2.0)
    label = str(int(count)) if count else ""
    font = QtGui.QFont("Segoe UI", 7, QtGui.QFont.Bold)
    painter.setFont(font)
    metrics = painter.fontMetrics()
    tag_w = max(14, metrics.horizontalAdvance(label) + 7)
    tag_h = max(12, metrics.height())
    tag_rect = QtCore.QRectF(x + bar_w / 2.0 - 2.0, top - tag_h * 0.6, tag_w, tag_h)
    painter.setBrush(QtGui.QBrush(color))
    painter.setPen(QtGui.QPen(QtGui.QColor(18, 12, 28), 1.0))
    painter.drawRoundedRect(tag_rect, tag_h / 2.0, tag_h / 2.0)
    painter.setPen(QtGui.QPen(QtGui.QColor(18, 12, 28)))
    painter.drawText(tag_rect, QtCore.Qt.AlignCenter, label)
    painter.restore()
    return QtCore.QRectF(x - bar_w / 2.0 - 2, top - tag_h, bar_w + tag_w * 0.5 + 4, total_h + tag_h + 4)


def _draw_footprint(painter, rect, *, kind, color, count, dims=None):
    """Shared drawing for a summary footprint (matrix file, or a detected
    grid/line/cloud group of loose points): solid border + translucent fill
    + a small icon-and-count chip in the corner. All footprint kinds share
    this one look now (they used to differ only by a dashed/dotted border
    style, which doesn't read at thumbnail scale) - the icon in the chip is
    what tells them apart."""
    color = QtGui.QColor(color)
    painter.save()
    shadow = QtGui.QColor(10, 10, 20, 75)
    painter.setPen(QtGui.QPen(shadow, 1.3))
    painter.setBrush(QtCore.Qt.NoBrush)
    painter.drawRoundedRect(rect.translated(2, 2), 7, 7)
    border = QtGui.QColor(color)
    border.setAlpha(235)
    painter.setPen(QtGui.QPen(border, 2.2))
    fill = QtGui.QColor(color.red(), color.green(), color.blue(), 55)
    painter.setBrush(QtGui.QBrush(fill))
    painter.drawRoundedRect(rect, 7, 7)

    chip_text = f"{count} pt" + ("s" if count != 1 else "") if count else (dims or "")
    icon_w = 13.0
    chip_font = QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold)
    painter.setFont(chip_font)
    metrics = painter.fontMetrics()
    text_w = metrics.horizontalAdvance(chip_text)
    chip_w = max(min(text_w + icon_w + 16, rect.width() - 8), 34)
    chip_h = max(metrics.height() + 6, 17)
    chip_rect = QtCore.QRectF(rect.left() + 6, rect.top() + 6, chip_w, chip_h)
    painter.setBrush(QtGui.QColor(border.red(), border.green(), border.blue(), 230))
    painter.setPen(QtGui.QPen(QtCore.Qt.white, 1.1))
    painter.drawRoundedRect(chip_rect, 6, 6)
    icon_cx = chip_rect.left() + icon_w / 2.0 + 4.0
    icon_cy = chip_rect.center().y()
    dark = QtGui.QColor(18, 14, 10)
    if kind == "matrix" or kind == "grid":
        _draw_grid_icon(painter, icon_cx, icon_cy, color=dark)
    elif kind in ("line", "cloud"):
        _draw_cluster_icon(painter, icon_cx, icon_cy, kind, color=dark)
    text_rect = QtCore.QRectF(chip_rect.left() + icon_w + 6, chip_rect.top(), chip_rect.width() - icon_w - 8, chip_rect.height())
    painter.setPen(QtGui.QPen(QtCore.Qt.white))
    painter.drawText(text_rect, QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, chip_text)
    painter.restore()
    return chip_rect


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


def _stack_badge_count(spec):
    try:
        count = int(spec.get("xy_stack_count") or 0)
    except Exception:
        count = 0
    return count if count > 1 else 0


def _stack_badge_tooltip(spec):
    text = str(spec.get("xy_stack_summary") or "").strip()
    if text:
        return text
    count = _stack_badge_count(spec)
    if not count:
        return ""
    if spec.get("xy_stack_z_varies"):
        return f"Z series (x{count})"
    return f"Repeated spectra (x{count})"


def _stack_badges_from_coords(coords):
    groups = OrderedDict()
    for spec, col, row in coords or []:
        count = _stack_badge_count(spec)
        if not count:
            continue
        key = str(spec.get("xy_stack_key") or f"{round(float(col), 3)}:{round(float(row), 3)}")
        groups.setdefault(key, {"spec": spec, "coords": [], "count": count})
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
            "count": group["count"],
            "tooltip": _stack_badge_tooltip(group["spec"]),
        })
    return badges


def _summarize_remaining_singles(specs):
    """Split the final (post-footprint) single-spectrum list into Z-stack /
    repeated-measurement sites vs. plain standalone points, for the
    corner-badge summary text. Each spec counts toward exactly one bucket."""
    seen_stack_keys = set()
    stack_sites = 0
    stack_traces = 0
    plain_points = 0
    for spec in specs or []:
        count = _stack_badge_count(spec)
        if count:
            key = str(spec.get("xy_stack_key") or id(spec))
            if key in seen_stack_keys:
                continue
            seen_stack_keys.add(key)
            stack_sites += 1
            stack_traces += count
        else:
            plain_points += 1
    return stack_sites, stack_traces, plain_points


def _summary_badge_text(matrices, drawn_groups, remaining_singles):
    """Plain-language replacement for the old 'S:132 M:1 (36x24)' shorthand -
    describes what's actually on the image (grids, clusters, sites, points)
    instead of internal single/matrix counts."""
    parts = []
    if matrices:
        if len(matrices) == 1:
            m_specs = next(iter(matrices.values()))
            gc = m_specs[0].get('grid_cols')
            gr = m_specs[0].get('grid_rows')
            parts.append(f"{gc}×{gr} grid" if gc and gr else "1 grid")
        else:
            parts.append(f"{len(matrices)} grids")
    grid_like = sum(1 for g in drawn_groups if (g.get('kind') or '') == 'grid')
    cluster_like = len(drawn_groups) - grid_like
    if grid_like:
        parts.append(f"{grid_like} grid" + ("s" if grid_like != 1 else ""))
    if cluster_like:
        parts.append(f"{cluster_like} cluster" + ("s" if cluster_like != 1 else ""))
    stack_sites, stack_traces, plain_points = _summarize_remaining_singles(remaining_singles)
    if stack_sites and not plain_points:
        if stack_sites == 1:
            parts.append(f"1 site · {stack_traces} traces")
        else:
            parts.append(f"{stack_sites} sites · {stack_traces} traces")
    elif stack_sites:
        parts.append(f"{stack_sites + plain_points} points ({stack_sites} sites)")
    elif plain_points:
        parts.append(f"{plain_points} point" + ("s" if plain_points != 1 else ""))
    return " · ".join(parts) if parts else "No spectroscopy"


def _draw_summary_badge(painter, bx, by, text, *, icon_kind=None, height=18):
    """Per-image corner badge describing what's on the image in plain
    language (see _summary_badge_text). Shared by the real thumbnail
    overlay and the legend dialog so the legend never drifts from what is
    actually drawn."""
    icon_w = 15.0 if icon_kind else 0.0
    font = QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold)
    painter.save()
    painter.setFont(font)
    metrics = painter.fontMetrics()
    text_w = metrics.horizontalAdvance(text)
    pad = 8
    badge_w = text_w + icon_w + pad * 2
    badge_h = max(height, metrics.height() + 6)
    rect = QtCore.QRectF(bx, by, badge_w, badge_h)
    painter.setPen(QtGui.QPen(QtCore.Qt.NoPen))
    painter.setBrush(QtGui.QColor(24, 26, 34, 215))
    painter.drawRoundedRect(rect, 8, 8)
    text_left = rect.left() + pad
    if icon_kind:
        icon_cx = text_left + icon_w / 2.0 - 2.0
        icon_cy = rect.center().y()
        if icon_kind == "grid":
            _draw_grid_icon(painter, icon_cx, icon_cy, color=QtGui.QColor(235, 237, 244))
        elif icon_kind in ("line", "cloud"):
            _draw_cluster_icon(painter, icon_cx, icon_cy, icon_kind, color=QtGui.QColor(235, 237, 244))
        text_left += icon_w
    text_rect = QtCore.QRectF(text_left, rect.top(), rect.right() - text_left - pad / 2.0, rect.height())
    painter.setPen(QtGui.QColor(240, 240, 240))
    painter.drawText(text_rect, QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, text)
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
    # Un-normalize by width/height *before* rotating (real nm, isotropic) -
    # this is the exact inverse of _map_spec_to_pixels's rotate-then-normalize
    # order, and must use the inverse (-theta) rotation to undo that
    # function's +theta forward rotation. Mirrors the fix applied there.
    u_nm = u * xspan
    v_nm = v * yspan
    angle_deg = viewer._header_scan_angle(header) if header is not None and hasattr(viewer, "_header_scan_angle") else 0.0
    if angle_deg:
        theta = math.radians(angle_deg)
        cos_t = math.cos(theta)
        sin_t = math.sin(theta)
        dx = u_nm * cos_t + v_nm * sin_t
        dy = -u_nm * sin_t + v_nm * cos_t
    else:
        dx, dy = u_nm, v_nm
    cx = 0.5 * (x0 + x1)
    cy = 0.5 * (y0 + y1)
    sx = cx + dx
    sy = cy + dy
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

    color_single = getattr(viewer, 'spectro_marker_color_single', QtGui.QColor(255, 160, 0, 200))
    color_matrix = getattr(viewer, 'spectro_marker_color_matrix', QtGui.QColor(64, 200, 255, 200))
    color_stack = getattr(viewer, 'spectro_marker_color_stack', QtGui.QColor(165, 141, 242, 235))
    color_cloud = getattr(viewer, 'spectro_marker_color_cloud', QtGui.QColor(95, 191, 147, 210))

    singles = []
    matrices = defaultdict(list)
    grouped_keys = set()
    for s in specs:
        midx = s.get('matrix_index')
        # is_matrix_file_entry alone only recognizes Omicron/Anfatec matrix
        # .dat files by filename - it misses Nanonis .3ds grid points, which
        # would otherwise force real grid points into `singles` and render
        # them as a dense raw-point swarm instead of one matrix footprint.
        # _is_matrix_spec additionally checks matrix_dataset/matrix_index.
        is_matrix_file = viewer._is_matrix_spec(s)
        force_points = matrix_as_points or not is_matrix_file
        if midx is None or force_points:
            singles.append(s)
        else:
            key = s.get('matrix_dataset') or str(s.get('path'))
            matrices[key].append(s)
        if s.get('spectro_group_key'):
            grouped_keys.add(str(s.get('spectro_group_key')))

    # When requested (e.g., matrix preview dialog), render matrix entries as points too.
    if matrix_as_points and matrices:
        for ms in matrices.values():
            singles.extend(ms)

    # Grid/line/cloud footprints for clustered loose .dat spectra (see
    # _detect_spectro_groups): a summary shape instead of a marker swarm, same
    # spirit as the matrix-file footprint below. Skipped when the caller wants
    # raw points (matrix_as_points) or the user asked to reveal every point on
    # this crowded thumbnail (reveal_points) - members then fall through to the
    # normal single-marker rendering below since they were never excluded from
    # `singles` in that case.
    if viewer.show_single_markers and grouped_keys and not matrix_as_points and not reveal_points:
        # "cloud" groups (loose, non-grid/line point clusters) are excluded
        # from the compact-footprint treatment - unlike an actual grid or a
        # line of points, a cloud has no shape that a single bounding-box
        # summary conveys, so the user wants every point visible rather
        # than merged into one rectangle. Grid/line groups keep the
        # existing footprint behavior.
        groups_here = {
            g["group_key"]: g
            for g in (getattr(viewer, "spectro_groups_by_image", {}) or {}).get(str(file_key), [])
            if g["group_key"] in grouped_keys and (g.get("kind") or "cloud") != "cloud"
        }
        drawn_group_keys = set()
        drawn_groups = []
        for group_key, group in groups_here.items():
            member_specs = [s for s in singles if s.get('spectro_group_key') == group_key]
            if not member_specs:
                continue
            rect = viewer._matrix_bbox_pixels(
                member_specs, header, xpix, ypix, w_scale, h_scale, file_key, thumb_crop=thumb_crop
            )
            if rect is None:
                continue
            kind = group.get("kind") or "cloud"
            footprint_color = color_matrix if kind == "grid" else color_cloud
            chip_rect = _draw_footprint(painter, rect, kind=kind, color=footprint_color, count=len(member_specs))
            markers.append({
                'rect': rect,
                'spec': member_specs[0],
                'label': 'spectro-group',
                'kind': kind,
                'group': group,
                'tooltip': group.get("summary") or group.get("display"),
            })
            drawn_group_keys.add(group_key)
            drawn_groups.append(group)
        if drawn_group_keys:
            singles = [s for s in singles if s.get('spectro_group_key') not in drawn_group_keys]
    else:
        drawn_groups = []

    # Matrix footprints (skip when explicitly rendering matrix entries as individual points)
    if viewer.show_matrix_markers and matrices and not matrix_as_points:
        matrix_color = QtGui.QColor(getattr(viewer, 'spectro_marker_color_matrix', QtGui.QColor(64, 200, 255, 200)))
        for m_specs in matrices.values():
            rect = viewer._matrix_bbox_pixels(
                m_specs, header, xpix, ypix, w_scale, h_scale, file_key, thumb_crop=thumb_crop
            )
            if rect is None:
                continue
            try:
                grid_cols = m_specs[0].get('grid_cols')
                grid_rows = m_specs[0].get('grid_rows')
                dims = f"{grid_cols}×{grid_rows}" if grid_cols and grid_rows else None
            except Exception:
                dims = None
            chip_rect = _draw_footprint(painter, rect, kind="matrix", color=matrix_color, count=len(m_specs), dims=dims)
            tooltip = Path(m_specs[0].get('path', '')).name
            if dims:
                tooltip = f"{tooltip}\nGrid: {dims}"
            markers.append({'rect': rect, 'spec': m_specs[0], 'label': dims or 'grid', 'kind': 'matrix', 'tooltip': tooltip})

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
            is_matrix_spec = viewer._is_matrix_spec(spec)
            base_color = color_matrix if is_matrix_spec else color_single
            low_conf = str(spec.get("assignment_confidence") or "").strip().lower() == "low"
            rect = _draw_marker_symbol(
                painter,
                x,
                y,
                marker_symbol,
                marker_size,
                base_color,
                highlight=highlight,
                pulse=pulse if highlight else 1.0,
                low_conf=low_conf,
            )
            markers.append({'rect': rect, 'spec': spec, 'label': ''})
        for badge in badge_defs:
            bx = float(badge["col"]) * w_scale
            by = float(badge["row"]) * h_scale
            rect = _draw_stack_icon(painter, bx, by, badge["count"], color=color_stack)
            markers.append({
                'rect': rect,
                'spec': badge.get("spec"),
                'label': 'stack-badge',
                'tooltip': badge.get("tooltip"),
            })
    # summary badge: plain-language description of what's on the image
    # (grids, clusters, sites, points) instead of the old "S:132 M:1" shorthand
    try:
        text = _summary_badge_text(matrices, drawn_groups, singles)
        icon_kind = "grid" if (matrices or any((g.get('kind') or '') == 'grid' for g in drawn_groups)) else None
        icon_w = 15.0 if icon_kind else 0.0
        metrics = QtGui.QFontMetrics(QtGui.QFont("Segoe UI", 8, QtGui.QFont.Bold))
        badge_w = metrics.horizontalAdvance(text) + icon_w + 16
        by = 6
        bx = pixmap.width() - badge_w - 6
        rect = _draw_summary_badge(painter, bx, by, text, icon_kind=icon_kind)
        markers.append({'rect': rect, 'spec': None, 'label': 'badge'})
    except Exception:
        pass

    painter.end()
    return markers


def _legend_pixmap(draw_callable, width=64, height=44):
    """Render one legend icon by invoking the same drawing code the real
    overlays use, so the legend always matches what is on screen."""
    pix = QtGui.QPixmap(width, height)
    pix.fill(QtCore.Qt.transparent)
    painter = QtGui.QPainter(pix)
    painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
    try:
        draw_callable(painter, width, height)
    except Exception:
        pass
    painter.end()
    return pix


def show_marker_legend_dialog(viewer):
    """Non-modal 'what do these markers mean?' legend for the spectroscopy
    overlays, rendered with the real drawing functions and the viewer's actual
    marker style/colors."""
    symbol = _normalized_symbol(viewer)
    color_single = QtGui.QColor(getattr(viewer, "spectro_marker_color_single", None) or QtGui.QColor(255, 160, 0, 200))
    color_matrix = QtGui.QColor(getattr(viewer, "spectro_marker_color_matrix", None) or QtGui.QColor(64, 200, 255, 200))
    color_stack = QtGui.QColor(getattr(viewer, "spectro_marker_color_stack", None) or QtGui.QColor(165, 141, 242, 235))
    color_cloud = QtGui.QColor(getattr(viewer, "spectro_marker_color_cloud", None) or QtGui.QColor(95, 191, 147, 210))

    rows = [
        (
            lambda p, w, h: _draw_marker_symbol(p, w / 2, h / 2, symbol, 6.0, color_single),
            "<b>Single point</b> - one spectrum measured at this position, standing alone. "
            "Click it to open the spectrum; Shift+click to add it to the comparison selection.",
        ),
        (
            lambda p, w, h: _draw_footprint(p, QtCore.QRectF(6, 6, w - 12, h - 12), kind="grid", color=color_matrix, count=0, dims="4×4"),
            "<b>Grid / matrix</b> - a .3ds map or dense point-grid, shown as one footprint "
            "with its point count instead of a marker per pixel. Click it to open the Grid map explorer.",
        ),
        (
            lambda p, w, h: _draw_stack_icon(p, w / 2, h * 0.82, 3, color=color_stack),
            "<b>Z-stack / repeated site</b> - several traces measured at the same position "
            "(e.g. a Z series). The number on the stack is how many. Click it for the position summary.",
        ),
        (
            lambda p, w, h: _draw_footprint(p, QtCore.QRectF(6, 6, w - 12, h - 12), kind="cloud", color=color_cloud, count=5),
            "<b>Loose cluster</b> - points that group by position but aren't a formal grid "
            "file, shown as one footprint (dotted icon vs. the grid's dot-grid icon).",
        ),
        (
            lambda p, w, h: _draw_marker_symbol(p, w / 2, h / 2, symbol, 6.0, color_single, low_conf=True),
            "<b>Low-confidence link</b> - the dashed ring means the spectrum-to-image "
            "assignment is a guess. Review it via Spectroscopy → Review low confidence, "
            "or right-click the marker to assign it manually.",
        ),
        (
            lambda p, w, h: _draw_marker_symbol(p, w / 2, h / 2, symbol, 6.0, color_single, highlight=True, pulse=1.0),
            "<b>Highlight glow</b> - the spectrum you last opened or selected pulses "
            "on its source image (toggle in Spectroscopy → Spectro highlight glow).",
        ),
        (
            lambda p, w, h: _draw_summary_badge(p, (w - 90) / 2, (h - 18) / 2, "4×4 grid · 3 points", icon_kind="grid"),
            "<b>Image summary badge</b> - corner badge describing what's on this image in "
            "plain words (grids, clusters, sites, points). Click it to open the per-image "
            "spectroscopy summary.",
        ),
    ]

    dlg = QtWidgets.QDialog(viewer)
    dlg.setWindowTitle("Spectroscopy markers explained")
    dlg.setWindowModality(QtCore.Qt.NonModal)
    dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
    layout = QtWidgets.QVBoxLayout(dlg)
    intro = QtWidgets.QLabel(
        "Markers drawn on thumbnails and previews, using your current marker style. "
        "Reopen this window after changing marker style/colors to see the update."
    )
    intro.setWordWrap(True)
    layout.addWidget(intro)
    grid = QtWidgets.QGridLayout()
    grid.setHorizontalSpacing(12)
    grid.setVerticalSpacing(10)
    for row_idx, (draw_fn, text) in enumerate(rows):
        icon_lbl = QtWidgets.QLabel()
        icon_lbl.setPixmap(_legend_pixmap(draw_fn))
        icon_lbl.setFixedSize(68, 48)
        icon_lbl.setAlignment(QtCore.Qt.AlignCenter)
        icon_lbl.setStyleSheet("background-color: rgba(90, 110, 140, 60); border-radius: 6px;")
        grid.addWidget(icon_lbl, row_idx, 0)
        text_lbl = QtWidgets.QLabel(text)
        text_lbl.setWordWrap(True)
        text_lbl.setTextFormat(QtCore.Qt.RichText)
        grid.addWidget(text_lbl, row_idx, 1)
    grid.setColumnStretch(1, 1)
    layout.addLayout(grid)
    close_btn = QtWidgets.QPushButton("Close")
    close_btn.clicked.connect(dlg.close)
    btn_row = QtWidgets.QHBoxLayout()
    btn_row.addStretch(1)
    btn_row.addWidget(close_btn)
    layout.addLayout(btn_row)
    dlg.resize(560, 460)
    dlg.show()
    try:
        if hasattr(viewer, "_popup_refs"):
            viewer._popup_refs.append(dlg)
            dlg.finished.connect(lambda _: viewer._popup_refs.remove(dlg) if dlg in viewer._popup_refs else None)
    except Exception:
        pass
    return dlg


__all__ = [
    "_spectros_near_thumb_pos",
    "_render_spectroscopy_overlays",
    "_spread_overlapping_marker_coords",
    "_stack_badges_from_coords",
    "show_marker_legend_dialog",
]



