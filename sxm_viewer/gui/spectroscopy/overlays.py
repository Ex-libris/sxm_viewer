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
    yspan = y0 - y1 if y0 != y1 else 1.0
    sx = x0 + (px / float(max(1, w - 1))) * xspan
    sy = y1 + ((py / float(max(1, h - 1))) * -yspan)
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


def _render_spectroscopy_overlays(viewer, pixmap, header, file_key, xpix, ypix, reveal_points_override=None, selected_spec=None, entries_override=None, matrix_as_points=False):
    """Render spectroscopy overlays with density + matrix footprints."""
    if not viewer.show_spectra:
        return []
    specs = entries_override if entries_override is not None else viewer.spectros_by_image.get(file_key, [])
    if not specs:
        return []
    markers = []
    painter = QtGui.QPainter(pixmap)
    painter.setRenderHint(QtGui.QPainter.Antialiasing, True)
    w_scale = pixmap.width() / max(1, xpix - 1)
    h_scale = pixmap.height() / max(1, ypix - 1)
    if reveal_points_override is None:
        reveal_points = hasattr(viewer, '_temp_reveal') and file_key in getattr(viewer, '_temp_reveal', set())
    else:
        reveal_points = bool(reveal_points_override)

    singles = [s for s in specs if s.get('matrix_index') is None]
    matrices = {}
    for s in specs:
        if s.get('matrix_index') is None:
            continue
        key = str(s.get('path'))
        matrices.setdefault(key, []).append(s)

    # When requested (e.g., matrix preview dialog), render matrix entries as points too.
    if matrix_as_points and matrices:
        flat_matrix_entries = []
        for ms in matrices.values():
            flat_matrix_entries.extend(ms)
        singles = singles + flat_matrix_entries

    # Matrix footprints (skip when explicitly rendering matrix entries as individual points)
    if viewer.show_matrix_markers and matrices and not matrix_as_points:
        for m_specs in matrices.values():
            rect = viewer._matrix_bbox_pixels(m_specs, header, xpix, ypix, w_scale, h_scale, file_key)
            if rect is None:
                continue
            fill = QtGui.QColor(0, 205, 255, 90)
            pen = QtGui.QPen(QtGui.QColor(0, 180, 230))
            pen.setWidth(3)
            painter.setBrush(fill)
            painter.setPen(pen)
            painter.drawRect(rect)
            try:
                grid_cols = m_specs[0].get('grid_cols')
                grid_rows = m_specs[0].get('grid_rows')
                label = f"{grid_cols}x{grid_rows}" if grid_cols and grid_rows else f"{len(m_specs)}"
            except Exception:
                label = "M"
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255)))
            painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
            painter.drawText(rect, QtCore.Qt.AlignCenter, label)
            markers.append({'rect': rect, 'spec': m_specs[0], 'label': label, 'kind': 'matrix'})

    color_single = getattr(viewer, 'spectro_marker_color_single', QtGui.QColor(255, 160, 0, 200))
    color_matrix = getattr(viewer, 'spectro_marker_color_matrix', QtGui.QColor(64, 200, 255, 200))

    # Single spectroscopies (density or points)
    if (viewer.show_single_markers or reveal_points or matrix_as_points) and singles:
        coords = []
        for idx, spec in enumerate(singles, 1):
            c = viewer._map_spec_to_pixels(spec, header, xpix, ypix, file_key)
            if c is None:
                c = viewer._fallback_spec_coords(idx, xpix, ypix)
            col, row = c
            coords.append((col * w_scale, row * h_scale, spec))

        coords_xy = np.array([(c[0], c[1]) for c in coords], dtype=float)
        count = coords_xy.shape[0]
        use_density = (not reveal_points) and viewer.use_density_markers and viewer._use_density_for(count, pixmap.width(), pixmap.height())
        if matrix_as_points:
            use_density = False

        if use_density:
            viewer._draw_density_overlay(painter, coords_xy, pixmap.width(), pixmap.height())
            # add count badge for interaction
            pad = 6
            size = 18
            rect = QtCore.QRectF(pixmap.width() - size - pad, pad, size, size)
            painter.setBrush(QtGui.QColor(255, 160, 0, 230))
            pen = QtGui.QPen(QtGui.QColor(40, 30, 20))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawEllipse(rect)
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255)))
            painter.setFont(QtGui.QFont("Segoe UI", 9, QtGui.QFont.Bold))
            painter.drawText(rect, QtCore.Qt.AlignCenter, f"{count}")
            markers.append({'rect': rect, 'spec': singles[0], 'label': f"{count}"})
        else:
            # choose a compact, low-occlusion marker style when crowd is large
            crowded = count > 200 or viewer.compact_markers
            for x, y, spec in coords:
                highlight = False
                try:
                    if selected_spec and viewer._spec_identity_key(spec) == viewer._spec_identity_key(selected_spec):
                        highlight = True
                except Exception:
                    highlight = False
                base_color = color_matrix if spec.get('matrix_index') is not None else color_single
                if reveal_points or crowded:
                    # small cross, minimal occlusion
                    arm = 2 if crowded and not reveal_points else 4
                    color = QtGui.QColor(base_color)
                    color.setAlpha(180 if crowded else base_color.alpha())
                    pen = QtGui.QPen(color)
                    pen.setWidth(1)
                    painter.setPen(pen)
                    painter.drawLine(QtCore.QPointF(x-arm, y), QtCore.QPointF(x+arm, y))
                    painter.drawLine(QtCore.QPointF(x, y-arm), QtCore.QPointF(x, y+arm))
                    if highlight:
                        painter.setPen(QtGui.QPen(QtGui.QColor(0, 230, 255), 2))
                        painter.drawEllipse(QtCore.QPointF(x, y), arm+3, arm+3)
                    rect = QtCore.QRectF(x-arm-3, y-arm-3, (arm+3)*2, (arm+3)*2)
                else:
                    radius = 3
                    color = QtGui.QColor(base_color)
                    color.setAlpha(160)
                    pen = QtGui.QPen(color)
                    pen.setWidth(1)
                    painter.setPen(pen)
                    bc = QtGui.QColor(base_color)
                    bc.setAlpha(180)
                    painter.setBrush(bc)
                    painter.drawEllipse(QtCore.QPointF(x, y), radius, radius)
                    if highlight:
                        painter.setPen(QtGui.QPen(QtGui.QColor(0, 230, 255), 2))
                        painter.drawEllipse(QtCore.QPointF(x, y), radius+2, radius+2)
                    rect = QtCore.QRectF(x-radius-2, y-radius-2, (radius+2)*2, (radius+2)*2)
                markers.append({'rect': rect, 'spec': spec, 'label': ''})
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
]



