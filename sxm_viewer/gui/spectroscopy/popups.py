"""Spectroscopy popup helpers for SXMGridViewer."""
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
from ...data.spectroscopy import _matrix_base_name, find_last_image_for_spec
from ..detail_panels import SpectroscopyPopup, SpectroscopyCompareDialog, MatrixSpectroViewer
from ..palettes import DEFAULT_COLOR_CYCLE

def _open_spectroscopy_popup(viewer, spec):
    if not spec:
        return None
    try:
        dlg = SpectroscopyPopup(spec, parent=viewer)
        try:
            dlg.setWindowModality(QtCore.Qt.NonModal)
            dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
            dlg.setWindowFlags(dlg.windowFlags() | QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)
            dlg.move(viewer._next_popup_pos())
        except Exception:
            pass
        dlg.show()
        viewer._spectro_popups.append(dlg)
        dlg.finished.connect(lambda _: viewer._remember_closed_spectro_dialog(dlg) if hasattr(viewer, "_remember_closed_spectro_dialog") else None)
        dlg.finished.connect(lambda _: viewer._spectro_popups.remove(dlg) if dlg in viewer._spectro_popups else None)
        return dlg
    except Exception as e:
        QtWidgets.QMessageBox.warning(viewer, "Spectroscopy", str(e))
        return None


def _open_multi_spectroscopy_popup(viewer):
    specs = list(viewer._multi_spec_selection)
    if len(specs) < 2:
        return
    dlg = viewer._multi_spectro_popups[0] if getattr(viewer, "_multi_spectro_popups", None) else None
    palette_name = getattr(viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE)
    if dlg is None or not dlg.isVisible():
        if dlg is not None:
            viewer._multi_spectro_popups = [dlg for dlg in viewer._multi_spectro_popups if dlg is not dlg]
        dlg = SpectroscopyCompareDialog(specs, parent=viewer, palette_name=palette_name)
        try:
            dlg.move(viewer._next_popup_pos())
        except Exception:
            pass
        dlg.show()
        viewer._multi_spectro_popups.append(dlg)
        dlg.finished.connect(lambda _: viewer._remember_closed_spectro_dialog(dlg) if hasattr(viewer, "_remember_closed_spectro_dialog") else None)
        dlg.finished.connect(lambda _: viewer._multi_spectro_popups.remove(dlg) if dlg in viewer._multi_spectro_popups else None)
    else:
        dlg.set_specs(specs)
        dlg.set_palette_name(palette_name)


def on_show_matrix_spectro_viewer(viewer):
    datasets = getattr(viewer, "matrix_datasets", {})
    if not datasets:
        QtWidgets.QMessageBox.information(viewer, "Matrix spectra", "No matrix spectroscopy datasets detected for this folder.")
        return
    bases = sorted(datasets.keys())
    item, ok = QtWidgets.QInputDialog.getItem(viewer, "Matrix spectroscopies", "Select matrix dataset:", bases, 0, False)
    if not ok or not item:
        return
    ds = datasets.get(item)
    if not ds:
        return
    specs = []
    for ch in ds.channels:
        specs.extend(_matrix_specs_for_file(viewer, ch['path']))
    if not specs:
        QtWidgets.QMessageBox.information(viewer, "Matrix spectra", "No spectra entries found for that dataset.")
        return
    anchor = _find_anchor_image_for_matrix(viewer, specs, ds.base)
    if anchor is None:
        QtWidgets.QMessageBox.warning(viewer, "Matrix spectra", "Could not find a preceding SXM image for this matrix dataset.")
        return
    entry = {'path': Path(anchor['path']), 'time': anchor.get('time')}
    dlg = MatrixSpectroViewer(
        viewer,
        entry,
        specs,
        dataset=ds,
        palette_name=getattr(viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE),
    )
    dlg.show()
    viewer._popup_refs.append(dlg)
    dlg.finished.connect(lambda _: viewer._remember_closed_spectro_dialog(dlg) if hasattr(viewer, "_remember_closed_spectro_dialog") else None)
    dlg.finished.connect(lambda _: viewer._popup_refs.remove(dlg) if dlg in viewer._popup_refs else None)


def _matrix_specs_for_file(viewer, dat_path):
    dat_path = str(dat_path)
    return [spec for spec in getattr(viewer, "matrix_spectros", []) if str(spec.get('path')) == dat_path]


def _find_anchor_image_for_matrix(viewer, specs, base_name):
    images = getattr(viewer, 'image_meta', [])
    if not images:
        return None
    base_name = (_matrix_base_name(base_name) or base_name).lower()
    candidates = [img for img in images if _matrix_base_name(Path(img['path']).stem).lower() == base_name]
    matrix_time = None
    for spec in specs:
        if spec.get('time'):
            matrix_time = spec['time']
            break
    if matrix_time is None:
        try:
            matrix_time = datetime.fromtimestamp(Path(specs[0].get('path')).stat().st_mtime)
        except Exception:
            matrix_time = None
    match = None
    if candidates:
        earlier = [img for img in candidates if img.get('time') and matrix_time and img['time'] <= matrix_time]
        if earlier:
            earlier.sort(key=lambda img: img['time'], reverse=True)
            match = earlier[0]
        else:
            candidates.sort(key=lambda img: abs((img.get('time') or datetime.min) - (matrix_time or datetime.min)))
            match = candidates[0]
    if not match:
        match = find_last_image_for_spec(matrix_time, images)
    return match
__all__ = [
    "_open_spectroscopy_popup",
    "_open_multi_spectroscopy_popup",
    "on_show_matrix_spectro_viewer",
]




