"""Spectroscopy popup helpers for SXMGridViewer."""
from __future__ import annotations

from ..._shared import *
from ...data.spectroscopy import _matrix_base_name, find_last_image_for_spec
from ..detail_panels import SpectroscopyPopup, SpectroscopyCompareDialog, MatrixSpectroViewer

def _open_spectroscopy_popup(viewer, spec):
    if not spec:
        return
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
        dlg.finished.connect(lambda _: viewer._spectro_popups.remove(dlg) if dlg in viewer._spectro_popups else None)
    except Exception as e:
        QtWidgets.QMessageBox.warning(viewer, "Spectroscopy", str(e))


def _open_multi_spectroscopy_popup(viewer):
    specs = list(viewer._multi_spec_selection)
    if len(specs) < 2:
        return
    # close previous multi popups
    for dlg in list(viewer._multi_spectro_popups):
        try:
            dlg.close()
        except Exception:
            pass
    viewer._multi_spectro_popups = []
    dlg = SpectroscopyCompareDialog(specs, parent=viewer)
    try:
        dlg.move(viewer._next_popup_pos())
    except Exception:
        pass
    dlg.show()
    viewer._multi_spectro_popups.append(dlg)
    dlg.finished.connect(lambda _: viewer._multi_spectro_popups.remove(dlg) if dlg in viewer._multi_spectro_popups else None)


def on_show_matrix_spectro_viewer(viewer):
    matrix_files = defaultdict(list)
    for spec in viewer.matrix_spectros:
        matrix_files[str(spec.get('path'))].append(spec)
    if not matrix_files:
        QtWidgets.QMessageBox.information(viewer, "Matrix spectra", "No matrix spectroscopy files detected for this folder.")
        return
    choices = sorted(matrix_files.items(), key=lambda item: Path(item[0]).name.lower())
    names = [Path(dat_path).name for dat_path, _ in choices]
    item, ok = QtWidgets.QInputDialog.getItem(viewer, "Matrix spectroscopies", "Select matrix file:", names, 0, False)
    if not ok or not item:
        return
    dat_key = None; target_specs = None
    for k, specs in choices:
        if Path(k).name == item:
            dat_key = k
            target_specs = specs
            break
    if not dat_key or not target_specs:
        return
    images = getattr(viewer, 'image_meta', [])
    if not images:
        QtWidgets.QMessageBox.information(viewer, "Matrix spectra", "No SXM images available to anchor matrix data.")
        return
    images = getattr(viewer, 'image_meta', [])
    if not images:
        QtWidgets.QMessageBox.information(viewer, "Matrix spectra", "No SXM images available to anchor matrix data.")
        return
    try:
        matrix_time = datetime.fromtimestamp(Path(dat_key).stat().st_mtime)
    except Exception:
        matrix_time = None
    if matrix_time is None:
        for spec in target_specs:
            if spec.get('time'):
                matrix_time = spec.get('time')
                break
    base_name = _matrix_base_name(Path(dat_key).stem).lower()
    candidates = [img for img in images if _matrix_base_name(Path(img['path']).stem).lower() == base_name]
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
    if not match:
        QtWidgets.QMessageBox.warning(viewer, "Matrix spectra", "Could not find a preceding SXM image for this matrix file.")
        return
    entry = {'path': Path(match['path']), 'time': match.get('time')}
    dlg = MatrixSpectroViewer(viewer, entry, target_specs)
    dlg.show()
    viewer._popup_refs.append(dlg)
    dlg.finished.connect(lambda _: viewer._popup_refs.remove(dlg) if dlg in viewer._popup_refs else None)
__all__ = [
    "_open_spectroscopy_popup",
    "_open_multi_spectroscopy_popup",
    "on_show_matrix_spectro_viewer",
]
