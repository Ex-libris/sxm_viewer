"""Interactive A/B image comparison for preview and popup canvases.

The controller keeps lightweight frozen copies of the selected views so a
comparison stays valid even if the source preview is changed or the source
popup is closed. The dialog focuses on rigid alignment controls that are
practical for microscopy troubleshooting: manual rotation/translation plus
auto-translate / auto-rigid helpers.
"""
from __future__ import annotations

from pathlib import Path

import numpy as np

from ..._shared import QtCore, QtGui, QtWidgets
from ..canvases.detail_preview_canvas import MultiPreviewCanvas

try:  # pragma: no cover - optional dependency
    from scipy import ndimage

    _HAS_SCIPY = True
except Exception:  # pragma: no cover - optional dependency
    ndimage = None
    _HAS_SCIPY = False

try:  # pragma: no cover - optional dependency
    from skimage.registration import phase_cross_correlation

    _HAS_SKIMAGE = True
except Exception:  # pragma: no cover - optional dependency
    phase_cross_correlation = None
    _HAS_SKIMAGE = False


def _safe_label(snapshot):
    if not snapshot:
        return "Empty"
    return str(snapshot.get("label") or "Unnamed view")


def _finite_fill(arr):
    data = np.asarray(arr, dtype=float)
    finite = np.isfinite(data)
    if not finite.any():
        return np.zeros_like(data, dtype=float), finite, 0.0
    fill = float(np.nanmedian(data[finite]))
    out = np.array(data, copy=True)
    out[~finite] = fill
    return out, finite, fill


def _resize_nearest(arr, shape):
    data = np.asarray(arr, dtype=float)
    if tuple(data.shape[:2]) == tuple(shape):
        return np.array(data, copy=True)
    rows = np.clip(np.rint(np.linspace(0, data.shape[0] - 1, shape[0])).astype(int), 0, data.shape[0] - 1)
    cols = np.clip(np.rint(np.linspace(0, data.shape[1] - 1, shape[1])).astype(int), 0, data.shape[1] - 1)
    return np.array(data[rows][:, cols], copy=True)


def _resize_to_shape(arr, shape):
    data = np.asarray(arr, dtype=float)
    if tuple(data.shape[:2]) == tuple(shape):
        return np.array(data, copy=True)
    if _HAS_SCIPY and ndimage is not None:
        zoom = (float(shape[0]) / float(max(1, data.shape[0])), float(shape[1]) / float(max(1, data.shape[1])))
        try:
            return np.array(ndimage.zoom(data, zoom=zoom, order=1, mode="nearest"), copy=True)
        except Exception:
            pass
    return _resize_nearest(data, shape)


def _window(shape):
    if len(shape) != 2:
        return 1.0
    wy = np.hanning(max(2, int(shape[0])))
    wx = np.hanning(max(2, int(shape[1])))
    return np.outer(wy, wx)


def _registration_image(arr):
    data, finite, _ = _finite_fill(arr)
    if finite.any():
        center = float(np.mean(data[finite]))
        scale = float(np.std(data[finite]))
    else:
        center = 0.0
        scale = 1.0
    scale = max(scale, 1e-6)
    normalized = (data - center) / scale
    return normalized * _window(normalized.shape)


def _phase_correlation_shift(reference, moving):
    ref = _registration_image(reference)
    mov = _registration_image(moving)
    if _HAS_SKIMAGE and phase_cross_correlation is not None:
        try:
            shift, error, _ = phase_cross_correlation(ref, mov, upsample_factor=10)
            dy, dx = shift
            score = 1.0 - float(error)
            return float(dx), float(dy), score
        except Exception:
            pass
    fft_ref = np.fft.fft2(ref)
    fft_mov = np.fft.fft2(mov)
    cross_power = fft_ref * np.conj(fft_mov)
    denom = np.abs(cross_power)
    denom = np.where(denom < 1e-12, 1.0, denom)
    corr = np.fft.ifft2(cross_power / denom)
    corr_abs = np.abs(corr)
    peak = np.unravel_index(int(np.argmax(corr_abs)), corr_abs.shape)
    shifts = []
    for idx, size in zip(peak, corr_abs.shape):
        shift = float(idx)
        if shift > (size / 2.0):
            shift -= float(size)
        shifts.append(shift)
    dy, dx = shifts
    return float(dx), float(dy), float(corr_abs[peak])


def _translate_integer(arr, shift_x, shift_y):
    data = np.asarray(arr, dtype=float)
    out = np.full(data.shape, np.nan, dtype=float)
    mask = np.zeros(data.shape, dtype=bool)
    dx = int(round(float(shift_x)))
    dy = int(round(float(shift_y)))
    src_y0 = max(0, -dy)
    src_y1 = min(data.shape[0], data.shape[0] - dy) if dy >= 0 else data.shape[0]
    src_x0 = max(0, -dx)
    src_x1 = min(data.shape[1], data.shape[1] - dx) if dx >= 0 else data.shape[1]
    dst_y0 = max(0, dy)
    dst_x0 = max(0, dx)
    height = max(0, src_y1 - src_y0)
    width = max(0, src_x1 - src_x0)
    if height <= 0 or width <= 0:
        return out, mask
    out[dst_y0 : dst_y0 + height, dst_x0 : dst_x0 + width] = data[src_y0:src_y1, src_x0:src_x1]
    mask[dst_y0 : dst_y0 + height, dst_x0 : dst_x0 + width] = True
    return out, mask


def _transform_with_mask(arr, rotation_deg, shift_x, shift_y):
    data, finite, fill = _finite_fill(arr)
    if _HAS_SCIPY and ndimage is not None:
        try:
            rotation_xy = _rotation_matrix(rotation_deg)
            rotation_rc = np.array(
                [
                    [rotation_xy[1, 1], rotation_xy[1, 0]],
                    [rotation_xy[0, 1], rotation_xy[0, 0]],
                ],
                dtype=float,
            )
            inverse_rc = rotation_rc.T
            center_rc = np.array(
                [
                    (max(1, int(data.shape[0])) - 1.0) / 2.0,
                    (max(1, int(data.shape[1])) - 1.0) / 2.0,
                ],
                dtype=float,
            )
            shift_rc = np.array([float(shift_y), float(shift_x)], dtype=float)
            offset = center_rc - inverse_rc @ (center_rc + shift_rc)
            transformed = ndimage.affine_transform(
                data,
                inverse_rc,
                offset=offset,
                order=1,
                mode="constant",
                cval=fill,
            )
            transformed_mask = ndimage.affine_transform(
                finite.astype(float),
                inverse_rc,
                offset=offset,
                order=0,
                mode="constant",
                cval=0.0,
            )
            valid = np.asarray(transformed_mask) > 0.5
            transformed = np.array(transformed, copy=True)
            transformed[~valid] = np.nan
            return transformed, valid
        except Exception:
            pass
    if abs(float(rotation_deg)) > 1e-9:
        raise RuntimeError("Rotation alignment requires scipy.")
    return _translate_integer(data, shift_x, shift_y)


def _match_intensity(reference, moving, mode):
    mode = str(mode or "None").strip().lower()
    valid = np.isfinite(reference) & np.isfinite(moving)
    if not valid.any():
        return np.array(moving, copy=True), 0.0
    ref = np.asarray(reference, dtype=float)[valid]
    mov = np.asarray(moving, dtype=float)[valid]
    if mode == "median match":
        offset = float(np.median(ref) - np.median(mov))
    elif mode == "mean match":
        offset = float(np.mean(ref) - np.mean(mov))
    else:
        offset = 0.0
    out = np.array(moving, copy=True)
    finite = np.isfinite(out)
    out[finite] += offset
    return out, offset


def _alignment_metrics(reference, moving):
    valid = np.isfinite(reference) & np.isfinite(moving)
    overlap = int(np.count_nonzero(valid))
    if overlap <= 1:
        return {
            "overlap": 0,
            "coverage": 0.0,
            "corr": float("nan"),
            "rmse": float("nan"),
            "mean_delta": float("nan"),
        }
    ref = np.asarray(reference, dtype=float)[valid]
    mov = np.asarray(moving, dtype=float)[valid]
    ref_center = ref - np.mean(ref)
    mov_center = mov - np.mean(mov)
    denom = float(np.linalg.norm(ref_center) * np.linalg.norm(mov_center))
    corr = float(np.dot(ref_center, mov_center) / denom) if denom > 1e-12 else float("nan")
    diff = ref - mov
    rmse = float(np.sqrt(np.mean(diff**2)))
    return {
        "overlap": overlap,
        "coverage": float(overlap) / float(reference.size or 1),
        "corr": corr,
        "rmse": rmse,
        "mean_delta": float(np.mean(diff)),
    }


def _image_center_xy(shape):
    """Return the pixel-space center used by scipy-style in-place rotations."""
    height = int(shape[0]) if shape else 0
    width = int(shape[1]) if shape else 0
    return np.array([(max(1, width) - 1.0) / 2.0, (max(1, height) - 1.0) / 2.0], dtype=float)


def _rotation_matrix(rotation_deg):
    """Return the point transform that matches scipy.ndimage.rotate(..., axes=(1, 0))."""
    radians = float(np.deg2rad(float(rotation_deg)))
    cos_a = float(np.cos(radians))
    sin_a = float(np.sin(radians))
    return np.array([[cos_a, -sin_a], [sin_a, cos_a]], dtype=float)


def _transform_points(points, shape, rotation_deg=0.0, shift_x=0.0, shift_y=0.0):
    """Apply the compare dialog transform to point pairs in image pixel coordinates."""
    pts = np.asarray(points, dtype=float)
    if pts.size == 0:
        return np.empty((0, 2), dtype=float)
    center = _image_center_xy(shape)
    rotated = (pts - center) @ _rotation_matrix(rotation_deg).T + center
    return rotated + np.array([float(shift_x), float(shift_y)], dtype=float)


def _inverse_transform_points(points, shape, rotation_deg=0.0, shift_x=0.0, shift_y=0.0):
    """Map displayed aligned-B landmark clicks back onto the original B pixel grid."""
    pts = np.asarray(points, dtype=float)
    if pts.size == 0:
        return np.empty((0, 2), dtype=float)
    center = _image_center_xy(shape)
    unshifted = pts - np.array([float(shift_x), float(shift_y)], dtype=float)
    return (unshifted - center) @ _rotation_matrix(rotation_deg) + center


def _fit_translation_from_points(points_a, points_b):
    """Estimate a pure translation from corresponding landmark pairs."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    if pts_a.shape != pts_b.shape or pts_a.ndim != 2 or pts_a.shape[1] != 2 or pts_a.shape[0] < 1:
        raise ValueError("Translation fit needs at least one complete A/B landmark pair.")
    delta = pts_a - pts_b
    shift = np.mean(delta, axis=0)
    residual = delta - shift
    rmse = float(np.sqrt(np.mean(np.sum(residual**2, axis=1)))) if residual.size else 0.0
    return 0.0, float(shift[0]), float(shift[1]), rmse


def _fit_shift_for_rotation(points_a, points_b, moving_shape, rotation_deg):
    """Estimate the post-rotation shift in the same parameterization used by the compare dialog."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    rotated = _transform_points(pts_b, moving_shape, rotation_deg=rotation_deg, shift_x=0.0, shift_y=0.0)
    delta = pts_a - rotated
    shift = np.mean(delta, axis=0)
    transformed = rotated + shift
    residual = pts_a - transformed
    rmse = float(np.sqrt(np.mean(np.sum(residual**2, axis=1)))) if residual.size else 0.0
    return float(shift[0]), float(shift[1]), rmse


def _wrap_angle_deg(angle_deg):
    """Normalize an angle to the [-180, 180) range used by the compare spin box."""
    wrapped = (float(angle_deg) + 180.0) % 360.0 - 180.0
    return 180.0 if wrapped == -180.0 else wrapped


def _refine_rigid_angle(points_a, points_b, moving_shape, seed_angle):
    """Refine the rigid-fit angle against the dialog's actual landmark residual model."""
    best = None
    for step, radius in ((1.0, 12.0), (0.2, 2.0), (0.05, 0.4)):
        if best is None:
            center = float(seed_angle)
        else:
            center = float(best[0])
        count = max(3, int(round((2.0 * radius) / step)) + 1)
        angles = np.linspace(center - radius, center + radius, count)
        current_best = best
        for angle in angles:
            trial_angle = _wrap_angle_deg(angle)
            shift_x, shift_y, rmse = _fit_shift_for_rotation(points_a, points_b, moving_shape, trial_angle)
            if current_best is None or rmse < current_best[3]:
                current_best = (trial_angle, shift_x, shift_y, rmse)
        best = current_best
    return best


def _fit_rigid_from_points(points_a, points_b, moving_shape):
    """Estimate rotation plus translation from corresponding landmark pairs."""
    pts_a = np.asarray(points_a, dtype=float)
    pts_b = np.asarray(points_b, dtype=float)
    if pts_a.shape != pts_b.shape or pts_a.ndim != 2 or pts_a.shape[1] != 2 or pts_a.shape[0] < 2:
        raise ValueError("Rigid fit needs at least two complete A/B landmark pairs.")
    centroid_a = np.mean(pts_a, axis=0)
    centroid_b = np.mean(pts_b, axis=0)
    aa = pts_a - centroid_a
    bb = pts_b - centroid_b
    covariance = bb.T @ aa
    u_mat, _singular, vt_mat = np.linalg.svd(covariance)
    rotation = vt_mat.T @ u_mat.T
    if np.linalg.det(rotation) < 0.0:
        vt_mat[-1, :] *= -1.0
        rotation = vt_mat.T @ u_mat.T
    angle = _wrap_angle_deg(np.degrees(np.arctan2(rotation[1, 0], rotation[0, 0])))
    refined = _refine_rigid_angle(pts_a, pts_b, moving_shape, angle)
    if refined is None:
        shift_x, shift_y, rmse = _fit_shift_for_rotation(pts_a, pts_b, moving_shape, angle)
        return angle, shift_x, shift_y, rmse
    return refined


class ImageCompareController:
    """Manage Compare A/B slots and the compare popup lifecycle."""

    def __init__(self, viewer):
        self.viewer = viewer
        self._slot_a = None
        self._slot_b = None
        self._dialog = None

    def menu_state(self):
        """Return lightweight state so canvas menus can enable/label compare actions."""
        return {
            "has_a": self._slot_a is not None,
            "has_b": self._slot_b is not None,
            "label_a": _safe_label(self._slot_a),
            "label_b": _safe_label(self._slot_b),
        }

    def handle_menu_action(self, action, view, canvas=None):
        """Dispatch compare actions coming from a preview canvas context menu."""
        action = str(action or "").strip().lower()
        try:
            if action == "set_a":
                self.set_slot("a", view, canvas=canvas)
            elif action == "set_b":
                self.set_slot("b", view, canvas=canvas)
            elif action == "compare_with_a":
                if self._slot_a is None:
                    self._show_missing_slot("A")
                    return
                self.set_slot("b", view, canvas=canvas, announce=False)
                self.open_compare_dialog()
            elif action == "compare_with_b":
                if self._slot_b is None:
                    self._show_missing_slot("B")
                    return
                self.set_slot("a", view, canvas=canvas, announce=False)
                self.open_compare_dialog()
            elif action == "open_compare":
                self.open_compare_dialog()
            elif action == "swap_compare":
                self.swap_slots()
            elif action == "clear_compare":
                self.clear_slots()
        except ValueError as exc:
            QtWidgets.QMessageBox.information(self.viewer, "Image comparison", str(exc))

    def set_slot(self, slot_name, view, canvas=None, announce=True):
        """Freeze the current view into Compare A or B."""
        snapshot = self._freeze_view(view, canvas=canvas)
        if slot_name == "a":
            self._slot_a = snapshot
        else:
            self._slot_b = snapshot
        if announce:
            self._show_feedback(f"Compare {slot_name.upper()}: {snapshot['label']}")
        self._sync_dialog()

    def swap_slots(self):
        """Swap A and B and refresh an open dialog."""
        if self._slot_a is None and self._slot_b is None:
            return
        self._slot_a, self._slot_b = self._slot_b, self._slot_a
        self._show_feedback("Swapped compare A and B")
        self._sync_dialog()

    def clear_slots(self):
        """Clear transient compare selections without closing the dialog."""
        self._slot_a = None
        self._slot_b = None
        self._show_feedback("Cleared compare A/B selection")
        self._sync_dialog()

    def open_compare_dialog(self):
        """Show or refresh the comparison popup when both slots are populated."""
        if self._slot_a is None or self._slot_b is None:
            self._show_missing_slot("A and B")
            return None
        if self._dialog is not None:
            try:
                if self._dialog.isVisible():
                    self._dialog.set_snapshots(self._slot_a, self._slot_b)
                    self._dialog.raise_()
                    self._dialog.activateWindow()
                    return self._dialog
            except Exception:
                self._dialog = None

        dlg = ImageCompareDialog(self, self.viewer, self._slot_a, self._slot_b)
        dlg.setAttribute(QtCore.Qt.WA_DeleteOnClose, True)
        dlg.finished.connect(lambda _=None: self._on_dialog_finished(dlg))
        self._dialog = dlg
        self._register_popup(dlg)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
        return dlg

    def _sync_dialog(self):
        if self._dialog is None:
            return
        try:
            if self._slot_a is None or self._slot_b is None:
                self._dialog.set_slots_pending(self._slot_a, self._slot_b)
            else:
                self._dialog.set_snapshots(self._slot_a, self._slot_b)
        except Exception:
            pass

    def _freeze_view(self, view, canvas=None):
        """Create a self-contained compare snapshot from a live view dictionary."""
        if not isinstance(view, dict):
            raise ValueError("Comparison requires a valid image view.")
        arr = view.get("arr")
        if arr is None:
            raise ValueError("Selected view does not contain image data.")
        data = np.asarray(arr, dtype=float)
        if data.ndim < 2 or data.size == 0:
            raise ValueError("Selected view is empty.")
        rel_axes = False
        if canvas is not None and hasattr(canvas, "_use_relative_axes"):
            try:
                rel_axes = bool(canvas._use_relative_axes(view))
            except Exception:
                rel_axes = False
        elif view.get("relative_axes"):
            rel_axes = True
        if rel_axes:
            data = np.flipud(data)
        raw_extent = view.get("extent_raw")
        if raw_extent is None:
            raw_extent = view.get("extent")
        extent = raw_extent
        if canvas is not None and hasattr(canvas, "_display_extent_for_view"):
            try:
                extent = canvas._display_extent_for_view(view, raw_extent)
            except Exception:
                extent = raw_extent
        if extent is not None:
            try:
                extent = tuple(float(v) for v in extent)
            except Exception:
                extent = None
        title = str(view.get("title") or "").strip()
        meta = dict(view.get("meta") or {})
        path_text = str(view.get("path") or meta.get("path") or meta.get("file_path") or "").strip()
        axis_unit = str(view.get("axis_unit") or meta.get("axis_unit") or ("px" if extent is None else "nm") or "px")
        color_unit = str(view.get("unit") or view.get("colorbar_label") or "").strip()
        channel_idx = view.get("channel_idx")
        if channel_idx is None:
            channel_idx = meta.get("channel_index")
        filename = Path(path_text).name if path_text else "Image"
        label_parts = [filename]
        if title and title != filename:
            label_parts.append(title)
        label = " | ".join(part for part in label_parts if part)
        if not label:
            label = title or "Image"
        return {
            "arr": np.array(data, copy=True),
            "extent": extent,
            "extent_raw": extent,
            "axis_unit": axis_unit,
            "unit": color_unit,
            "cmap": view.get("cmap", "viridis"),
            "clim": tuple(view.get("clim")) if view.get("clim") else None,
            "path": path_text,
            "channel_idx": channel_idx,
            "title": title or filename,
            "label": label,
        }

    def _show_feedback(self, text):
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self.viewer, self.viewer.rect(), 2200)
        except Exception:
            pass

    def _show_missing_slot(self, slot_text):
        QtWidgets.QMessageBox.information(
            self.viewer,
            "Image comparison",
            f"Set compare {slot_text} first from a preview or popup image.",
        )

    def _register_popup(self, dlg):
        refs = getattr(self.viewer, "_popup_refs", None)
        if refs is not None and dlg not in refs:
            refs.append(dlg)
        controller = getattr(self.viewer, "quick_crop_controller", None)
        if controller is not None:
            try:
                controller.update_popup_actions()
            except Exception:
                pass

    def _on_dialog_finished(self, dlg):
        if getattr(self.viewer, "_popup_refs", None) and dlg in self.viewer._popup_refs:
            self.viewer._popup_refs.remove(dlg)
        controller = getattr(self.viewer, "quick_crop_controller", None)
        if controller is not None:
            try:
                controller.update_popup_actions()
            except Exception:
                pass
        if self._dialog is dlg:
            self._dialog = None


class ImageCompareDialog(QtWidgets.QDialog):
    """Popup that aligns B onto A and renders the comparison views."""

    def __init__(self, controller, viewer, snapshot_a, snapshot_b):
        super().__init__(viewer)
        self.controller = controller
        self.viewer = viewer
        self._snapshot_a = None
        self._snapshot_b = None
        self._base_a = None
        self._base_b = None
        self._compare_axes = {}
        self._landmark_pairs = []
        self._landmark_overlay_artists = []
        self._last_landmark_rmse = None
        self._update_timer = QtCore.QTimer(self)
        self._update_timer.setSingleShot(True)
        self._update_timer.setInterval(40)
        self._update_timer.timeout.connect(self._rebuild_views)
        self.setWindowTitle("Compare A/B")
        self.setMinimumSize(860, 720)
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
            | QtCore.Qt.WindowSystemMenuHint
        )

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header_row = QtWidgets.QHBoxLayout()
        header_row.setSpacing(8)
        header_row.addWidget(QtWidgets.QLabel("A:", self))
        self.label_a = QtWidgets.QLabel("", self)
        self.label_a.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        header_row.addWidget(self.label_a, 1)
        header_row.addWidget(QtWidgets.QLabel("B:", self))
        self.label_b = QtWidgets.QLabel("", self)
        self.label_b.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        header_row.addWidget(self.label_b, 1)
        layout.addLayout(header_row)

        controls = QtWidgets.QHBoxLayout()
        controls.setSpacing(8)
        controls.addWidget(QtWidgets.QLabel("Auto mode:", self))
        self.auto_mode_combo = QtWidgets.QComboBox(self)
        self.auto_mode_combo.addItems(["Translate", "Rigid"])
        self.auto_mode_combo.setToolTip("Translate estimates x/y shift. Rigid searches rotation + shift.")
        controls.addWidget(self.auto_mode_combo)
        controls.addWidget(QtWidgets.QLabel("Intensity:", self))
        self.intensity_combo = QtWidgets.QComboBox(self)
        self.intensity_combo.addItems(["None", "Median match", "Mean match"])
        self.intensity_combo.setToolTip("Match the baseline of B to A before computing difference maps.")
        controls.addWidget(self.intensity_combo)
        controls.addWidget(QtWidgets.QLabel("Rotation:", self))
        self.rotation_spin = QtWidgets.QDoubleSpinBox(self)
        self.rotation_spin.setRange(-180.0, 180.0)
        self.rotation_spin.setDecimals(2)
        self.rotation_spin.setSingleStep(1.0)
        self.rotation_spin.setSuffix(" deg")
        self.rotation_spin.setKeyboardTracking(False)
        self.rotation_spin.setToolTip("Manual rotation applied to B before comparison.")
        controls.addWidget(self.rotation_spin)
        controls.addWidget(QtWidgets.QLabel("Shift X:", self))
        self.shift_x_spin = QtWidgets.QDoubleSpinBox(self)
        self.shift_x_spin.setRange(-4096.0, 4096.0)
        self.shift_x_spin.setDecimals(2)
        self.shift_x_spin.setSingleStep(0.5)
        self.shift_x_spin.setSuffix(" px")
        self.shift_x_spin.setKeyboardTracking(False)
        self.shift_x_spin.setToolTip("Horizontal shift on the resampled A grid.")
        controls.addWidget(self.shift_x_spin)
        controls.addWidget(QtWidgets.QLabel("Shift Y:", self))
        self.shift_y_spin = QtWidgets.QDoubleSpinBox(self)
        self.shift_y_spin.setRange(-4096.0, 4096.0)
        self.shift_y_spin.setDecimals(2)
        self.shift_y_spin.setSingleStep(0.5)
        self.shift_y_spin.setSuffix(" px")
        self.shift_y_spin.setKeyboardTracking(False)
        self.shift_y_spin.setToolTip("Vertical shift on the resampled A grid.")
        controls.addWidget(self.shift_y_spin)

        self.auto_btn = QtWidgets.QPushButton("Auto fit", self)
        self.auto_btn.clicked.connect(self._auto_fit)
        controls.addWidget(self.auto_btn)
        self.reset_btn = QtWidgets.QPushButton("Reset", self)
        self.reset_btn.clicked.connect(self._reset_transform)
        controls.addWidget(self.reset_btn)
        self.swap_btn = QtWidgets.QPushButton("Swap A/B", self)
        self.swap_btn.clicked.connect(self._swap_slots)
        controls.addWidget(self.swap_btn)
        controls.addStretch(1)
        layout.addLayout(controls)

        landmark_row = QtWidgets.QHBoxLayout()
        landmark_row.setSpacing(8)
        self.pick_landmarks_btn = QtWidgets.QPushButton("Pick landmarks", self)
        self.pick_landmarks_btn.setCheckable(True)
        self.pick_landmarks_btn.setToolTip(
            "Alternate clicks between the A and B panels to define matching reference points."
        )
        self.pick_landmarks_btn.toggled.connect(self._on_pick_landmarks_toggled)
        landmark_row.addWidget(self.pick_landmarks_btn)
        landmark_row.addWidget(QtWidgets.QLabel("Point fit:", self))
        self.landmark_mode_combo = QtWidgets.QComboBox(self)
        self.landmark_mode_combo.addItems(["Translate", "Rigid"])
        self.landmark_mode_combo.setToolTip(
            "Translate uses one or more pairs. Rigid uses two or more pairs to solve rotation and shift."
        )
        self.landmark_mode_combo.currentIndexChanged.connect(lambda *_args: self._update_landmark_status())
        landmark_row.addWidget(self.landmark_mode_combo)
        self.fit_landmarks_btn = QtWidgets.QPushButton("Fit from points", self)
        self.fit_landmarks_btn.clicked.connect(self._fit_from_landmarks)
        landmark_row.addWidget(self.fit_landmarks_btn)
        self.undo_landmarks_btn = QtWidgets.QPushButton("Undo point", self)
        self.undo_landmarks_btn.clicked.connect(self._undo_landmark)
        landmark_row.addWidget(self.undo_landmarks_btn)
        self.clear_landmarks_btn = QtWidgets.QPushButton("Clear points", self)
        self.clear_landmarks_btn.clicked.connect(self._clear_landmarks)
        landmark_row.addWidget(self.clear_landmarks_btn)
        self.landmark_status = QtWidgets.QLabel("", self)
        self.landmark_status.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        landmark_row.addWidget(self.landmark_status, 1)
        layout.addLayout(landmark_row)

        self.metrics_label = QtWidgets.QLabel("", self)
        self.metrics_label.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        layout.addWidget(self.metrics_label)

        self.canvas = MultiPreviewCanvas(self, figsize=(8.5, 7.0))
        try:
            self.canvas.set_render_suspended(True)
        except Exception:
            pass
        self.canvas.set_compact_size_hints(True)
        self.canvas.set_window_arrange_callback(viewer.on_arrange_popouts)
        self.canvas.set_window_minimize_callback(viewer.on_minimize_popouts)
        self.canvas.set_window_restore_callback(viewer.on_restore_popouts)
        self.canvas.set_window_close_callback(viewer.on_close_popouts)
        self.canvas.show_molecules = False
        self.canvas._show_acquisition_overlay = False
        self.canvas._show_shortcut_hint = False
        self.canvas._show_profile_overlays = False
        self.canvas._show_angle_overlays = False
        self.canvas.scale_bar_enabled = False
        try:
            self.canvas.set_measurement_shortcuts_enabled(False)
        except Exception:
            self.canvas._measurement_shortcuts_enabled = False
        try:
            self.canvas.set_plot_font_family_callback(lambda fam: viewer.set_plot_font_family(fam))
            self.canvas.set_plot_font_family(getattr(viewer, "_plot_font_family", "sans-serif"))
        except Exception:
            pass
        self.canvas._detail_dark = bool(getattr(viewer, "detail_dark_view", False))
        self.canvas._detail_grid = bool(getattr(viewer, "detail_grid_view", False))
        self._landmark_click_cid = self.canvas.mpl_connect("button_press_event", self._on_canvas_click)
        layout.addWidget(self.canvas, 1)

        self.intensity_combo.currentIndexChanged.connect(lambda *_args: self._schedule_update())
        self.rotation_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.shift_x_spin.valueChanged.connect(lambda *_args: self._schedule_update())
        self.shift_y_spin.valueChanged.connect(lambda *_args: self._schedule_update())

        self.set_snapshots(snapshot_a, snapshot_b)
        self._update_landmark_status()
        try:
            self.canvas.set_render_suspended(False)
        except Exception:
            pass
        self.resize(1020, 820)

    def set_slots_pending(self, snapshot_a, snapshot_b):
        """Keep the dialog open but show that one of the slots is missing."""
        self._snapshot_a = snapshot_a
        self._snapshot_b = snapshot_b
        self._base_a = None
        self._base_b = None
        self._compare_axes = {}
        self._clear_landmark_overlays()
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self.label_a.setText(_safe_label(snapshot_a))
        self.label_b.setText(_safe_label(snapshot_b))
        self.metrics_label.setText("Both compare slots must be populated to render the comparison.")
        self.canvas.clear_views()
        self._update_landmark_status()

    def set_snapshots(self, snapshot_a, snapshot_b):
        """Replace A/B sources and rebuild the comparison views."""
        self._snapshot_a = snapshot_a
        self._snapshot_b = snapshot_b
        self.label_a.setText(_safe_label(snapshot_a))
        self.label_b.setText(_safe_label(snapshot_b))
        if snapshot_a is None or snapshot_b is None:
            self.set_slots_pending(snapshot_a, snapshot_b)
            return
        self._base_a = np.array(snapshot_a.get("arr"), copy=True)
        self._base_b = _resize_to_shape(snapshot_b.get("arr"), self._base_a.shape)
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._clear_landmark_overlays()
        span = float(max(self._base_a.shape[:2]))
        for spin in (self.shift_x_spin, self.shift_y_spin):
            spin.blockSignals(True)
            spin.setRange(-span, span)
            spin.blockSignals(False)
        self.setWindowTitle(f"Compare A/B - {_safe_label(snapshot_a)} vs {_safe_label(snapshot_b)}")
        self._update_landmark_status()
        self._schedule_update(immediate=True)

    def _schedule_update(self, immediate=False):
        if immediate:
            self._update_timer.stop()
            self._rebuild_views()
            return
        self._update_timer.start()

    def _reset_transform(self):
        """Reset manual alignment without clearing the current landmark pairs."""
        self.rotation_spin.blockSignals(True)
        self.shift_x_spin.blockSignals(True)
        self.shift_y_spin.blockSignals(True)
        self.rotation_spin.setValue(0.0)
        self.shift_x_spin.setValue(0.0)
        self.shift_y_spin.setValue(0.0)
        self.rotation_spin.blockSignals(False)
        self.shift_x_spin.blockSignals(False)
        self.shift_y_spin.blockSignals(False)
        self._schedule_update(immediate=True)

    def _swap_slots(self):
        """Delegate slot swapping to the controller so menus and popup state stay in sync."""
        self.controller.swap_slots()

    def _auto_fit(self):
        """Estimate a transform from the image content instead of user-supplied landmarks."""
        if self._base_a is None or self._base_b is None:
            return
        mode = str(self.auto_mode_combo.currentText() or "Translate")
        try:
            if mode == "Rigid":
                rotation, shift_x, shift_y = self._estimate_rigid_transform(self._base_a, self._base_b)
                self.rotation_spin.setValue(rotation)
            else:
                rotation = float(self.rotation_spin.value())
                shift_x, shift_y = self._estimate_translation(self._base_a, self._base_b, rotation)
            self.shift_x_spin.setValue(shift_x)
            self.shift_y_spin.setValue(shift_y)
            self._schedule_update(immediate=True)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to estimate alignment.\n{exc}")

    def _estimate_translation(self, reference, moving, rotation_deg):
        rotated, _ = _transform_with_mask(moving, rotation_deg, 0.0, 0.0)
        shift_x, shift_y, _ = _phase_correlation_shift(reference, rotated)
        return float(shift_x), float(shift_y)

    def _estimate_rigid_transform(self, reference, moving):
        if not _HAS_SCIPY or ndimage is None:
            raise RuntimeError("Rigid auto-alignment requires scipy.")
        ref_ds, mov_ds, scale = self._downsample_pair(reference, moving)
        coarse_angles = np.arange(-180.0, 180.0, 15.0, dtype=float)
        best = None
        for angles in (coarse_angles, None):
            if angles is None and best is not None:
                angles = np.arange(best[0] - 15.0, best[0] + 15.01, 3.0, dtype=float)
            for angle in angles:
                try:
                    shift_x, shift_y = self._estimate_translation(ref_ds, mov_ds, float(angle))
                    aligned, _ = _transform_with_mask(mov_ds, float(angle), shift_x, shift_y)
                    metrics = _alignment_metrics(ref_ds, aligned)
                    score = metrics.get("corr")
                    if score is None or not np.isfinite(score):
                        continue
                    if best is None or score > best[1]:
                        best = (float(angle), float(score), float(shift_x), float(shift_y))
                except Exception:
                    continue
        if best is None:
            raise RuntimeError("No valid rigid alignment candidate was found.")
        angle, _score, shift_x, shift_y = best
        if abs(scale - 1.0) > 1e-9:
            shift_x /= scale
            shift_y /= scale
        return float(angle), float(shift_x), float(shift_y)

    def _downsample_pair(self, reference, moving, max_size=256):
        ref = np.asarray(reference, dtype=float)
        mov = np.asarray(moving, dtype=float)
        longest = float(max(ref.shape[:2]))
        if longest <= float(max_size):
            return ref, mov, 1.0
        scale = float(max_size) / longest
        new_shape = (
            max(48, int(round(ref.shape[0] * scale))),
            max(48, int(round(ref.shape[1] * scale))),
        )
        return _resize_to_shape(ref, new_shape), _resize_to_shape(mov, new_shape), float(new_shape[1]) / float(ref.shape[1])

    def _rebuild_views(self):
        if self._snapshot_a is None or self._snapshot_b is None or self._base_a is None or self._base_b is None:
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        try:
            aligned_b, _mask = _transform_with_mask(self._base_b, rotation, shift_x, shift_y)
        except Exception as exc:
            self.metrics_label.setText(str(exc))
            return
        aligned_b, offset = _match_intensity(self._base_a, aligned_b, self.intensity_combo.currentText())
        diff = np.array(self._base_a, copy=True) - np.array(aligned_b, copy=True)
        abs_diff = np.abs(diff)
        metrics = _alignment_metrics(self._base_a, aligned_b)
        self.metrics_label.setText(
            "Overlap: {coverage:.1%} | Corr: {corr:.4f} | RMSE: {rmse:.4g} | Mean(A-B): {mean_delta:.4g} | "
            "Offset(B): {offset:.4g} | rot={rotation:.2f} deg, dx={shift_x:.2f} px, dy={shift_y:.2f} px".format(
                coverage=metrics.get("coverage", 0.0),
                corr=metrics.get("corr", float("nan")),
                rmse=metrics.get("rmse", float("nan")),
                mean_delta=metrics.get("mean_delta", float("nan")),
                offset=offset,
                rotation=rotation,
                shift_x=shift_x,
                shift_y=shift_y,
            )
        )
        views = self._build_compare_views(aligned_b, diff, abs_diff, offset)
        self.canvas.set_views(views, preserve_profiles=False)
        self._refresh_compare_axes()
        self._draw_landmark_overlays()

    def _build_compare_views(self, aligned_b, diff, abs_diff, offset):
        extent = self._snapshot_a.get("extent")
        unit = str(self._snapshot_a.get("unit") or "").strip()
        axis_unit = self._snapshot_a.get("axis_unit")
        view_a = self._view_payload(
            self._base_a,
            title=f"A: {_safe_label(self._snapshot_a)}",
            cmap=self._snapshot_a.get("cmap", "viridis"),
            clim=self._snapshot_a.get("clim"),
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=unit,
        )
        clim_b = self._snapshot_b.get("clim")
        if clim_b is not None:
            try:
                clim_b = (float(clim_b[0]) + float(offset), float(clim_b[1]) + float(offset))
            except Exception:
                clim_b = None
        view_b = self._view_payload(
            aligned_b,
            title=f"B aligned: {_safe_label(self._snapshot_b)}",
            cmap=self._snapshot_b.get("cmap", "viridis"),
            clim=clim_b,
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=unit,
        )
        finite_diff = diff[np.isfinite(diff)]
        if finite_diff.size:
            diff_span = float(np.nanpercentile(np.abs(finite_diff), 98))
        else:
            diff_span = 1.0
        diff_span = max(diff_span, 1e-9)
        finite_abs = abs_diff[np.isfinite(abs_diff)]
        if finite_abs.size:
            abs_span = float(np.nanpercentile(finite_abs, 98))
        else:
            abs_span = 1.0
        abs_span = max(abs_span, 1e-9)
        view_diff = self._view_payload(
            diff,
            title="A - B",
            cmap="coolwarm",
            clim=(-diff_span, diff_span),
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=f"diff {unit}".strip(),
        )
        view_abs = self._view_payload(
            abs_diff,
            title="|A - B|",
            cmap="inferno",
            clim=(0.0, abs_span),
            extent=extent,
            axis_unit=axis_unit,
            colorbar_label=f"abs diff {unit}".strip(),
        )
        return [view_a, view_b, view_diff, view_abs]

    @staticmethod
    def _view_payload(arr, *, title, cmap, clim, extent, axis_unit, colorbar_label):
        payload = {
            "arr": np.array(arr, copy=True),
            "title": title,
            "cmap": cmap,
            "axis_unit": axis_unit,
            "colorbar_label": colorbar_label,
            "unit": colorbar_label,
        }
        if clim is not None:
            payload["clim"] = clim
        if extent is not None:
            payload["extent"] = extent
            payload["extent_raw"] = extent
        return payload

    def _on_pick_landmarks_toggled(self, _checked):
        """Update the instruction label when interactive landmark picking is toggled."""
        self._update_landmark_status()
        if self.pick_landmarks_btn.isChecked():
            self._show_landmark_hint("Click a reference point in A, then the matching point in B.")

    def _fit_from_landmarks(self):
        """Solve the manual alignment from the currently completed landmark pairs."""
        if self._base_a is None or self._base_b is None:
            return
        points_a, points_b = self._complete_landmark_arrays()
        mode = str(self.landmark_mode_combo.currentText() or "Translate")
        try:
            if mode == "Rigid":
                rotation, shift_x, shift_y, rmse = _fit_rigid_from_points(points_a, points_b, self._base_b.shape)
            else:
                rotation, shift_x, shift_y, rmse = _fit_translation_from_points(points_a, points_b)
            self.rotation_spin.blockSignals(True)
            self.shift_x_spin.blockSignals(True)
            self.shift_y_spin.blockSignals(True)
            self.rotation_spin.setValue(rotation)
            self.shift_x_spin.setValue(shift_x)
            self.shift_y_spin.setValue(shift_y)
            self.rotation_spin.blockSignals(False)
            self.shift_x_spin.blockSignals(False)
            self.shift_y_spin.blockSignals(False)
            self._last_landmark_rmse = float(rmse)
            self._update_landmark_status()
            self._schedule_update(immediate=True)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Image comparison", f"Unable to fit the selected landmarks.\n{exc}")

    def _undo_landmark(self):
        """Remove the last picked point so a mis-click can be corrected quickly."""
        if not self._landmark_pairs:
            return
        last = self._landmark_pairs[-1]
        if last.get("b") is not None:
            last["b"] = None
        else:
            self._landmark_pairs.pop()
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _clear_landmarks(self):
        """Drop all manually picked points while leaving the current transform untouched."""
        self._landmark_pairs = []
        self._last_landmark_rmse = None
        self._clear_landmark_overlays()
        self._update_landmark_status()

    def _update_landmark_status(self):
        """Refresh the compact status text and enablement for landmark-related controls."""
        complete_pairs = sum(1 for pair in self._landmark_pairs if pair.get("a") is not None and pair.get("b") is not None)
        total_pairs = len(self._landmark_pairs)
        expected = self._expected_landmark_target()
        next_index = total_pairs + 1 if expected == "a" else total_pairs
        required_pairs = 2 if str(self.landmark_mode_combo.currentText() or "Translate") == "Rigid" else 1
        if self._base_a is None or self._base_b is None:
            text = "Landmarks unavailable until both compare slots are set."
        elif self.pick_landmarks_btn.isChecked():
            text = f"{complete_pairs} pair(s) stored. Next click: {expected.upper()}{max(1, next_index)}."
        elif complete_pairs:
            text = f"{complete_pairs} pair(s) stored."
        else:
            text = "Toggle 'Pick landmarks' to add matching A/B reference points."
        if self._base_a is not None and complete_pairs < required_pairs:
            text += f" {required_pairs} complete pair(s) required for {self.landmark_mode_combo.currentText().lower()} fit."
        if self._last_landmark_rmse is not None and np.isfinite(self._last_landmark_rmse):
            text += f" Fit RMSE: {self._last_landmark_rmse:.2f} px."
        self.landmark_status.setText(text)
        self.fit_landmarks_btn.setEnabled(
            complete_pairs >= required_pairs and self._base_a is not None and self._base_b is not None
        )
        self.undo_landmarks_btn.setEnabled(bool(self._landmark_pairs))
        self.clear_landmarks_btn.setEnabled(bool(self._landmark_pairs))

    def _expected_landmark_target(self):
        """Return which panel should receive the next click in the alternating A/B workflow."""
        if not self._landmark_pairs or self._landmark_pairs[-1].get("b") is not None:
            return "a"
        return "b"

    def _complete_landmark_arrays(self):
        """Return completed landmark pairs as two Nx2 float arrays."""
        points_a = []
        points_b = []
        for pair in self._landmark_pairs:
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is None or point_b is None:
                continue
            points_a.append(point_a)
            points_b.append(point_b)
        if not points_a:
            raise ValueError("Pick at least one complete A/B landmark pair first.")
        return np.asarray(points_a, dtype=float), np.asarray(points_b, dtype=float)

    def _refresh_compare_axes(self):
        """Cache the first two compare axes so clicks and overlays only target A and B."""
        axes = list(getattr(self.canvas, "_ax_view_map", {}).keys())
        self._compare_axes = {
            "a": axes[0] if len(axes) > 0 else None,
            "b": axes[1] if len(axes) > 1 else None,
        }

    def _axis_role(self, ax):
        """Map a matplotlib axis back to the A or B compare panel."""
        if not self._compare_axes:
            self._refresh_compare_axes()
        for role, target_ax in self._compare_axes.items():
            if ax is target_ax:
                return role
        return None

    def _on_canvas_click(self, event):
        """Capture landmark clicks on the A/B panels without touching the diff panels."""
        if not self.pick_landmarks_btn.isChecked() or self._base_a is None or self._base_b is None:
            return
        if getattr(event, "button", None) != 1 or event.xdata is None or event.ydata is None:
            return
        role = self._axis_role(getattr(event, "inaxes", None))
        if role not in {"a", "b"}:
            return
        expected = self._expected_landmark_target()
        if role != expected:
            self._show_landmark_hint(f"Click the next reference point in {expected.upper()} first.")
            return
        try:
            point = self._event_to_landmark_point(role, event)
        except ValueError as exc:
            self._show_landmark_hint(str(exc))
            return
        if role == "a":
            self._landmark_pairs.append({"a": point, "b": None})
        else:
            self._landmark_pairs[-1]["b"] = point
        self._last_landmark_rmse = None
        self._update_landmark_status()
        self._draw_landmark_overlays()

    def _event_to_landmark_point(self, role, event):
        """Convert a canvas click into stored A-grid or raw-B pixel coordinates."""
        ax = getattr(event, "inaxes", None)
        view = getattr(self.canvas, "_ax_view_map", {}).get(ax)
        if view is None:
            raise ValueError("The clicked panel does not contain a comparison image.")
        arr = np.asarray(view.get("arr")) if view.get("arr") is not None else None
        if arr is None or arr.ndim < 2 or arr.size == 0:
            raise ValueError("The clicked comparison image is empty.")
        height, width = arr.shape[:2]
        point = np.array(
            [
                float(self.canvas._axis_coord_to_pixel_float(view, event.xdata, width, "x", ax=ax)),
                float(self.canvas._axis_coord_to_pixel_float(view, event.ydata, height, "y", ax=ax)),
            ],
            dtype=float,
        )
        if role == "a":
            if not self._point_in_bounds(point, self._base_a.shape):
                raise ValueError("A landmark click must stay inside the visible image.")
            return tuple(point)
        base_point = _inverse_transform_points(
            [point],
            self._base_b.shape,
            rotation_deg=float(self.rotation_spin.value()),
            shift_x=float(self.shift_x_spin.value()),
            shift_y=float(self.shift_y_spin.value()),
        )[0]
        if not self._point_in_bounds(base_point, self._base_b.shape):
            raise ValueError("The clicked B point maps outside the source image. Reset or refine the transform first.")
        return tuple(base_point)

    def _point_in_bounds(self, point, shape):
        """Return True when a pixel-space landmark lies inside the target image."""
        if point is None or shape is None or len(shape) < 2:
            return False
        x_val = float(point[0])
        y_val = float(point[1])
        height = int(shape[0])
        width = int(shape[1])
        return 0.0 <= x_val <= float(max(0, width - 1)) and 0.0 <= y_val <= float(max(0, height - 1))

    def _clear_landmark_overlays(self):
        """Remove existing landmark markers before re-drawing them on the latest axes."""
        for artist in self._landmark_overlay_artists:
            try:
                artist.remove()
            except Exception:
                pass
        self._landmark_overlay_artists = []

    def _draw_landmark_overlays(self):
        """Repaint numbered landmark markers on the A and aligned-B panels."""
        self._clear_landmark_overlays()
        if not self._landmark_pairs:
            try:
                self.canvas.draw_idle()
            except Exception:
                pass
            return
        self._refresh_compare_axes()
        ax_a = self._compare_axes.get("a")
        ax_b = self._compare_axes.get("b")
        if ax_a is None or ax_b is None:
            return
        view_a = self.canvas._ax_view_map.get(ax_a)
        view_b = self.canvas._ax_view_map.get(ax_b)
        if view_a is None or view_b is None:
            return
        rotation = float(self.rotation_spin.value())
        shift_x = float(self.shift_x_spin.value())
        shift_y = float(self.shift_y_spin.value())
        for index, pair in enumerate(self._landmark_pairs, start=1):
            point_a = pair.get("a")
            point_b = pair.get("b")
            if point_a is not None:
                self._add_landmark_marker(ax_a, view_a, point_a, index, color="#ffd24d")
            if point_b is not None:
                aligned_point = _transform_points(
                    [point_b],
                    self._base_b.shape,
                    rotation_deg=rotation,
                    shift_x=shift_x,
                    shift_y=shift_y,
                )[0]
                self._add_landmark_marker(ax_b, view_b, aligned_point, index, color="#7ee0ff")
        try:
            self.canvas.draw_idle()
        except Exception:
            pass

    def _add_landmark_marker(self, ax, view, point, index, *, color):
        """Draw a single numbered landmark marker on the provided compare axis."""
        coords = self._pixel_to_axis_coords(view, point)
        if coords is None:
            return
        x_val, y_val = coords
        scatter = ax.scatter(
            [x_val],
            [y_val],
            s=64,
            c=color,
            edgecolors="black",
            linewidths=0.9,
            zorder=30,
        )
        text = ax.text(
            x_val,
            y_val,
            f" {index}",
            color="white",
            fontsize=9,
            fontweight="bold",
            ha="left",
            va="bottom",
            zorder=31,
        )
        self._landmark_overlay_artists.extend([scatter, text])

    def _pixel_to_axis_coords(self, view, point):
        """Convert stored pixel-space landmarks back into the current axis coordinate system."""
        if view is None or point is None:
            return None
        arr = np.asarray(view.get("arr")) if view.get("arr") is not None else None
        if arr is None or arr.ndim < 2 or arr.size == 0:
            return None
        height, width = arr.shape[:2]
        x_idx = float(point[0])
        y_idx = float(point[1])
        try:
            extent = self.canvas._view_extent(view)
        except Exception:
            extent = None
        if extent is None:
            return x_idx, y_idx
        x_val = float(self.canvas._index_to_axis_coord(x_idx, extent[0], extent[1], width))
        y_val = float(self.canvas._index_to_axis_coord(y_idx, extent[2], extent[3], height))
        return x_val, y_val

    def _show_landmark_hint(self, text):
        """Show lightweight point-picking feedback without interrupting the dialog workflow."""
        try:
            QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text, self, self.rect(), 2200)
        except Exception:
            pass
