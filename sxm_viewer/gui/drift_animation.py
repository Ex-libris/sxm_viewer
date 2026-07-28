"""Drift correction, animation export, and the alignment preview dialog.

One workflow: take a stack of images of the same area, cross-correlate to
find inter-frame drift, then either review the alignment or export it as
an animation. Extracted from `main_window.py` (688 lines) as part of the
god-class reduction - see `docs/refactor/GOD_CLASS_PLAN.md`.

Follows the `gui/viewer/*.py` convention: plain module functions taking
the `SXMGridViewer` as `viewer`. Coupling back to the viewer is
deliberately small - four methods (`_get_channel_array`,
`show_file_channel`, `populate_thumbnails_for_channel`,
`_inherit_star_for_virtual_copy`) and a handful of attributes
(`headers`, `files`, `channel_dropdown`, `_processed_views`,
`thumb_cmap`) - which is what made this group safe to lift out whole.
"""
from __future__ import annotations

from pathlib import Path

from .._shared import QtCore, QtGui, QtWidgets, np
from .. import cmap_registry
from ..data.io import parse_header


def on_drift_correct(viewer, paths):
    if not paths:
        return
    try:
        from scipy import ndimage  # type: ignore
    except Exception:
        QtWidgets.QMessageBox.warning(viewer, "Drift correction", "scipy is required for alignment interpolation.")
        return
    # Prefer skimage phase_cross_correlation; fall back to OpenCV ECC; else zeros
    try:
        from skimage.registration import phase_cross_correlation  # type: ignore
    except Exception:
        phase_cross_correlation = None  # type: ignore
    try:
        import cv2  # type: ignore
        has_cv = True
    except Exception:
        has_cv = False
    channel_idx = viewer.channel_dropdown.currentIndex()
    images = []
    names_full = []
    names_display = []
    missing = 0
    # Preserve user selection order while dropping duplicates
    seen = set()
    ordered_paths = []
    for p in paths:
        ps = str(Path(p))
        if ps in seen:
            continue
        seen.add(ps)
        ordered_paths.append(ps)
    for p in ordered_paths:
        try:
            header, fds = viewer.headers.get(p, (None, None))
            if header is None or fds is None:
                header, fds = parse_header(Path(p))
            if not fds:
                continue
            # Prefer current channel, but fall back to any available channel with data
            indices = [channel_idx] + [i for i in range(len(fds)) if i != channel_idx]
            arr = None
            for idx in indices:
                if idx < 0 or idx >= len(fds):
                    continue
                try:
                    arr = viewer._get_channel_array(p, idx, header, fds[idx])
                except Exception:
                    arr = None
                if arr is not None:
                    break
            if arr is None:
                missing += 1
                continue
            names_full.append(str(p))
            names_display.append(Path(p).stem)
            images.append(np.array(arr, dtype=float))
        except Exception:
            missing += 1
            continue
    if len(images) < 2:
        QtWidgets.QMessageBox.information(
            viewer,
            "Drift correction",
            f"Need at least two images to align.\nLoaded: {len(images)} / Selected: {len(set(paths))}\n"
            f"Skipped/missing: {missing}",
        )
        return
    # Align relative to the first frame
    ref_idx = 0
    reference = images[ref_idx]
    shifts = np.zeros((len(images), 2), dtype=float)
    ref_gray = reference.astype(np.float32)
    ref_gray = (ref_gray - ref_gray.min()) / max(ref_gray.max() - ref_gray.min(), 1e-6)
    # apply a Hann window to reduce edge effects
    try:
        win_y = np.hanning(ref_gray.shape[0])
        win_x = np.hanning(ref_gray.shape[1])
        window = np.sqrt(np.outer(win_y, win_x))
        ref_gray *= window
    except Exception:
        pass
    for i, img in enumerate(images):
        if i == ref_idx:
            continue
        target = img.astype(np.float32)
        target = (target - target.min()) / max(target.max() - target.min(), 1e-6)
        try:
            target *= window
        except Exception:
            pass
        try:
            if phase_cross_correlation is not None:
                shift, _, _ = phase_cross_correlation(ref_gray, target, upsample_factor=20, normalization="phase")
                shifts[i] = [float(shift[0]), float(shift[1])]  # dy, dx mapping target -> ref
            elif has_cv:
                import cv2  # type: ignore
                criteria = (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, 100, 1e-6)
                warp_matrix = np.eye(2, 3, dtype=np.float32)
                _, warp_matrix = cv2.findTransformECC(ref_gray, target, warp_matrix, cv2.MOTION_TRANSLATION, criteria)
                shifts[i] = [warp_matrix[1, 2], warp_matrix[0, 2]]  # dy, dx that map target -> ref
            else:
                shifts[i] = [0.0, 0.0]
        except Exception:
            shifts[i] = [0.0, 0.0]
    
    H, W = images[0].shape[:2]
    
    # Calculate the intersection crop region after alignment
    # After shifting image i by (dy, dx), its valid region is constrained
    top = int(np.ceil(max(0, np.max(shifts[:, 0]))))
    bottom = int(np.floor(min(H, H + np.min(shifts[:, 0]))))
    left = int(np.ceil(max(0, np.max(shifts[:, 1]))))
    right = int(np.floor(min(W, W + np.min(shifts[:, 1]))))
    
    # Ensure valid bounds
    top = max(0, min(top, H - 1))
    left = max(0, min(left, W - 1))
    bottom = max(top + 1, min(bottom, H))
    right = max(left + 1, min(right, W))
    
    # REMOVED: Square enforcement logic that was causing severe overcropping
    # The intersection crop is sufficient - no need to force square dimensions
    
    aligned = []
    for img, shift in zip(images, shifts):
        dy, dx = shift
        try:
            # FIXED: Apply shift directly (removed negation)
            warped = ndimage.shift(img, [dy, dx], order=3, mode="reflect", cval=0.0)
        except Exception:
            warped = img
        aligned.append(warped[top:bottom, left:right])
    
    show_alignment_preview(viewer, names_display, aligned, shifts, channel_idx, names_full, crop_bounds=(top, bottom, left, right))

def on_create_animation(viewer, paths):
    if not paths:
        return
    try:
        import imageio.v3 as iio  # type: ignore
    except Exception:
        QtWidgets.QMessageBox.warning(viewer, "Animation", "imageio is required to create GIF/MP4 animations.")
        return

    def _resize_frame(arr: np.ndarray, target_w: int, target_h: int) -> np.ndarray:
        if arr.shape[0] == target_h and arr.shape[1] == target_w:
            return arr
        # Prefer Pillow if available
        try:
            from PIL import Image
            mode = "L" if arr.ndim == 2 else "RGB"
            im = Image.fromarray(arr, mode=mode if mode else None)
            im = im.resize((target_w, target_h), Image.BILINEAR)
            return np.array(im)
        except Exception:
            pass
        # Fallback to scipy if present
        try:
            from scipy import ndimage as _ndi  # type: ignore
            zoom = (target_h / arr.shape[0], target_w / arr.shape[1]) + (() if arr.ndim == 2 else (1,))
            return _ndi.zoom(arr, zoom, order=1)
        except Exception:
            pass
        # Last resort: nearest-neighbor using numpy repeat
        y_idx = np.linspace(0, arr.shape[0] - 1, target_h).astype(int)
        x_idx = np.linspace(0, arr.shape[1] - 1, target_w).astype(int)
        if arr.ndim == 2:
            return arr[np.ix_(y_idx, x_idx)]
        else:
            return arr[np.ix_(y_idx, x_idx, np.arange(arr.shape[2]))]

    # Build a rich export dialog with preview and options
    dlg = QtWidgets.QDialog(viewer)
    dlg.setWindowTitle("Export animation")
    dlg.resize(800, 640)
    vbox = QtWidgets.QVBoxLayout(dlg); vbox.setContentsMargins(10, 10, 10, 10); vbox.setSpacing(8)

    # Gather frames
    channel_idx = viewer.channel_dropdown.currentIndex()
    frames = []
    missing = 0
    names = []
    for p in sorted({str(Path(p)) for p in paths}):
        try:
            header, fds = viewer.headers.get(p, (None, None))
            if header is None or fds is None:
                header, fds = parse_header(Path(p))
            if not fds:
                continue
            indices = [channel_idx] + [i for i in range(len(fds)) if i != channel_idx]
            arr = None
            for idx in indices:
                if idx < 0 or idx >= len(fds):
                    continue
                try:
                    arr = viewer._get_channel_array(p, idx, header, fds[idx])
                except Exception:
                    arr = None
                if arr is not None:
                    break
            if arr is None:
                missing += 1
                continue
            frames.append(np.array(arr, dtype=float))
            names.append(Path(p).name)
        except Exception:
            missing += 1
            continue
    if not frames:
        QtWidgets.QMessageBox.information(
            viewer,
            "Animation",
            f"No frames could be loaded. Selected: {len(set(paths))}, skipped: {missing}",
        )
        return

    # Controls row
    controls = QtWidgets.QHBoxLayout(); controls.setSpacing(12)
    controls.addWidget(QtWidgets.QLabel("Format:"))
    fmt_combo = QtWidgets.QComboBox(); fmt_combo.addItems(["gif", "mp4", "png-seq"]); controls.addWidget(fmt_combo)
    controls.addWidget(QtWidgets.QLabel("FPS:"))
    fps_spin = QtWidgets.QSpinBox(); fps_spin.setRange(1, 60); fps_spin.setValue(6); controls.addWidget(fps_spin)
    controls.addWidget(QtWidgets.QLabel("Duration (s):"))
    dur_spin = QtWidgets.QDoubleSpinBox(); dur_spin.setRange(0.1, 120.0); dur_spin.setDecimals(1); dur_spin.setSingleStep(0.5); controls.addWidget(dur_spin)
    dur_spin.setValue(max(0.1, len(frames) / fps_spin.value()))
    def _update_duration():
        dur_spin.setValue(max(0.1, len(frames) / max(1, fps_spin.value())))
    fps_spin.valueChanged.connect(_update_duration)
    controls.addStretch(1)
    vbox.addLayout(controls)

    # Overlay toggles
    overlay_row = QtWidgets.QHBoxLayout(); overlay_row.setSpacing(12)
    scale_cb = QtWidgets.QCheckBox("Include scale bar"); scale_cb.setChecked(True)
    markers_cb = QtWidgets.QCheckBox("Include markers/overlays"); markers_cb.setChecked(True)
    mol_cb = QtWidgets.QCheckBox("Include molecules"); mol_cb.setChecked(True)
    overlay_row.addWidget(scale_cb); overlay_row.addWidget(markers_cb); overlay_row.addWidget(mol_cb); overlay_row.addStretch(1)
    vbox.addLayout(overlay_row)

    # Resolution
    res_row = QtWidgets.QHBoxLayout(); res_row.setSpacing(12)
    res_row.addWidget(QtWidgets.QLabel("Resolution:"))
    res_combo = QtWidgets.QComboBox(); res_combo.addItems(["Auto", "720p", "1080p", "Custom"]); res_row.addWidget(res_combo)
    w_spin = QtWidgets.QSpinBox(); w_spin.setRange(256, 4096); w_spin.setValue(frames[0].shape[1]); res_row.addWidget(QtWidgets.QLabel("W")); res_row.addWidget(w_spin)
    h_spin = QtWidgets.QSpinBox(); h_spin.setRange(256, 4096); h_spin.setValue(frames[0].shape[0]); res_row.addWidget(QtWidgets.QLabel("H")); res_row.addWidget(h_spin)
    def _on_res_change(text):
        presets = {"720p": (1280, 720), "1080p": (1920, 1080)}
        if text in presets:
            w_spin.setValue(presets[text][0]); h_spin.setValue(presets[text][1])
        elif text == "Auto":
            w_spin.setValue(frames[0].shape[1]); h_spin.setValue(frames[0].shape[0])
    res_combo.currentTextChanged.connect(_on_res_change)
    vbox.addLayout(res_row)

    # Preview canvas
    prev_label = QtWidgets.QLabel(); prev_label.setAlignment(QtCore.Qt.AlignCenter)
    prev_label.setMinimumHeight(260)
    vbox.addWidget(prev_label, 1)

    # Buttons
    btn_row = QtWidgets.QHBoxLayout(); btn_row.addStretch(1)
    save_btn = QtWidgets.QPushButton("Save…"); cancel_btn = QtWidgets.QPushButton("Cancel")
    btn_row.addWidget(save_btn); btn_row.addWidget(cancel_btn)
    vbox.addLayout(btn_row)

    def _render_frame(arr):
        # simple normalization for preview
        a = np.asarray(arr, dtype=float)
        rng = a.max() - a.min()
        if rng <= 0:
            norm = np.zeros_like(a, dtype=np.uint8)
        else:
            norm = ((a - a.min()) / rng * 255.0).clip(0, 255).astype(np.uint8)
        h, w = norm.shape[:2]
        if norm.ndim == 2:
            rgb = np.stack([norm]*3, axis=-1)
        else:
            rgb = norm
        qimg = QtGui.QImage(rgb.data, w, h, 3*w, QtGui.QImage.Format_RGB888)
        pm = QtGui.QPixmap.fromImage(qimg.copy())
        return pm

    def _fit_resize_with_pad(rgb: np.ndarray, tw: int, th: int, fill=255):
        rgb = np.asarray(rgb)
        if rgb.ndim == 2:
            rgb = np.stack([rgb]*3, axis=-1)
        h, w = rgb.shape[:2]
        if h == 0 or w == 0:
            return np.zeros((th, tw, 3), dtype=np.uint8)
        scale = min(tw / w, th / h)
        new_w = max(1, int(round(w * scale)))
        new_h = max(1, int(round(h * scale)))
        resized = _resize_frame(rgb, new_w, new_h)
        canvas = np.full((th, tw, 3), fill, dtype=np.uint8)
        off_x = (tw - new_w) // 2
        off_y = (th - new_h) // 2
        canvas[off_y:off_y+new_h, off_x:off_x+new_w, :] = resized if resized.ndim == 3 else np.stack([resized]*3, axis=-1)
        return canvas

    def _update_preview(idx=0):
        pm = _render_frame(frames[idx % len(frames)])
        if not pm.isNull():
            pm = pm.scaled(prev_label.width(), prev_label.height(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
            prev_label.setPixmap(pm)
    _update_preview(0)

    # Auto-play preview timer
    timer = QtCore.QTimer(dlg); timer.setInterval(400)
    idx_ref = {"i": 0}
    def _tick():
        idx_ref["i"] = (idx_ref["i"] + 1) % len(frames)
        _update_preview(idx_ref["i"])
    timer.timeout.connect(_tick); timer.start()

    def _save():
        fmt = fmt_combo.currentText()
        default_name = f"animation.{ 'gif' if fmt=='gif' else ('mp4' if fmt=='mp4' else 'png') }"
        filter_str = "GIF (*.gif);;MP4 (*.mp4);;PNG sequence (*.png)"
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(dlg, "Save animation", default_name, filter_str)
        if not out_path:
            return
        # render each frame via the preview canvas so overlays (scale bar) are honored
        target_w, target_h = w_spin.value(), h_spin.value()
        canvas = getattr(viewer, "preview_canvas", None)
        orig_last = getattr(viewer, "last_preview", None)
        orig_views = list(getattr(canvas, "views", [])) if canvas else []
        norm_frames = []

        def _render_path(path_str: str):
            if not canvas:
                return None
            try:
                # toggle scale bar; drop molecules/markers for clean export
                prev_scale = canvas.scale_bar_enabled
                prev_ticks = canvas._show_ticks
                prev_mols = list(getattr(canvas, "molecules", []))
                canvas.scale_bar_enabled = scale_cb.isChecked()
                canvas._show_ticks = prev_ticks  # keep ticks as-is
                canvas.molecules = []  # always exclude molecules for animation per requirement
                # show file/channel and draw a fresh figure (with scale bar)
                viewer.show_file_channel(path_str, channel_idx)
                QtWidgets.QApplication.processEvents(QtCore.QEventLoop.AllEvents, 50)
                if not getattr(canvas, "views", None):
                    return None
                view = canvas.views[0]
                fig = canvas._render_view_figure(view)
                fig.set_dpi(100)
                fig.set_size_inches(target_w / 100.0, target_h / 100.0)
                fig.canvas.draw()
                buf = fig.canvas.buffer_rgba()
                if buf is None:
                    return None
                arr = np.asarray(buf)
                rgb = arr[:, :, :3].copy()
                rgb = _fit_resize_with_pad(rgb, target_w, target_h, fill=255)
                try:
                    import matplotlib.pyplot as _plt  # type: ignore
                    _plt.close(fig)
                except Exception:
                    pass
                # restore
                canvas.scale_bar_enabled = prev_scale
                canvas._show_ticks = prev_ticks
                canvas.molecules = prev_mols
                return rgb
            except Exception:
                return None

        for p in sorted({str(Path(p)) for p in paths}):
            rendered = _render_path(p)
            if rendered is not None:
                norm_frames.append(rendered)

        # fallback to raw frames if rendering failed
        if not norm_frames:
            for arr in frames:
                a = np.asarray(arr, dtype=float)
                rng = a.max() - a.min()
                if rng <= 0:
                    norm = np.zeros_like(a, dtype=np.uint8)
                else:
                    norm = ((a - a.min()) / rng * 255.0).clip(0, 255).astype(np.uint8)
                if norm.shape[0] != target_h or norm.shape[1] != target_w:
                    norm = _resize_frame(norm, target_w, target_h)
                norm_frames.append(norm)

        # restore original view
        try:
            if orig_last:
                viewer.show_file_channel(orig_last[0], orig_last[1])
            elif canvas and orig_views:
                canvas.views = orig_views
                canvas._redraw()
        except Exception:
            pass
        try:
            if fmt == "gif":
                iio.imwrite(out_path, norm_frames, plugin="pillow", loop=0, duration=1000.0 / max(1, fps_spin.value()))
            elif fmt == "mp4":
                iio.imwrite(out_path, norm_frames, plugin="ffmpeg", fps=max(1, fps_spin.value()))
            else:
                stem = Path(out_path).with_suffix("")
                for i, fr in enumerate(norm_frames):
                    iio.imwrite(f"{stem}_{i:03d}.png", fr)
            QtWidgets.QMessageBox.information(dlg, "Animation", f"Saved animation to {out_path}")
            dlg.accept()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(dlg, "Animation", f"Failed to save animation: {exc}")

    save_btn.clicked.connect(_save)
    cancel_btn.clicked.connect(dlg.reject)
    dlg.exec_()

def show_alignment_preview(viewer, names, aligned, shifts, channel_idx, source_paths=None, crop_bounds=None):
    """Preview aligned/cropped images and optionally save outputs/animation."""
    dlg = QtWidgets.QDialog(viewer)
    dlg.setWindowTitle("Drift correction preview")
    dlg.resize(900, 720)
    layout = QtWidgets.QVBoxLayout(dlg)

    info = QtWidgets.QPlainTextEdit()
    info.setReadOnly(True)
    info.setMaximumHeight(140)
    text_lines = []
    max_shift = 0.0
    for name, shift in zip(names, shifts):
        mag = float(np.hypot(shift[0], shift[1]))
        max_shift = max(max_shift, mag)
        text_lines.append(f"{name}: dy={shift[0]:.3f} px, dx={shift[1]:.3f} px | |d|={mag:.3f} px")
    if crop_bounds:
        top, bottom, left, right = crop_bounds
        crop_h = max(0, bottom - top)
        crop_w = max(0, right - left)
        text_lines.append(f"\nCrop: top={top}, bottom={bottom}, left={left}, right={right}  -> size={crop_w}x{crop_h}px")
    text_lines.append(f"Max shift magnitude: {max_shift:.3f} px")
    info.setPlainText("\n".join(text_lines))
    layout.addWidget(info)

    # Controls row: cmap + speed slider
    controls = QtWidgets.QHBoxLayout()
    controls.addWidget(QtWidgets.QLabel("Colormap:"))
    cmap_combo = QtWidgets.QComboBox()
    cmap_combo.addItems(cmap_registry.all_cmap_names())
    if hasattr(viewer, "thumb_cmap"):
        idx = cmap_combo.findText(viewer.thumb_cmap)
        if idx >= 0:
            cmap_combo.setCurrentIndex(idx)
    controls.addWidget(cmap_combo)
    controls.addSpacing(12)
    controls.addWidget(QtWidgets.QLabel("Speed (fps):"))
    speed_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
    speed_slider.setRange(1, 30)
    speed_slider.setValue(6)
    controls.addWidget(speed_slider)
    controls.addStretch(1)
    layout.addLayout(controls)

    preview_label = QtWidgets.QLabel("")
    preview_label.setAlignment(QtCore.Qt.AlignCenter)
    preview_label.setMinimumHeight(320)
    layout.addWidget(preview_label, 1)

    btn_row = QtWidgets.QHBoxLayout()
    save_imgs_btn = QtWidgets.QPushButton("Save aligned PNGs...")
    save_gif_btn = QtWidgets.QPushButton("Save animation...")
    save_virtual_btn = QtWidgets.QPushButton("Save corrected copies to thumbnails")
    btn_row.addWidget(save_imgs_btn)
    btn_row.addWidget(save_gif_btn)
    btn_row.addWidget(save_virtual_btn)
    btn_row.addStretch(1)
    layout.addLayout(btn_row)

    # Build simple QTimer-based preview using RGB frames to avoid GIF issues
    preview_timer = QtCore.QTimer(dlg)
    preview_timer.setSingleShot(False)
    frames_rgb = []

    def _build_frames(cmap_name):
        nonlocal frames_rgb
        frames_rgb = []
        try:
            import matplotlib.cm as mcm
        except Exception:
            mcm = None
        cmap_lookup = getattr(mcm, "cmap_d", None)
        cmap = None
        if mcm:
            try:
                if (cmap_lookup and cmap_name in cmap_lookup) or hasattr(mcm, "get_cmap"):
                    cmap = mcm.get_cmap(cmap_name)
            except Exception:
                cmap = None
        for arr in aligned:
            arr = np.asarray(arr, dtype=float)
            rng = arr.max() - arr.min()
            if rng <= 0:
                base = np.zeros_like(arr, dtype=float)
            else:
                base = (arr - arr.min()) / rng
            if cmap is not None:
                rgb = (cmap(base)[:, :, :3] * 255.0).astype(np.uint8)
            else:
                rgb = np.repeat((base * 255.0).astype(np.uint8)[..., None], 3, axis=2)
            frames_rgb.append(rgb)

    def _update_preview():
        if not frames_rgb:
            preview_label.setText("Preview unavailable")
            return
        idx = (preview_timer.property("frame_idx") or 0) % len(frames_rgb)
        frame = frames_rgb[idx]
        h, w, _ = frame.shape
        qimg = QtGui.QImage(frame.data, w, h, 3 * w, QtGui.QImage.Format_RGB888)
        preview_label.setPixmap(QtGui.QPixmap.fromImage(qimg))
        preview_timer.setProperty("frame_idx", (idx + 1) % len(frames_rgb))

    def _render_preview(cmap_name, fps):
        _build_frames(cmap_name)
        interval = max(30, int(1000 / max(1, fps)))
        preview_timer.setInterval(interval)
        preview_timer.setProperty("frame_idx", 0)
        preview_timer.start()
        _update_preview()

    _render_preview(cmap_combo.currentText(), speed_slider.value())
    cmap_combo.currentTextChanged.connect(lambda name: _render_preview(name, speed_slider.value()))
    speed_slider.valueChanged.connect(lambda val: _render_preview(cmap_combo.currentText(), val))

    def _save_imgs():
        out_dir = QtWidgets.QFileDialog.getExistingDirectory(viewer, "Select output folder")
        if not out_dir:
            return
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        for name, arr in zip(names, aligned):
            out_path = out_dir / f"{name}_aligned.png"
            try:
                import imageio.v3 as iio  # type: ignore
                iio.imwrite(out_path, arr.astype(np.float32))
            except Exception:
                try:
                    from matplotlib import pyplot as plt  # type: ignore
                    plt.imsave(out_path, arr, cmap=cmap_combo.currentText())
                except Exception:
                    np.savetxt(out_path.with_suffix(".txt"), arr)
        QtWidgets.QMessageBox.information(viewer, "Drift correction", f"Saved aligned images to {out_dir}")

    def _save_anim():
        try:
            import imageio.v3 as iio  # type: ignore
        except Exception:
            QtWidgets.QMessageBox.warning(dlg, "Animation", "imageio is required to save animations.")
            return
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(dlg, "Save animation", "aligned.gif", "GIF (*.gif);;MP4 (*.mp4)")
        if not out_path:
            return
        try:
            import matplotlib.cm as mcm
        except Exception:
            mcm = None
        cmap_lookup = getattr(mcm, "cmap_d", None)
        frames_out = []
        for arr in aligned:
            arr = np.asarray(arr, dtype=float)
            rng = arr.max() - arr.min()
            if rng <= 0:
                base = np.zeros_like(arr, dtype=float)
            else:
                base = (arr - arr.min()) / rng
            if mcm and ((cmap_lookup and cmap_combo.currentText() in cmap_lookup) or hasattr(mcm, "get_cmap")):
                cmap = mcm.get_cmap(cmap_combo.currentText())
                frames_out.append((cmap(base)[:, :, :3] * 255.0).astype(np.uint8))
            else:
                frames_out.append((base * 255.0).astype(np.uint8))
        suffix = Path(out_path).suffix.lower()
        fps = max(1, speed_slider.value())
        try:
            if suffix == ".mp4":
                iio.imwrite(out_path, frames_out, plugin="ffmpeg", fps=fps)
            else:
                iio.imwrite(out_path, frames_out, plugin="pillow", loop=0, duration=max(20, int(1000 / fps)))
        except Exception as exc:
            QtWidgets.QMessageBox.warning(dlg, "Animation", f"Failed to save animation: {exc}")
            return
        QtWidgets.QMessageBox.information(dlg, "Animation", f"Saved animation to {out_path}")

    def _save_virtual():
        added = 0
        existing = set(str(p) for p in viewer.files)
        top, bottom, left, right = crop_bounds if crop_bounds else (None, None, None, None)
        for idx, arr in enumerate(aligned):
            orig = str(source_paths[idx] if source_paths and idx < len(source_paths) else names[idx])
            header_fds = viewer.headers.get(orig)
            if not header_fds:
                continue
            header, fds = header_fds
            if not fds:
                continue
            fd_src = fds[channel_idx if 0 <= channel_idx < len(fds) else 0]
            header_new = dict(header)
            header_new['xPixel'] = arr.shape[1]
            header_new['yPixel'] = arr.shape[0]
            arr_by_channel = {}
            try:
                from scipy import ndimage as _ndi  # type: ignore
            except Exception:
                _ndi = None
            dy, dx = shifts[idx]
            for ch_idx, fd_ch in enumerate(fds):
                try:
                    if ch_idx == channel_idx:
                        raw_arr = arr  # already aligned/cropped for the primary channel
                    else:
                        raw_arr = viewer._get_channel_array(orig, ch_idx, header, fd_ch)
                except Exception:
                    continue
                try:
                    if ch_idx == channel_idx:
                        shifted = raw_arr
                    elif _ndi is not None:
                        shifted = _ndi.shift(raw_arr, [-dy, -dx], order=1, mode="reflect", cval=0.0)
                    else:
                        shifted = raw_arr
                    if all(v is not None for v in (top, bottom, left, right)):
                        shifted = shifted[top:bottom, left:right]
                    arr_by_channel[ch_idx] = np.array(shifted, copy=True)
                except Exception:
                    continue
            if not arr_by_channel:
                continue
            # adjust header dims to cropped size
            sample_arr = next(iter(arr_by_channel.values()))
            header_new['xPixel'] = sample_arr.shape[1]
            header_new['yPixel'] = sample_arr.shape[0]
            fds_new = [dict(fd) for fd in fds]
            caption_base = fds[channel_idx].get('Caption') or Path(orig).name if 0 <= channel_idx < len(fds) else Path(orig).name
            for i, fd_new in enumerate(fds_new):
                fd_new['FileName'] = f"{Path(orig).name}_drift_ch{i}"
                fd_new['Caption'] = f"{caption_base} [drift]"
            processed_key = f"processed_{Path(orig).stem}_drift"
            viewer._processed_views[processed_key] = {
                'arr_by_channel': arr_by_channel,
                'header': header_new,
                'fds': fds_new,
                'channel_idx': channel_idx,
                'source': orig,
            }
            viewer.headers[processed_key] = (header_new, fds_new)
            viewer._inherit_star_for_virtual_copy(processed_key, orig)
            if processed_key not in existing:
                viewer.files.append(Path(processed_key))
                existing.add(processed_key)
            added += 1
        if added:
            try:
                # insert new processed entries right after the last selected source in current ordering
                cur_files = [str(p) for p in viewer.files]
                inserted = []
                for idx, src in enumerate(source_paths or []):
                    src_str = str(src)
                    pk = f"processed_{Path(src_str).stem}_drift"
                    try:
                        pos = cur_files.index(src_str)
                    except ValueError:
                        pos = len(viewer.files)
                    if pk not in cur_files:
                        viewer.files.insert(pos + 1, Path(pk))
                        cur_files.insert(pos + 1, pk)
                        inserted.append(pk)
                if not inserted:
                    for pk in list(viewer._processed_views.keys()):
                        if pk not in cur_files:
                            viewer.files.append(Path(pk))
                viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
            except Exception:
                pass
            QtWidgets.QMessageBox.information(dlg, "Drift correction", f"Added {added} drift-corrected copy(ies) to thumbnails.\nLook for entries tagged [drift].")
        else:
            QtWidgets.QMessageBox.information(dlg, "Drift correction", "No corrected copies were created (missing headers or channels).")

    save_imgs_btn.clicked.connect(_save_imgs)
    save_gif_btn.clicked.connect(_save_anim)
    save_virtual_btn.clicked.connect(_save_virtual)
    dlg.exec_()

# ---------- Virtual copies (channels, crops, drift) ----------

