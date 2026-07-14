"""Folder report generation: payload collection + worker orchestration.

Two-phase design (see sxm_viewer/reporting/__init__.py): this controller
collects a plain-data payload on the GUI thread - the only place the live
viewer state (headers, channel caches, _map_spec_to_pixels) may be touched -
then hands it to ReportWorker, which builds the model and renders the PDF
entirely off-thread.
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

from ..._shared import QtCore, QtWidgets, log_status, np
from ...processing.detection import _find_topography_channel
from ...reporting.channels import pick_image_channels, pick_signal_channels
from ..palettes import get_color_cycle
from ..viewer import loader as viewer_loader
from ..viewer import thumbnail_ui as viewer_thumb_ui
from ..workers.report_worker import ReportWorker

# Topo arrays are downsampled before entering the payload: full resolution
# is pointless at report page sizes and 200+ scans of float64 would not be.
_TOPO_MAX_DIM_FEATURED = 512   # images that host spectra / anchor a grid
_TOPO_MAX_DIM_THUMB = 96       # contact-sheet-only images
_CONTACT_SHEET_LIMIT = 32


def _downsample(arr, max_dim):
    arr = np.asarray(arr, dtype=float)
    step = max(1, int(np.ceil(max(arr.shape) / float(max_dim))))
    return arr[::step, ::step].copy() if step > 1 else arr.copy()


def _spec_display_name(spec):
    try:
        name = Path(str(spec.get("path") or "")).name or "spectrum"
    except Exception:
        name = "spectrum"
    idx = spec.get("matrix_index")
    if idx is not None:
        name = f"{name} [{idx}]"
    return name


def _image_meta_summary(header, xpix, ypix, angle):
    """Small dict of short display strings for the report's metadata panel."""
    meta = {}
    try:
        xr = header.get("XScanRange") or header.get("XRange")
        yr = header.get("YScanRange") or header.get("YRange")
        if xr and yr:
            meta["size"] = f"{float(xr):.3g} x {float(yr):.3g} nm"
    except Exception:
        pass
    meta["pixels"] = f"{xpix} x {ypix}"
    if angle:
        meta["angle"] = f"{angle:.1f} deg"
    for label, keys in (
        ("bias", ("GapVoltage", "BiasVoltage", "Bias", "UBias", "Voltage")),
        ("setpoint", ("SetPoint", "Setpoint", "FeedbackValue", "Current")),
    ):
        for key in keys:
            val = header.get(key)
            if val not in (None, ""):
                meta[label] = str(val)
                break
    return meta


class ReportController:
    """Owns the "Generate folder report" workflow.

    Public surface: ``generate_folder_report()`` (interactive entry point,
    asks for a save path) and ``generate_report_to(out_path)`` (headless/
    programmatic). Expects the viewer to provide: files, headers,
    spectros, ensure_spectros_loaded, _header_extent, _header_scan_angle,
    _header_datetime_dt, _map_spec_to_pixels, _get_filtered_channel_array.
    """

    def __init__(self, viewer):
        self.viewer = viewer
        self._worker = None
        self._progress = None

    # ------------------------------------------------------------------
    # Entry points
    # ------------------------------------------------------------------

    def generate_folder_report(self):
        viewer = self.viewer
        if not getattr(viewer, "files", None):
            QtWidgets.QMessageBox.information(
                viewer, "Folder report", "Load a folder first.")
            return
        if self._worker is not None:
            QtWidgets.QMessageBox.information(
                viewer, "Folder report", "A report is already being generated.")
            return
        folder = getattr(viewer, "last_dir", None)
        stem = Path(folder).name if folder else "sxm"
        default = str(Path(folder or ".") / f"{stem}_report.pdf")
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            viewer, "Save folder report", default, "PDF files (*.pdf)")
        if not out_path:
            return
        self.generate_report_to(out_path)

    def generate_report_to(self, out_path):
        viewer = self.viewer
        out_path = str(out_path)
        if not out_path.lower().endswith(".pdf"):
            out_path += ".pdf"
        viewer.ensure_spectros_loaded(refresh=False)
        busy = QtWidgets.QProgressDialog(
            "Collecting folder data...", None, 0, 0, viewer)
        busy.setWindowTitle("Folder report")
        busy.setWindowModality(QtCore.Qt.WindowModal)
        busy.setMinimumDuration(0)
        busy.show()
        QtWidgets.QApplication.processEvents()
        try:
            payload = self.collect_payload()
        except Exception as exc:
            busy.close()
            QtWidgets.QMessageBox.warning(
                viewer, "Folder report", f"Could not collect folder data:\n{exc}")
            return
        busy.setLabelText("Rendering report...")
        busy.setMaximum(100)
        self._progress = busy

        worker = ReportWorker(payload, out_path)
        worker.signals.progress.connect(self._on_progress)
        worker.signals.finished.connect(self._on_finished)
        worker.signals.error.connect(self._on_error)
        self._worker = worker
        QtCore.QThreadPool.globalInstance().start(worker)

    # ------------------------------------------------------------------
    # Worker callbacks
    # ------------------------------------------------------------------

    def _on_progress(self, done, total, desc):
        if self._progress is not None:
            self._progress.setMaximum(max(1, int(total)))
            self._progress.setValue(int(done))
            self._progress.setLabelText(f"Rendering page {done}/{total}: {desc}")

    def _close_progress(self):
        if self._progress is not None:
            try:
                self._progress.close()
            except Exception:
                pass
            self._progress = None
        self._worker = None

    def _on_finished(self, out_path, pages):
        self._close_progress()
        log_status(f"[Report] Wrote {pages} page(s) to {out_path}")
        try:
            os.startfile(out_path)  # noqa: S606 - deliberate "open the PDF"
        except Exception:
            QtWidgets.QMessageBox.information(
                self.viewer, "Folder report", f"Report saved to:\n{out_path}")

    def _on_error(self, message):
        self._close_progress()
        QtWidgets.QMessageBox.warning(
            self.viewer, "Folder report", f"Report generation failed:\n{message}")

    # ------------------------------------------------------------------
    # Collect phase (GUI thread only)
    # ------------------------------------------------------------------

    def collect_payload(self):
        viewer = self.viewer
        specs_src = list(getattr(viewer, "spectros", []) or [])
        # Fill in lazily-deferred curve payloads (V/channels). Read-only
        # consumption after this point: we copy what we need into plain
        # dicts and never write derived fields back onto the spec dicts
        # (see the _merge_payload_into_spec whitelist trap in CLAUDE.md).
        try:
            viewer_loader.hydrate_spectro_entries(viewer, specs_src)
        except Exception:
            pass

        headers = getattr(viewer, "headers", {}) or {}

        # Which images deserve full-resolution topo in the payload?
        featured = set()
        for spec in specs_src:
            key = spec.get("image_key")
            if key:
                featured.add(str(key))

        specs = []
        grids_by_dataset = {}
        for spec in specs_src:
            image_key = str(spec.get("image_key") or "") or None
            # Store the marker as a raster *fraction*: the payload's topo
            # arrays are downsampled, so absolute pixel coords wouldn't
            # land on the drawn array - pdf.py rescales by the displayed
            # array's own shape.
            marker_frac = None
            if image_key and image_key in headers:
                header, _fds = headers[image_key]
                try:
                    xpix = int(header.get("xPixel", 128))
                    ypix = int(header.get("yPixel", xpix))
                    mapped = viewer._map_spec_to_pixels(
                        spec, header, xpix, ypix, file_key=image_key)
                    if mapped is not None:
                        col, row = mapped
                        marker_frac = (float(col) / max(1, xpix - 1),
                                       float(row) / max(1, ypix - 1))
                except Exception:
                    marker_frac = None
            channels = spec.get("channels") or {}
            # Sweep axis: the parser/adapter already prefers a Z axis for
            # z-spectroscopy files; additionally, if the default axis is
            # constant (e.g. fixed bias during a sweep the filename didn't
            # flag), fall back to whichever AxisChoice actually varies -
            # a flat x axis makes every curve meaningless.
            v = spec.get("V")
            x_label = "Bias (V)"
            try:
                axis_label = str(spec.get("AxisLabel") or "Bias").strip() or "Bias"
                axis_unit = str(spec.get("AxisUnit") or "V").strip() or "V"
                x_label = f"{axis_label} ({axis_unit})"
                v_arr = np.asarray(v, dtype=float) if v is not None else None
                if v_arr is not None and v_arr.size >= 2 and float(np.nanmax(v_arr) - np.nanmin(v_arr)) == 0.0:
                    for choice in (spec.get("AxisChoices") or []):
                        vals = np.asarray(choice.get("values"), dtype=float)
                        if vals.size == v_arr.size and float(np.nanmax(vals) - np.nanmin(vals)) > 0.0:
                            v = vals
                            x_label = f"{choice.get('label') or choice.get('key')} ({choice.get('unit') or ''})"
                            break
            except Exception:
                pass
            # Signal channels, most meaningful first (freq shift > current >
            # lock-in ...); flat/axis-like channels are excluded entirely.
            curves = {}
            for name, _cls in pick_signal_channels(channels):
                try:
                    curves[name] = np.asarray(channels[name], dtype=float)
                except Exception:
                    continue
            curve = None
            curve_channel = ""
            if curves:
                preferred = str(spec.get("channel_name") or "").strip()
                curve_channel = preferred if preferred in curves else next(iter(curves))
                curve = curves[curve_channel]
            elif channels:
                # Nothing classified as signal - keep the old first-channel
                # behavior rather than dropping the spec from the report.
                curve_channel = next(iter(channels))
                try:
                    curve = np.asarray(channels[curve_channel], dtype=float)
                except Exception:
                    curve = None
            entry = {
                "name": _spec_display_name(spec),
                "path": str(spec.get("path") or ""),
                "x": spec.get("x"),
                "y": spec.get("y"),
                "time": spec.get("time") if isinstance(spec.get("time"), datetime) else None,
                "image_key": image_key,
                "marker_frac": marker_frac,
                "V": np.asarray(v, dtype=float) if v is not None else None,
                "x_label": x_label,
                "curve": curve,
                "curve_channel": curve_channel,
                "curves": curves,
                "off_frame_direction": spec.get("off_frame_direction"),
                "off_frame_distance_nm": spec.get("off_frame_distance_nm"),
                "assignment_reason_label": str(spec.get("assignment_reason_label") or ""),
                "assignment_confidence": str(spec.get("assignment_confidence") or ""),
                "is_matrix": spec.get("matrix_index") is not None,
                "matrix_dataset": spec.get("matrix_dataset"),
                "grid_row": spec.get("grid_row"),
                "grid_col": spec.get("grid_col"),
                "grid_rows": spec.get("grid_rows"),
                "grid_cols": spec.get("grid_cols"),
                "matrix_index": spec.get("matrix_index"),
            }
            idx = len(specs)
            specs.append(entry)
            if entry["is_matrix"] and entry["matrix_dataset"]:
                grids_by_dataset.setdefault(str(entry["matrix_dataset"]), {
                    "dataset": str(entry["matrix_dataset"]),
                    "name": entry["name"].split(" [")[0],
                    "image_key": image_key,
                    "spec_indices": [],
                })["spec_indices"].append(idx)

        images = []
        files = [str(p) for p in (getattr(viewer, "files", []) or [])]
        for order, key in enumerate(files):
            if key not in headers:
                continue
            header, fds = headers[key]
            try:
                xpix = int(header.get("xPixel", 128))
                ypix = int(header.get("yPixel", xpix))
            except Exception:
                xpix = ypix = 128
            angle = float(viewer._header_scan_angle(header) or 0.0)
            try:
                extent = viewer._header_extent(header)
            except Exception:
                extent = None
            try:
                time = viewer._header_datetime_dt(header, Path(key))
            except Exception:
                time = None
            is_featured = key in featured
            # Which channels represent this image best: constant-height scans
            # get current + frequency-shift/lock-in panels, everything else
            # its topography channel (user heuristic - a CH image's "topo"
            # is just the constant plane, the signal lives elsewhere).
            try:
                tag = viewer_thumb_ui.effective_tag(viewer, key)
            except Exception:
                tag = None
            panels = []
            if is_featured or order < _CONTACT_SHEET_LIMIT:
                topo_idx = _find_topography_channel(fds)
                picks = pick_image_channels(fds, tag, topo_idx)
                if not is_featured:
                    picks = picks[:1]  # contact-sheet thumbs only need one
                for chan_idx, cls in picks:
                    fd = fds[chan_idx]
                    try:
                        unit, arr = viewer._get_filtered_channel_array(
                            key, chan_idx, header, fd)
                        panels.append({
                            "label": str(fd.get("Caption") or fd.get("FileName") or f"ch{chan_idx}"),
                            "unit": str(unit or ""),
                            "cls": cls,
                            "data": _downsample(
                                arr,
                                _TOPO_MAX_DIM_FEATURED if is_featured else _TOPO_MAX_DIM_THUMB),
                        })
                    except Exception:
                        continue
            # An untagged constant-height scan sneaks past the tag check with
            # a dead-flat Z channel (all the signal lives in current/df) - a
            # solid black panel. Detect that and re-pick as if CH-tagged.
            if panels and tag != "constant-height":
                first_arr = panels[0]["data"]
                finite = first_arr[np.isfinite(first_arr)]
                if finite.size and float(finite.max() - finite.min()) <= 1e-12 * max(1.0, abs(float(finite.max()))):
                    ch_picks = pick_image_channels(fds, "constant-height", None)
                    retry = []
                    for chan_idx, cls in ch_picks[:2 if is_featured else 1]:
                        fd = fds[chan_idx]
                        try:
                            unit, arr = viewer._get_filtered_channel_array(
                                key, chan_idx, header, fd)
                        except Exception:
                            continue
                        arr = _downsample(
                            arr,
                            _TOPO_MAX_DIM_FEATURED if is_featured else _TOPO_MAX_DIM_THUMB)
                        finite = arr[np.isfinite(arr)]
                        if not finite.size or float(finite.max() - finite.min()) <= 0.0:
                            continue
                        retry.append({
                            "label": str(fd.get("Caption") or fd.get("FileName") or f"ch{chan_idx}"),
                            "unit": str(unit or ""),
                            "cls": cls,
                            "data": arr,
                        })
                    if retry:
                        panels = retry
            images.append({
                "key": key,
                "name": Path(key).stem,
                "time": time,
                "extent": list(extent) if extent else None,
                "angle": angle,
                "xpix": xpix,
                "ypix": ypix,
                "tag": tag,
                "panels": panels,
                # First panel doubles as the generic "picture of this image"
                # (contact sheet, overview, grid anchor pane).
                "topo": panels[0]["data"] if panels else None,
                "topo_unit": panels[0]["unit"] if panels else "",
                "meta": _image_meta_summary(header, xpix, ypix, angle),
            })

        folder = getattr(viewer, "last_dir", None)
        # User display preferences the report should honor: the persisted
        # color cycle (resolved to concrete colors so the Qt-free renderer
        # never touches palettes.py) and any starred favourite colormaps.
        get_fav = getattr(viewer, "get_favorite_cmap", None)
        prefs = {
            "color_cycle": list(get_color_cycle(
                getattr(viewer, "spectro_color_cycle", None))),
            "grid_metric_cmap": get_fav("grid_metric", None) if callable(get_fav) else None,
            "preview_cmap": get_fav("preview", None) if callable(get_fav) else None,
        }
        return {
            "folder": str(folder or ""),
            "images": images,
            "specs": specs,
            "grids": list(grids_by_dataset.values()),
            "prefs": prefs,
        }
