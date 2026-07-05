"""Review dialog for FFT-based periodic-noise removal.

Deliberately not a slider-based SingleFilterDialog: which frequency regions
to remove is a judgment call that needs a look at the actual spectrum, not
a one-size-fits-all default - a genuine atomic lattice or moire pattern
produces sharp FFT peaks too, and this module's detection can only offer a
plain-language tip (does a region match a known pump/mains frequency, does
it look like scan-line banding, or does it look like real periodic
structure?), not a safe default to auto-apply. Detected regions are only
suggestions - the user can also draw an arbitrary rectangular or
elliptical region directly on the spectrum, in either "remove" (attenuate)
or "protect" (force back to unattenuated, even where an overlapping
remove-region would otherwise touch it) mode. Protect exists specifically
for the case where genuine low-frequency topography and noise close to
zero frequency can't be separated - a small protected zone around DC lets
a user draw one broad, sloppy removal region without needing its boundary
to be surgically precise near the center. Nothing here runs unless the
user explicitly checks a suggested region, draws a custom one, and clicks
OK.
"""
from __future__ import annotations

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.patches import Ellipse, Rectangle
from matplotlib.widgets import RectangleSelector

from ..._shared import QtCore, QtWidgets, np
from ...processing import periodic_noise as pn
from ...processing.filters import FILTER_DEFINITIONS
from ...processing.scan_timing import estimate_scan_timing

_CATEGORY_COLORS = {
    "noise": "#d9534f",     # red - matches a known pump/mains frequency, or looks like banding
    "caution": "#f0ad4e",   # amber - looks lattice-like, review carefully
    "unclear": "#777777",   # gray - no strong evidence either way
}
_MODE_COLORS = {
    "remove": "#5bc0de",   # cyan - user-drawn removal region
    "protect": "#5cb85c",  # green - user-drawn protected region
}


class PeriodicNoiseDialog(QtWidgets.QDialog):
    """Shows the FFT magnitude spectrum with candidate periodic-noise
    regions marked and tagged, lets the user check which suggestions to
    remove and/or drag out arbitrary custom regions (rectangle or ellipse,
    remove or protect) directly on the spectrum, and shows a live
    before/after preview of the real-space result inside the dialog
    itself - not relying solely on the caller's preview_callback, which
    updates the main canvas that may well be hidden behind this modal
    dialog and give the impression nothing is happening."""

    def __init__(self, parent, base_image, header=None, preview_callback=None,
                 preview_target_text="current image", initial_taper=None):
        super().__init__(parent)
        self.setWindowTitle("Remove periodic noise")
        self.resize(1060, 660)
        self._preview_callback = preview_callback
        self._base_image = np.asarray(base_image, dtype=float)
        default_taper = FILTER_DEFINITIONS.get("periodic_noise", {}).get("default_taper", 0.01)
        self._taper = float(initial_taper or default_taper)
        self._custom_regions = []  # list of (shape, mode, fx_min, fx_max, fy_min, fy_max)

        timing = estimate_scan_timing(header or {})
        expected = None
        if timing is not None:
            expected = pn.expected_peak_frequencies(timing["line_time_s"], timing["pixel_dwell_s"])
        self._timing = timing

        self._freqs_x, self._freqs_y, self._magnitude = pn.fft_magnitude(self._base_image)
        self._regions = pn.detect_candidate_regions(self._magnitude, self._freqs_x, self._freqs_y, expected=expected)

        layout = QtWidgets.QVBoxLayout(self)

        target_label = QtWidgets.QLabel(f"Spectrum of: {preview_target_text}")
        target_label.setStyleSheet("color: gray;")
        layout.addWidget(target_label)

        if timing is None:
            note_text = (
                "Scan timing not available for this file - regions can't be matched against "
                "known pump/mains frequencies. Review each region's shape and position before removing."
            )
        else:
            note_text = (
                f"Scan timing from {timing['source']} header - matching against 50/60 Hz and "
                "500/1000/1500 Hz pump harmonics."
            )
        note = QtWidgets.QLabel(note_text)
        note.setWordWrap(True)
        note.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(note)

        body = QtWidgets.QHBoxLayout()
        layout.addLayout(body, 1)

        self.figure = Figure(figsize=(9, 5))
        self.canvas = FigureCanvas(self.figure)
        self.spectrum_ax = self.figure.add_subplot(1, 3, (1, 2))
        self.preview_ax = self.figure.add_subplot(1, 3, 3)
        self.canvas.mpl_connect("button_press_event", self._on_plot_click)
        body.addWidget(self.canvas, 3)
        self._selector = RectangleSelector(
            self.spectrum_ax, self._on_region_drawn, useblit=True,
            button=[1], minspanx=0.002, minspany=0.002, spancoords="data",
        )

        right = QtWidgets.QVBoxLayout()
        body.addLayout(right, 1)

        right.addWidget(QtWidgets.QLabel(
            "Detected candidate regions (click one to toggle, or drag on the\n"
            "spectrum to mask any other frequency region yourself):"
        ))

        draw_row = QtWidgets.QFormLayout()
        self.draw_shape_combo = QtWidgets.QComboBox()
        self.draw_shape_combo.addItems(["Rectangle", "Ellipse"])
        draw_row.addRow("Draw shape", self.draw_shape_combo)
        self.draw_mode_combo = QtWidgets.QComboBox()
        self.draw_mode_combo.addItem("Remove (attenuate)", "remove")
        self.draw_mode_combo.addItem("Protect (keep, even if a remove region overlaps)", "protect")
        draw_row.addRow("Draw mode", self.draw_mode_combo)
        right.addLayout(draw_row)

        self.region_list = QtWidgets.QListWidget()
        self.region_list.itemChanged.connect(self._on_selection_changed)
        self.region_list.currentRowChanged.connect(self._on_row_focused)
        right.addWidget(self.region_list, 1)

        remove_custom_btn = QtWidgets.QPushButton("Remove selected custom region")
        remove_custom_btn.clicked.connect(self._remove_selected_custom_region)
        right.addWidget(remove_custom_btn)

        self.tip_label = QtWidgets.QLabel("Select a region to see details.")
        self.tip_label.setWordWrap(True)
        right.addWidget(self.tip_label)

        taper_row = QtWidgets.QFormLayout()
        self.taper_spin = QtWidgets.QDoubleSpinBox()
        self.taper_spin.setDecimals(4)
        self.taper_spin.setRange(0.001, 0.05)
        self.taper_spin.setSingleStep(0.002)
        self.taper_spin.setValue(self._taper)
        self.taper_spin.valueChanged.connect(self._on_taper_changed)
        taper_row.addRow("Edge softness (cycles/pixel)", self.taper_spin)
        right.addLayout(taper_row)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._rebuild_region_list()
        self._render()

    def _rebuild_region_list(self):
        # Called every time a custom region is drawn/removed, not just at
        # init - preserve whichever suggestions were already checked, or
        # drawing a second custom region would silently un-check the first
        # selection the user made.
        previously_checked = {
            self.region_list.item(row).data(QtCore.Qt.UserRole)
            for row in range(self.region_list.count())
            if self.region_list.item(row).checkState() == QtCore.Qt.Checked
        }
        self.region_list.blockSignals(True)
        self.region_list.clear()
        if not self._regions and not self._custom_regions:
            self.region_list.addItem("No strong periodic regions detected - drag on the spectrum to add one.")
            self.region_list.setEnabled(False)
        else:
            self.region_list.setEnabled(True)
            for i, region in enumerate(self._regions):
                shape_label = "band" if region["is_streak"] else "spot"
                item = QtWidgets.QListWidgetItem(
                    f"[{shape_label}] ({region['fx']:.4f}, {region['fy']:.4f}) - {region['tip_category']}"
                )
                item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
                checked = ("suggested", i) in previously_checked
                item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
                item.setData(QtCore.Qt.UserRole, ("suggested", i))
                self.region_list.addItem(item)
            for i, bounds in enumerate(self._custom_regions):
                shape, mode = bounds[0], bounds[1]
                item = QtWidgets.QListWidgetItem(f"[custom {i + 1}: {shape}, {mode}]")
                item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
                item.setCheckState(QtCore.Qt.Checked)
                item.setData(QtCore.Qt.UserRole, ("custom", i))
                self.region_list.addItem(item)
        self.region_list.blockSignals(False)

    def _render(self):
        self._render_spectrum()
        self._render_preview()
        self.canvas.draw_idle()

    def _add_patch(self, ax, shape, fx_min, fx_max, fy_min, fy_max, color, selected):
        if shape == "ellipse":
            cx, cy = (fx_min + fx_max) / 2.0, (fy_min + fy_max) / 2.0
            patch = Ellipse(
                (cx, cy), fx_max - fx_min or 1e-4, fy_max - fy_min or 1e-4,
                fill=selected, facecolor=color, edgecolor=color,
                alpha=0.55 if selected else 0.9, linewidth=2 if selected else 1.2,
            )
        else:
            patch = Rectangle(
                (fx_min, fy_min), fx_max - fx_min or 1e-4, fy_max - fy_min or 1e-4,
                fill=selected, facecolor=color, edgecolor=color,
                alpha=0.55 if selected else 0.9, linewidth=2 if selected else 1.2,
            )
        ax.add_patch(patch)

    def _render_spectrum(self):
        ax = self.spectrum_ax
        ax.clear()
        log_mag = np.log1p(self._magnitude)
        extent = [self._freqs_x[0], self._freqs_x[-1], self._freqs_y[0], self._freqs_y[-1]]
        ax.imshow(log_mag, origin="lower", extent=extent, cmap="magma", aspect="auto")
        for i, region in enumerate(self._regions):
            color = _CATEGORY_COLORS.get(region["tip_category"], "#ffffff")
            selected = self._is_suggested_checked(i)
            self._add_patch(ax, "rect", region["fx_min"], region["fx_max"], region["fy_min"], region["fy_max"], color, selected)
        for shape, mode, fx_min, fx_max, fy_min, fy_max in self._custom_regions:
            self._add_patch(ax, shape, fx_min, fx_max, fy_min, fy_max, _MODE_COLORS.get(mode, "#ffffff"), True)
        ax.set_xlabel("fx (cycles/pixel)")
        ax.set_ylabel("fy (cycles/pixel)")
        ax.set_title("Click a region to toggle it, or drag to draw your own")

    def _render_preview(self):
        ax = self.preview_ax
        ax.clear()
        active = self._active_regions()
        if active:
            result = pn.apply_mask_filter(self._base_image, active, taper=self._taper)
            n_remove = sum(1 for r in active if r[1] == "remove")
            n_protect = sum(1 for r in active if r[1] == "protect")
            title = f"Preview ({n_remove} removed, {n_protect} protected)"
        else:
            result = self._base_image
            title = "Preview (no regions selected)"
        vmin, vmax = np.nanpercentile(self._base_image, [1, 99])
        ax.imshow(result, origin="lower", cmap="gray", vmin=vmin, vmax=vmax, aspect="auto")
        ax.set_title(title, fontsize=9)
        ax.set_xticks([])
        ax.set_yticks([])

    def _is_suggested_checked(self, index):
        for row in range(self.region_list.count()):
            item = self.region_list.item(row)
            data = item.data(QtCore.Qt.UserRole)
            if data == ("suggested", index):
                return item.checkState() == QtCore.Qt.Checked
        return False

    def _on_row_focused(self, row):
        if row < 0:
            return
        item = self.region_list.item(row)
        if item is None:
            return
        data = item.data(QtCore.Qt.UserRole)
        if data is None:
            return
        kind, index = data
        if kind == "suggested" and 0 <= index < len(self._regions):
            self.tip_label.setText(self._regions[index]["tip"])
        elif kind == "custom":
            _shape, mode = self._custom_regions[index][0], self._custom_regions[index][1]
            if mode == "protect":
                self.tip_label.setText(
                    "Custom protected region - forces this frequency area back to unattenuated, "
                    "even where a remove region overlaps it. Always active until removed."
                )
            else:
                self.tip_label.setText("Custom removal region - always active until removed.")

    def _on_plot_click(self, event):
        if event.inaxes is not self.spectrum_ax or not self._regions:
            return
        if event.xdata is None or event.ydata is None:
            return
        for i, region in enumerate(self._regions):
            if region["fx_min"] <= event.xdata <= region["fx_max"] and region["fy_min"] <= event.ydata <= region["fy_max"]:
                for row in range(self.region_list.count()):
                    item = self.region_list.item(row)
                    if item.data(QtCore.Qt.UserRole) == ("suggested", i):
                        item.setCheckState(
                            QtCore.Qt.Unchecked if item.checkState() == QtCore.Qt.Checked else QtCore.Qt.Checked
                        )
                        self.region_list.setCurrentRow(row)
                        break
                return

    def _on_region_drawn(self, eclick, erelease):
        if eclick.xdata is None or erelease.xdata is None:
            return
        fx_min, fx_max = sorted((eclick.xdata, erelease.xdata))
        fy_min, fy_max = sorted((eclick.ydata, erelease.ydata))
        if (fx_max - fx_min) < 1e-4 or (fy_max - fy_min) < 1e-4:
            return
        shape = "ellipse" if self.draw_shape_combo.currentText() == "Ellipse" else "rect"
        mode = self.draw_mode_combo.currentData()
        self._custom_regions.append((shape, mode, fx_min, fx_max, fy_min, fy_max))
        self._rebuild_region_list()
        self._update_preview()

    def _remove_selected_custom_region(self):
        item = self.region_list.currentItem()
        if item is None:
            return
        data = item.data(QtCore.Qt.UserRole)
        if data is None or data[0] != "custom":
            return
        del self._custom_regions[data[1]]
        self._rebuild_region_list()
        self._update_preview()

    def _active_regions(self):
        """Returns a list of (shape, mode, fx_min, fx_max, fy_min, fy_max)."""
        active = []
        for row in range(self.region_list.count()):
            item = self.region_list.item(row)
            if item.checkState() != QtCore.Qt.Checked:
                continue
            data = item.data(QtCore.Qt.UserRole)
            if data is None:
                continue
            kind, index = data
            if kind == "suggested":
                r = self._regions[index]
                active.append(("rect", "remove", r["fx_min"], r["fx_max"], r["fy_min"], r["fy_max"]))
            elif kind == "custom":
                active.append(self._custom_regions[index])
        return active

    def _current_step(self):
        active = self._active_regions()
        if not active:
            return None
        return {"key": "periodic_noise", "params": {"regions": tuple(active), "taper": float(self._taper)}}

    def _update_preview(self):
        self._render()
        if self._preview_callback is None:
            return
        step = self._current_step()
        steps = [step] if step else None
        self._preview_callback(steps, label="Remove periodic noise" if step else None)

    def _on_selection_changed(self, _item):
        self._update_preview()

    def _on_taper_changed(self, value):
        self._taper = float(value)
        self._update_preview()

    def accepted_step(self):
        """Returns the final {"key": "periodic_noise", ...} step, or None if
        nothing was selected (caller should treat that as "no change")."""
        return self._current_step()


__all__ = ["PeriodicNoiseDialog"]
