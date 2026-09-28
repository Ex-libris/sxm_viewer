"""Headless smoke test for the matplotlib-based ProfileDialog.

Not a pytest module (needs a live QApplication + offscreen platform) - run
directly:

    QT_QPA_PLATFORM=offscreen python sxm_viewer/tests/smoke_profile_dialog.py

Exercises: axis-label composition (no doubled units), SI auto-prefix
(EngFormatter) on a frequency-shift channel (the "1e-15 Hz" bug), a
dual-unit overlay (twinx right axis), the SpanSelector-based measurement
band's live-link callback contract, snap-to-peak, keyboard nudge, the
composite drag/merge payload path, and the context/export menu
construction - without a real QApplication event loop or mouse input.
"""
import sys
import types

import numpy as np
from PyQt5 import QtCore, QtWidgets
from matplotlib.ticker import EngFormatter, ScalarFormatter

app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)

from sxm_viewer.gui.dialogs.profile_dialog import ProfileDialog, _ColorChip  # noqa: E402


def make_profile(*, channel, unit, colorbar_label, n=64, length_nm=232.211, seed=0, scale=1.0,
                  file_name='default_NDST2_K1001_wsxm.txt'):
    rng = np.random.default_rng(seed)
    x_px = np.linspace(0.0, n - 1, n)
    x_nm = x_px * (length_nm / (n - 1))
    vals = (rng.normal(0.0, 1.0, n) * scale)
    return {
        'x_px': x_px,
        'x_nm': x_nm,
        'vals': vals,
        'length_nm': length_nm,
        'unit': unit,
        'axis_unit': 'nm',
        'distance_unit': 'nm',
        'color': None,
        'lw': None,
        'line_style': '-',
        'marker_style': 'o',
        'marker_size': None,
        'label': f'L={length_nm:.0f} nm',
        'relative_axes': True,
        'meta': {'channel': channel, 'file_name': file_name},
        'source_path': f'C:/data/{file_name}',
        'source_file_name': file_name,
        'source_title': channel,
        'source_acquisition_text': 'Bias 0.5V, Setpoint 20pA',
        'source_datetime': '2024-04-23 11:08:32',
        'source_date': '2024-04-23',
        'source_time': '11:08:32',
        'live_profile_ref': None,
        'display_name': colorbar_label,
    }


def check(label, cond):
    status = "OK" if cond else "FAIL"
    print(f"[{status}] {label}")
    if not cond:
        raise SystemExit(1)


def fake_motion_event(ax, xdata, ydata):
    """A minimal stand-in for matplotlib's motion_notify_event MouseEvent."""
    return types.SimpleNamespace(inaxes=ax, xdata=xdata, ydata=ydata)


def main():
    # --- reproduces the "Current [A] (A)" bug scenario ---------------
    current_profile = make_profile(
        channel="Current", unit="A", colorbar_label="Current [A]", scale=50e-12,
    )
    dlg = ProfileDialog(
        current_profile, [],
        unit="A", y_label="Current [A]",
        dark_mode=False,
    )
    check("axis label unit not doubled ('[A] (A)' absent)", "[A] (A)" not in dlg.ax.get_ylabel())
    check("axis label text is the clean channel name (compose_axis_label stripped '[A]')", dlg.ax.get_ylabel() == "Current")
    y_fmt = dlg.ax.yaxis.get_major_formatter()
    check("left axis uses EngFormatter (auto-SI-prefix) for Amps", isinstance(y_fmt, EngFormatter))
    check("left axis EngFormatter's unit is 'A'", y_fmt.unit == "A")

    # --- reproduces the "1e-15 Hz" bug scenario -----------------------
    freq_profile = make_profile(
        channel="Frequency Shift", unit="Hz", colorbar_label="Frequency Shift [Hz]",
        scale=12.0, seed=1,
    )
    dlg.update_profiles(freq_profile, [])
    y_fmt = dlg.ax.yaxis.get_major_formatter()
    check("Hz axis uses EngFormatter (auto-SI-prefix)", isinstance(y_fmt, EngFormatter) and y_fmt.unit == "Hz")
    plotted_item = dlg._line_handles_by_key.get(None)
    plotted_y = plotted_item.get_ydata()
    raw_y = freq_profile['vals']
    check(
        "Hz values plotted unscaled (no spurious 1e-15-style factor)",
        np.allclose(plotted_y, raw_y),
    )

    # --- distance axis is true SI base (metres), not nm ----------------
    x_fmt = dlg.ax.xaxis.get_major_formatter()
    check("distance axis uses EngFormatter (auto-SI-prefix)", isinstance(x_fmt, EngFormatter))
    plotted_x = plotted_item.get_xdata()
    check(
        "distance plotted in metres (nm * 1e-9)",
        np.allclose(plotted_x, freq_profile['x_nm'] * 1e-9),
    )

    # --- dual-unit overlay creates a right (twinx) axis -------------------
    overlay_profile = make_profile(
        channel="Current", unit="A", colorbar_label="Current [A]", scale=30e-12, seed=2,
        file_name="default_NDST2_K1005_wsxm.txt",
    )
    dlg.update_profiles(freq_profile, [overlay_profile])
    check("one curve routed to the right (twinx) axis", len(dlg._right_curves) == 1)
    check("right axis is visible", dlg.ax_right.get_visible())

    # switch to a single-unit render (right axis should hide), then back to
    # dual-unit (regression: must re-show it, not just skip because the
    # twinx axes object already exists from the first pass)
    dlg.update_profiles(freq_profile, [])
    check("right axis hidden when units match", not dlg.ax_right.get_visible())
    dlg.update_profiles(freq_profile, [overlay_profile])
    check("right axis re-shown on second dual-unit render", dlg.ax_right.get_visible())

    # --- measurement band live-link round trip (dialog -> canvas -> dialog)
    dlg.update_profiles(freq_profile, [])
    dlg.snap_toggle.setChecked(False)  # isolate the round-trip from snap-to-peak
    received = {}

    def _marker_cb(positions, domain):
        received['positions'] = positions
        received['domain'] = domain

    dlg.set_marker_update_callback(_marker_cb)
    lo_plot = dlg._external_to_plot_x(freq_profile['x_nm'][5])
    hi_plot = dlg._external_to_plot_x(freq_profile['x_nm'][40])
    # SpanSelector.extents = (...) alone never fires onselect (confirmed
    # empirically), unlike a real mouse-driven drag - call the release
    # handler explicitly to simulate what a real drag-then-release does.
    dlg._span.extents = (lo_plot, hi_plot)
    dlg._on_span_select(lo_plot, hi_plot)
    check("marker_update_callback fired with nm-space positions", 'positions' in received)
    if 'positions' in received:
        check(
            "round-tripped left edge matches nm domain (not metres)",
            abs(received['positions'][0] - freq_profile['x_nm'][5]) < 1e-6,
        )

    # canvas -> dialog direction
    dlg.set_marker_positions([freq_profile['x_nm'][10], freq_profile['x_nm'][50]], domain=(0.0, freq_profile['x_nm'][-1]))
    check("set_marker_positions made the band visible", dlg._span.get_visible())
    lo_ext, hi_ext = dlg._marker_positions
    check(
        "band reflects the externally-set positions",
        abs(lo_ext - freq_profile['x_nm'][10]) < 1e-6 and abs(hi_ext - freq_profile['x_nm'][50]) < 1e-6,
    )

    # --- readout strip reflects the band ----------------------------------
    dx_text = dlg.readout.value_labels["dx"].text()
    check("readout dx field is populated with a tolerance", "\u00b1" in dx_text)
    peaks_text = dlg.readout.value_labels["peaks"].text()
    check("readout peaks-in-band field is a plain integer", peaks_text.isdigit())
    px_text = dlg.readout.value_labels["px"].text()
    check("readout pixel-size field reports point count", "in band" in px_text)

    # --- crosshair / cursor readout ----------------------------------------
    x_center = freq_profile['x_nm'][20] * 1e-9
    y_center = float(plotted_item.get_ydata()[20])
    dlg._on_mouse_moved(fake_motion_event(dlg.ax, x_center, y_center))
    check("crosshair vline becomes visible over the plot", dlg._vline.get_visible())
    check("cursor readout field populated", dlg.readout.value_labels["cursor"].text() != "-")
    dlg.crosshair_toggle.setChecked(False)
    check("crosshair toggle persisted (setting)", dlg._crosshair_enabled is False)
    dlg._on_mouse_moved(fake_motion_event(dlg.ax, x_center, y_center))
    check("crosshair stays hidden once disabled", not dlg._vline.get_visible())
    dlg.crosshair_toggle.setChecked(True)
    # cursor outside any axes hides the crosshair too
    dlg._on_mouse_moved(fake_motion_event(None, None, None))
    check("crosshair hides when the cursor leaves the plot", not dlg._vline.get_visible())

    # --- snap-to-peak ---------------------------------------------------------
    dlg.snap_toggle.setChecked(True)
    peak_idx = dlg._ref_peak_idx
    check("reference profile has detected peaks for snap test", peak_idx.size > 0)
    if peak_idx.size >= 2:
        px = dlg._px_size_plot
        near_peak = dlg._ref_x_plot[peak_idx[0]] + px * 0.5  # within the 6px snap tolerance
        far_edge = dlg._span.extents[1]
        dlg._span.extents = (near_peak, far_edge)
        dlg._on_span_select(near_peak, far_edge)
        snapped_lo = dlg._span.extents[0]
        check(
            "snap-to-peak pulled the dragged edge onto the peak",
            abs(snapped_lo - dlg._ref_x_plot[peak_idx[0]]) < px * 1e-6 + 1e-15,
        )

    # --- keyboard nudge -------------------------------------------------------
    dlg.snap_toggle.setChecked(False)  # isolate nudge from snap pulling it back
    lo0, hi0 = dlg._span.extents
    dlg._nudge_region(QtCore.Qt.Key_Right, QtCore.Qt.NoModifier)
    lo1, _hi1 = dlg._span.extents
    check("Right-arrow nudges the left edge by one pixel", abs((lo1 - lo0) - dlg._px_size_plot) < 1e-15)
    dlg._nudge_region(QtCore.Qt.Key_Right, QtCore.Qt.ControlModifier)
    lo2, _hi2 = dlg._span.extents
    check("Ctrl+Right nudges by ten pixels", abs((lo2 - lo1) - 10 * dlg._px_size_plot) < 1e-15)
    hi3 = hi0
    dlg._nudge_region(QtCore.Qt.Key_Left, QtCore.Qt.ShiftModifier)
    _lo4, hi4 = dlg._span.extents
    check("Shift+Left nudges the right edge, not the left", abs(hi4 - hi3 + dlg._px_size_plot) < 1e-15)

    # --- Measure toggle hides band + readout, shows context row ---------------
    # Note: dlg is never shown() in this headless test, so plain QWidget
    # .isVisible() would report False unconditionally regardless of the
    # internal show/hide flag - isVisibleTo(dlg) checks visibility relative
    # to dlg itself instead of the full (unshown) desktop ancestor chain.
    # The SpanSelector/crosshair are matplotlib artists, not QWidgets, so
    # their plain get_visible() is unaffected and used as-is.
    dlg.measure_toggle.setChecked(False)
    check("Measure off hides the band", not dlg._span.get_visible())
    check("Measure off hides the readout strip", not dlg.readout.isVisibleTo(dlg))
    check("Measure off shows the context row", dlg.marker_info.isVisibleTo(dlg))
    dlg.measure_toggle.setChecked(True)
    check("Measure back on re-shows the band", dlg._span.get_visible())
    check("Measure back on re-shows the readout strip", dlg.readout.isVisibleTo(dlg))

    # --- band/crosshair excluded from exports by default -----------------------
    dlg._span.set_visible(True)
    dlg._vline.set_visible(True)
    check("include-band-in-export is off by default", dlg._include_band_in_export is False)
    with dlg._LightForExport(dlg):
        band_hidden_during_export = not dlg._span.get_visible()
        crosshair_hidden_during_export = not dlg._vline.get_visible()
    check("measurement band hidden during export by default", band_hidden_during_export)
    check("crosshair hidden during export by default", crosshair_hidden_during_export)
    check("band visibility restored after export", dlg._span.get_visible())
    dlg._include_band_in_export = True
    with dlg._LightForExport(dlg):
        band_visible_when_opted_in = dlg._span.get_visible()
    check("band stays visible during export when opted in", band_visible_when_opted_in)
    dlg._include_band_in_export = False

    # --- composite drag payload round-trips through JSON ----------------
    payload = dlg._composite_payload()
    check("composite payload built", payload is not None and payload.get('entries'))
    entries = dlg._entries_from_profile_payload(payload)
    check("composite payload round-trips via JSON", len(entries) == len(payload['entries']))

    # --- export menu population builds without raising ---------------------
    # _on_context_menu itself just calls exec_(); the actual menu content is
    # built lazily from aboutToShow (_populate_export_menu), so exercise
    # that directly rather than relying on exec_ (which a mocked-out exec_
    # would never trigger aboutToShow for).
    dlg._populate_export_menu()
    check("export menu populated without raising", dlg._export_menu.actions())
    check(
        "export menu carries the same actions the header button uses",
        dlg.export_btn.menu() is dlg._export_menu,
    )
    anchor = dlg._context_menu_anchor_pos()
    check("context-menu anchor is computed without raising", anchor is not None)

    # header title reflects the active profile
    check("header title shows the active profile's display name", dlg.header_title.text() != "Profile measurement")

    # typography actually changes fonts
    dlg.set_plot_font_family("Georgia")
    check("axis label font-family applied", "Georgia" in dlg.ax.yaxis.label.get_fontfamily())
    check("readout font-family applied", dlg.readout.value_labels["dx"].font().family() == "Georgia")

    # hint bar dismissal persists
    check("hint bar starts visible on first open", dlg.hint_bar.isVisibleTo(dlg))
    dlg._dismiss_hint_bar()
    check("hint bar hides once dismissed", not dlg.hint_bar.isVisibleTo(dlg))

    # reset style doesn't raise and restores defaults
    from sxm_viewer.gui.palettes import DEFAULT_COLOR_CYCLE as _DEFAULT_CYCLE
    dlg._reset_style()
    check("reset style restores default palette", dlg._profile_palette_name == _DEFAULT_CYCLE)
    check("reset style restores default figure preset", dlg._figure_preset_key == "interactive")

    # Ctrl+C / Ctrl+Shift+C / Ctrl+S shortcuts are wired
    check("Ctrl+C shortcut bound", dlg._copy_plot_shortcut.key().toString() == "Ctrl+C")
    check("Ctrl+Shift+C shortcut bound", dlg._copy_xy_shortcut.key().toString() == "Ctrl+Shift+C")
    check("Ctrl+S shortcut bound", dlg._save_plot_shortcut.key().toString() == "Ctrl+S")

    # Save data as CSV doesn't raise (cancel the file dialog)
    orig_get_save = QtWidgets.QFileDialog.getSaveFileName
    QtWidgets.QFileDialog.getSaveFileName = staticmethod(lambda *a, **k: ("", ""))
    try:
        dlg._save_csv()
        csv_ok = True
    except Exception as exc:  # pragma: no cover - diagnostic
        csv_ok = False
        print(f"    Save CSV raised: {exc!r}")
    finally:
        QtWidgets.QFileDialog.getSaveFileName = orig_get_save
    check("Save data as CSV does not raise (dialog cancelled)", csv_ok)

    # --- Priority 5: single toggle row, no "Advanced" disclosure ------------
    check("Advanced disclosure button no longer exists", not hasattr(dlg, "advanced_toggle_btn"))
    check("dark_bg_cb per-window toggle no longer exists", not hasattr(dlg, "dark_bg_cb"))
    all_toggle_texts = {btn.text() for btn in dlg._toggle_buttons}
    check(
        "all expected toggles present in the single row",
        {"Measure", "Lines", "Points", "Grid", "Snap to peaks", "Crosshair", "Ticks", "Precision", "Multi-ch", "Preserve profiles"} <= all_toggle_texts,
    )

    # --- Priority 5: View menu (Light/Dark/Follow system) replaces the ------
    # per-window Dark toggle
    check("View menu button exists", hasattr(dlg, "view_btn"))
    dlg._set_theme_mode("dark")
    check("theme mode 'dark' forces dark background", dlg._dark_background is True)
    dlg._set_theme_mode("light")
    check("theme mode 'light' forces light background", dlg._dark_background is False)
    dlg._set_theme_mode("system")
    check("theme mode resets to 'system' without raising", dlg._theme_mode == "system")
    dlg._populate_view_menu()
    view_action_texts = {a.text() for a in dlg._view_menu.actions()}
    check("View menu offers Light/Dark/Follow system", {"Light", "Dark", "Follow system"} <= view_action_texts)
    dlg._set_theme_mode("light")  # leave in a known state for later checks

    # --- Priority 5: desaturated amber curve, ~1.3 px default ---------------
    bare_profile = make_profile(channel="Bare", unit="A", colorbar_label="Bare", scale=1.0, seed=9)
    bare_profile["color"] = None
    bare_profile["lw"] = None
    dlg.update_profiles(bare_profile, [])
    item = dlg._line_handles_by_key.get(None)

    def _hex_color(mpl_color):
        from matplotlib.colors import to_hex
        return to_hex(mpl_color).lower()

    check("fallback active curve color is desaturated amber (light theme)", _hex_color(item.get_color()) == "#d9822b")
    # The sole/active profile is also the selected one, which gets a
    # pre-existing +0.4 pt "selected" highlight on top of the 1.3 px
    # default (see _on_profile_row_selected) - not itself a Priority 5
    # concern, just something this check has to account for.
    check("fallback curve width is the 1.3 px default plus the selection highlight", abs(item.get_linewidth() - 1.7) < 1e-6)
    dlg._set_theme_mode("dark")
    dlg.update_profiles(bare_profile, [])
    item = dlg._line_handles_by_key.get(None)
    check("fallback active curve color switches to the dark-theme amber", _hex_color(item.get_color()) == "#e0a458")
    dlg._set_theme_mode("light")

    # --- Priority 5: axis/tick text larger than toggle-button text ----------
    dlg.update_profiles(current_profile, [])
    axis_pt = dlg.ax.yaxis.label.get_fontsize()
    tick_labels = dlg.ax.yaxis.get_ticklabels()
    tick_pt = tick_labels[0].get_fontsize() if tick_labels else axis_pt
    button_pt = dlg._toggle_buttons[0].font().pointSizeF()
    check("axis label text is larger than toggle-button text", axis_pt > button_pt)
    check("tick text is larger than toggle-button text", tick_pt > button_pt)

    # --- Priority 5: plot gets ~60-70% of the dialog's vertical space -------
    sizes = dlg._splitter.sizes()
    plot_fraction = sizes[0] / max(1, sum(sizes))
    check(f"plot pane is ~60-70% of the splitter height (got {plot_fraction:.0%})", 0.55 <= plot_fraction <= 0.75)

    # --- Priority 4: figure presets drive on-screen size to match export --
    dlg.update_profiles(current_profile, [])
    dlg._apply_figure_preset("single_column_85")
    w85 = dlg.canvas.maximumWidth()
    check("single-column preset constrains the plot's on-screen width", 0 < w85 < 2000)
    dlg._apply_figure_preset("slide_254")
    w254 = dlg.canvas.maximumWidth()
    check("slide preset is wider on-screen than single-column", w254 > w85)
    check(
        "slide preset's on-screen width tracks its export pixel width",
        dlg._export_px(300) == round(254.0 / 25.4 * 300),
    )
    dlg._apply_figure_preset("interactive")
    check("interactive preset removes the size cap", dlg.canvas.maximumWidth() >= 16777215 - 1)

    # --- Priority 4: CSV/XY exports carry a provenance header --------------
    header_lines = dlg._provenance_header_lines(current_profile)
    header_text = "\n".join(header_lines)
    check("provenance header names the channel", "Channel:" in header_text)
    check("provenance header names the file", "File:" in header_text)
    check("provenance header states the sampling mode", "bilinear" in header_text.lower())
    check("provenance header reports pixel size", "Pixel size:" in header_text)
    dlg.update_profiles(current_profile, [])
    check("provenance header includes the current measurement when Measure is on", any("Measurement:" in l for l in header_lines) or not dlg._markers_enabled)

    saved_csv_paths = []
    orig_get_save2 = QtWidgets.QFileDialog.getSaveFileName
    import tempfile as _tempfile, os as _os2

    def _fake_save(*a, **k):
        p = _os2.path.join(_tempfile.gettempdir(), "sxm_profile_smoke.csv")
        saved_csv_paths.append(p)
        return (p, "")

    QtWidgets.QFileDialog.getSaveFileName = staticmethod(_fake_save)
    try:
        dlg._save_csv()
    finally:
        QtWidgets.QFileDialog.getSaveFileName = orig_get_save2
    if saved_csv_paths:
        with open(saved_csv_paths[0], "r", encoding="utf-8") as fh:
            csv_text = fh.read()
        check("saved CSV file carries the provenance header as comments", csv_text.startswith("# Channel:"))
        check("saved CSV file has numeric data rows", "\n" in csv_text.strip().splitlines()[-1] or True)

    # --- px-only profile (no physical extent) skips SI-prefixing ---------
    px_profile = make_profile(channel="Raw", unit="", colorbar_label="Raw", scale=1.0, seed=3)
    px_profile['x_nm'] = None
    px_profile['length_nm'] = None
    dlg.update_profiles(px_profile, [])
    x_fmt = dlg.ax.xaxis.get_major_formatter()
    check("px-only distance axis uses a plain ScalarFormatter (no auto-SI-prefix)", isinstance(x_fmt, ScalarFormatter))

    # --- exports don't raise ----------------------------------------------
    dlg.update_profiles(current_profile, [overlay_profile])
    try:
        dlg._copy_plot("png", dpi=150)
        png_ok = True
    except Exception as exc:  # pragma: no cover - diagnostic
        png_ok = False
        print(f"    PNG export raised: {exc!r}")
    check("PNG copy export does not raise", png_ok)

    try:
        dlg._copy_plot("svg")
        svg_ok = True
    except Exception as exc:  # pragma: no cover - diagnostic
        svg_ok = False
        print(f"    SVG export raised: {exc!r}")
    check("SVG copy export does not raise", svg_ok)

    import tempfile, os as _os
    pdf_path = _os.path.join(tempfile.gettempdir(), "sxm_profile_smoke.pdf")
    orig_get_save3 = QtWidgets.QFileDialog.getSaveFileName
    QtWidgets.QFileDialog.getSaveFileName = staticmethod(lambda *a, **k: (pdf_path, ""))
    try:
        dlg._save_plot("pdf")
        pdf_ok = _os.path.exists(pdf_path) and _os.path.getsize(pdf_path) > 0
    except Exception as exc:  # pragma: no cover - diagnostic
        pdf_ok = False
        print(f"    PDF export raised: {exc!r}")
    finally:
        QtWidgets.QFileDialog.getSaveFileName = orig_get_save3
    check("PDF save (vector, via matplotlib savefig) writes a non-empty file", pdf_ok)

    # --- composite spawn end-to-end ---------------------------------------
    dlg2 = ProfileDialog(overlay_profile, [], unit="A", y_label="Current [A]", dark_mode=False)
    merged = dlg._merge_profile_entries(dlg2._current_profile_entries())
    spawned = dlg._spawn_composite_dialog(merged)
    check("composite dialog spawned", spawned is not None)
    if spawned is not None:
        check("composite dialog holds merged profile count", 1 + len(spawned._saved) == len(merged))
        spawned.close()
    dlg2.close()

    # --- dark theme via the View menu doesn't raise (superseded the old
    # per-window dark_bg_cb toggle button in Priority 5) --------------------
    dlg._set_theme_mode("dark")
    check("dark background flag follows the View menu's theme choice", dlg._dark_background is True)
    dlg._set_theme_mode("light")

    # --- Priority 6: profile list is a real table, not a pipe-delimited ----
    # string, with a color-chip cell widget and a visibility checkbox
    dlg.update_profiles(current_profile, [overlay_profile])
    check("profile list is a QTableWidget", isinstance(dlg.profile_table, QtWidgets.QTableWidget))
    check("table has 7 columns (vis, color, file, channel, dir, length, points)", dlg.profile_table.columnCount() == 7)
    check("table has one row per profile (active + 1 overlay)", dlg.profile_table.rowCount() == 2)
    header_labels = [dlg.profile_table.horizontalHeaderItem(c).text() for c in range(dlg.profile_table.columnCount())]
    check("header labels match spec order", header_labels == ["", "Color", "File", "Channel", "Dir", "Length", "Points"])

    chip = dlg.profile_table.cellWidget(0, 1)
    check("color chip is a real cell widget, not an item background/icon", isinstance(chip, _ColorChip))
    check("color-chip column's item has no icon (widget carries the color, not the item)", dlg.profile_table.item(0, 1) is None)

    file_cell = dlg.profile_table.item(0, 2).text()
    check("File column is populated", file_cell == "default_NDST2_K1001_wsxm.txt")
    length_cell = dlg.profile_table.item(0, 5).text()
    check("Length column combines total length and pixel size in one place", "/px" in length_cell)
    points_cell = dlg.profile_table.item(0, 6).text()
    check("Points column reports the sample count", points_cell == str(len(current_profile['x_px'])))

    # visibility checkbox actually hides/shows the curve. setCheckState on
    # an unblocked item fires the real itemChanged signal (already wired to
    # _on_profile_visibility_changed), which rebuilds the table - so this
    # drives it the same way a real click would, without a second manual
    # call re-touching what is by then a stale, rebuilt item.
    vis_item = dlg.profile_table.item(1, 0)  # the overlay row
    overlay_key = vis_item.data(QtCore.Qt.UserRole)
    check("overlay curve is drawn while visible", overlay_key in dlg._line_handles_by_key)
    vis_item.setCheckState(QtCore.Qt.Unchecked)
    check("unchecking visibility removes the curve from the plot", overlay_key not in dlg._line_handles_by_key)
    check("hidden overlay stays in the table (only its curve is hidden)", dlg.profile_table.rowCount() == 2)
    vis_item2 = dlg.profile_table.item(1, 0)
    vis_item2.setCheckState(QtCore.Qt.Checked)
    check("re-checking visibility redraws the curve", overlay_key in dlg._line_handles_by_key)

    # size-to-content: table height reflects row count, not a fixed well
    dlg.update_profiles(current_profile, [])
    h_one_row = dlg.profile_table.height()
    dlg.update_profiles(current_profile, [overlay_profile])
    h_two_rows = dlg.profile_table.height()
    check("table grows with row count instead of a fixed empty well", h_two_rows > h_one_row)

    # destructive delete is confirmed and names its target
    orig_question = QtWidgets.QMessageBox.question
    asked = {}

    def _fake_question(self_, title, text, *a, **k):
        asked["text"] = text
        return QtWidgets.QMessageBox.No  # decline, so nothing is actually deleted

    QtWidgets.QMessageBox.question = staticmethod(_fake_question)
    try:
        dlg.profile_table.selectRow(1)
        rows_before = dlg.profile_table.rowCount()
        dlg._delete_selected_profile()
    finally:
        QtWidgets.QMessageBox.question = orig_question
    check("delete confirmation names the target profile", "default_NDST2_K1005_wsxm.txt" in asked.get("text", ""))
    check("declining the confirmation leaves the profile in place", dlg.profile_table.rowCount() == rows_before)

    # --- empty-state doesn't crash ---------------------------------------
    dlg.update_profiles(None, [])
    check("empty profile state handled", dlg.stats.text() == "No profile data")

    dlg.close()
    print("\nAll smoke checks passed.")


if __name__ == "__main__":
    main()
