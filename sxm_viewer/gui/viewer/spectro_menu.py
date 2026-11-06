"""Spectroscopy-thumbnail right-click menu.

Module-function convention (see ``gui/viewer/*.py``): ``SXMGridViewer.
_on_spectro_thumb_context_menu`` is a one-line shim into
``show_spectro_thumb_context_menu``. Built from the shared
``gui/menu_kit.py`` pieces like the image menus:

- greyed header (the file name, or "N spectra selected");
- **Open**: spectrum / compare selected, compare at this position,
  position summary, metadata;
- **Display**: miniature channel;
- **Export**: menu_kit export section, scope ``"spectra"``. The plot is
  rendered here, off-screen, from the spec data (the miniature's channel and
  axis), so no Spectrum window has to be opened first; several selected
  spectra are overlaid in one plot;
- **Assignment**: assign to the current image / clear a manual assignment
  (only when there is one);
- **Organise**: source file.
"""
from __future__ import annotations

import io
import itertools
from pathlib import Path

import numpy as np

from ..._shared import QtGui, QtWidgets, log_status
from ...config import save_config
from ...data.channel_units import guess_channel_unit
from ...utils.units import _auto_display_unit
from ..menu_kit import MenuPrefs, Route, add_export_section, again_folder, section
from ..palettes import DEFAULT_COLOR_CYCLE, get_color_cycle
from ..ppt_bridge import powerpoint_support_status, send_rendered_to_ppt
from ..system_open import add_source_file_menu
from . import export_routes
from . import loader as viewer_loader
from . import thumbnail_ui
from ..spectroscopy import popups as spectro_popups

DEFAULT_PINNED = ("copy_png", "ppt_new")


def menu_prefs(viewer):
    return MenuPrefs("spectra", DEFAULT_PINNED, config=getattr(viewer, "config", None), save=save_config)


# ---------------------------------------------------------------------------
# data + plot
# ---------------------------------------------------------------------------
def _spec_by_path(viewer, path):
    for s in getattr(viewer, "spectros", []) or []:
        if str(s.get("path", "") or "") == path:
            return s
    return None


def _ensure_data(viewer, specs):
    """Hydrate lazily-loaded specs so channels/axis are present."""
    lazy = [s for s in specs if not (s.get("channels") or s.get("data"))]
    if lazy:
        try:
            viewer_loader.hydrate_spectro_entries(viewer, lazy)
        except Exception as exc:
            log_status(f"[Spectro menu] hydrate failed: {exc}")
    return specs


def _axis(spec):
    axes = spec.get("AxisChoices") or []
    label = spec.get("AxisLabel") or "Bias"
    unit = spec.get("AxisUnit") or ""
    vals = None
    if axes:
        first = axes[0] or {}
        vals = np.asarray(first.get("values", []), dtype=float)
        label = first.get("label") or label
        unit = first.get("unit") or unit
    if vals is None or vals.size == 0:
        vals = np.asarray(spec.get("V", []), dtype=float)
        if not spec.get("AxisUnit") and not axes:
            unit = unit or "V"
    return vals, str(label), str(unit)


def _traces(viewer, specs, channel):
    """[(spec, x, y, channel_used)] with the requested channel where the
    spec has it, else its own miniature channel."""
    out = []
    for spec in specs:
        ch = channel if channel in thumbnail_ui._spectro_available_channels(spec) else \
            thumbnail_ui._spectro_display_channel(viewer, spec)
        y = thumbnail_ui._spectro_values_for_channel(spec, ch)
        x, _label, _unit = _axis(spec)
        n = min(x.size, y.size)
        if n >= 2:
            out.append((spec, x[:n], y[:n], ch))
    return out


def _channel_unit(spec, channel):
    unit_map = (spec or {}).get("unit_map") or {}
    return str(unit_map.get(channel) or guess_channel_unit(channel) or "")


def render_plot(viewer, specs, channel, *, dpi=300):
    """A publication-style matplotlib Figure of the spectra (one trace each)."""
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg

    traces = _traces(viewer, specs, channel)
    if not traces:
        raise ValueError("No spectrum data to plot")
    fig = Figure(figsize=(5.0, 3.6), dpi=dpi)
    FigureCanvasAgg(fig)
    ax = fig.add_subplot(1, 1, 1)
    cycle = get_color_cycle(getattr(viewer, "spectro_color_cycle", DEFAULT_COLOR_CYCLE)) or ["C0"]
    ch_name = str(traces[0][3])
    unit = _channel_unit(traces[0][0], ch_name)
    y_unit, factor = _auto_display_unit(unit, np.concatenate([t[2] for t in traces]))
    for (spec, x, y, _ch), color in zip(traces, itertools.cycle(cycle)):
        ax.plot(x, y * factor, lw=1.2, color=color, label=Path(str(spec.get("path", ""))).name)
    _x, xlabel, xunit = _axis(traces[0][0])
    ax.set_xlabel(f"{xlabel} ({xunit})" if xunit else xlabel)
    has_unit = ch_name.endswith(")") and "(" in ch_name
    ax.set_ylabel(f"{ch_name} ({y_unit})" if y_unit and not has_unit else ch_name)
    if len(traces) > 1:
        ax.legend(fontsize=7, frameon=False)
    else:
        ax.set_title(Path(str(traces[0][0].get("path", ""))).name, fontsize=9)
    ax.tick_params(labelsize=8)
    fig.tight_layout()
    return fig


def _png_bytes(viewer, specs, channel):
    buf = io.BytesIO()
    render_plot(viewer, specs, channel).savefig(buf, format="png", dpi=300, bbox_inches="tight")
    return buf.getvalue()


def _base_name(specs, channel, ext):
    stem = Path(str(specs[0].get("path", ""))).stem or "spectrum"
    if len(specs) > 1:
        stem += f"_and_{len(specs) - 1}_more"
    clean = "".join(c if c.isalnum() or c in "-_." else "_" for c in f"{stem}__{channel}")
    return clean.strip("_") + ext


def _unique(path):
    path = Path(path)
    if not path.exists():
        return path
    for n in itertools.count(2):
        cand = path.with_name(f"{path.stem}_{n}{path.suffix}")
        if not cand.exists():
            return cand


def _ask_folder(viewer, title):
    start = str(getattr(viewer, "last_export_dir", None) or getattr(viewer, "last_dir", "") or "")
    folder = QtWidgets.QFileDialog.getExistingDirectory(viewer, title, start)
    if folder:
        viewer.last_export_dir = folder
    return folder or None


def _data_table(viewer, specs, channel):
    """Tab-separated text: a header block per spectrum, then columns."""
    blocks = []
    for spec, x, y, ch in _traces(viewer, specs, channel):
        _x, xlabel, xunit = _axis(spec)
        lines = [f"File\t{Path(str(spec.get('path', ''))).name}", f"Channel\t{ch}",
                 f"Position (nm)\t{spec.get('x', '?')}\t{spec.get('y', '?')}",
                 f"Acquired\t{spec.get('time', '')}", "",
                 f"{xlabel} ({xunit})\t{ch}"]
        lines += [f"{float(a):.9g}\t{float(b):.9g}" for a, b in zip(x, y)]
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


def _write_csvs(viewer, specs, folder):
    """One CSV per spectrum with the axis and every channel."""
    written = []
    for spec in specs:
        x, xlabel, xunit = _axis(spec)
        cols = [(f"{xlabel} ({xunit})" if xunit else xlabel, x)]
        for ch in thumbnail_ui._spectro_available_channels(spec):
            vals = thumbnail_ui._spectro_values_for_channel(spec, ch)
            if vals.size == x.size and ch.strip().lower() != "bias":
                cols.append((ch, vals))
        if x.size < 2 or len(cols) < 2:
            continue
        out = _unique(Path(folder) / (Path(str(spec.get("path", "spectrum"))).stem + ".csv"))
        with open(out, "w", encoding="utf-8", newline="") as fh:
            fh.write(",".join('"' + name.replace('"', "'") + '"' for name, _ in cols) + "\n")
            for row in zip(*(c for _, c in cols)):
                fh.write(",".join(f"{float(v):.9g}" for v in row) + "\n")
        written.append(out)
    return written


# ---------------------------------------------------------------------------
# export routes (ctx = {"specs": [...], "channel": str})
# ---------------------------------------------------------------------------
def _tip(text):
    QtWidgets.QToolTip.showText(QtGui.QCursor.pos(), text)


def spectro_routes(viewer):
    def copy_png(ctx, _dest):
        img = QtGui.QImage.fromData(_png_bytes(viewer, ctx["specs"], ctx["channel"]))
        QtWidgets.QApplication.clipboard().setImage(img)
        _tip("Plot copied")

    def copy_data(ctx, _dest):
        QtWidgets.QApplication.clipboard().setText(_data_table(viewer, ctx["specs"], ctx["channel"]))
        _tip("Data copied")

    def save_png(ctx, dest):
        folder = (dest or {}).get("out_dir")
        if not folder or not Path(folder).is_dir():
            folder = _ask_folder(viewer, "Save spectrum plot to folder")
            if not folder:
                return None
        out = _unique(Path(folder) / _base_name(ctx["specs"], ctx["channel"], ".png"))
        render_plot(viewer, ctx["specs"], ctx["channel"]).savefig(str(out), dpi=300, bbox_inches="tight")
        _tip(f"Saved {out.name}")
        return {"out_dir": str(folder)}

    def save_csv(ctx, dest):
        folder = (dest or {}).get("out_dir")
        if not folder or not Path(folder).is_dir():
            folder = _ask_folder(viewer, "Export spectrum data (CSV) to folder")
            if not folder:
                return None
        written = _write_csvs(viewer, ctx["specs"], folder)
        _tip(f"Wrote {len(written)} CSV file(s)")
        return {"out_dir": str(folder)} if written else None

    def ppt(new_slide):
        def run(ctx, _dest):
            pix = QtGui.QPixmap()
            pix.loadFromData(_png_bytes(viewer, ctx["specs"], ctx["channel"]), "PNG")
            label = Path(str(ctx["specs"][0].get("path", ""))).stem
            try:
                send_rendered_to_ppt(pix, label=label, new_slide=new_slide)
            except Exception as exc:
                QtWidgets.QMessageBox.warning(viewer, "PowerPoint", f"Could not send to PowerPoint:\n{exc}")
        return run

    def ppt_ok(_ctx):
        ok, reason = powerpoint_support_status()
        return ok, reason or "PowerPoint is unavailable"

    def obsidian_run(ctx, note, vault):
        specs, channel = ctx["specs"], ctx["channel"]

        def render(img_dir):
            Path(img_dir).mkdir(parents=True, exist_ok=True)
            out = _unique(Path(img_dir) / _base_name(specs, channel, ".png"))
            render_plot(viewer, specs, channel).savefig(str(out), dpi=300, bbox_inches="tight")
            return [{"path": out, "file_key": str(specs[0].get("path", "")), "channel_idx": None,
                     "caption": _spectro_caption(specs, channel)}], [], False

        return viewer.export_to_obsidian(note, vault, render=render)

    def hydrated(fn):
        def run(ctx, *args):
            _ensure_data(viewer, ctx["specs"])
            try:
                return fn(ctx, *args)
            except ValueError as exc:  # no plottable data
                QtWidgets.QMessageBox.information(viewer, "Spectroscopy export", str(exc))
                return None
        return run

    copy_png, copy_data, save_png, save_csv = map(hydrated, (copy_png, copy_data, save_png, save_csv))
    ppt_new, ppt_current = hydrated(ppt(True)), hydrated(ppt(False))
    obsidian_hydrated = lambda ctx, note, vault: (_ensure_data(viewer, ctx["specs"]), obsidian_run(ctx, note, vault))[1]

    return [
        Route("copy_png", "Copy plot as PNG", copy_png, tip="The spectra as a 300 dpi plot on the clipboard"),
        Route("copy_data", "Copy data (tab-separated)", copy_data, tip="Axis and channel values, ready to paste"),
        Route("save_png", "Save plot as PNG...", save_png, again=again_folder("plot PNG")),
        Route("csv", "Export data (CSV)...", save_csv, tip="One CSV per spectrum with every channel",
              again=again_folder("CSV")),
        Route("ppt_new", "Send plot to PowerPoint", ppt_new, available=ppt_ok),
        Route("ppt_current", "Send plot to current PowerPoint slide", ppt_current, available=ppt_ok),
        *export_routes._obsidian_routes(obsidian_hydrated),
    ]


def _spectro_caption(specs, channel):
    if len(specs) > 1:
        return f"{len(specs)} spectra · {channel}"
    spec = specs[0]
    parts = [f"`{Path(str(spec.get('path', ''))).name}`", str(channel)]
    if spec.get("x") is not None and spec.get("y") is not None:
        try:
            parts.append(f"({float(spec['x']):.2f}, {float(spec['y']):.2f}) nm")
        except Exception:
            pass
    if spec.get("time"):
        parts.append(str(spec.get("time")))
    return " · ".join(parts)


# ---------------------------------------------------------------------------
# menu
# ---------------------------------------------------------------------------
def build_spectro_thumb_context_menu(viewer, label_widget):
    spec = label_widget.property("spectro_entry")
    if not spec:
        return None
    key = str(spec.get("path", "") or "")
    # The label carries a converted copy; work on the viewer's own entry.
    spec = _spec_by_path(viewer, key) or spec
    selected = sorted(set(getattr(viewer, "spectro_thumb_multi_select", set()) or []) | {key})
    targets = selected if len(selected) > 1 else [key]
    # One entry per path, the clicked spectrum first.
    target_specs = [spec] + [s for s in (_spec_by_path(viewer, p) for p in targets if p != key) if s is not None]
    image_key = str(spec.get("image_key") or spec.get("primary_image_key") or "")

    menu = QtWidgets.QMenu(viewer)
    menu.setToolTipsVisible(True)
    section(menu, Path(key).name if len(targets) == 1 else f"{len(targets)} spectra selected", first=True)

    # ---- Open
    if len(target_specs) > 1:
        menu.addAction(f"Compare selected ({len(target_specs)})...").triggered.connect(
            lambda _=False, ss=list(target_specs): spectro_popups._open_spectroscopy_compare_popup(viewer, ss))
        menu.addAction("Open this spectrum").triggered.connect(lambda _=False: viewer._open_spectroscopy_popup(spec))
    else:
        menu.addAction("Open spectrum").triggered.connect(lambda _=False: viewer._open_spectroscopy_popup(spec))
    menu.addAction("Compare spectra at this position").triggered.connect(
        lambda _=False: viewer.spectro_compare_controller.open_site_popup(spec, file_key=image_key))
    menu.addAction("Position summary").triggered.connect(
        lambda _=False: viewer._open_spectro_summary_for_site(spec, file_key=image_key, quiet=True))
    menu.addAction("Show metadata in Details").triggered.connect(lambda _=False: viewer.show_spectroscopy_details(spec))

    # ---- Display
    channels = thumbnail_ui._spectro_available_channels(spec)
    common = set(channels)
    for s in target_specs:
        common &= set(thumbnail_ui._spectro_available_channels(s))
    channel_list = sorted(common) if common else channels
    current = thumbnail_ui._spectro_display_channel(viewer, spec)
    if channel_list:
        section(menu, "Display")
        ch_menu = menu.addMenu(f"Miniature channel: {current}" if current else "Miniature channel")
        group = QtWidgets.QActionGroup(ch_menu)
        for ch_name in channel_list:
            act = ch_menu.addAction(ch_name)
            act.setCheckable(True)
            act.setChecked(ch_name == current)
            group.addAction(act)
            act.triggered.connect(lambda _=False, ch=ch_name, paths=list(targets):
                                  viewer.on_set_spectro_thumbnail_channel(ch, paths))

    # ---- Export
    try:
        ctx = {"specs": list(target_specs), "channel": current}
        add_export_section(menu, spectro_routes(viewer), menu_prefs(viewer), ctx, parent=viewer)
    except Exception as exc:
        log_status(f"[Spectro menu] export section failed: {exc}")

    # ---- Assignment
    current_image_key = viewer._current_spectro_assignment_target_image_key()
    overridden = [s for s in target_specs if str(s.get("assignment_override_image_key") or "").strip()]
    if current_image_key or overridden:
        section(menu, "Assignment")
        if current_image_key:
            menu.addAction(f"Assign to current image ({Path(current_image_key).name})").triggered.connect(
                lambda _=False, ss=list(target_specs): viewer._apply_spectro_assignment_override(ss, current_image_key))
        if overridden:
            menu.addAction("Clear manual assignment").triggered.connect(
                lambda _=False, ss=list(overridden): viewer._clear_spectro_assignment_override(ss))

    # ---- Organise
    section(menu, "Organise")
    add_source_file_menu(menu, spec.get("path"), viewer)
    return menu


def show_spectro_thumb_context_menu(viewer, label_widget, pos):
    menu = build_spectro_thumb_context_menu(viewer, label_widget)
    if menu is not None:
        menu.exec_(label_widget.mapToGlobal(pos))
