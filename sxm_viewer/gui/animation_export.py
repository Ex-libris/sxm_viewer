"""Animation export: a series of images as GIF / MP4 / PNG frames, as displayed.

Every frame is rendered by the main preview canvas, the same route as PNG
export (`viewer/export.py:render_targets_as_displayed`): per-file colormap,
filters, contrast, Zero, quarter turn, display preset, scale bar, molecules
and overlays all apply. On top of that the dialog offers display-time
overrides - one colormap for every frame, one colour range shared by all
frames, overlay switches, a per-frame label - which are applied to the
drawn artists only for the capture and undone right after. Nothing the
user set is changed; the previously shown image comes back on close.

Molecules are per file (`viewer.molecule_overlays`, stored/loaded by
`SXMGridViewer.show_file_channel` on every switch). So the canvas' molecule
list is never cleared here - the old export did that and the next switch
saved the empty list over the image's models - and the "current image's
models on every frame" option reloads each frame's own models after the
capture, before the next switch stores them.

Timing, ordering and padding are Qt-free in `utils/animation.py`.
"""
from __future__ import annotations

import io
from pathlib import Path

from matplotlib.transforms import Bbox
import matplotlib.patheffects as PathEffects

from .._shared import QtCore, QtGui, QtWidgets, np
from .. import cmap_registry
from ..config import save_config
from ..utils import animation as anim
from . import cmap_menu

FORMATS = (("GIF", "gif"), ("MP4 video", "mp4"), ("PNG frames", "png"))
ORDERS = (("Thumbnail order", "grid"), ("Acquisition time", "time"), ("Bias", "bias"), ("File name", "name"))
LABELS = (("No label", "none"), ("File name", "name"), ("Bias", "bias"), ("CH height", "z"),
          ("Date / time", "time"), ("Frame n / N", "index"))
LOOPS = (("Loop", "loop"), ("Forth and back", "pingpong"), ("Play once", "once"))
CONTENT = (("Whole preview", "figure"), ("Image only", "image"))
RANGES = (("As displayed (each image)", "displayed"), ("Same for all frames", "shared"))
MOL_MODES = (("Each image's own", "own"), ("Current image's on every frame", "same"))

CONFIG_KEY = "animation_export"
DEFAULTS = {
    "format": "gif", "order": "grid", "reverse": False, "label": "none", "range": "displayed",
    "content": "figure", "width": 800, "fps": 4.0, "hold": 0.0, "loop": "loop",
    "overlays": True, "mol_mode": "own",
}
# Options that change the rendered pixels (the rest is timing / file type).
_RENDER_KEYS = ("cmap", "range", "label", "content", "width", "scale_bar", "molecules", "mol_mode", "overlays")


# ---------------------------------------------------------------------- #
# Frame metadata                                                          #
# ---------------------------------------------------------------------- #
def _header(viewer, key):
    return ((getattr(viewer, "headers", {}) or {}).get(str(key)) or (None, None))[0] or {}


def _bias_v(header):
    try:
        from .controllers.report import _header_bias_v
        return _header_bias_v(header)
    except Exception:
        return None


def _format_bias(v):
    if v is None:
        return ""
    return f"V = {v * 1e3:+.0f} mV" if abs(v) < 1 else f"V = {v:+.3g} V"


def _ch_label(viewer, key):
    try:
        from .viewer import ch_height
        z = ch_height.ch_abs_z_nm(viewer, key)
        if z is None:
            return ""
        return ch_height.label_text(viewer, key, z).replace("CH", "", 1).strip()
    except Exception:
        return ""


def frame_label(viewer, key, mode, index, count):
    if mode == "name":
        return Path(str(key)).stem
    if mode == "bias":
        return _format_bias(_bias_v(_header(viewer, key)))
    if mode == "z":
        return _ch_label(viewer, key)
    if mode == "time":
        h = _header(viewer, key)
        return " ".join(t for t in (str(h.get("Date", "") or "").strip(), str(h.get("Time", "") or "").strip()) if t)
    if mode == "index":
        return f"{index + 1} / {count}"
    return ""


def order_keys(viewer, keys, mode, reverse=False):
    keys = list(keys)
    if mode == "time":
        def k(key):
            try:
                return (0, viewer._header_datetime_dt(_header(viewer, key), Path(str(key))).timestamp())
            except Exception:
                return (1, 0.0)
        keys = sorted(keys, key=k)
    elif mode == "bias":
        def k(key):
            v = _bias_v(_header(viewer, key))
            return (v is None, v if v is not None else 0.0)
        keys = sorted(keys, key=k)
    elif mode == "name":
        keys = sorted(keys, key=lambda key: Path(str(key)).name.lower())
    if reverse:
        keys.reverse()
    return keys


def _entry_artists(entries):
    """Artists in molecule-overlay entries (3-D and SVG share the dict shape)."""
    out = []
    for e in entries or []:
        for val in (e or {}).values():
            for a in (val if isinstance(val, (list, tuple)) else [val]):
                if hasattr(a, "set_visible") and hasattr(a, "get_visible") and hasattr(a, "axes"):
                    out.append(a)
    return out


# ---------------------------------------------------------------------- #
# Rendering through the preview canvas                                    #
# ---------------------------------------------------------------------- #
class _FrameRenderer:
    """Loads each image into the (render-suspended) main preview and grabs
    it; restores the canvas flags and the shown image in ``end``."""

    def __init__(self, viewer):
        self.viewer = viewer
        self.canvas = getattr(viewer, "preview_canvas", None)
        self._active = False
        self._orig_last = None
        self.ref_key = None
        self.ref_molecules = []
        self.shared = None  # (lo, hi, SI unit)

    def begin(self):
        c = self.canvas
        if c is None or self._active:
            return
        self._orig_last = getattr(self.viewer, "last_preview", None)
        self.ref_key = str(self._orig_last[0]) if self._orig_last else None
        try:
            self.ref_molecules = list(c.export_molecule_state() or [])
        except Exception:
            self.ref_molecules = []
        self._saved = {
            "scale_bar_enabled": bool(getattr(c, "scale_bar_enabled", False)),
            "show_molecules": bool(getattr(c, "show_molecules", True)),
            "_overlays_hidden": bool(getattr(c, "_overlays_hidden", False)),
        }
        c.set_render_suspended(True)
        self._active = True

    def end(self):
        if not self._active:
            return
        self._active = False
        c = self.canvas
        for name, val in self._saved.items():
            setattr(c, name, val)
        try:
            if self._orig_last:
                self.viewer.show_file_channel(self._orig_last[0], self._orig_last[1], use_local_cmap=True)
            else:
                c._redraw()
        except Exception:
            pass
        finally:
            c.set_render_suspended(False)
            try:
                c.draw_idle()
            except Exception:
                pass

    # -- shared colour range ------------------------------------------- #
    def compute_shared_range(self, keys):
        from .thumbnail_render import convert_to_si
        v = self.viewer
        ranges, units = [], set()
        for key in keys:
            try:
                ch = v._thumbnail_channel_idx_for_file(key)
                _h, exports = v._collect_channel_exports(key, ch)
                if not exports:
                    continue
                arr, unit = convert_to_si(exports[0]["arr"], exports[0].get("unit"))
                a = np.asarray(arr, dtype=float)
                a = a[np.isfinite(a)]
                if a.size == 0:
                    continue
                lo, hi = np.percentile(a, [1, 99])
                ranges.append((float(lo), float(hi)))
                units.add(unit)
            except Exception:
                continue
        r = anim.shared_range(ranges) if len(units) == 1 else None
        self.shared = (r[0], r[1], units.pop()) if r else None
        return self.shared

    def _shared_clim(self, ax):
        if not self.shared:
            return None
        from .thumbnail_render import convert_to_si
        view = (getattr(self.canvas, "_ax_view_map", {}) or {}).get(ax) or {}
        one, unit = convert_to_si(np.array([1.0]), view.get("unit"))
        if unit != self.shared[2] or not one[0]:
            return None
        return self.shared[0] / one[0], self.shared[1] / one[0]

    # -- one frame ------------------------------------------------------ #
    def render(self, key, opts, index, count):
        v, c = self.viewer, self.canvas
        c.scale_bar_enabled = bool(opts["scale_bar"])
        c.show_molecules = bool(opts["molecules"])
        c._overlays_hidden = False
        ch = v._thumbnail_channel_idx_for_file(key)
        v.show_file_channel(key, ch, use_local_cmap=True)
        same = bool(opts["molecules"] and opts["mol_mode"] == "same" and self.ref_molecules
                    and str(key) != self.ref_key)
        if same:
            c.import_molecule_state(self.ref_molecules)
        undo = []
        try:
            ax = c.main_ax
            if ax is None or not ax.images:
                raise RuntimeError("nothing to render")
            images = list(ax.images)
            if opts["cmap"]:
                name = cmap_registry.effective_cmap_name(opts["cmap"])
                for im in images:
                    undo.append(lambda im=im, cm=im.get_cmap(): im.set_cmap(cm))
                    im.set_cmap(name)
            clim = self._shared_clim(ax) if opts["range"] == "shared" else None
            if clim:
                im = images[0]
                undo.append(lambda im=im, lim=im.get_clim(): im.set_clim(*lim))
                im.set_clim(*clim)
            if opts["cmap"] or clim:
                c._recolor_scale_bars()  # auto black/white under the new colours
            molecule_artists = _entry_artists(getattr(c, "_molecule_artists", None)) + \
                _entry_artists(getattr(c, "_svg_molecule_artists", None))
            hide = []
            if not opts["molecules"]:
                hide += [a for a in molecule_artists if a.get_visible()]
            if not opts["overlays"]:
                keep = {id(a) for a in list(getattr(c, "_scale_bar_artists", []) or []) + molecule_artists}
                hide += [a for a in c._overlay_artists() if id(a) not in keep]
            for a in hide:
                a.set_visible(False)
                undo.append(lambda a=a: a.set_visible(True))
            text = frame_label(v, key, opts["label"], index, count)
            if text:
                # Bottom-left: the acquisition box sits top, the scale bar
                # bottom-right; sized like the scale-bar text.
                try:
                    fs = float(c._scale_bar_font_size(getattr(c, "_view_font_scale", 1.0)))
                except Exception:
                    fs = 12.0
                t = ax.text(0.03, 0.03, text, transform=ax.transAxes, ha="left", va="bottom",
                            color="white", fontsize=fs, fontweight="bold", zorder=60)
                t.set_path_effects([PathEffects.withStroke(linewidth=max(1.5, fs / 5.0), foreground="black")])
                undo.append(t.remove)
            return self._capture(ax, opts["content"], int(opts["width"]))
        finally:
            for fn in reversed(undo):
                try:
                    fn()
                except Exception:
                    pass
            if same:
                # Back to this image's own models before the next switch
                # stores the canvas state for it.
                v._load_molecule_overlay(key)

    def _capture(self, ax, content, width):
        c = self.canvas
        fig = c.fig
        fig.draw_without_rendering()  # layout + aspect, the canvas itself is suspended
        kw = {}
        if content == "image":
            box = ax.get_window_extent()
            try:
                box = Bbox.intersection(box, ax.images[0].get_window_extent()) or box
            except Exception:
                pass
            inches = box.transformed(fig.dpi_scale_trans.inverted())
            dpi = width / max(inches.width, 1e-3)
            kw = {"bbox_inches": inches, "pad_inches": 0}
        else:
            # Same dpi for every frame (scale never jumps between frames),
            # cropped to what is drawn - the canvas' empty margins go, a
            # title wider than the figure is kept whole. Frames of slightly
            # different size are centred by `pad_to_common` when saving.
            dpi = width / max(fig.get_figwidth(), 1e-3)
            kw = {"bbox_inches": "tight", "pad_inches": 0.06}
        buf = io.BytesIO()
        c._save_current_figure_without_shortcut_hint(
            lambda: fig.savefig(buf, format="png", dpi=dpi, **kw))
        buf.seek(0)
        from PIL import Image
        return np.asarray(Image.open(buf).convert("RGB"))

    def background_rgb(self):
        try:
            r, g, b = self.canvas.fig.get_facecolor()[:3]
            return int(r * 255), int(g * 255), int(b * 255)
        except Exception:
            return 255, 255, 255


# ---------------------------------------------------------------------- #
# Dialog                                                                  #
# ---------------------------------------------------------------------- #
def _combo(items, current):
    box = QtWidgets.QComboBox()
    for text, data in items:
        box.addItem(text, data)
    i = box.findData(current)
    box.setCurrentIndex(i if i >= 0 else 0)
    return box


class AnimationExportDialog(QtWidgets.QDialog):
    def __init__(self, viewer, keys, iio):
        super().__init__(viewer)
        self.viewer = viewer
        self.iio = iio
        self.setWindowTitle("Export animation")
        self.resize(860, 760)
        self._keys_in = list(keys)
        self._cache = {}
        self._gen = 0
        self._pos = 0
        self._playing = True
        self._cmap = None  # None = each image's own colormap
        self.renderer = _FrameRenderer(viewer)
        saved = dict(DEFAULTS)
        saved.update((getattr(viewer, "config", {}) or {}).get(CONFIG_KEY) or {})
        canvas = self.renderer.canvas
        self._build(saved, canvas)
        self.renderer.begin()
        self._apply_order()
        self._invalidate()
        self._play_timer = QtCore.QTimer(self)
        self._play_timer.setSingleShot(True)
        self._play_timer.timeout.connect(self._advance)
        self._schedule_play()

    # -- UI ------------------------------------------------------------ #
    def _build(self, s, canvas):
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(6)
        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        lay.addLayout(grid)

        def row(r, title, widgets):
            lab = QtWidgets.QLabel(title)
            lab.setStyleSheet("font-weight: 600;")
            grid.addWidget(lab, r, 0)
            h = QtWidgets.QHBoxLayout()
            h.setSpacing(8)
            for w in widgets:
                if isinstance(w, str):
                    h.addWidget(QtWidgets.QLabel(w))
                else:
                    h.addWidget(w)
            h.addStretch(1)
            grid.addLayout(h, r, 1)

        self.order_combo = _combo(ORDERS, s["order"])
        self.reverse_cb = QtWidgets.QCheckBox("Reverse")
        self.reverse_cb.setChecked(bool(s["reverse"]))
        self.count_label = QtWidgets.QLabel()
        self.count_label.setStyleSheet("color: palette(mid);")
        row(0, "Frames", [self.order_combo, self.reverse_cb, self.count_label])

        self.cmap_btn = QtWidgets.QPushButton()
        self.cmap_btn.setToolTip("One colormap for every frame, or each image's own (as in the viewer)")
        self.cmap_btn.clicked.connect(self._open_cmap_menu)
        self.range_combo = _combo(RANGES, s["range"])
        self.range_combo.setToolTip("Same for all frames: one colour range over the whole series, "
                                    "so brightness changes between frames are real")
        self.label_combo = _combo(LABELS, s["label"])
        row(1, "Look", [self.cmap_btn, "Colour range", self.range_combo, "Label", self.label_combo])
        self._sync_cmap_button()

        self.scale_cb = QtWidgets.QCheckBox("Scale bar")
        self.scale_cb.setChecked(bool(getattr(canvas, "scale_bar_enabled", True)))
        self.mol_cb = QtWidgets.QCheckBox("Molecules")
        self.mol_cb.setChecked(bool(getattr(canvas, "show_molecules", True)))
        self.mol_combo = _combo(MOL_MODES, s["mol_mode"])
        has_ref = bool(canvas is not None and getattr(canvas, "molecules", None))
        if not has_ref:
            self.mol_combo.setCurrentIndex(0)
            self.mol_combo.model().item(1).setEnabled(False)
            self.mol_combo.setToolTip("Load a molecule on the shown image to put the same model on every frame")
        self.overlay_cb = QtWidgets.QCheckBox("Other overlays (markers, profiles, labels)")
        self.overlay_cb.setChecked(bool(s["overlays"]))
        row(2, "Show", [self.scale_cb, self.mol_cb, self.mol_combo, self.overlay_cb])

        self.format_combo = _combo(FORMATS, s["format"])
        self.content_combo = _combo(CONTENT, s["content"])
        self.width_spin = QtWidgets.QSpinBox()
        self.width_spin.setRange(120, 4000)
        self.width_spin.setSingleStep(100)
        self.width_spin.setSuffix(" px wide")
        self.width_spin.setValue(int(s["width"]))
        self.width_spin.setToolTip("Image only: width of the image. Whole preview: width of the preview "
                                   "before its empty margins are trimmed (same scale on every frame).")
        row(3, "Output", [self.format_combo, self.content_combo, self.width_spin])

        self.fps_spin = QtWidgets.QDoubleSpinBox()
        self.fps_spin.setRange(0.2, 60.0)
        self.fps_spin.setDecimals(1)
        self.fps_spin.setSingleStep(0.5)
        self.fps_spin.setSuffix(" fps")
        self.fps_spin.setValue(float(s["fps"]))
        self.hold_spin = QtWidgets.QDoubleSpinBox()
        self.hold_spin.setRange(0.0, 30.0)
        self.hold_spin.setDecimals(1)
        self.hold_spin.setSingleStep(0.5)
        self.hold_spin.setSuffix(" s")
        self.hold_spin.setValue(float(s["hold"]))
        self.hold_spin.setToolTip("Extra time on the last frame (the turnaround when going forth and back)")
        self.loop_combo = _combo(LOOPS, s["loop"])
        self.duration_label = QtWidgets.QLabel()
        row(4, "Timing", [self.fps_spin, "pause on last frame", self.hold_spin, self.loop_combo,
                          self.duration_label])

        self.preview = QtWidgets.QLabel()
        self.preview.setAlignment(QtCore.Qt.AlignCenter)
        self.preview.setMinimumSize(320, 260)
        self.preview.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        lay.addWidget(self.preview, 1)

        play_row = QtWidgets.QHBoxLayout()
        self.play_btn = QtWidgets.QToolButton()
        self.play_btn.clicked.connect(self._toggle_play)
        self.slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.slider.valueChanged.connect(self._on_slider)
        self.frame_label = QtWidgets.QLabel()
        play_row.addWidget(self.play_btn)
        play_row.addWidget(self.slider, 1)
        play_row.addWidget(self.frame_label)
        lay.addLayout(play_row)

        bottom = QtWidgets.QHBoxLayout()
        self.status = QtWidgets.QLabel()
        self.status.setStyleSheet("color: palette(mid);")
        bottom.addWidget(self.status, 1)
        self.save_btn = QtWidgets.QPushButton("Save…")
        self.save_btn.setDefault(True)
        self.save_btn.clicked.connect(self._save)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.reject)
        bottom.addWidget(self.save_btn)
        bottom.addWidget(close_btn)
        lay.addLayout(bottom)

        for w in (self.range_combo, self.label_combo, self.content_combo, self.mol_combo):
            w.currentIndexChanged.connect(self._invalidate)
        for w in (self.scale_cb, self.mol_cb, self.overlay_cb):
            w.toggled.connect(self._invalidate)
        self.mol_cb.toggled.connect(lambda on: self.mol_combo.setEnabled(on))
        self.mol_combo.setEnabled(self.mol_cb.isChecked())
        self.order_combo.currentIndexChanged.connect(self._on_order_changed)
        self.reverse_cb.toggled.connect(self._on_order_changed)
        self._width_timer = QtCore.QTimer(self)
        self._width_timer.setSingleShot(True)
        self._width_timer.setInterval(450)
        self._width_timer.timeout.connect(self._invalidate)
        self.width_spin.valueChanged.connect(lambda _=0: self._width_timer.start())
        for w in (self.fps_spin, self.hold_spin):
            w.valueChanged.connect(self._timing_changed)
        self.loop_combo.currentIndexChanged.connect(self._timing_changed)
        self.format_combo.currentIndexChanged.connect(self._timing_changed)
        self._sync_play_button()

    def _sync_cmap_button(self):
        name = self._cmap
        self.cmap_btn.setText(f"Colormap: {name}" if name else "Colormap: as displayed")
        icon = cmap_menu._icon(name, 16, 16) if name else None
        self.cmap_btn.setIcon(icon if icon is not None else QtGui.QIcon())

    def _cmap_prefs(self):
        return cmap_menu.CmapPrefs(getattr(self.viewer, "config", None), save_config)

    def _open_cmap_menu(self):
        menu = QtWidgets.QMenu(self)
        a = menu.addAction("As displayed (each image's own colormap)")
        a.setCheckable(True)
        a.setChecked(self._cmap is None)
        a.triggered.connect(lambda _=False: self._set_cmap(None))
        menu.addSeparator()
        cmap_menu.populate_cmap_menu(menu, self._cmap, self._set_cmap, prefs=self._cmap_prefs(),
                                     context="general", on_gallery=self._cmap_gallery)
        menu.exec_(self.cmap_btn.mapToGlobal(QtCore.QPoint(0, self.cmap_btn.height())))

    def _cmap_gallery(self):
        start = self._cmap or str(getattr(self.viewer, "preview_cmap", "") or "viridis")
        name = cmap_menu.pick_from_gallery(self, start, prefs=self._cmap_prefs())
        if name:
            self._set_cmap(name)

    def _set_cmap(self, name):
        self._cmap = name or None
        self._sync_cmap_button()
        self._invalidate()

    # -- state ---------------------------------------------------------- #
    def options(self):
        return {
            "cmap": self._cmap,
            "range": self.range_combo.currentData(),
            "label": self.label_combo.currentData(),
            "content": self.content_combo.currentData(),
            "width": self.width_spin.value(),
            "scale_bar": self.scale_cb.isChecked(),
            "molecules": self.mol_cb.isChecked(),
            "mol_mode": self.mol_combo.currentData(),
            "overlays": self.overlay_cb.isChecked(),
            "format": self.format_combo.currentData(),
            "order": self.order_combo.currentData(),
            "reverse": self.reverse_cb.isChecked(),
            "fps": self.fps_spin.value(),
            "hold": self.hold_spin.value(),
            "loop": self.loop_combo.currentData(),
        }

    def _remember(self):
        cfg = getattr(self.viewer, "config", None)
        if not isinstance(cfg, dict):
            return
        o = self.options()
        cfg[CONFIG_KEY] = {k: o[k] for k in DEFAULTS}
        try:
            save_config(cfg)
        except Exception:
            pass

    def _apply_order(self):
        o = self.options()
        self.keys = order_keys(self.viewer, self._keys_in, o["order"], o["reverse"])
        n = len(self.keys)
        self.slider.blockSignals(True)
        self.slider.setRange(0, max(0, n - 1))
        self.slider.blockSignals(False)
        self._pos = min(self._pos, max(0, n - 1))
        chans = set()
        for k in self.keys:
            try:
                ch = self.viewer._thumbnail_channel_idx_for_file(k)
                fds = (self.viewer.headers.get(str(k)) or (None, None))[1] or []
                chans.add(fds[ch].get("Caption", fds[ch].get("FileName", "")) if 0 <= ch < len(fds) else "")
            except Exception:
                pass
        chans.discard("")
        chan_txt = (", channel " + next(iter(chans))) if len(chans) == 1 else (
            f", {len(chans)} channels" if chans else "")
        self.count_label.setText(f"{n} image{'s' if n != 1 else ''}{chan_txt}")
        self._timing_changed()

    def _on_order_changed(self, *_):
        self._apply_order()
        if self.label_combo.currentData() == "index":
            self._invalidate()
        else:
            self._show(self._pos)

    def _timing_changed(self, *_):
        o = self.options()
        total = anim.total_duration_s(len(getattr(self, "keys", [])), o["fps"], o["hold"], o["loop"])
        self.duration_label.setText(f"= {total:.1f} s per pass")
        fmt = o["format"]
        self.loop_combo.setEnabled(fmt != "png")
        self.hold_spin.setEnabled(fmt != "png")

    def _invalidate(self, *_):
        self._gen += 1
        self._cache = {}
        o = self.options()
        if o["range"] == "shared":
            self.renderer.compute_shared_range(self._keys_in)
            if not self.renderer.shared:
                self.status.setText("Same colour range: the frames have different units or no data - "
                                    "showing each image's own range.")
        QtCore.QTimer.singleShot(0, lambda g=self._gen: self._render_next(g))

    def _render_next(self, gen):
        if gen != self._gen:
            return
        missing = [i for i, k in enumerate(self.keys) if k not in self._cache]
        if not missing:
            self.status.setText(f"Ready - {len(self.keys)} frames.")
            self._show(self._pos)
            return
        # Render the frame the preview is about to show first.
        order = sorted(missing, key=lambda i: (i - self._pos) % max(1, len(self.keys)))
        i = order[0]
        key = self.keys[i]
        self.status.setText(f"Rendering frame {len(self.keys) - len(missing) + 1} of {len(self.keys)}…")
        try:
            self._cache[key] = self.renderer.render(key, self.options(), i, len(self.keys))
        except Exception as exc:
            self._cache[key] = None
            self.status.setText(f"{Path(str(key)).name}: {exc}")
        if i == self._pos:
            self._show(i)
        QtCore.QTimer.singleShot(0, lambda g=gen: self._render_next(g))

    def _render_all(self, parent_title="Export animation"):
        """Synchronously fill the cache (before saving)."""
        missing = [i for i, k in enumerate(self.keys) if k not in self._cache]
        if not missing:
            return True
        prog = QtWidgets.QProgressDialog("Rendering frames…", "Cancel", 0, len(missing), self)
        prog.setWindowTitle(parent_title)
        prog.setWindowModality(QtCore.Qt.WindowModal)
        prog.setMinimumDuration(200)
        gen = self._gen
        for n, i in enumerate(missing):
            prog.setValue(n)
            QtWidgets.QApplication.processEvents()
            if prog.wasCanceled():
                prog.close()
                return False
            key = self.keys[i]
            if key in self._cache:
                continue
            try:
                self._cache[key] = self.renderer.render(key, self.options(), i, len(self.keys))
            except Exception:
                self._cache[key] = None
        prog.setValue(len(missing))
        prog.close()
        return gen == self._gen

    # -- playback ------------------------------------------------------- #
    def _sequence(self):
        o = self.options()
        seq = anim.playback_sequence(len(self.keys), "pingpong" if o["loop"] == "pingpong" else "loop")
        return seq, anim.frame_durations_ms(seq, o["fps"], o["hold"])

    def _schedule_play(self):
        if not self._playing or not self.keys:
            return
        seq, durs = self._sequence()
        try:
            j = seq.index(self._pos)
        except ValueError:
            j = 0
        self._play_timer.start(int(max(15, durs[j] if durs else 250)))

    def _advance(self):
        seq, _ = self._sequence()
        if not seq:
            return
        cur = getattr(self, "_seq_pos", 0)
        if cur >= len(seq) or seq[cur] != self._pos:
            cur = seq.index(self._pos) if self._pos in seq else 0
        cur = (cur + 1) % len(seq)
        self._seq_pos = cur
        self._show(seq[cur])
        self._schedule_play()

    def _toggle_play(self):
        self._playing = not self._playing
        self._sync_play_button()
        if self._playing:
            self._schedule_play()
        else:
            self._play_timer.stop()

    def _sync_play_button(self):
        style = self.style()
        icon = QtWidgets.QStyle.SP_MediaPause if self._playing else QtWidgets.QStyle.SP_MediaPlay
        self.play_btn.setIcon(style.standardIcon(icon))
        self.play_btn.setToolTip("Pause" if self._playing else "Play")

    def _on_slider(self, value):
        if self._playing:
            self._toggle_play()
        self._show(int(value))

    def _show(self, i):
        if not self.keys:
            return
        i = max(0, min(int(i), len(self.keys) - 1))
        self._pos = i
        self.slider.blockSignals(True)
        self.slider.setValue(i)
        self.slider.blockSignals(False)
        key = self.keys[i]
        self.frame_label.setText(f"{i + 1} / {len(self.keys)}  {Path(str(key)).name}")
        frame = self._cache.get(key)
        if frame is None:
            return
        h, w = frame.shape[:2]
        img = QtGui.QImage(np.ascontiguousarray(frame).data, w, h, 3 * w, QtGui.QImage.Format_RGB888).copy()
        pm = QtGui.QPixmap.fromImage(img)
        size = self.preview.size()
        if pm.width() > size.width() or pm.height() > size.height():
            pm = pm.scaled(size, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.preview.setPixmap(pm)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if getattr(self, "keys", None):
            self._show(self._pos)

    # -- save ----------------------------------------------------------- #
    def _save(self):
        o = self.options()
        fmt = o["format"]
        ext = {"gif": ".gif", "mp4": ".mp4", "png": ".png"}[fmt]
        filters = {"gif": "GIF (*.gif)", "mp4": "MP4 video (*.mp4)", "png": "PNG frames (*.png)"}[fmt]
        start_dir = str(getattr(self.viewer, "last_export_dir", None) or getattr(self.viewer, "last_dir", "") or "")
        stem = Path(str(self.keys[0])).stem if self.keys else "animation"
        default = str(Path(start_dir) / f"{stem}_animation{ext}") if start_dir else f"{stem}_animation{ext}"
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save animation", default, filters)
        if not out_path:
            return
        out_path = Path(out_path)
        if out_path.suffix.lower() != ext:
            out_path = out_path.with_suffix(ext)
        if not self._render_all():
            return
        frames = [self._cache.get(k) for k in self.keys]
        frames = [f for f in frames if f is not None]
        if not frames:
            QtWidgets.QMessageBox.warning(self, "Animation", "No frame could be rendered.")
            return
        fill = self.renderer.background_rgb()
        try:
            if fmt == "mp4":
                padded = anim.pad_to_common(frames, fill, multiple=2)
                seq = anim.constant_rate_sequence(
                    anim.playback_sequence(len(padded), o["loop"]), o["fps"], o["hold"])
                self.iio.imwrite(out_path, [padded[i] for i in seq], plugin="FFMPEG",
                                 fps=float(o["fps"]), macro_block_size=1)
                written = out_path
            elif fmt == "gif":
                padded = anim.pad_to_common(frames, fill)
                seq = anim.playback_sequence(len(padded), o["loop"])
                kw = {"duration": anim.frame_durations_ms(seq, o["fps"], o["hold"])}
                if o["loop"] != "once":
                    kw["loop"] = 0
                self.iio.imwrite(out_path, [padded[i] for i in seq], plugin="pillow", **kw)
                written = out_path
            else:
                padded = anim.pad_to_common(frames, fill)
                stem_path = out_path.with_suffix("")
                for i, fr in enumerate(padded):
                    self.iio.imwrite(f"{stem_path}_{i + 1:03d}.png", fr)
                written = out_path.parent
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Animation", f"Failed to save the animation:\n{exc}")
            return
        self.viewer.last_export_dir = str(out_path.parent)
        self._remember()
        h, w = padded[0].shape[:2]
        self.status.setText(f"Saved {len(frames)} frames ({w} x {h} px) to {written}")
        QtWidgets.QMessageBox.information(self, "Animation", f"Saved {len(frames)} frames ({w} x {h} px) to\n{written}")

    def done(self, result):
        try:
            self._play_timer.stop()
        except Exception:
            pass
        self._gen += 1  # stop the background render chain
        self.renderer.end()
        super().done(result)


def on_create_animation(viewer, paths):
    keys = []
    try:
        from .viewer.export import export_targets
        keys = export_targets(viewer, paths)
    except Exception:
        keys = list(dict.fromkeys(str(p) for p in paths or []))
    if not keys:
        QtWidgets.QMessageBox.information(viewer, "Animation", "Select the images to animate first.")
        return
    if getattr(viewer, "preview_canvas", None) is None:
        QtWidgets.QMessageBox.warning(viewer, "Animation", "The preview canvas is unavailable.")
        return
    try:
        import imageio.v3 as iio  # type: ignore
    except Exception:
        QtWidgets.QMessageBox.warning(
            viewer, "Animation",
            "Saving animations needs the imageio package (and imageio-ffmpeg for MP4).\n\n"
            "Install them into the environment that runs SXM Viewer:\n"
            "    python -m pip install imageio imageio-ffmpeg")
        return
    dlg = AnimationExportDialog(viewer, keys, iio)
    dlg.exec_()
