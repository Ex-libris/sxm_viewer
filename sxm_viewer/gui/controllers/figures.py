"""Figure recipes: GUI-thread data gathering + window orchestration.

The recipe logic and renderer are Qt-free (``sxm_viewer/figures/``); this
controller is the only place that touches live viewer state for them:

* ``ensure_series_index()`` - which loaded images belong to a height series
  (drives the ``z×N`` thumbnail chip). Cheap and cached on a signature of
  the CH tags, so it can run at the top of every thumbnail repopulate.
* ``member_inputs()`` / ``panel()`` - plain dicts/arrays for the spec and
  renderer, taken through the same display pipeline as the preview
  (filters, display adjustments, quarter turns, Zero) so a figure looks
  like the images you already tuned.
* ``open_height_series(key)`` / ``open_selection(keys)`` - one window per
  series, reused if already open; per-series user edits (reference,
  dropped frames, row channels, ranges, toggles) are remembered in
  ``viewer.figure_states`` (config + sessions).
"""
from __future__ import annotations

from pathlib import Path

from ..._shared import QtCore, QtWidgets, log_status, np
from ...config import save_config
from ...figures import height_series as hs
from ...processing.detection import _find_topography_channel
from ...reporting.channels import classify_channel, pick_image_channels
from ..viewer import ch_height as viewer_ch_height
from ..viewer import thumbnail_ui as viewer_thumb_ui

_STATE_LIMIT = 200          # remembered figure states kept in the config
_ABORTED_FRAC = 0.25        # < this fraction of rows scanned = stopped scan
_IDLE_BATCH = 4             # frames checked per idle tick
_CLASS_OF_UNIT = {"A": "current", "mA": "current", "uA": "current", "µA": "current",
                  "nA": "current", "pA": "current", "fA": "current",
                  "Hz": "freq", "kHz": "freq", "mHz": "freq",
                  "m": "z", "nm": "z", "pm": "z", "Å": "z"}
_LENGTH_TO_NM = {"nm": 1.0, "Å": 0.1, "A": 0.1, "ang": 0.1, "pm": 1e-3,
                 "um": 1e3, "µm": 1e3, "m": 1e9}


def _mask_aborted_rows(arr):
    """NaN out the rows a stopped scan never reached (filled with a constant
    - often 0 next to real topography at -185 nm, which would flatten the
    whole shared range). Same detector as the thumbnails and the preview's
    auto contrast; geometry is kept, so panels still line up."""
    a = np.array(arr, dtype=float, copy=True)
    if a.ndim != 2:
        return a
    try:
        from ..thumbnail_render import detect_valid_scan_region
        region = detect_valid_scan_region(a)
    except Exception:
        region = None
    if region:
        r0, r1 = region
        a[:r0, :] = np.nan
        a[r1 + 1:, :] = np.nan
    return a


def _header_bias_v(header):
    from .report import _header_bias_v as _bias
    return _bias(header)


_SP_PREFIX = {"f": 1e-15, "p": 1e-12, "n": 1e-9, "u": 1e-6, "\u00b5": 1e-6, "\u03bc": 1e-6,
              "m": 1e-3, "": 1.0, "k": 1e3}


def _header_setpoint(header):
    """(value in base SI unit, base unit) of the feedback setpoint, e.g.
    Anfatec ``SetPoint: 50`` + ``SetPointPhysUnit: pA`` -> (5e-11, "A")."""
    try:
        val = float(header.get("SetPoint"))
    except (TypeError, ValueError):
        return None, ""
    unit = str(header.get("SetPointPhysUnit") or "").strip()
    for base in ("Hz", "A", "V", "m"):
        if unit.endswith(base) and unit[:-len(base)] in _SP_PREFIX:
            return val * _SP_PREFIX[unit[:-len(base)]], base
    return val, unit


class FigureController:
    """Owns the figure windows and the height-series index."""

    def __init__(self, viewer):
        self.viewer = viewer
        self._index_sig = None
        self._series = {}          # id -> series dict (keys ordered by z)
        self._series_of_key = {}   # image key -> series id
        self._windows = {}         # id -> window
        # Fraction of rows actually scanned, per series frame - checked in
        # idle time after a folder load (reading every frame up front would
        # slow loading), so chips stop counting stopped scans once known.
        self._scan_frac = {}
        self._idle_queue = []
        self._idle_timer = QtCore.QTimer()
        self._idle_timer.setSingleShot(True)
        self._idle_timer.setInterval(0)
        self._idle_timer.timeout.connect(self._idle_check_scans)
        states = viewer.config.get("figure_states")
        viewer.figure_states = dict(states) if isinstance(states, dict) else {}

    # ------------------------------------------------------------------
    # Series index (thumbnail chips)
    # ------------------------------------------------------------------
    def _index_signature(self):
        """Everything detection depends on: the loaded files (bias/footprint
        come from their headers) and the CH tags (tag, height, Z range)."""
        viewer = self.viewer
        tags = getattr(viewer, "tags", {}) or {}
        sig = []
        for p in getattr(viewer, "files", []) or []:
            key = str(p)
            info = tags.get(key) or {}
            sig.append((key, info.get("tag"), info.get("abs_z_pm"), info.get("rng_nm"),
                        self._scan_frac.get(key)))
        return tuple(sig)

    def ensure_series_index(self):
        """Recompute the series index if the files or CH tags changed.

        Height series (CH frames at a ladder of tip heights) and bias series
        (same frame at a ladder of biases, CC or CH) come from the folder
        report's own sequence detector; a frame belongs to at most one."""
        sig = self._index_signature()
        if sig == self._index_sig:
            return False
        self._index_sig = sig
        series, by_key = {}, {}
        if len(sig) >= 2:
            try:
                images = self.image_inputs([s[0] for s in sig])
                found = hs.find_height_series(images) + hs.find_bias_series(images)
                found = self._prune_stopped_scans(found, {i["key"]: i for i in images})
                for s in found:
                    if any(k in by_key for k in s["keys"]):
                        continue  # never two chips on one frame
                    series[s["id"]] = s
                    for k in s["keys"]:
                        by_key[k] = s["id"]
            except Exception as exc:
                log_status(f"[Figures] series detection failed: {exc}")
        changed = (by_key != self._series_of_key or
                   {k: v.get("levels") for k, v in series.items()} !=
                   {k: v.get("levels") for k, v in self._series.items()})
        self._series, self._series_of_key = series, by_key
        self._schedule_scan_checks()
        if series:
            n_h = sum(1 for s in series.values() if s.get("recipe") == "height")
            log_status(f"[Figures] {n_h} height / {len(series) - n_h} bias series detected "
                       f"({sum(len(s['keys']) for s in series.values())} images)")
        return changed

    def _prune_stopped_scans(self, found, by_key):
        """Move frames whose scan stopped early out of each series (listed
        as excluded, reason "aborted"); drop a series left with < 2
        acquisition conditions. Frames not checked yet stay in for now."""
        out = []
        for s in found:
            stopped = [k for k in s["keys"]
                       if self._scan_frac.get(k) is not None and self._scan_frac[k] < _ABORTED_FRAC]
            if stopped:
                keys = [k for k in s["keys"] if k not in stopped]
                levels = len({hs.condition_key(by_key[k]) for k in keys})
                if levels < 2:
                    continue
                s = dict(s, keys=keys, levels=levels,
                         excluded=list(s.get("excluded") or []) + [
                             {"key": k, "reason": "aborted", "frac": self._scan_frac[k]} for k in stopped])
            out.append(s)
        return out

    def _schedule_scan_checks(self):
        files = {str(p) for p in (getattr(self.viewer, "files", []) or [])}
        todo = [k for s in self._series.values() for k in s["keys"]
                if k in files and k not in self._scan_frac]
        if todo:
            self._idle_queue = todo
            self._idle_timer.start()

    def _idle_check_scans(self):
        files = {str(p) for p in (getattr(self.viewer, "files", []) or [])}
        batch, self._idle_queue = self._idle_queue[:_IDLE_BATCH], self._idle_queue[_IDLE_BATCH:]
        for key in batch:
            if key not in files or key in self._scan_frac:
                continue
            header, fds = (self.viewer.headers or {}).get(key, (None, None))
            try:
                tag = viewer_thumb_ui.effective_tag(self.viewer, key)
                topo = _find_topography_channel(list(fds or []))
                picks = pick_image_channels(list(fds or []), tag, topo)
                idx = picks[0][0] if picks else 0
                self._scan_frac[key] = float(self.valid_fraction(key, idx))
            except Exception:
                self._scan_frac[key] = 1.0
        if self._idle_queue:
            self._idle_timer.start()
            return
        before = {sid: (len(s["keys"]), s.get("levels")) for sid, s in self._series.items()}
        if any(f < _ABORTED_FRAC for f in self._scan_frac.values()) and self.ensure_series_index():
            pass
        after = {sid: (len(s["keys"]), s.get("levels")) for sid, s in self._series.items()}
        if after != before:
            try:
                self.viewer.populate_thumbnails_for_channel(self.viewer.channel_dropdown.currentIndex())
            except Exception:
                pass

    def series_for_key(self, key):
        sid = self._series_of_key.get(str(key))
        return self._series.get(sid) if sid else None

    # ------------------------------------------------------------------
    # Inputs
    # ------------------------------------------------------------------
    def _z_nm(self, key, tag=None):
        # Only constant-height frames have a controlled tip height; asking
        # the HUD for others may sample a file per image for nothing.
        if tag is not None and tag != "constant-height":
            return None
        try:
            return viewer_ch_height.ch_abs_z_nm(self.viewer, key)
        except Exception:
            return None

    def image_inputs(self, keys):
        viewer = self.viewer
        out = []
        for key in keys:
            key = str(key)
            header, fds = (viewer.headers or {}).get(key, (None, None))
            if not header:
                continue
            try:
                extent = viewer._header_extent(header)
            except Exception:
                extent = None
            if extent and list(extent) == [0.0, 1.0, 0.0, 1.0]:
                extent = None  # unit-square fallback: no real frame size
            try:
                time = viewer._header_datetime_dt(header, Path(key))
            except Exception:
                time = None
            tag = viewer_thumb_ui.effective_tag(viewer, key)
            setpoint, setpoint_unit = _header_setpoint(header)
            out.append({
                "key": key,
                "name": Path(key).stem,
                "time": time,
                "extent": list(extent) if extent else None,
                "angle": float(viewer._header_scan_angle(header) or 0.0),
                "tag": tag,
                "bias_v": _header_bias_v(header),
                "setpoint": setpoint,
                "setpoint_unit": setpoint_unit,
                "z_nm": self._z_nm(key, tag),
                "z_range_nm": (viewer.tags.get(key) or {}).get("rng_nm"),
            })
        return out

    def member_inputs(self, keys):
        viewer = self.viewer
        members = []
        for img in self.image_inputs(keys):
            key = img["key"]
            header, fds = viewer.headers.get(key, (None, None))
            fds = list(fds or [])
            channels = []
            for idx, fd in enumerate(fds):
                label = str(fd.get("Caption") or fd.get("FileName") or f"Channel {idx}")
                cls = classify_channel(f"{fd.get('Caption', '')} {fd.get('FileName', '')}")
                if cls == "other":
                    # Vendor naming we don't know yet: the unit still says
                    # what the channel is (A = current, Hz = Δf, m = height).
                    cls = _CLASS_OF_UNIT.get(str(fd.get("PhysUnit") or "").strip(), "other")
                channels.append({"idx": idx, "label": label, "cls": cls})
            tag = img.get("tag")
            try:
                topo_idx = _find_topography_channel(fds)
            except Exception:
                topo_idx = None
            picks = pick_image_channels(fds, tag, topo_idx)
            members.append(dict(img, channels=channels, picks=picks))
        return members

    def panel(self, key, channel_idx):
        """Displayed data of one channel, in the channel's normalised unit
        (the renderer picks one display unit per row)."""
        viewer = self.viewer
        key = str(key)
        header, fds = viewer.headers.get(key, (None, None))
        if not header or not fds or not (0 <= channel_idx < len(fds)):
            return None
        fd = fds[channel_idx]
        unit, arr = viewer._get_filtered_channel_array(key, channel_idx, header, fd)
        arr = _mask_aborted_rows(arr)
        base_extent = viewer._header_extent(header)
        arr, extent = viewer._apply_adjustments_for_channel(key, channel_idx, arr, base_extent)
        arr = np.asarray(arr, dtype=float)
        if viewer._is_image_zeroed(key):
            ref = viewer._zero_reference(arr)
            if ref is not None:
                arr = arr - ref
        width_nm = height_nm = None
        scale = _LENGTH_TO_NM.get(str(header.get("XPhysUnit") or "nm").strip(), None)
        if extent and list(base_extent) != [0.0, 1.0, 0.0, 1.0] and scale:
            width_nm = abs(float(extent[1]) - float(extent[0])) * scale
            height_nm = abs(float(extent[3]) - float(extent[2])) * scale
            if not (width_nm > 0 and height_nm > 0):
                width_nm = height_nm = None
        return {"data": arr, "unit": unit or "", "width_nm": width_nm,
                "height_nm": height_nm, "caption": str(fd.get("Caption") or "")}

    def valid_fraction(self, key, channel_idx):
        """Fraction of the frame's rows that were actually scanned (1.0 when
        unknown or when the channel has no variation to judge by)."""
        viewer = self.viewer
        header, fds = viewer.headers.get(str(key), (None, None))
        if not header or not fds or not (0 <= int(channel_idx) < len(fds)):
            return 1.0
        try:
            from ..thumbnail_render import detect_valid_scan_region
            raw = np.asarray(viewer._get_channel_array(str(key), int(channel_idx), header,
                                                       fds[int(channel_idx)]), dtype=float)
            region = detect_valid_scan_region(raw)
        except Exception:
            return 1.0
        if raw.ndim != 2 or not raw.shape[0] or not region:
            return 1.0
        return (region[1] - region[0] + 1) / float(raw.shape[0])

    def data_signature(self, keys):
        """Cheap fingerprint of everything that changes a panel's pixels."""
        viewer = self.viewer
        sig = []
        for key in keys:
            key = str(key)
            try:
                filt = viewer._thumbnail_filter_signature(key)
            except Exception:
                filt = None
            try:
                adj = repr(viewer.image_adjustments.get(key))
            except Exception:
                adj = None
            sig.append((key, repr(filt), adj, viewer._is_image_zeroed(key),
                        viewer._get_image_orientation(key),
                        (viewer.tags.get(key) or {}).get("abs_z_pm")))
        return tuple(sig)

    def default_reference(self, keys):
        ref = getattr(self.viewer, "ch_z_reference", None) or {}
        key = ref.get("file_key")
        return str(key) if key and str(key) in {str(k) for k in keys} else None

    # ------------------------------------------------------------------
    # Remembered per-figure state
    # ------------------------------------------------------------------
    def get_state(self, fid):
        state = (self.viewer.figure_states or {}).get(fid)
        return dict(state) if isinstance(state, dict) else {}

    def set_state(self, fid, state):
        states = self.viewer.figure_states
        states.pop(fid, None)
        states[fid] = dict(state)
        while len(states) > _STATE_LIMIT:
            states.pop(next(iter(states)))
        self.viewer.config["figure_states"] = states
        try:
            save_config(self.viewer.config)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Windows
    # ------------------------------------------------------------------
    def open_series(self, key):
        """The figure of the series ``key`` belongs to (height or bias)."""
        self.ensure_series_index()
        series = self.series_for_key(key)
        if series is None:
            return self.open_selection([key])
        return self._open(series["id"], list(series["keys"]), mode=series.get("recipe", "height"),
                          excluded=series.get("excluded"))

    open_height_series = open_series

    def open_selection(self, keys):
        keys = [str(k) for k in keys if str(k) in (self.viewer.headers or {})]
        if not keys:
            return None
        self.ensure_series_index()
        # A selection that is exactly (part of) one series opens that series.
        sids = {self._series_of_key.get(k) for k in keys}
        if len(sids) == 1 and None not in sids and len(keys) >= 2:
            series = self._series[next(iter(sids))]
            if set(keys) == set(series["keys"]):
                return self._open(series["id"], list(series["keys"]),
                                  mode=series.get("recipe", "height"),
                                  excluded=series.get("excluded"))
        return self._open("sel-" + hs.series_id(keys), keys, mode="selection")

    def _open(self, fid, keys, mode, excluded=None):
        from ..dialogs.figure_window import FigureWindow

        win = self._windows.get(fid)
        if win is not None:
            try:
                win.showNormal()
                win.raise_()
                win.activateWindow()
                return win
            except RuntimeError:
                self._windows.pop(fid, None)
        try:
            win = FigureWindow(self, fid, keys, mode=mode, excluded=excluded, parent=self.viewer)
        except Exception as exc:
            # This runs from a thumbnail click: an exception escaping a Qt
            # event handler aborts the whole application (PyQt >= 5.5) and
            # would take the user's unsaved session with it.
            import traceback
            log_status(f"[Figures] could not open figure: {exc}\n{traceback.format_exc()}")
            QtWidgets.QMessageBox.warning(self.viewer, "Figure", f"Could not open the figure:\n{exc}")
            return None
        self._windows[fid] = win
        win.destroyed.connect(lambda *_a, f=fid: self._windows.pop(f, None))
        win.show()
        return win

    # The user's format (letters, scale bar, colour bars, layout ...) is
    # remembered for new figures: the last style edited anywhere.
    def default_style(self):
        style = self.viewer.config.get("figure_style_default")
        return dict(style) if isinstance(style, dict) else {}

    def set_default_style(self, style):
        # The header is per recipe ("auto" = dz for height, V for bias), so
        # it never carries over from one figure to the next.
        style = {k: v for k, v in dict(style or {}).items() if k != "header"}
        self.viewer.config["figure_style_default"] = style
        try:
            save_config(self.viewer.config)
        except Exception:
            pass

    def featured_cmaps(self):
        try:
            from ...cmap_registry import featured_cmap_names
            names = list(featured_cmap_names("preview"))
        except Exception:
            names = []
        base = ["Blues_r", "gray", "afmhot", "viridis", "inferno", "RdBu_r", "bwr",
                "cividis", "magma", "twilight"]
        for n in base:
            if n not in names:
                names.append(n)
        return names[:24]

    def forget_window(self, fid, win):
        if self._windows.get(fid) is win:
            self._windows.pop(fid, None)

    # ------------------------------------------------------------------
    # Thin hooks back into the viewer (window stays viewer-agnostic)
    # ------------------------------------------------------------------
    def show_image(self, key, channel_idx):
        try:
            self.viewer.on_thumbnail_clicked(str(key), int(channel_idx))
            self.viewer._scroll_to_thumbnail(str(key))
        except Exception:
            pass

    def theme_is_dark(self):
        return bool(getattr(self.viewer, "dark_mode", False))

    def default_save_dir(self, keys):
        for k in keys:
            try:
                return str(Path(k).parent)
            except Exception:
                continue
        return str(getattr(self.viewer, "last_dir", "") or "")

    def open_in_canvas(self, cells, n_cols):
        """``cells``: [(row, col, key, channel_idx, cmap, kind)]."""
        viewer = self.viewer
        try:
            viewer._on_open_canvas()
            win = viewer._canvas_window_ref()
        except Exception:
            win = None
        if win is None:
            return
        kind_map = {"current": "current", "df": "df", "z": "topo"}
        placed = []
        for r, c, key, idx, cmap, kind in cells:
            try:
                item = win._add_view_from_header(Path(key), int(idx), cmap_override=cmap,
                                                 place=False, kind=kind_map.get(kind))
            except Exception:
                item = None
            if item is not None:
                placed.append((r, c, item))
        if not placed:
            return
        # Same rows x columns as the figure (the canvas' own reflow is a
        # square-ish grid that would scramble channel rows).
        w = max(i.boundingRect().width() for _, _, i in placed)
        h = max(i.boundingRect().height() for _, _, i in placed)
        gap, margin = 16.0, 40.0
        for r, c, item in placed:
            item.setPos(margin + c * (w + gap), margin + r * (h + gap))

