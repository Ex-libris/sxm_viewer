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

from ..._shared import QtWidgets, log_status, np
from ...config import save_config
from ...figures import height_series as hs
from ...processing.detection import _find_topography_channel
from ...reporting.channels import classify_channel, pick_image_channels
from ..viewer import ch_height as viewer_ch_height
from ..viewer import thumbnail_ui as viewer_thumb_ui

_STATE_LIMIT = 200          # remembered figure states kept in the config
_LENGTH_TO_NM = {"nm": 1.0, "Å": 0.1, "A": 0.1, "ang": 0.1, "pm": 1e-3,
                 "um": 1e3, "µm": 1e3, "m": 1e9}


def _header_bias_v(header):
    from .report import _header_bias_v as _bias
    return _bias(header)


class FigureController:
    """Owns the figure windows and the height-series index."""

    def __init__(self, viewer):
        self.viewer = viewer
        self._index_sig = None
        self._series = {}          # id -> series dict (keys ordered by z)
        self._series_of_key = {}   # image key -> series id
        self._windows = {}         # id -> window
        states = viewer.config.get("figure_states")
        viewer.figure_states = dict(states) if isinstance(states, dict) else {}

    # ------------------------------------------------------------------
    # Series index (thumbnail chips)
    # ------------------------------------------------------------------
    def _index_signature(self):
        viewer = self.viewer
        tags = getattr(viewer, "tags", {}) or {}
        sig = []
        for p in getattr(viewer, "files", []) or []:
            key = str(p)
            info = tags.get(key) or {}
            if info.get("tag") != "constant-height":
                continue
            sig.append((key, info.get("abs_z_pm"), info.get("rng_nm")))
        return tuple(sig)

    def ensure_series_index(self):
        """Recompute the series index if CH tags/heights changed."""
        sig = self._index_signature()
        if sig == self._index_sig:
            return False
        self._index_sig = sig
        series, by_key = {}, {}
        if len(sig) >= 2:
            try:
                images = self.image_inputs([s[0] for s in sig])
                for s in hs.find_height_series(images):
                    series[s["id"]] = s
                    for k in s["keys"]:
                        by_key[k] = s["id"]
            except Exception as exc:
                log_status(f"[Figures] height-series detection failed: {exc}")
        changed = by_key != self._series_of_key
        self._series, self._series_of_key = series, by_key
        if series:
            log_status(f"[Figures] {len(series)} height series detected "
                       f"({sum(len(s['keys']) for s in series.values())} images)")
        return changed

    def series_for_key(self, key):
        sid = self._series_of_key.get(str(key))
        return self._series.get(sid) if sid else None

    # ------------------------------------------------------------------
    # Inputs
    # ------------------------------------------------------------------
    def _z_nm(self, key):
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
            out.append({
                "key": key,
                "name": Path(key).stem,
                "time": time,
                "extent": list(extent) if extent else None,
                "angle": float(viewer._header_scan_angle(header) or 0.0),
                "tag": viewer_thumb_ui.effective_tag(viewer, key),
                "bias_v": _header_bias_v(header),
                "z_nm": self._z_nm(key),
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
    def open_height_series(self, key):
        self.ensure_series_index()
        series = self.series_for_key(key)
        if series is None:
            return self.open_selection([key])
        return self._open(series["id"], list(series["keys"]), mode="height",
                          excluded=series.get("excluded"))

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
                return self._open(series["id"], list(series["keys"]), mode="height",
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
        win = FigureWindow(self, fid, keys, mode=mode, excluded=excluded, parent=self.viewer)
        self._windows[fid] = win
        win.destroyed.connect(lambda *_a, f=fid: self._windows.pop(f, None))
        win.show()
        return win

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

