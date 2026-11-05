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
* Custom figures (``open_custom``) - any images, in the user's order, under
  a stable ``custom-<id>`` with a name: frames can be added, removed and
  reordered without losing the figure's edits. Listed by
  ``recent_custom_figures``; ``save_figure_file`` / ``open_figure_file``
  write / read one figure as ``*.sxmfig.json``.
* Per-panel processing (``panel(..., edits)``): a panel can carry its own
  filter steps and Zero instead of following the viewer's per-file ones.
"""
from __future__ import annotations

import json
import uuid
from pathlib import Path

from ..._shared import QtCore, QtWidgets, log_status, np
from ...config import save_config
from ...data.io import normalize_unit_and_data
from ...figures import height_series as hs
from ...processing.detection import _find_topography_channel
from ...reporting.channels import classify_channel, pick_image_channels
from ..viewer import ch_height as viewer_ch_height
from ..viewer import thumbnail_ui as viewer_thumb_ui

_STATE_LIMIT = 200          # remembered figure states kept in the config
FIGURE_FILE_KIND = "sxm_figure"
FIGURE_FILE_SUFFIX = ".sxmfig.json"
_RECENT_LIMIT = 12
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
        self._series_of_key = {}   # image key -> series id (height / bias)
        self._zoom_of_key = {}     # overview key -> zoom-family id
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
        series, by_key, zoom_of = {}, {}, {}
        if len(sig) >= 2:
            try:
                images = self.image_inputs([s[0] for s in sig])
                img_by_key = {i["key"]: i for i in images}
                found = hs.find_height_series(images) + hs.find_bias_series(images)
                found = self._prune_stopped_scans(found, img_by_key)
                for s in found:
                    if any(k in by_key for k in s["keys"]):
                        continue  # never two series chips on one frame
                    series[s["id"]] = s
                    for k in s["keys"]:
                        by_key[k] = s["id"]
                # Overview + zooms: its own chip on the overview, so a frame
                # can be both part of a series and an overview.
                for fam in self._prune_stopped_zooms(hs.find_zoom_families(images), img_by_key):
                    series[fam["id"]] = fam
                    zoom_of[fam["root"]] = fam["id"]
            except Exception as exc:
                log_status(f"[Figures] series detection failed: {exc}")
        changed = (by_key != self._series_of_key or zoom_of != self._zoom_of_key or
                   {k: v.get("levels") for k, v in series.items()} !=
                   {k: v.get("levels") for k, v in self._series.items()})
        self._series, self._series_of_key, self._zoom_of_key = series, by_key, zoom_of
        self._schedule_scan_checks()
        if series:
            n = {r: sum(1 for s in series.values() if s.get("recipe") == r)
                 for r in ("height", "bias", "zoom")}
            log_status(f"[Figures] {n['height']} height / {n['bias']} bias series, "
                       f"{n['zoom']} overviews with zooms detected")
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

    def _prune_stopped_zooms(self, families, by_key):
        out = []
        for fam in families:
            stopped = [k for k in fam["keys"]
                       if self._scan_frac.get(k) is not None and self._scan_frac[k] < _ABORTED_FRAC]
            if fam["root"] in stopped:
                continue
            if stopped:
                kids = [by_key[k] for k in fam["keys"][1:] if k not in stopped and k in by_key]
                if not kids:
                    continue
                kept, _hidden = hs._collapse_same_footprint(kids)
                fam = dict(fam, keys=[fam["root"]] + [str(m["key"]) for m in kids], levels=len(kept),
                           excluded=[{"key": k, "reason": "aborted", "frac": self._scan_frac[k]}
                                     for k in stopped])
            out.append(fam)
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

    def zoom_for_key(self, key):
        """The overview + zooms family whose overview is ``key``, or None."""
        fid = self._zoom_of_key.get(str(key))
        return self._series.get(fid) if fid else None

    def open_zoom(self, key):
        self.ensure_series_index()
        fam = self.zoom_for_key(key)
        if fam is None:
            return None
        return self._open(fam["id"], list(fam["keys"]), mode="zoom",
                          excluded=fam.get("excluded"), root_key=fam["root"])

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

    def panel(self, key, channel_idx, edits=None):
        """Displayed data of one channel, in the channel's normalised unit
        (the renderer picks one display unit per row).

        ``edits`` - the figure's own processing for this one panel:
        ``{"filters": [steps], "zero": bool}``; a key that is absent follows
        the viewer (its per-file filter pipeline / Zero). ``filters: []``
        means "no filters" even when the viewer has some."""
        viewer = self.viewer
        key = str(key)
        header, fds = viewer.headers.get(key, (None, None))
        if not header or not fds or not (0 <= channel_idx < len(fds)):
            return None
        fd = fds[channel_idx]
        edits = edits or {}
        steps = edits.get("filters")
        if steps is None:
            unit, arr = viewer._get_filtered_channel_array(key, channel_idx, header, fd)
        else:
            unit, arr = self.raw_array(key, channel_idx)
            if steps:
                arr = viewer._apply_filter_pipeline(arr, list(steps))
        arr = _mask_aborted_rows(arr)
        base_extent = viewer._header_extent(header)
        arr, extent = viewer._apply_adjustments_for_channel(key, channel_idx, arr, base_extent)
        arr = np.asarray(arr, dtype=float)
        zero = edits.get("zero")
        if zero is None:
            zero = viewer._is_image_zeroed(key)
        if zero:
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
        try:
            turns = int(viewer._get_image_orientation(key) or 0)
        except Exception:
            turns = 0
        return {"turns": turns, "data": arr, "unit": unit or "", "width_nm": width_nm,
                "height_nm": height_nm, "caption": str(fd.get("Caption") or "")}

    def raw_array(self, key, channel_idx):
        """``(unit, array)`` of one channel before any filter, in the
        channel's normalised unit (the base a panel's own filters run on)."""
        viewer = self.viewer
        header, fds = viewer.headers.get(str(key), (None, None))
        fd = fds[int(channel_idx)]
        raw = viewer._get_channel_array(str(key), int(channel_idx), header, fd)
        unit, arr = normalize_unit_and_data(raw, fd.get("PhysUnit", ""))
        return unit, np.asarray(arr, dtype=float)

    def viewer_filter_steps(self, key):
        """The viewer's filter pipeline of a file (what a panel follows
        until it gets its own)."""
        spec = (getattr(self.viewer, "thumbnail_filters", {}) or {}).get(str(key)) or {}
        return [dict(s) for s in (spec.get("steps") or []) if isinstance(s, dict)]

    def viewer_zero(self, key):
        return bool(self.viewer._is_image_zeroed(str(key)))

    def run_filter_step(self, arr, step):
        return self.viewer.filter_controller._run_filter_step(arr, step)

    def filter_label(self, steps):
        try:
            return self.viewer.filter_controller._filter_pipeline_label_from_steps(steps, default="")
        except Exception:
            return ", ".join(str(s.get("key")) for s in steps or [])

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
        return self.open_custom(keys)

    def known_keys(self, keys):
        """``{given: image key}`` for the given paths that are loaded images
        (drag payloads, dropped files - a path spelled with / still maps)."""
        headers = self.viewer.headers or {}
        out = {}
        for k in keys:
            k = str(k)
            if k in headers:
                out[k] = k
                continue
            try:
                norm = str(Path(k))
            except Exception:
                continue
            if norm in headers:
                out[k] = norm
        return out

    def previewed_key(self):
        last = getattr(self.viewer, "last_preview", None)
        key = str(last[0]) if last else ""
        return key if key in (self.viewer.headers or {}) else None

    def channel_label(self, key, idx):
        """Caption of channel ``idx`` of an image (how figure rows name it)."""
        _header, fds = (self.viewer.headers or {}).get(str(key), (None, None))
        try:
            fd = list(fds or [])[int(idx)]
        except Exception:
            return None
        return str(fd.get("Caption") or fd.get("FileName") or f"Channel {idx}")

    def selection_keys(self):
        """The thumbnail selection in grid order, else the previewed image."""
        from ..viewer import export as viewer_export
        try:
            keys = viewer_export.export_targets(self.viewer)
        except Exception:
            keys = []
        return [k for k in keys if k in (self.viewer.headers or {})]

    def open_custom(self, keys, name=None):
        """A new custom figure of ``keys``, in that order. Its id stays the
        same whatever frames are added, removed or reordered later."""
        keys = list(dict.fromkeys(str(k) for k in keys if str(k) in (self.viewer.headers or {})))
        if not keys:
            return None
        fid = "custom-" + uuid.uuid4().hex[:10]
        self.set_state(fid, {"mode": "custom", "keys": keys, "name": name or self._default_name()})
        return self._open(fid, keys, mode="custom")

    def _default_name(self):
        taken = {str(st.get("name")) for st in (self.viewer.figure_states or {}).values()
                 if isinstance(st, dict)}
        n = 1
        while f"Figure {n}" in taken:
            n += 1
        return f"Figure {n}"

    def recent_custom_figures(self, limit=_RECENT_LIMIT):
        """``[{fid, name, keys, missing}]``, most recently edited first;
        ``missing`` = frames whose file is not loaded now."""
        loaded = self.viewer.headers or {}
        out = []
        for fid, st in reversed(list((self.viewer.figure_states or {}).items())):
            if not str(fid).startswith("custom-") or not isinstance(st, dict):
                continue
            keys = [str(k) for k in st.get("keys") or []]
            if not keys:
                continue
            out.append({"fid": fid, "name": str(st.get("name") or "Figure"), "keys": keys,
                        "missing": [k for k in keys if k not in loaded]})
            if len(out) >= limit:
                break
        return out

    def open_saved(self, fid):
        """Reopen a remembered figure (custom or series) by its id."""
        st = self.get_state(fid)
        keys = [str(k) for k in st.get("keys") or []]
        if not keys:
            return None
        missing = [k for k in keys if k not in (self.viewer.headers or {})]
        if missing and not self._load_missing(missing):
            return None
        return self._open(fid, keys, mode=st.get("mode") or "custom", root_key=st.get("root_key"))

    def _load_missing(self, missing):
        """Offer to load the folder holding figure frames that are not
        loaded. True when every frame is available afterwards."""
        folders = sorted({str(Path(k).parent) for k in missing})
        names = ", ".join(Path(k).name for k in missing[:4]) + (" ..." if len(missing) > 4 else "")
        if len(folders) != 1 or not Path(folders[0]).is_dir():
            QtWidgets.QMessageBox.information(
                self.viewer, "Figure",
                f"{len(missing)} image(s) of this figure are not loaded ({names}) and cannot be "
                "loaded together:\n" + "\n".join(folders))
            return False
        ans = QtWidgets.QMessageBox.question(
            self.viewer, "Figure",
            f"{len(missing)} image(s) of this figure are not loaded ({names}).\n\n"
            f"Load their folder?\n{folders[0]}")
        if ans != QtWidgets.QMessageBox.Yes:
            return False
        try:
            self.viewer.load_folder(Path(folders[0]))
        except Exception as exc:
            log_status(f"[Figures] could not load {folders[0]}: {exc}")
            return False
        still = [k for k in missing if k not in (self.viewer.headers or {})]
        if still:
            QtWidgets.QMessageBox.information(self.viewer, "Figure",
                                              f"{len(still)} image(s) are still missing.")
        return not still

    # Figure files (*.sxmfig.json): one figure - its frames (absolute
    # paths), recipe and every edit - to keep next to the data or share.
    def save_figure_file(self, fid, path):
        st = self.get_state(fid)
        payload = {"kind": FIGURE_FILE_KIND, "version": 1, "fid": fid,
                   "mode": st.get("mode") or "custom", "keys": list(st.get("keys") or []),
                   "root_key": st.get("root_key"), "state": st}
        Path(path).write_text(json.dumps(payload, indent=1, default=str), encoding="utf-8")

    def open_figure_file(self, path):
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self.viewer, "Figure", f"Could not read the figure file:\n{exc}")
            return None
        if not isinstance(payload, dict) or payload.get("kind") != FIGURE_FILE_KIND:
            QtWidgets.QMessageBox.warning(self.viewer, "Figure", "Not an SXM Viewer figure file.")
            return None
        st = dict(payload.get("state") or {})
        st["keys"] = [str(k) for k in payload.get("keys") or st.get("keys") or []]
        st["mode"] = payload.get("mode") or st.get("mode") or "custom"
        if payload.get("root_key"):
            st["root_key"] = payload["root_key"]
        fid = str(payload.get("fid") or ("custom-" + uuid.uuid4().hex[:10]))
        win = self._windows.pop(fid, None)
        if win is not None:
            try:
                win.close()
            except RuntimeError:
                pass
        self.set_state(fid, st)   # the file wins over a remembered copy
        return self.open_saved(fid)

    def _open(self, fid, keys, mode, excluded=None, root_key=None):
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
            win = FigureWindow(self, fid, keys, mode=mode, excluded=excluded, root_key=root_key,
                               parent=self.viewer)
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

    # The user's format is remembered for new figures: letters, fonts,
    # outlines, page ... globally; the recipe-dependent keys
    # (hs.RECIPE_KEYS: ranges, colourbars, scale bar, header, layout) per
    # recipe, on top of that recipe's own defaults (hs.RECIPE_DEFAULTS).
    def default_style(self, recipe=None):
        cfg = self.viewer.config
        shared = cfg.get("figure_style_default")
        style = dict(hs.RECIPE_DEFAULTS.get(recipe or "", {}))
        if isinstance(shared, dict):
            style.update({k: v for k, v in shared.items() if k not in hs.RECIPE_KEYS})
        per = (cfg.get("figure_style_recipe") or {}).get(recipe or "")
        if isinstance(per, dict):
            style.update({k: v for k, v in per.items() if k in hs.RECIPE_KEYS and k != "header"})
        return style

    def set_default_style(self, style, recipe=None):
        style = dict(style or {})
        cfg = self.viewer.config
        if not style:
            cfg["figure_style_default"] = {}
            cfg["figure_style_recipe"] = {}
        else:
            cfg["figure_style_default"] = {k: v for k, v in style.items() if k not in hs.RECIPE_KEYS}
            per = dict(cfg.get("figure_style_recipe") or {})
            # The header is per recipe ("auto" = dz / V / size) and never
            # carried over; the other recipe keys are remembered per recipe.
            per[recipe or ""] = {k: v for k, v in style.items() if k in hs.RECIPE_KEYS and k != "header"}
            cfg["figure_style_recipe"] = per
        try:
            save_config(cfg)
        except Exception:
            pass

    # Named formats ("My Nature style", "Group style"): a full style dict
    # each, in config ``figure_named_styles``; files share them in a group.
    STYLE_FILE_KIND = "sxm_figure_styles"

    def named_styles(self):
        styles = self.viewer.config.get("figure_named_styles")
        return {str(k): dict(v) for k, v in (styles or {}).items() if isinstance(v, dict)}

    def save_named_style(self, name, style):
        styles = self.named_styles()
        styles[str(name)] = {k: v for k, v in dict(style).items() if not str(k).startswith("_")}
        self.viewer.config["figure_named_styles"] = styles
        self._save()

    def delete_named_style(self, name):
        styles = self.named_styles()
        styles.pop(str(name), None)
        self.viewer.config["figure_named_styles"] = styles
        self._save()

    def export_named_styles(self, path, names=None):
        styles = self.named_styles()
        if names:
            styles = {k: v for k, v in styles.items() if k in names}
        Path(path).write_text(json.dumps({"kind": self.STYLE_FILE_KIND, "version": 1, "styles": styles},
                                         indent=1), encoding="utf-8")
        return len(styles)

    def import_named_styles(self, path):
        """Merge the styles of a file (same name = replaced). Returns names."""
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        if not isinstance(payload, dict) or payload.get("kind") != self.STYLE_FILE_KIND:
            raise ValueError("not an SXM Viewer figure-styles file")
        new = {str(k): dict(v) for k, v in (payload.get("styles") or {}).items() if isinstance(v, dict)}
        styles = self.named_styles()
        styles.update(new)
        self.viewer.config["figure_named_styles"] = styles
        self._save()
        return list(new)

    def other_figure_styles(self, exclude_fid=None, limit=12):
        """``[(title, style)]`` of open figures first, then remembered ones -
        for "Copy format from"."""
        out, seen = [], set()
        for fid, win in list(self._windows.items()):
            if fid == exclude_fid:
                continue
            try:
                title = win.windowTitle()
                out.append((f"{title}  (open)", dict(win.state.get("style") or {})))
                seen.add(fid)
            except RuntimeError:
                continue
        for fid, st in reversed(list((self.viewer.figure_states or {}).items())):
            if fid == exclude_fid or fid in seen or not isinstance(st, dict) or not st.get("style"):
                continue
            title = st.get("name") or {"height": "Height series", "bias": "Bias series",
                                       "zoom": "Overview + zooms"}.get(st.get("mode"), "Figure")
            keys = st.get("keys") or []
            if keys:
                title += f"  ({Path(str(keys[0])).stem}{' …' if len(keys) > 1 else ''})"
            out.append((title, dict(st["style"])))
            if len(out) >= limit:
                break
        return out

    def _save(self):
        try:
            save_config(self.viewer.config)
        except Exception:
            pass

    # Export preferences (dpi, transparent background, PowerPoint mode, ...)
    def export_settings(self):
        st = self.viewer.config.get("figure_export")
        return dict(st) if isinstance(st, dict) else {}

    def set_export_settings(self, **values):
        st = self.export_settings()
        st.update(values)
        self.viewer.config["figure_export"] = st
        try:
            save_config(self.viewer.config)
        except Exception:
            pass

    def obsidian_menu(self, menu, render, parent=None):
        """One-click "append to the active note" entries + "Choose note..."
        for a figure. ``render(img_dir) -> (results, errors, cancelled)``."""
        from ..viewer import export as viewer_export
        try:
            from ...integrations import obsidian
        except Exception:
            menu.addAction("Obsidian support unavailable").setEnabled(False)
            return
        settings = viewer_export._obsidian_settings(self.viewer)
        vaults = obsidian.list_vaults()
        shown = [v for v in vaults if v.is_open] or vaults[:1]
        added = set()
        for vault in shown:
            note = obsidian.active_note(vault.path)
            if note is None:
                continue
            rel = note.relative_to(vault.path).as_posix()
            act = menu.addAction(f"Active note: {rel}  ({vault.name})")
            act.triggered.connect(lambda _=False, n=note, v=vault.path: viewer_export.export_to_obsidian(
                self.viewer, n, v, render=render, parent=parent))
            added.add(str(note))
        last = settings.get("note")
        if last and last not in added and Path(last).is_file():
            act = menu.addAction(f"Last used: {Path(last).name}")
            act.triggered.connect(lambda _=False, n=Path(last), v=settings.get("vault"):
                                  viewer_export.export_to_obsidian(self.viewer, n, v, render=render,
                                                                   parent=parent))
        if not menu.isEmpty():
            menu.addSeparator()
        menu.addAction("Choose note / options…").triggered.connect(
            lambda _=False: viewer_export.export_to_obsidian(self.viewer, None, None, render=render,
                                                             parent=parent))

    def cmap_prefs(self):
        """Colormap favourites / recent (the viewer's live config)."""
        from ..cmap_menu import CmapPrefs
        return CmapPrefs(self.viewer.config, save_config)

    # ------------------------------------------------------------------
    # Molecular models (shared with the preview / pop-ups)
    # ------------------------------------------------------------------
    def molecule_settings(self):
        """The preview's molecule look (palette, bonds, shadows, H), so a
        figure's model matches what the user tuned in the preview."""
        pc = getattr(self.viewer, "preview_canvas", None)
        return {
            "palette": str(getattr(pc, "molecule_palette", "avogadro") or "avogadro"),
            "bond_color_mode": getattr(pc, "_bond_color_mode", "default"),
            "default_bond_color": getattr(pc, "_default_bond_color", (0.9, 0.9, 0.9)),
            "show_shadows": bool(getattr(pc, "_show_molecule_shadow", False)),
            "show_hydrogens": bool(getattr(pc, "_show_hydrogens", True)),
        }

    def _frame_origin_nm(self, key):
        """Lower-left corner (display data coords) of a frame as the preview
        draws it - subtracting it maps preview molecule positions into a
        figure panel's local nm (both have y up)."""
        viewer = self.viewer
        header, _fds = (viewer.headers or {}).get(str(key), (None, None))
        if not header:
            return 0.0, 0.0
        try:
            base = viewer._header_extent(header)
            _arr, ext = viewer._apply_adjustments_for_channel(str(key), 0, np.zeros((2, 2)), base)
            ext = viewer._display_extent(ext, header)
            return min(float(ext[0]), float(ext[1])), min(float(ext[2]), float(ext[3]))
        except Exception:
            return 0.0, 0.0

    def preview_molecules(self, keys):
        """Molecules already placed on one of these frames in the preview
        (live preview first, then the per-file overlays), as dicts in panel
        nm. Returns ``(molecules, source_key)``."""
        viewer = self.viewer
        keys = [str(k) for k in keys]
        found, src = [], None
        last = getattr(viewer, "last_preview", None)
        pc = getattr(viewer, "preview_canvas", None)
        if last and str(last[0]) in keys and getattr(pc, "molecules", None):
            found = [m.to_dict() for m in pc.molecules]
            src = str(last[0])
        if not found:
            for k in keys:
                mols = (getattr(viewer, "molecule_overlays", {}) or {}).get(k)
                if mols:
                    found = [dict(m) if isinstance(m, dict) else m.to_dict() for m in mols]
                    src = k
                    break
        if not found:
            return [], None
        ox, oy = self._frame_origin_nm(src)
        out = []
        for m in found:
            m = dict(m)
            off = list(m.get("offset") or [0.0, 0.0, 0.0])
            off[0] = float(off[0]) - ox
            off[1] = float(off[1]) - oy
            m["offset"] = off
            out.append(m)
        return out, src

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

