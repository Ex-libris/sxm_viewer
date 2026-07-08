"""Configuration persistence helpers for the SXM viewer."""
from __future__ import annotations

import copy
import json

from PyQt5 import QtCore

from .config_defaults import (
    CONFIG_PATH,
    HEADER_CACHE_PATH,
    HEADER_CACHE_VERSION,
    COLLECTIONS_INDEX_PATH,
    COLLECTIONS_INDEX_VERSION,
)

# save_config is called from 90+ GUI call sites (essentially every UI toggle:
# thumb sort/filter, tags, starring, ...), each historically doing a
# synchronous ~50ms full rewrite of the whole config dict (dominated by
# `tags`/`starred`, measured ~787KB on a real profile) on the GUI thread with
# no debounce. Debounced here the same way the spectro manifest already is
# (see main_window.py's _schedule_spectro_manifest_save): the write is
# delayed and coalesced, but the LATEST value is held in-memory immediately
# so load_config() never returns stale data even before the delayed write
# lands. That matters concretely: gui/dialogs/spectroscopy_dialogs.py and
# gui/canvases/detail_preview_canvas.py both do a direct
# read-modify-write against this file (load_config() -> mutate one key ->
# save_config()) independent of the viewer's own in-memory config - a naive
# "just delay the write" debounce would let those read stale disk data and
# clobber a pending change once the delayed write eventually landed.
_pending_cfg = None
_save_timer = None


def _flush_pending_write():
    global _pending_cfg
    if _pending_cfg is None:
        return
    cfg = _pending_cfg
    _pending_cfg = None
    try:
        # Not meant to be human-read; measured 5x faster to dump compact vs
        # pretty-printed on this codebase's other big JSON caches.
        CONFIG_PATH.write_text(json.dumps(cfg, separators=(',', ':')), encoding="utf-8")
    except Exception:
        pass


def flush_pending_config_save():
    """Force any debounced save_config() write to happen immediately.
    Call before the app exits so a change right before quitting isn't
    silently dropped (this exact failure mode already happened once for the
    similarly-debounced spectro manifest before its own close-time flush was
    added)."""
    global _save_timer
    if _save_timer is not None:
        _save_timer.stop()
    _flush_pending_write()


def load_config():
    """Load persisted viewer configuration from disk - or the latest
    not-yet-flushed value if save_config() ran more recently than the
    debounced write has landed. Returns an independent copy either way, so
    callers mutating the result (a common pattern: load_config(), set one
    key, save_config()) never accidentally mutate the pending/shared state
    out from under a concurrent caller."""
    if _pending_cfg is not None:
        return copy.deepcopy(_pending_cfg)
    try:
        s = CONFIG_PATH.read_text(encoding="utf-8")
        return json.loads(s)
    except Exception:
        return {}


def save_config(cfg):
    """Persist configuration dictionary to disk, debounced ~400ms."""
    global _pending_cfg, _save_timer
    _pending_cfg = cfg
    if _save_timer is None:
        _save_timer = QtCore.QTimer()
        _save_timer.setSingleShot(True)
        _save_timer.timeout.connect(_flush_pending_write)
    _save_timer.start(400)


def load_header_cache():
    """Load cached headers parsed in previous sessions - or the latest
    not-yet-flushed value if save_header_cache() ran more recently than the
    debounced write has landed (see save_header_cache's docstring)."""
    if _pending_header_cache is not None:
        return copy.deepcopy(_pending_header_cache)
    try:
        # Explicit encoding matters here beyond correctness: without it,
        # read_text()/write_text() fall back to the locale-default encoding
        # (cp1252 on typical Windows setups), whose generic codec is slower
        # to decode than UTF-8's optimized ASCII/UTF-8 fast path in CPython —
        # relevant given this file accumulates every folder ever opened and
        # can reach tens of MB.
        s = HEADER_CACHE_PATH.read_text(encoding="utf-8")
        data = json.loads(s)
        if not isinstance(data, dict):
            return {}
        if data.get("_version") != HEADER_CACHE_VERSION:
            return {}
        return data.get("entries", {})
    except Exception:
        return {}


_pending_header_cache = None
_header_cache_save_timer = None


def _flush_pending_header_cache_write():
    global _pending_header_cache
    if _pending_header_cache is None:
        return
    cache = _pending_header_cache
    _pending_header_cache = None
    try:
        payload = {"_version": HEADER_CACHE_VERSION, "entries": cache}
        HEADER_CACHE_PATH.write_text(json.dumps(payload), encoding="utf-8")
    except Exception:
        pass


def flush_pending_header_cache_save():
    """Force any debounced save_header_cache() write to happen immediately -
    call before the app exits, same reasoning as flush_pending_config_save()."""
    global _header_cache_save_timer
    if _header_cache_save_timer is not None:
        _header_cache_save_timer.stop()
    _flush_pending_header_cache_write()


def save_header_cache(cache):
    """Persist header cache (used to speed up future loads), debounced
    ~400ms - this file accumulates every folder ever opened and can reach
    tens of MB, so a synchronous rewrite (confirmed 630ms on a real 34MB/
    2971-entry cache) is worth keeping off the immediate call path,
    especially if the user opens several folders in quick succession. Same
    read-through design as save_config/load_config (see save_config's
    docstring) even though load_header_cache currently has only one,
    startup-time caller - keeps both header-cache functions safe against a
    future mid-session caller being added without anyone re-auditing this."""
    global _pending_header_cache, _header_cache_save_timer
    _pending_header_cache = cache
    if _header_cache_save_timer is None:
        _header_cache_save_timer = QtCore.QTimer()
        _header_cache_save_timer.setSingleShot(True)
        _header_cache_save_timer.timeout.connect(_flush_pending_header_cache_write)
    _header_cache_save_timer.start(400)


def load_collections_index():
    """Load the folder -> collections usage index (which collections reference files from which
    folder). Global file, mirrors the header-cache pattern - avoids writing into user data
    folders and keeps folder-keyed lookups simple."""
    try:
        s = COLLECTIONS_INDEX_PATH.read_text(encoding="utf-8")
        data = json.loads(s)
        if not isinstance(data, dict):
            return {}
        if data.get("_version") != COLLECTIONS_INDEX_VERSION:
            return {}
        return data.get("folders", {})
    except Exception:
        return {}


def save_collections_index(index):
    """Persist the folder -> collections usage index."""
    try:
        payload = {"_version": COLLECTIONS_INDEX_VERSION, "folders": index}
        COLLECTIONS_INDEX_PATH.write_text(json.dumps(payload, separators=(',', ':')), encoding="utf-8")
    except Exception:
        pass


__all__ = [
    "load_config",
    "save_config",
    "flush_pending_config_save",
    "load_header_cache",
    "save_header_cache",
    "flush_pending_header_cache_save",
    "load_collections_index",
    "save_collections_index",
]



