"""Configuration persistence helpers for the SXM viewer."""
from __future__ import annotations

import json

from .config_defaults import (
    CONFIG_PATH,
    HEADER_CACHE_PATH,
    HEADER_CACHE_VERSION,
    COLLECTIONS_INDEX_PATH,
    COLLECTIONS_INDEX_VERSION,
)


def load_config():
    """Load persisted viewer configuration from disk."""
    try:
        s = CONFIG_PATH.read_text(encoding="utf-8")
        return json.loads(s)
    except Exception:
        return {}


def save_config(cfg):
    """Persist configuration dictionary to disk."""
    try:
        # Not meant to be human-read; measured 5x faster to dump compact vs
        # pretty-printed on this codebase's other big JSON caches, and this
        # file (dominated by `tags`/`starred`) gets rewritten in full on
        # essentially every UI toggle.
        CONFIG_PATH.write_text(json.dumps(cfg, separators=(',', ':')), encoding="utf-8")
    except Exception:
        pass


def load_header_cache():
    """Load cached headers parsed in previous sessions."""
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


def save_header_cache(cache):
    """Persist header cache (used to speed up future loads)."""
    try:
        payload = {"_version": HEADER_CACHE_VERSION, "entries": cache}
        HEADER_CACHE_PATH.write_text(json.dumps(payload), encoding="utf-8")
    except Exception:
        pass


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
    "load_header_cache",
    "save_header_cache",
    "load_collections_index",
    "save_collections_index",
]



