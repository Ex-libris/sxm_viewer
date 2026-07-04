"""Configuration persistence helpers for the SXM viewer."""
from __future__ import annotations

import json

from .config_defaults import CONFIG_PATH, HEADER_CACHE_PATH, HEADER_CACHE_VERSION


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
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
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


__all__ = [
    "load_config",
    "save_config",
    "load_header_cache",
    "save_header_cache",
]



