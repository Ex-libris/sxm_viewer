"""Configuration persistence helpers for the SXM viewer."""
from __future__ import annotations

import json

from .config_defaults import CONFIG_PATH, HEADER_CACHE_PATH, HEADER_CACHE_VERSION


def load_config():
    """Load persisted viewer configuration from disk."""
    try:
        s = CONFIG_PATH.read_text()
        return json.loads(s)
    except Exception:
        return {}


def save_config(cfg):
    """Persist configuration dictionary to disk."""
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2))
    except Exception:
        pass


def load_header_cache():
    """Load cached headers parsed in previous sessions."""
    try:
        s = HEADER_CACHE_PATH.read_text()
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
        HEADER_CACHE_PATH.write_text(json.dumps(payload))
    except Exception:
        pass


__all__ = [
    "load_config",
    "save_config",
    "load_header_cache",
    "save_header_cache",
]



