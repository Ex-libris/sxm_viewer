"""Configuration persistence and cache constants for the SXM viewer."""
from __future__ import annotations

import json
from pathlib import Path


CONFIG_PATH = Path.home() / ".sxm_viewer_config.json"
HEADER_CACHE_PATH = Path.home() / ".sxm_viewer_header_cache.json"
CH_EQUALITY_TOL_NM = 0.001    # 1 pm tolerance for "flat" topo samples
CH_SAMPLE_POINTS = 16         # number of points to probe when classifying CH/CC
CHANNEL_DATA_CACHE_LIMIT = 24  # max channel arrays cached in-memory
FILTERED_CACHE_LIMIT = 32      # max filtered arrays cached in-memory
THUMB_DISK_CACHE_DIR = Path.home() / ".sxm_thumb_cache"

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
        return json.loads(s)
    except Exception:
        return {}

def save_header_cache(cache):
    """Persist header cache (used to speed up future loads)."""
    try:
        HEADER_CACHE_PATH.write_text(json.dumps(cache))
    except Exception:
        pass


__all__ = [
    "CONFIG_PATH",
    "HEADER_CACHE_PATH",
    "CH_EQUALITY_TOL_NM",
    "CH_SAMPLE_POINTS",
    "CHANNEL_DATA_CACHE_LIMIT",
    "FILTERED_CACHE_LIMIT",
    "THUMB_DISK_CACHE_DIR",
    "load_config",
    "save_config",
    "load_header_cache",
    "save_header_cache",
]
