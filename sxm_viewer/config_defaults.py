"""Default configuration constants for the SXM viewer."""
from __future__ import annotations

from pathlib import Path

CONFIG_PATH = Path.home() / ".sxm_viewer_config.json"
HEADER_CACHE_PATH = Path.home() / ".sxm_viewer_header_cache.json"
HEADER_CACHE_VERSION = 2
CH_EQUALITY_TOL_NM = 0.001    # 1 pm tolerance for "flat" topo samples
CH_SAMPLE_POINTS = 16         # number of points to probe when classifying CH/CC
CHANNEL_DATA_CACHE_LIMIT = 24  # max channel arrays cached in-memory
FILTERED_CACHE_LIMIT = 32      # max filtered arrays cached in-memory
THUMB_DISK_CACHE_DIR = Path.home() / ".sxm_thumb_cache"

__all__ = [
    "CONFIG_PATH",
    "HEADER_CACHE_PATH",
    "HEADER_CACHE_VERSION",
    "CH_EQUALITY_TOL_NM",
    "CH_SAMPLE_POINTS",
    "CHANNEL_DATA_CACHE_LIMIT",
    "FILTERED_CACHE_LIMIT",
    "THUMB_DISK_CACHE_DIR",
]



