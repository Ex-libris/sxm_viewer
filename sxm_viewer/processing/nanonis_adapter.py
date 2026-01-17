"""Compatibility shim. The Nanonis adapter now lives under
``sxm_viewer.providers.nanonis`` to decouple parsers from GUI logic.
"""
from __future__ import annotations

from sxm_viewer.providers.nanonis import (  # type: ignore F401
    parse_nanonis_spectroscopy,
    prepare_nanonis_folder,
)

__all__ = ["prepare_nanonis_folder", "parse_nanonis_spectroscopy"]
