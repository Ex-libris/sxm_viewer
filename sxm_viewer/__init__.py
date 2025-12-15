"""Top-level package for the modular SXM viewer."""
from __future__ import annotations

from .cli import main
from .gui.main_window import SXMGridViewer

__all__ = ["main", "SXMGridViewer"]
__version__ = "0.1.0"
