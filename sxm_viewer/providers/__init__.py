"""Provider registry for format-specific loaders."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, List

# Expose Nanonis provider
from .nanonis import prepare_nanonis_folder, parse_nanonis_spectroscopy  # noqa: F401


def convert_nanonis(folder: Path | str) -> List[Path]:
    """Convert Nanonis scans within ``folder`` into viewer-compatible headers."""
    return prepare_nanonis_folder(folder)


__all__: Iterable[str] = [
    "convert_nanonis",
    "prepare_nanonis_folder",
    "parse_nanonis_spectroscopy",
]
