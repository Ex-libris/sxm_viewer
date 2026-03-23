"""Provider registry for format-specific loaders."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, List

# Expose vendor providers
from .nanonis import prepare_nanonis_folder, prepare_nanonis_files, parse_nanonis_spectroscopy, parse_nanonis_3ds  # noqa: F401
from .mtrx import classify_mtrx_file, is_mtrx_file, parse_mtrx_spectroscopy, prepare_mtrx_files, prepare_mtrx_folder  # noqa: F401


def convert_nanonis(folder: Path | str) -> List[Path]:
    """Convert Nanonis scans within ``folder`` into viewer-compatible headers."""
    return prepare_nanonis_folder(folder)


def convert_nanonis_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit Nanonis scan files into viewer-compatible headers."""
    return prepare_nanonis_files(paths)


def convert_mtrx(folder: Path | str) -> List[Path]:
    """Convert MATRIX image-like scans within ``folder`` into viewer-compatible headers."""
    return prepare_mtrx_folder(folder)


def convert_mtrx_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit MATRIX image files into viewer-compatible headers."""
    return prepare_mtrx_files(paths)


__all__: Iterable[str] = [
    "convert_nanonis",
    "convert_nanonis_files",
    "convert_mtrx",
    "convert_mtrx_files",
    "prepare_nanonis_folder",
    "prepare_nanonis_files",
    "prepare_mtrx_folder",
    "prepare_mtrx_files",
    "parse_nanonis_spectroscopy",
    "parse_nanonis_3ds",
    "classify_mtrx_file",
    "is_mtrx_file",
    "parse_mtrx_spectroscopy",
]
