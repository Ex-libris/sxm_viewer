"""Provider registry for format-specific loaders."""
from __future__ import annotations

from pathlib import Path
from typing import Iterable, List

# Expose Nanonis provider
from .nanonis import prepare_nanonis_folder, prepare_nanonis_files, parse_nanonis_spectroscopy, parse_nanonis_3ds  # noqa: F401
from .omicronscala import prepare_scala_folder, prepare_scala_files  # noqa: F401
from .rhksm4 import prepare_sm4_folder, prepare_sm4_files  # noqa: F401


def convert_nanonis(folder: Path | str) -> List[Path]:
    """Convert Nanonis scans within ``folder`` into viewer-compatible headers."""
    return prepare_nanonis_folder(folder)


def convert_nanonis_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit Nanonis scan files into viewer-compatible headers."""
    return prepare_nanonis_files(paths)


def convert_omicronscala(folder: Path | str) -> List[Path]:
    """Convert Omicron SCALA (.par) exports within ``folder`` into headers."""
    return prepare_scala_folder(folder)


def convert_omicronscala_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit Omicron SCALA (.par) files into headers."""
    return prepare_scala_files(paths)


def convert_rhksm4(folder: Path | str) -> List[Path]:
    """Convert RHK SM4 (.sm4) exports within ``folder`` into headers."""
    return prepare_sm4_folder(folder)


def convert_rhksm4_files(paths: Iterable[Path | str]) -> List[Path]:
    """Convert explicit RHK SM4 (.sm4) files into headers."""
    return prepare_sm4_files(paths)


__all__: Iterable[str] = [
    "convert_nanonis",
    "convert_nanonis_files",
    "convert_omicronscala",
    "convert_omicronscala_files",
    "convert_rhksm4",
    "convert_rhksm4_files",
    "prepare_nanonis_folder",
    "prepare_nanonis_files",
    "parse_nanonis_spectroscopy",
    "parse_nanonis_3ds",
    "prepare_scala_folder",
    "prepare_scala_files",
    "prepare_sm4_folder",
    "prepare_sm4_files",
]
