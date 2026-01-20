"""
Data helpers for the SXM viewer.

The GUI expects ``sxm_viewer.data.io`` and ``sxm_viewer.data.spectroscopy`` to
be importable.  This package simply forwards the public helpers so callers can
``from sxm_viewer.data import parse_header`` without caring about submodules.
"""
from .io import normalize_unit_and_data, parse_header, read_channel_file
from .spectroscopy import (
    find_last_image_for_spec,
    fit_parabola_bias,
    parse_spectroscopy_file,
    SpectroscopyParseError,
    MatrixDatError,
    _matrix_base_name,
)

__all__ = [
    "parse_header",
    "read_channel_file",
    "normalize_unit_and_data",
    "parse_spectroscopy_file",
    "SpectroscopyParseError",
    "MatrixDatError",
    "fit_parabola_bias",
    "find_last_image_for_spec",
    "_matrix_base_name",
]



