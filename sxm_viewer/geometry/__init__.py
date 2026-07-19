"""Qt-free coordinate geometry shared by GUI, reporting and providers.

Like ``providers/`` and ``reporting/``, nothing here may import Qt or any
``gui`` module - that is what makes this logic unit-testable without a
display, and it is enforced by ``scripts/analysis`` parse checks and the
layering rule in CLAUDE.md.

Modules:

* ``spec_mapping`` - absolute nm <-> image raster pixel transforms for
  spectroscopy positions. Previously methods on ``SXMGridViewer``.
"""
from .spec_mapping import (  # noqa: F401
    header_extent,
    header_scan_angle,
    fallback_spec_coords,
    apply_thumb_crop,
    map_spec_by_grid,
    map_spec_by_cloud_bounds,
    map_spec_to_pixels,
)

__all__ = [
    "header_extent",
    "header_scan_angle",
    "fallback_spec_coords",
    "apply_thumb_crop",
    "map_spec_by_grid",
    "map_spec_by_cloud_bounds",
    "map_spec_to_pixels",
]
