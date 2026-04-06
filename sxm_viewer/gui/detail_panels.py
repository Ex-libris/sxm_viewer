"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

# Re-export shims for backward-compatible imports.
from .detail_panels_export import BatchExportSignals, BatchExportWorker
from .detail_panels_filters import CustomFilterDialog, SingleFilterDialog
from .detail_panels_image import ImageAdjustPreviewPanel, ImageAdjustDialog
from .detail_panels_matrix import MatrixFitWorker, MatrixFitDialog
from .detail_panels_preview import MultiPreviewCanvas, SafeFigureCanvas
from .detail_panels_profile import ProfileDialog
from .detail_panels_spectro import (
    SpectroscopyPopup,
    MatrixSpectroViewer,
    _SpectroFitWorker,
    SpectroscopyCompareDialog,
)

__all__ = [
    "MultiPreviewCanvas",
    "SafeFigureCanvas",
    "ProfileDialog",
    "SpectroscopyPopup",
    "MatrixSpectroViewer",
    "_SpectroFitWorker",
    "SpectroscopyCompareDialog",
    "MatrixFitWorker",
    "MatrixFitDialog",
    "CustomFilterDialog",
    "SingleFilterDialog",
    "ImageAdjustPreviewPanel",
    "ImageAdjustDialog",
    "BatchExportSignals",
    "BatchExportWorker",
]



