"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

# Re-export shims for backward-compatible imports.
from .canvases.multi_preview import MultiPreviewCanvas, SafeFigureCanvas
from .dialogs.profile import ProfileDialog
from .dialogs.spectroscopy import (
    SpectroscopyPopup,
    MatrixSpectroViewer,
    _SpectroFitWorker,
    SpectroscopyCompareDialog,
)
from .dialogs.matrix_fit import MatrixFitWorker, MatrixFitDialog
from .dialogs.filters import CustomFilterDialog
from .dialogs.image_adjust import ImageAdjustPreviewPanel, ImageAdjustDialog
from .workers.batch_export import BatchExportSignals, BatchExportWorker

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
    "ImageAdjustPreviewPanel",
    "ImageAdjustDialog",
    "BatchExportSignals",
    "BatchExportWorker",
]