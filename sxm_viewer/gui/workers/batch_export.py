"""Detail canvases and spectroscopy dialogs."""
from __future__ import annotations

import itertools
import json
import math

import numpy as np
from matplotlib import patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter
from matplotlib.widgets import RectangleSelector
from mpl_toolkits.axes_grid1.anchored_artists import AnchoredSizeBar
from mpl_toolkits.axes_grid1.inset_locator import inset_axes

from ..._shared import *
from ...config import *
from ...data.io import *
from ...data.spectroscopy import *
from ..thumbnails import *

class BatchExportSignals(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int, str)
    finished = QtCore.pyqtSignal(list, list, bool)

class BatchExportWorker(QtCore.QRunnable):
    def __init__(self, parent, paths, config, out_dir):
        super().__init__()
        self.parent = parent
        self.paths = [str(p) for p in paths]
        self.config = config
        self.out_dir = Path(out_dir)
        self.signals = BatchExportSignals()
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        saved = []
        errors = []
        total = len(self.paths)
        for idx, path in enumerate(self.paths, 1):
            if self._cancelled:
                break
            try:
                result = self.parent.render_and_save_file_using_config(Path(path), self.config, self.out_dir)
                saved.extend(result)
            except Exception as e:
                errors.append(f"{Path(path).name}: {e}")
            self.signals.progress.emit(idx, total, path)
        self.signals.finished.emit(saved, errors, self._cancelled)
