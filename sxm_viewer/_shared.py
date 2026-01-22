"""Shared imports and utility helpers for the modular SXM viewer."""
from __future__ import annotations

import io
import itertools
import json
import math
import os
import sys
import threading
from collections import OrderedDict, defaultdict
from datetime import datetime
from pathlib import Path

import hashlib
import matplotlib
import numpy as np
from matplotlib import colormaps
from matplotlib.backends import backend_qt5agg as _backend_qt5agg
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QBrush, QIcon, QImage, QPainter, QPen, QPixmap

matplotlib.use("Agg")

try:
    from scipy import ndimage as _scipy_ndimage
except Exception:  # pragma: no cover - optional dependency
    _scipy_ndimage = None


class LogEmitter(QtCore.QObject):
    message_logged = QtCore.pyqtSignal(str)


log_emitter = LogEmitter()

_orig_resize_event = FigureCanvas.resizeEvent


def _safe_resize_event(self, event):
    try:
        size = event.size()
    except Exception:
        return _orig_resize_event(self, event)
    safe_w = max(1, int(size.width()))
    safe_h = max(1, int(size.height()))
    if not math.isfinite(safe_w):
        safe_w = 1
    if not math.isfinite(safe_h):
        safe_h = 1
    if safe_w != size.width() or safe_h != size.height():
        event = QtGui.QResizeEvent(QtCore.QSize(safe_w, safe_h), event.oldSize())
    try:
        return _orig_resize_event(self, event)
    except ValueError:
        fallback = QtGui.QResizeEvent(QtCore.QSize(max(10, safe_w), max(10, safe_h)), event.oldSize())
        try:
            return _orig_resize_event(self, fallback)
        except ValueError:
            return


FigureCanvas.resizeEvent = _safe_resize_event
_backend_qt5agg.FigureCanvasQTAgg.resizeEvent = _safe_resize_event


def log_status(message: str):
    """Emit startup/progress info to the terminal."""
    try:
        print(f"[SXMViewer] {message}", flush=True)
    except Exception:
        pass
    try:
        log_emitter.message_logged.emit(str(message))
    except Exception:
        pass


__all__ = [
    "QtWidgets",
    "QtCore",
    "QtGui",
    "QIcon",
    "QPixmap",
    "QImage",
    "QPainter",
    "QPen",
    "QBrush",
    "FigureCanvas",
    "Figure",
    "Line2D",
    "colormaps",
    "np",
    "Path",
    "defaultdict",
    "OrderedDict",
    "datetime",
    "hashlib",
    "itertools",
    "io",
    "json",
    "math",
    "os",
    "sys",
    "threading",
    "_scipy_ndimage",
    "log_status",
    "log_emitter",
    "matplotlib",
]



