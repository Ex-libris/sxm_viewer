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


def log_status(message: str):
    """Emit startup/progress info to the terminal."""
    try:
        print(f"[SXMViewer] {message}", flush=True)
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
    "matplotlib",
]
