"""Folder report generation (PDF).

Non-GUI package: like ``providers/``, nothing here may import Qt or any
``sxm_viewer.gui`` module, so report building stays headlessly testable.
The GUI hands in a plain-data payload (see ``gui/controllers/report.py``)
collected on the GUI thread; ``model.build_report_model`` derives the
report structure (session chapters via K-Means, per-grid curve clusters,
flagged items) and ``pdf.render_report_pdf`` renders it with matplotlib.
"""

from .model import build_report_model
from .pdf import render_report_pdf

__all__ = ["build_report_model", "render_report_pdf"]
