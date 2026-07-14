"""Background worker for folder-report generation.

Runs model building + PDF rendering off the GUI thread against a plain-data
payload only (collected beforehand on the GUI thread by ReportController) -
it must never touch the viewer or any Qt widget. Same QRunnable + signals
QObject pattern as BatchExportWorker (batch_export.py).
"""
from __future__ import annotations

from ..._shared import QtCore
from ...reporting import build_report_model, render_report_pdf


class ReportSignals(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, int, str)   # done, total, page description
    finished = QtCore.pyqtSignal(str, int)        # out_path, page count
    error = QtCore.pyqtSignal(str)


class ReportWorker(QtCore.QRunnable):
    def __init__(self, payload, out_path):
        super().__init__()
        self.payload = payload
        self.out_path = str(out_path)
        self.signals = ReportSignals()

    def run(self):
        try:
            model = build_report_model(self.payload)
            pages = render_report_pdf(
                model, self.out_path,
                progress_cb=lambda done, total, desc: self.signals.progress.emit(done, total, str(desc)),
            )
        except Exception as exc:
            self.signals.error.emit(f"{type(exc).__name__}: {exc}")
            return
        self.signals.finished.emit(self.out_path, int(pages))
