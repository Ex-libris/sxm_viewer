"""CLI entrypoint for launching the Qt viewer."""
from __future__ import annotations

from ._shared import QtGui, QtWidgets, sys
from .app_meta import configure_application
from .gui.main_window import SXMGridViewer


def main():
    app = QtWidgets.QApplication(sys.argv)
    configure_application(app)
    try: app.setFont(QtGui.QFont("Segoe UI", 11))
    except Exception: pass
    w = SXMGridViewer(); w.show(); sys.exit(app.exec_())

if __name__ == "__main__":
    main()



