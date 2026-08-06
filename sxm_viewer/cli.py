"""CLI entrypoint for launching the Qt viewer."""
from __future__ import annotations

import faulthandler
import traceback

from ._shared import QtGui, QtWidgets, sys
from .app_meta import configure_application
from .gui.main_window import SXMGridViewer


def _install_crash_guards():
    """Keep the app alive on stray errors instead of hard-aborting.

    PyQt5 (>=5.5) calls ``qFatal()`` — a process abort — whenever a Python
    exception propagates out of a slot or a reimplemented Qt virtual
    (e.g. a matplotlib mouse-motion callback dispatched from within
    ``mouseMoveEvent``). Installing our own ``sys.excepthook`` makes PyQt
    log the traceback and continue rather than abort, turning what would be
    a full-window crash into a single logged error. ``faulthandler`` still
    dumps a native traceback for genuine C-level faults.
    """
    faulthandler.enable()

    def _hook(exc_type, exc, tb):
        try:
            traceback.print_exception(exc_type, exc, tb)
        except Exception:
            pass

    sys.excepthook = _hook


def main():
    _install_crash_guards()
    app = QtWidgets.QApplication(sys.argv)
    configure_application(app)
    try: app.setFont(QtGui.QFont("Segoe UI", 11))
    except Exception: pass
    w = SXMGridViewer(); w.show(); sys.exit(app.exec_())

if __name__ == "__main__":
    main()



