"""Headless smoke test: launch the app offscreen and exercise the main flows.

This is the project's safety net for refactoring. There is no pytest suite
for the GUI (see docs/TESTING.md), so "did I break it?" is answered by
driving the real ``SXMGridViewer`` under Qt's offscreen platform and
asserting the core paths still work end to end.

    python scripts/smoke_test.py --folder "C:\\DATA\\some_folder"
    python scripts/smoke_test.py                 # UI-only checks, no data

Exit code 0 = all checks passed.

Hard-won setup notes (each of these cost real debugging time):

* ``QT_QPA_PLATFORM=offscreen`` renders **no text at all** unless
  ``QT_QPA_FONTDIR`` is also set - PyQt5 ships no fonts. Screenshots then
  look like a theme regression that isn't there.
* The offscreen viewer is the *real* app: anything persisting config writes
  to the user's real ``~/.sxm_viewer_config.json``. This script redirects
  ``config_io``'s module-level paths to a temp dir **before** importing the
  viewer, so a run can never touch real user settings.
* ``_maybe_offer_recovery_session`` opens a modal dialog from a startup
  timer if a recovery snapshot exists, which hard-crashes offscreen
  (0xC0000005) at a point that looks random. It is stubbed out below.
* ``os._exit`` at the end dodges Qt teardown crashes corrupting the exit
  code - but stdout must be flushed first or buffered output vanishes.
"""
from __future__ import annotations

import argparse
import faulthandler
import os
import shutil
import sys
import tempfile
import time
import traceback
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

faulthandler.enable()
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")


class Checks:
    def __init__(self):
        self.passed, self.failed = [], []

    def check(self, name, fn):
        try:
            result = fn()
        except Exception as exc:
            self.failed.append((name, f"{type(exc).__name__}: {exc}"))
            traceback.print_exc()
            return None
        if result is False:
            self.failed.append((name, "returned False"))
        else:
            self.passed.append(name)
        return result

    def report(self):
        print()
        for name in self.passed:
            print(f"  PASS  {name}")
        for name, why in self.failed:
            print(f"  FAIL  {name}: {why}")
        print(f"\n{len(self.passed)} passed, {len(self.failed)} failed")
        return not self.failed


def pump(app, seconds):
    """Run the Qt event loop for a while (folder load is async)."""
    end = time.time() + seconds
    while time.time() < end:
        app.processEvents()
        time.sleep(0.01)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder", default=None,
                    help="data folder to load (skipped if omitted)")
    ap.add_argument("--report", action="store_true",
                    help="also generate a PDF folder report")
    ap.add_argument("--settle", type=float, default=20.0,
                    help="seconds to wait for folder load")
    args = ap.parse_args()

    tmp = Path(tempfile.mkdtemp(prefix="sxm_smoke_"))
    # Redirect config BEFORE importing the viewer - these are module-level
    # globals read at call time, so patching the attributes fully isolates us.
    from sxm_viewer import config_io
    config_io.CONFIG_PATH = tmp / "config.json"
    config_io.HEADER_CACHE_PATH = tmp / "header_cache.json"
    if hasattr(config_io, "COLLECTIONS_INDEX_PATH"):
        config_io.COLLECTIONS_INDEX_PATH = tmp / "collections.json"

    from sxm_viewer._shared import QtWidgets
    from sxm_viewer.gui.main_window import SXMGridViewer
    # See module docstring - modal recovery dialog crashes offscreen.
    SXMGridViewer._maybe_offer_recovery_session = lambda self, *a, **k: None

    checks = Checks()
    app = QtWidgets.QApplication([])
    viewer = checks.check("construct SXMGridViewer", lambda: SXMGridViewer())
    if viewer is None:
        checks.report()
        sys.stdout.flush()
        os._exit(1)

    checks.check("toolbar built",
                 lambda: hasattr(viewer, "toolbar_report_act"))
    checks.check("preview canvas built",
                 lambda: viewer.preview_canvas is not None)
    checks.check("colormap combos populated",
                 lambda: viewer.preview_cmap_combo.count() > 10)

    # Widget-silencing behaviour: setting a combo programmatically must not
    # fire its handler (this is what set_silent/blockSignals protects).
    def combo_silent():
        fired = []
        combo = viewer.preview_cmap_combo
        combo.currentIndexChanged.connect(lambda _i: fired.append(1))
        prev = combo.blockSignals(True)
        combo.setCurrentIndex((combo.currentIndex() + 1) % combo.count())
        combo.blockSignals(prev)
        return not fired
    checks.check("blockSignals suppresses handler", combo_silent)

    # ActivityLog batches through a timer, so a message is only visible in
    # the widget after a flush - assert the whole round trip, not just that
    # append() did not raise.
    def activity_log_roundtrip():
        log = getattr(viewer, "activity_log", None)
        if log is None:
            return False
        before = viewer.activity_log_box.toPlainText()
        log.append("smoke-test-marker")
        log.flush()
        after = viewer.activity_log_box.toPlainText()
        if "smoke-test-marker" not in after:
            return False
        log.clear()
        return viewer.activity_log_box.toPlainText() == "" and before is not None
    checks.check("activity log append/flush/clear", activity_log_roundtrip)

    # Debouncer: both payload policies, plus the rearm path used when a
    # request lands while a render is already running.
    def debouncer_semantics():
        from sxm_viewer.gui.debounce import ACCUMULATE, LATEST, Debouncer
        fired = []
        latest = Debouncer(lambda: fired.append("x"), 0, LATEST)
        latest.schedule(("a", 1))
        latest.schedule(("b", 2))
        if latest.take() != ("b", 2):        # latest wins
            return False
        if latest.take() is not None:        # take() clears
            return False
        acc = Debouncer(lambda: None, 0, ACCUMULATE)
        acc.schedule({"p1"})
        acc.schedule({"p2", "p1"})
        if acc.take() != {"p1", "p2"}:       # union, deduped
            return False
        # rearm must not resurrect an empty payload
        acc.cancel()
        acc.rearm()
        if acc.is_active:
            return False
        # rearm DOES re-arm when a payload is still outstanding (the
        # "request arrived mid-render" path).
        acc.schedule({"p3"})
        acc.cancel()
        acc.schedule({"p4"})
        acc.rearm()
        if acc.peek() != {"p4"}:
            return False
        # flush() invokes the callback; a callback that does not take()
        # deliberately leaves the payload for the next round.
        taken = []
        drain = Debouncer(lambda: taken.append(drain.take()), 0, LATEST)
        drain.schedule(("c", 3))
        drain.flush()
        return bool(fired is not None) and taken == [("c", 3)]
    checks.check("debouncer latest/accumulate/rearm", debouncer_semantics)

    if args.folder:
        folder = Path(args.folder)
        if not folder.exists():
            print(f"!! folder not found: {folder}")
            checks.failed.append(("folder exists", str(folder)))
        else:
            checks.check("load_folder", lambda: viewer.load_folder(folder) or True)
            pump(app, args.settle)
            checks.check("files loaded", lambda: len(viewer.files) > 0)
            checks.check("headers parsed", lambda: len(viewer.headers) > 0)
            checks.check("thumbnails populated",
                         lambda: len(getattr(viewer, "thumb_widgets", {})) > 0)

            checks.check("ensure_spectros_loaded",
                         lambda: viewer.ensure_spectros_loaded(refresh=False) or True)
            pump(app, 5)

            def preview_first():
                viewer.show_file_channel(str(viewer.files[0]), 0)
                app.processEvents()
                return viewer.last_preview is not None
            checks.check("preview an image", preview_first)

            checks.check("report action enabled after load",
                         lambda: viewer.toolbar_report_act.isEnabled())

            if args.report:
                out = tmp / "smoke_report.pdf"

                def gen_report():
                    payload = viewer.report_controller.collect_payload()
                    from sxm_viewer.reporting.model import build_report_model
                    from sxm_viewer.reporting.pdf import render_report_pdf
                    model = build_report_model(payload)
                    pages = render_report_pdf(model, str(out))
                    print(f"       report: {pages} pages, "
                          f"{model['summary']['n_sequences']} sequence(s)")
                    return pages > 0
                checks.check("generate folder report", gen_report)

    ok = checks.report()
    shutil.rmtree(tmp, ignore_errors=True)
    sys.stdout.flush()
    os._exit(0 if ok else 1)


if __name__ == "__main__":
    main()
