"""Visual regression fingerprint for the preview canvas.

The smoke test proves the app runs. It does **not** prove the pixels are
unchanged - and every bug this repo has recorded from refactoring
`detail_preview_canvas.py` was visual: smeared overlays, an axis dragged
back to a stale range, a mirrored grid. A green smoke test would have
passed through all of them.

This renders a deterministic set of canvas states to PNG and hashes each
one, so two git revisions can be compared pixel-for-pixel:

    # on the baseline revision
    python scripts/render_fingerprint.py --folder DATA --out before

    # after refactoring
    python scripts/render_fingerprint.py --folder DATA --out after
    python scripts/render_fingerprint.py --compare before after

Any hash difference is a real rendering change; the PNGs are kept so you
can look at what moved rather than guessing.

**Run this before and after any change to `MultiPreviewCanvas` render
state** - the profile / molecule overlay / crop-template domains are all
render-coupled and cannot be verified any other way.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QT_QPA_FONTDIR", r"C:\Windows\Fonts")


def _hash_png(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16]


def capture(folder: str, out_dir: Path, n_images: int, settle: float):
    out_dir.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="sxm_render_"))
    from sxm_viewer import config_io
    config_io.CONFIG_PATH = tmp / "config.json"
    config_io.HEADER_CACHE_PATH = tmp / "header_cache.json"

    from sxm_viewer._shared import QtWidgets
    from sxm_viewer.gui.main_window import SXMGridViewer
    SXMGridViewer._maybe_offer_recovery_session = lambda self, *a, **k: None

    app = QtWidgets.QApplication([])
    viewer = SXMGridViewer()
    viewer.load_folder(Path(folder))
    end = time.time() + settle
    while time.time() < end:
        app.processEvents()
        time.sleep(0.01)
    viewer.ensure_spectros_loaded(refresh=False)
    end = time.time() + 5
    while time.time() < end:
        app.processEvents()
        time.sleep(0.01)

    def settle(seconds=0.6):
        """Let debounced reflows/redraws finish before grabbing pixels.

        Without this the capture races with layout: an early grab caught a
        half-drawn second view strip in one run and not the next, making
        two *identical* revisions compare as different. A visual check with
        false positives is worse than none - it trains you to ignore it.
        """
        deadline = time.time() + seconds
        while time.time() < deadline:
            app.processEvents()
            time.sleep(0.01)

    fingerprint = {}
    files = [str(f) for f in viewer.files][:n_images]

    # Prime the canvas with a throwaway render first. The *first* preview
    # after a folder load can still carry startup state (leftover view
    # count from the load-time preview), which made image 000 - and only
    # image 000 - differ between otherwise identical runs. Priming puts
    # the canvas in the same steady state before anything is captured.
    if files:
        try:
            viewer.show_file_channel(files[0], 0)
            settle(1.0)
            viewer.preview_canvas.draw()
            settle(0.3)
        except Exception:
            pass

    for idx, key in enumerate(files):
        try:
            viewer.show_file_channel(key, 0)
            settle()
            # Force any pending draw to complete, then settle again - the
            # canvas batches redraws and a grab can otherwise land between
            # the data update and the repaint.
            try:
                viewer.preview_canvas.draw()
            except Exception:
                pass
            settle(0.3)
        except Exception as exc:
            fingerprint[f"{idx:03d}_preview"] = f"ERROR {exc}"
            continue
        canvas = viewer.preview_canvas
        # 1. The LIVE figure - everything the interactive canvas has drawn
        #    (image, axes, colorbar, overlays).
        #
        #    Deliberately savefig() on canvas.figure rather than a Qt
        #    widget grab: the widget's bottom few rows pick up neighbouring
        #    chrome and were not reproducible run to run (2416 pixels in
        #    rows 584-591 differed between two identical revisions), while
        #    the canvas's own state - view count, axes, size - was
        #    identical. The figure is the thing whose correctness we care
        #    about, and it renders deterministically.
        png = out_dir / f"{idx:03d}_canvas.png"
        try:
            canvas.figure.savefig(str(png), dpi=72)
            fingerprint[f"{idx:03d}_canvas"] = _hash_png(png)
        except Exception as exc:
            fingerprint[f"{idx:03d}_canvas"] = f"ERROR {exc}"
        # 2. the export figure, which uses a separate rendering path
        try:
            views = list(getattr(canvas, "views", []) or [])
            if views:
                fig = canvas._render_view_figure(views[0])
                fpath = out_dir / f"{idx:03d}_export.png"
                fig.savefig(fpath, dpi=72)
                fingerprint[f"{idx:03d}_export"] = _hash_png(fpath)
        except Exception as exc:
            fingerprint[f"{idx:03d}_export"] = f"ERROR {exc}"

    (out_dir / "fingerprint.json").write_text(
        json.dumps(fingerprint, indent=1, sort_keys=True), encoding="utf-8")
    print(f"captured {len(fingerprint)} renders -> {out_dir}")
    sys.stdout.flush()
    os._exit(0)


def compare(before: Path, after: Path):
    a = json.loads((before / "fingerprint.json").read_text(encoding="utf-8"))
    b = json.loads((after / "fingerprint.json").read_text(encoding="utf-8"))
    keys = sorted(set(a) | set(b))
    same = diff = missing = 0
    for k in keys:
        va, vb = a.get(k), b.get(k)
        if va is None or vb is None:
            print(f"  MISSING  {k}: before={va} after={vb}")
            missing += 1
        elif va == vb:
            same += 1
        else:
            print(f"  CHANGED  {k}")
            print(f"           {before / (k + '.png')}")
            print(f"           {after / (k + '.png')}")
            diff += 1
    print(f"\n{same} identical, {diff} changed, {missing} missing")
    if diff or missing:
        print("\nA changed hash is a real pixel difference. Open both PNGs "
              "before assuming it is benign - anti-aliasing does not change "
              "on its own.")
        return 1
    print("No rendering change.")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder")
    ap.add_argument("--out")
    ap.add_argument("--images", type=int, default=6)
    ap.add_argument("--settle", type=float, default=20.0)
    ap.add_argument("--compare", nargs=2, metavar=("BEFORE", "AFTER"))
    args = ap.parse_args()

    if args.compare:
        return compare(Path(args.compare[0]), Path(args.compare[1]))
    if not (args.folder and args.out):
        ap.error("need --folder and --out (or --compare BEFORE AFTER)")
    capture(args.folder, Path(args.out), args.images, args.settle)
    return 0


if __name__ == "__main__":
    sys.exit(main())
