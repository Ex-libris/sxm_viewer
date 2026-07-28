"""Spectroscopy load lifecycle: lazy scan, async autoload, manifest save.

Follows the `gui/spectroscopy/*.py` module-function convention (`viewer`
first). Extracted from `main_window.py` as part of the god-class reduction
(`docs/refactor/GOD_CLASS_PLAN.md`).

Three related concerns live here:

* **Lazy loading** - a folder load defers the spectroscopy scan so the
  window appears immediately; `ensure_loaded` forces it on demand.
* **Async autoload** - the deferred trigger runs the scan off the GUI
  thread. Everything *after* the scan is applied back on the GUI thread by
  `apply_scan_results`, which the synchronous path shares, so both routes
  produce identical state.
* **Manifest persistence** - debounced, single-flight background save.

**Why only the autoload path is async**: several callers of
`ensure_loaded`/`reload` rely on `viewer.spectros` being populated the
moment the call returns. Making those asynchronous would silently break
them, so they stay synchronous on purpose; only the post-folder-load
autoload trigger (which nothing waits on) moved off-thread, because that
scan froze the whole UI for 1-2+ seconds on real folders.
"""
from __future__ import annotations

import threading
import time
from pathlib import Path

from ..._shared import QtCore, log_status
from ..viewer import loader as viewer_loader

MANIFEST_SAVE_DELAY_MS = 400


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def ensure_loaded(viewer, refresh: bool = True):
    """Load spectroscopies now if they were deferred. Synchronous.

    Returns True when the data is (or already was) available, False when a
    load is already in flight.
    """
    if viewer._spectros_loaded:
        return True
    if getattr(viewer, "_spectros_loading", False):
        return False
    _stop_autoload_timer(viewer)
    viewer._spectros_loading = True
    viewer._spectros_pending = False
    try:
        log_status("[Lazy] Loading spectroscopy references...")
        reload_spectros(viewer, refresh=refresh)
        if not refresh and viewer._spectros_loaded:
            _refresh_markers_and_preview(viewer)
    finally:
        viewer._spectros_loading = False
    return True


def schedule_pending_load(viewer, delay_ms: int = 1200):
    """Arm the deferred autoload timer, if a load is still outstanding."""
    if viewer._spectros_loaded or not getattr(viewer, "_spectros_pending", False):
        return
    if getattr(viewer, "_spectros_loading", False):
        return
    try:
        viewer._spectro_autoload_timer.start(max(0, int(delay_ms)))
    except (AttributeError, RuntimeError):
        pass


def run_pending_load(viewer):
    """Synchronous fallback, used when async worker setup itself fails."""
    if viewer._spectros_loaded or not getattr(viewer, "_spectros_pending", False):
        return
    if getattr(viewer, "_spectros_loading", False):
        return
    ensure_loaded(viewer,
                  refresh=bool(getattr(viewer, "show_spectro_miniatures", False)))


def run_pending_load_async(viewer, worker_factory):
    """Background-threaded autoload (see module docstring).

    ``worker_factory(viewer, folder)`` builds the scan worker; injected so
    this module does not depend on `main_window`'s private worker class.
    """
    if viewer._spectros_loaded or not getattr(viewer, "_spectros_pending", False):
        return
    if getattr(viewer, "_spectros_loading", False):
        return

    refresh = bool(getattr(viewer, "show_spectro_miniatures", False))
    viewer._spectros_loading = True
    viewer._spectros_pending = False
    viewer._spectros_loaded = False
    viewer._spectro_miniature_cache.clear()

    folder = _spectro_folder(viewer)
    log_status("[Lazy] Loading spectroscopy references...")
    log_status(f"Scanning spectroscopy files in: {folder}")
    t_scan_start = time.perf_counter()

    thread = QtCore.QThread(viewer)
    worker = worker_factory(viewer, folder)
    worker.moveToThread(thread)
    thread.started.connect(worker.run)

    def _cleanup_thread():
        thread.quit()
        thread.wait()
        viewer._spectro_scan_thread = None
        viewer._spectro_scan_worker = None

    def _on_finished(specs, spec_stats):
        scan_ms = (time.perf_counter() - t_scan_start) * 1000.0
        _cleanup_thread()
        apply_scan_results(viewer, specs, spec_stats,
                           refresh=refresh, scan_ms=scan_ms)
        viewer._spectros_loading = False
        if not refresh and viewer._spectros_loaded:
            _refresh_markers_and_preview(viewer)

    def _on_failed(error_msg):
        log_status(f"Spectroscopy scan failed: {error_msg}")
        _cleanup_thread()
        viewer._spectros_loading = False

    worker.finished.connect(_on_finished)
    worker.failed.connect(_on_failed)
    # Keep references alive for the thread's lifetime - a local-only
    # QThread/QObject can be garbage-collected out from under itself.
    viewer._spectro_scan_thread = thread
    viewer._spectro_scan_worker = worker
    thread.start()


def reload_spectros(viewer, refresh=True):
    """Full synchronous rescan of the spectroscopy folder."""
    # Until a reload completes successfully the cache is stale.
    viewer._spectros_loaded = False
    viewer._spectros_pending = False
    viewer._spectro_miniature_cache.clear()
    t_scan_start = time.perf_counter()
    folder = _spectro_folder(viewer)
    log_status(f"Scanning spectroscopy files in: {folder}")
    specs, spec_stats = viewer._scan_spectros(folder)
    scan_ms = (time.perf_counter() - t_scan_start) * 1000.0
    apply_scan_results(viewer, specs, spec_stats,
                       refresh=refresh, scan_ms=scan_ms)


def apply_scan_results(viewer, specs, spec_stats, *, refresh, scan_ms):
    """Apply scan output to viewer state. **Always on the GUI thread.**

    Shared by the synchronous and async paths so both produce identical
    state - the async completion handler calls this from the GUI thread.
    """
    viewer.spectros = specs
    if not spec_stats:
        # Keep stats for the UI but avoid duplicate terminal spam (the
        # loader already logged its own summary).
        log_status(f"Loaded {len(viewer.spectros)} spectroscopy entries")
    if getattr(viewer, "_spectro_manifest_pending_save", False):
        viewer._spectro_manifest_pending_save = False
        schedule_manifest_save(viewer)

    t_assign_start = time.perf_counter()
    viewer._assign_spectros_to_images()
    t_assign_end = time.perf_counter()

    viewer.matrix_spectros = [spec for spec in viewer.spectros
                              if spec.get("matrix_index") is not None]
    viewer._clear_multi_spec_selection()
    viewer._update_spectro_stats_label(spec_stats)
    viewer._spectros_loaded = True
    viewer._update_matrix_summary_banner()

    if refresh:
        t_thumb_start = time.perf_counter()
        viewer.populate_thumbnails_for_channel(viewer.channel_dropdown.currentIndex())
        if viewer.last_preview:
            viewer.show_file_channel(viewer.last_preview[0], viewer.last_preview[1])
        t_thumb_end = time.perf_counter()
    else:
        t_thumb_start = t_thumb_end = t_assign_end

    log_status(
        f"[Perf] Spectros: scan {scan_ms:.0f} ms | "
        f"assign {(t_assign_end - t_assign_start) * 1000.0:.0f} ms | "
        f"thumbs {(t_thumb_end - t_thumb_start) * 1000.0:.0f} ms")


# ---------------------------------------------------------------------------
# Manifest persistence
# ---------------------------------------------------------------------------

def schedule_manifest_save(viewer):
    """Debounced save; falls back to an immediate flush without a timer."""
    viewer._spectro_manifest_save_pending = True
    try:
        viewer._spectro_manifest_save_timer.start(MANIFEST_SAVE_DELAY_MS)
    except (AttributeError, RuntimeError):
        flush_manifest_save(viewer)


def flush_manifest_save(viewer):
    """Write the manifest on a background thread, single-flight.

    If a save is already running the request is re-queued rather than
    started concurrently - two writers would race on the same file.
    """
    if viewer._spectro_manifest_save_inflight:
        viewer._spectro_manifest_save_pending = True
        return
    folder = (getattr(viewer, "spec_folder_path", None)
              or getattr(viewer, "last_dir", None))
    manifest_entries = dict(getattr(viewer, "_spectro_manifest_entries", {}) or {})
    if not folder or not manifest_entries:
        viewer._spectro_manifest_save_pending = False
        return
    viewer._spectro_manifest_save_pending = False
    viewer._spectro_manifest_save_inflight = True

    def _persist(snapshot_folder, snapshot_manifest):
        try:
            viewer_loader.save_spectro_manifest_snapshot(
                snapshot_folder, snapshot_manifest)
        finally:
            # Hop back to the GUI thread to clear the in-flight flag.
            QtCore.QTimer.singleShot(0,
                                     lambda: on_manifest_save_finished(viewer))

    threading.Thread(
        target=_persist,
        args=(folder, manifest_entries),
        name="spectro-manifest-save",
        daemon=True,
    ).start()


def on_manifest_save_finished(viewer):
    viewer._spectro_manifest_save_inflight = False
    if viewer._spectro_manifest_save_pending:
        schedule_manifest_save(viewer)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _spectro_folder(viewer):
    try:
        return Path(getattr(viewer, "spec_folder_path", None) or viewer.last_dir)
    except (TypeError, ValueError):
        return viewer.last_dir


def _stop_autoload_timer(viewer):
    try:
        viewer._spectro_autoload_timer.stop()
    except (AttributeError, RuntimeError):
        pass


def _refresh_markers_and_preview(viewer):
    """Post-load refresh used when the scan ran without a full repopulate."""
    viewer._schedule_marker_refresh()
    if viewer.last_preview:
        try:
            viewer.show_file_channel(viewer.last_preview[0], viewer.last_preview[1])
        except Exception:
            pass


__all__ = [
    "ensure_loaded",
    "schedule_pending_load",
    "run_pending_load",
    "run_pending_load_async",
    "reload_spectros",
    "apply_scan_results",
    "schedule_manifest_save",
    "flush_manifest_save",
    "on_manifest_save_finished",
]
