# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

SXM Viewer is a Python/PyQt5 desktop application for scientific SPM (Scanning
Probe Microscopy) data analysis and visualization. It natively supports
Anfatec/Omicron file formats and, via a conversion adapter, Nanonis `.sxm`
files.

## Commands

Setup and run (from repo root, PowerShell):

```powershell
conda create -n sxmviewer python=3.11
conda activate sxmviewer
cd .\scripts
python -m pip install -r .\requirements.txt
cd ..
python -m sxm_viewer
```

- Installed as a package: `python -m pip install .` then `sxm-viewer`
  (`pyproject.toml`; version from `sxm_viewer.__version__`).
- Legacy launch shim: `python scripts/sxm_grid_viewer.py` (forwards to `sxm_viewer.cli.main`).
- Checks (all run in CI, `.github/workflows/tests.yml`, Windows, Python
  3.10 + 3.12) - see `docs/TESTING.md` for what each catches:
  - `ruff check sxm_viewer scripts` - narrow gate configured in
    `pyproject.toml`: syntax errors, undefined names, unused imports.
  - `python -m pytest sxm_viewer/tests` - unit tests (profile measuring/units).
  - `python scripts/check_imports.py` - imports every package module offscreen.
  - `python scripts/smoke_test.py [--folder DATA --report]` - drives the
    real viewer offscreen; with a data folder it also loads, previews and
    renders a PDF report. Isolates config to a temp dir.
  - `python scripts/render_fingerprint.py` - pixel fingerprints of preview
    renders, compared across revisions (needs a data folder; not in CI).
  The vendored `nanonispy2` tests (`providers/nanonis/vendor/.../tests`) are
  upstream's and import the upstream name `nanonispy`, so they do not run
  as shipped; leave them untouched like the rest of the vendor folder.
- Headless smoke tests work with `$env:QT_QPA_PLATFORM = "offscreen"` before
  constructing `SXMGridViewer` in a script. **Trap**: the offscreen viewer is
  the real app — anything that persists config (e.g. `set_ui_theme`, toggle
  handlers calling `save_config`) writes to the user's real
  `~/.sxm_viewer_config.json`. Either avoid config-persisting calls in test
  scripts or restore the file afterwards.
- Headless testing notes (learned building the figure recipes): back up
  and restore `~/.sxm_viewer_config.json` around every script (figure
  windows save styles/state on each edit); the offscreen platform has no
  font glyphs and inflated font metrics, so widget widths measured there
  are ~2x real and grabbed screenshots show blank text (matplotlib text
  renders fine - check figures via `savefig`, layouts via numbers); lazy
  thumbnails stay placeholders unless you pump the event loop; stdout is
  lost on a hard abort, so run with `python -u` and print with `flush=True`.
- PowerShell 5.1 mangles embedded double quotes when passing args to native
  executables — write `git commit` messages without `"` characters (or use a
  `-F` message file).
- `scripts/install.py` / `scripts/install_sxm_viewer.bat` /
  `scripts/run_sxm_viewer.bat` are Windows installer/launcher helpers, not
  part of the app's runtime import path.

## Architecture

### Package layout (`sxm_viewer/`)

- `config.py`, `config_defaults.py`, `config_io.py` — user config file
  (`~/.sxm_viewer_config.json`), header cache, and in-memory cache size
  limits (`CHANNEL_DATA_CACHE_LIMIT`, `FILTERED_CACHE_LIMIT`).
- `header_cache.py` — Qt-free SQLite store of parsed `(header, fds)` per
  header file (`~/.sxm_viewer_header_cache.sqlite3`, opened via
  `config_io.open_header_cache`). Folder loads do one batched
  `lookup_many` and write only misses (`store_many`), so cost scales with
  the folder, not with every folder ever opened — it replaced a single JSON
  file that was parsed in full at startup and grew to 50 MB. Rows are
  validated by file mtime; `PRAGMA user_version` holds
  `HEADER_CACHE_VERSION` (bump it on breaking format changes - rows are
  wiped). The legacy `.json` is imported once on first open and otherwise
  ignored. Stale rows are pruned per loaded folder only, so unplugged
  drives keep their entries.
- `_shared.py` — common Qt/matplotlib/numpy imports re-exported for the GUI
  layer, plus a process-wide patch of `FigureCanvasQTAgg.resizeEvent` to guard
  against non-finite/zero resize events. `matplotlib.use("Agg")` is set here.
- `app_meta.py` — app name/org constants and window-icon discovery
  (packaged `assets/app_icon.ico`; an `app_icon.*` in `samples/` overrides
  it in a source checkout).
- `cli.py` — the real entry point (`python -m sxm_viewer` → `__main__.py` →
  `cli.main()`), which builds the `QApplication` and shows `SXMGridViewer`.
- `data/` — format-agnostic parsing: `io.py` (Omicron/Anfatec header + channel
  parsing, `parse_header`/`read_channel_file`/`normalize_unit_and_data`),
  `spectroscopy.py` (`.dat` metadata, axis helpers, matrix-scan detection),
  `matrix.py` (`MatrixDataset` representation).
- `processing/` — `filters.py` (image filter definitions/pipeline),
  `detection.py` (topography-channel auto-detection). GUI code and
  everything else import Nanonis support directly from `providers/nanonis`
  (a since-removed `processing/nanonis_adapter.py` re-export shim had zero
  importers anywhere in the repo and was deleted).
- `providers/` — format-specific adapters, kept import-isolated from GUI/Qt:
  - `nanonis/adapter.py` converts Nanonis `.sxm` scans into Omicron-style
    header/channel caches (`.sxmviewer_nanonis/` folders next to the data,
    gitignored). Channel caches are binary `.npy`; CH/CC tag auto-detection
    reuses cached results when the header and topography source are
    unchanged, so warm reloads skip redundant topography re-reads.
  - `nanonis/vendor/` bundles `nanonispy2` upstream — **do not edit**, it
    mirrors an external package and is the only place with its own tests.
- `reporting/` — PDF folder-report generation, kept import-isolated from
  GUI/Qt like `providers/` (headlessly testable): `model.py` (K-Means sample
  regions, per-grid curve clustering, KPFM fit maps, flagged items),
  `pdf.py` (matplotlib `PdfPages` page rendering), `channels.py` (channel-
  name classification heuristics — see "Folder reports" below).
- `molecules/` — the molecular-model overlay (.xyz/.pdb...), Qt-free and
  shared by every window: `model.py` (`Molecule`, palettes, radii, render
  styles), `render.py` (`draw_molecules`, THE renderer), `interaction.py`
  (gesture map + drag math, orientation/gizmo). See "Molecular overlays".
- `utils/` — small helpers (`units.py` unit parsing/formatting, `logging.py`).
- `gui/` — the Qt UI (see below).

### GUI structure (`gui/`)

`main_window.py` (`SXMGridViewer`) is the top-level coordinator: it builds the
shared widgets (thumbnail list, preview canvas, toolbars), instantiates
feature controllers, and wires high-level events (folder load, theme switch).
It is intentionally large/monolithic as the historical center of the app —
new feature logic should generally *not* be added directly here once it grows
non-trivial; instead follow the extension pattern below.

```
gui/
├── main_window.py            # SXMGridViewer: composition root
├── main_window_layout.py     # layout helpers, shortcuts panel
├── main_window_toolbar.py    # toolbar actions, theme selector button
├── main_window_spectro.py    # spectro dock wiring
├── theme.py                  # named UI themes (light/dark/amber): tokens, QPalette, chrome QSS — never applied to data imagery (sole exception: the user-opted "Full amber imagery" toggle, which recolors image artists display-time-only via cmap_registry.effective_cmap; per-file cmaps and exports keep true colormaps)
├── controllers/              # feature controllers, see table below
├── viewer/                   # thumbnail loading/rendering, preview, loader, measurement, state
├── dialogs/                  # modal dialogs (histogram/profile/filters/spectroscopy/matrix-fit)
├── canvases/                 # canvas workspace window, tiles, rendering, molecule overlays
├── spectroscopy/             # spectroscopy browser/controller/overlays/popups
├── workers/                  # background workers (batch export, folder-report rendering)
└── thumbnail_render.py, minimap.py, palettes.py, styles.py, ...  # shared widgets/utilities
```

Controllers in `gui/controllers/` expose a compact API (e.g.
`show_histogram()`, `handle_thumbnail_event()`) and receive only the widgets/
callbacks they need — dependencies are explicit, not implicit via shared
mutable state:

| Controller | Responsibilities |
| --- | --- |
| `preview_popup.py` | pop-out window spawning, toolbar sync |
| `thumbnail_controller.py` | thumbnail clicks, drag-to-canvas, keyboard nav, CH/CC tagging |
| `histogram.py` | Histogram & Range dialog, auto CLIM, live preview updates |
| `profile.py` | wires canvases to `ProfileDialog`, keeps measurements/overlays in sync |
| `spectro_compare.py` | spectroscopy comparison workflows, trace export, minima |
| `quick_crop.py` | quick crop panel, overlays, pop-out history |
| `collection.py` | collection/session-level state across loaded scans |
| `filter_controller.py` | image filter pipeline application/state |
| `image_compare.py` | side-by-side/overlay image comparison workflows |
| `session.py` | serializes/deserializes full viewer state to session JSON files |
| `report.py` | folder-report workflow: GUI-thread payload collection, save dialog, progress, launches `ReportWorker` |

**Extending the GUI**: prototype new cross-widget features inline in
`main_window.py` first; once the workflow solidifies, move it into a new
controller module with a documented public surface (expected widgets,
signals) and wire it up via composition in the main window constructor. This
is how the current controller set came to exist and keeps `main_window.py`
navigable despite its size.

Widgets/canvases (`viewer/`, `canvases/`) and dialogs (`dialogs/`) stay pure
UI: they emit Qt signals but do not reach back into `main_window` directly;
controllers subscribe to those signals instead.

### The `gui/viewer/*.py` module-function convention

`gui/viewer/thumbnail_ui.py`, `preview.py`, `loader.py`, `measurement.py`, and
`export.py` do **not** define classes. Every top-level function takes the
`SXMGridViewer` instance as its first parameter (conventionally named
`viewer`) and reads/writes `viewer.*` attributes directly — e.g.
`populate_thumbnails_for_channel(viewer, channel_idx)`,
`show_file_channel(viewer, header_path_str, channel_idx)`,
`load_folder(viewer, folder)`. `main_window.py` exposes each as a one-line
bound-method shim (e.g. `SXMGridViewer.on_thumb_sort_changed` just calls
`viewer_thumb_ui.on_thumb_sort_changed(self, idx)`). This is a mixin achieved
via plain functions instead of actual mixin classes — when adding a thumbnail/
preview/loader/measurement/export feature, add a function here rather than a
method on `SXMGridViewer`. Note this differs from `gui/controllers/*.py`,
where `SessionController`/`CollectionController`/etc. are real classes
holding `self.viewer` — that inconsistency is historical, not a rule to
replicate deliberately.

## Feature reference: where to find things

Line numbers below are approximate (as of the code at documentation time) —
files like `main_window.py` and `detail_preview_canvas.py` change often, so
re-grep the method name to confirm before relying on an exact line.

### Menus & toolbars
- `SXMGridViewer` is a `QtWidgets.QWidget`, not `QMainWindow` — there is no
  classic `menuBar()`. "Menus" are `QToolButton`s with attached `QMenu`s.
  Toolbar assembly lives in `gui/main_window_toolbar.py`
  (`create_main_toolbar`, `update_toolbar_actions`), invoked from
  `SXMGridViewer.__init__`.
- Menu population helpers in `main_window.py`: `_populate_browse_molecules_menu`,
  `_refresh_recent_dirs_menu`, `_refresh_recent_session_dirs_menu`.
- Display presets (Focus/Analysis/Publication) are transient treatments on
  `MultiPreviewCanvas` (`apply_display_preset`, `_active_display_preset`,
  "Back to ..." = `restore_previous_display`). One menu builder,
  `populate_display_preset_menu`, feeds both the image right-click menu
  (top-level "Display preset: X") and toolbar Display > Display preset. The
  user's default lives in config `default_display_preset` (absent = built-in
  Analysis) via `SXMGridViewer.get/set_default_display_preset`; an explicit
  default is applied transiently at startup.
- Right-click/context menus (not the toolbar) are separate methods:
  `_on_thumb_context_menu`, `_on_spectro_thumb_context_menu`,
  `_populate_marker_style_menu` in `main_window.py`.
- **Image right-click menus share one mechanism, `gui/menu_kit.py`**
  (viewer-agnostic): `section()` greyed headers (QLabel `QWidgetAction`s -
  `addSection` text isn't drawn by the native Windows style),
  `checkable`/`choice_group`, and the export section (`Route` registry +
  `add_export_section`: "Again: ..." / user-pinned routes / "More export"
  with "Customise pinned..."). `MenuPrefs` stores pinned routes and the last
  export per *scope* in config `context_menus.<scope>`. Rules for every
  such menu: sections Display / (Tools) / Export / Analyse & create /
  Organise, at most one submenu level, hide items that don't apply.
  - Thumbnails: `gui/viewer/thumb_menu.py` (`_on_thumb_context_menu` is a
    shim), routes in `gui/viewer/export_routes.py:thumbnail_routes`, scope
    `thumbnails` (migrates the old `thumb_menu` config key once).
  - Main preview + every pop-up: `MultiPreviewCanvas._show_context_menu`
    (one builder for both), built-in routes in `_view_export_routes`;
    the owner calls `export_routes.install_canvas_menu_support` (main
    window init + `preview_popup.spawn_preview_popup`) to pass the scope
    `view` prefs and viewer-backed routes (Obsidian). Other canvases (e.g.
    the matrix-fit dialog) fall back to in-memory prefs.
  - Spectroscopy thumbnails: `gui/viewer/spectro_menu.py`
    (`_on_spectro_thumb_context_menu` is a shim), scope `spectra`, ctx
    `{"specs", "channel"}`. Exports render the plot off-screen from the spec
    data (miniature channel + axis, SI-prefixed y unit via
    `_auto_display_unit`; several selected spectra overlaid) so no Spectrum
    window is needed; specs are hydrated lazily inside each route, not at
    menu open. Obsidian results may carry their own `caption`.
  - Route `run(ctx, dest)` returns a destination record (`{"out_dir"}` /
    `{"note","vault"}`) to make "Again" possible; export functions take
    `paths`/`out_dir` and return the destination on success for this.
  New per-image actions go into the right section of these builders, not
  as new top-level items.
- The publication-canvas toolbar (`gui/canvases/canvas_window_ui.py:build_toolbar`)
  and its tabbed side panel are built independently of the main window's.
- Spectroscopy dialogs (`gui/dialogs/spectroscopy_dialogs.py`) each build their
  own menus, e.g. `SpectroscopyPopup.trace_style_menu`/`legend_menu`/`filters_menu`.

### Keyboard shortcuts
- `SXMGridViewer.keyPressEvent` — Ctrl+A (select all thumbnails, only when
  focus is in the thumbnail area), Ctrl+C (copy thumbnail selection to
  clipboard as files), Ctrl+D (duplicate preview popup), arrow keys (thumbnail
  navigation).
- `QShortcut`s registered in `SXMGridViewer.__init__`: Ctrl+Shift+C/Z/W/R/T/M
  for quick-crop toggle/undo/close/real-size/template/minimize,
  Ctrl+Shift+P (recall popups), Ctrl+S (save session).
- `_init_mode_shortcuts` wires per-mode shortcuts for Browse/Measure/Spectro
  mode switching.
- The main preview canvas (`MultiPreviewCanvas.keyPressEvent` in
  `detail_preview_canvas.py`) has its own independent key map: Ctrl+1..5 toggle
  profile/angle/molecule/scale-bar/acquisition overlays, Ctrl+H toggles the
  shortcut hint, Ctrl+C copies the displayed image, Ctrl+Z undoes, `A` triggers
  histogram auto-contrast, `M` opens molecule load, `Home` resets zoom,
  `R`/`Shift+R` turn the image 90° CW/CCW (main preview only, via
  `_rotate_image_callback`), X/Y/Z rotate the selected molecule
  (`Shift+Home` resets it), `H` hides every overlay (`set_overlays_hidden`:
  display-time only - overlay artists on image axes are hidden inside
  `draw()`/copy and restored, so per-overlay state is untouched; blit
  fast-paths check `_overlays_hidden` too; from the thumbnail grid the main
  window forwards `H` and `A` to the preview), `G`/`T` add the `plane2`/`tilt`
  filter step (Shift removes; `set_level_filter` ->
  `FilterController.set_quick_filter_step`, undo entry `("filters", old_specs)`,
  also from the thumbnail grid), otherwise `Z` toggles per-image Zero
  (`_zero_image_callback` → `toggle_zero_images`; popups use their own
  `_popup_relative_zero_setter`), Enter/Esc apply/cancel crop-transform mode.
- Per-image Zero: `viewer.image_zero` (`file_key -> bool`, only overrides of
  the global `display_units_relative` "Zero all images" default); always
  query `_is_image_zeroed(file_key)` and pass `file_key` to
  `_scale_unit_for_display`. Changes go through `_restore_zero_state`
  (undo entry `("zero", default, overrides)`, session key `image_zero`).
- The publication canvas has yet another independent key map
  (`CanvasGraphicsView.keyPressEvent` / `ExperimentalCanvasWindow._handle_canvas_key`
  in `gui/canvases/canvas_view.py` / `canvas_window.py`): arrows nudge
  selected items, Ctrl+A selects all, Escape clears selection,
  Delete/Backspace deletes, Ctrl+Z/Ctrl+Y undo/redo.
- **Global undo** (`gui/undo_journal.py`): Ctrl+Z in the main window is one
  chronological order, pop-ups included. Each undo stack records its pushes
  in `viewer._undo_journal` as `(source, weakref target)` (sources `view` =
  a canvas' `push_undo_state`, `crop` = quick-crop entries, `adjust` =
  `_adjustment_undo_stack`, a `JournaledList`, `collection`);
  `_on_global_undo_requested` pops the newest token and runs that source's
  undo on its target, skipping tokens whose entry or canvas is gone.
  `install_canvas_undo_journal(canvas, window)` puts a canvas in the order
  (main preview at startup, every `spawn_preview_popup` pop-up after its
  setup, plus a Ctrl+Z `QShortcut` on the pop-up window - the main window's
  shortcut doesn't reach other top-level windows). Display options mirrored
  to other canvases (`_apply_canvas_display_options`) are suspended there:
  one click = one step, on the canvas that was changed (the preview for
  toolbar changes). The main preview hands Ctrl+Z over (`_global_undo_callback`) and, with
  `_undo_keeps_current_image`, restores only display settings from a
  snapshot taken on another image (snapshots carry `views`; restoring them
  used to swap the image or blank the preview). Startup/session-load preset
  snapshots are kept out of the journal but stay in the canvas history for
  "Back to ...". Canvases not installed (e.g. the matrix-fit dialog) keep
  local undo. A new undo stack in the main window must journal its pushes, or Ctrl+Z only reaches it after the journal is
  empty (fallback chain).
- Full user-facing shortcut list: `docs/SHORTCUTS.md`.

### Mouse events & drag-drop
- Main window: `dragEnterEvent`/`dropEvent` accept dropped folders/files;
  `eventFilter` intercepts drag events on the thumbnail viewport for in-grid
  thumbnail reordering (`_handle_thumbnail_drag_event`); rectangle-select is
  `_update_rubber_band_selection`.
- Main preview canvas (`MultiPreviewCanvas` in `detail_preview_canvas.py`)
  uses matplotlib's `mpl_connect` (not Qt event overrides) for most
  interaction: `_on_base_click` is the primary dispatcher (scale-bar drag,
  molecule drag, crop-handle drag, double-click pop-out, right-click context
  menu, Ctrl-drag = quick profile line, Ctrl+Alt-drag = quick angle,
  Shift-drag = crop rect/square, Alt-drag = outline extraction);
  `_on_press`/`_on_motion`/`_on_release` handle profile-line editing;
  `_on_angle_*` handle angle-measurement frames; `_on_sb_*` handle scale-bar
  repositioning; `_on_scroll_zoom` handles cursor-centered zoom. Qt-level
  `mouseMoveEvent`/`mouseReleaseEvent` only handle starting an external QDrag
  (dragging the image out to another widget).
- Publication canvas (`gui/canvases/canvas_view.py` `CanvasGraphicsView`,
  `canvas_items.py` `CanvasImageItem`): rubber-band select and panning on the
  view; item drag/resize/rotate and Alt-drag-to-duplicate on individual tiles.

### Sorting & filtering thumbnails
- Sort/filter combos live on `SXMGridViewer` (`thumb_sort_combo`,
  `thumb_filter_combo`); changes are handled by
  `on_thumb_sort_changed`/`on_thumb_filter_changed` in `main_window.py`,
  which are thin shims into `gui/viewer/thumbnail_ui.py`'s
  `on_thumb_sort_changed`/`on_thumb_filter_changed` — the actual sort/filter
  logic lives in `populate_thumbnails_for_channel`.
- Sort modes: Name (A-Z, natural sort), Date (new→old / old→new, via
  `sort_key_date`/header datetime), Tag (constant-height → constant-current →
  untagged). No manual drag-order or size-based sort exists.
- Filter modes: All, Constant height, Constant current, Untagged, Matrix
  datasets — tag/type-only, no free-text search.

### Sessions vs. collections (two distinct persistence mechanisms)
- **Session** (`gui/controllers/session.py`, `SessionController`) — the *entire*
  app state: loaded folder, headers, per-file channel/cmap/adjustments,
  thumbnail filters/sort, preview + canvas + popup snapshots, window layout,
  zoom state. Saved as one `*.json` file (`save_session`/`save_session_as`/
  `load_session`) plus a sidecar `<name>_data/` folder holding `.npy` arrays
  and view snapshots. Payload has a `version` field but no migration logic is
  implemented yet (single format so far). Autosave/recovery snapshots
  (`_save_recovery_snapshot`, `_maybe_offer_recovery_session` in
  `main_window.py`) reuse the same `save_session`/`load_session` machinery.
  Collections are **not** part of the session payload.
  **JSON key trap**: `_jsonify` stringifies every dict key, so tuple/int-keyed
  state must be encoded explicitly and decoded on load -
  `per_file_channel_cmap` is saved as `[[file, ch, name], ...]`
  (`_decode_cmap_map` also reads the old `"('file', 0)"` form) and
  `image_adjustments` channel keys go back to int (`_decode_adjustments`);
  both used to come back silently unmatched. Virtual copies persist their
  arrays plus `insert_after`/`virtual_copy_order`, and restore bumps
  `viewer._virtual_counter` past restored keys so new copies can't reuse
  one. Both load paths share `_restore_processed_views_payload` +
  `_apply_basic_session_payload`.
- **Collection** (`gui/controllers/collection.py`, `CollectionController`) — a
  standalone, user-named "scrapbook" of view/popup/crop snapshots gathered
  from anywhere, independent of which folder/session is currently loaded.
  Saved as `*.sxmcoll.json` (`KIND = "sxm_collection"`) plus a sidecar
  `<stem>_collection_data/views/` folder. Item-level add/remove only — no
  whole-file rename/delete method (left to the OS file dialog). UI is the
  floating `_CollectionTrayWindow`/`_CollectionTrayList` in `main_window.py`.

### Main preview canvas: image display & redraw speed
`MultiPreviewCanvas` (`gui/canvases/detail_preview_canvas.py`, ~12k lines) is
the core image-rendering canvas (a matplotlib `FigureCanvas` embedded in Qt).
Performance-relevant entry points, since this file is the usual target for
responsiveness work (see recent commits on thumbnail/preview perf):
- `_redraw` — full rebuild of the figure (all `imshow` axes, cmap/clim,
  colorbar, overlays); expensive, used when structure changes (layout,
  view count, tool mode).
- `_fast_update_single_view` — cheap path: mutates an existing image's
  `set_data`/`set_cmap`/`set_extent`/`set_clim` in place instead of a full
  `_redraw`, used when exactly one view changes and no profile/angle/molecule
  tool is active.
- `draw`/`draw_idle`/`set_render_suspended` — suspend/batch redraws during
  bulk updates, replaying once unsuspended.
- `resizeEvent`/`_reflow_after_resize`/`_finalize_after_resize` — debounced
  resize: a cheap `draw_idle()` reflow while actively resizing, full
  `_redraw()` only once resizing settles.
- Throttled drag paths: `_on_motion_value` throttles fixed-crop template drag
  refresh (`_fixed_crop_drag_throttle_ms`); profile/angle artists have
  blitting fast-paths (`_blit_profile_artists`, `_blit_angle_frames`).

`_redraw` and `_fast_update_single_view` both configure a single axes/image
for a view via the same shared helpers in `gui/canvases/preview_axes_sync.py`
(`sync_axes_to_view`, `style_colorbar`) rather than duplicating that logic —
when the fast path was a hand-rolled second copy of this setup, it missed
matplotlib invalidation rules the full rebuild got for free, causing three
separate bugs in one session: `ax.dataLim` doesn't update when you call
`image.set_extent()` (only when an artist is first added), the axes'
aspect-locked box/attached colorbar can lag behind an extent change without
an explicit `ax.apply_aspect()`, and `ax.set_xticks(ax.get_xticks())` freezes
the axis onto a stale `FixedLocator` that silently drags `xlim`/`ylim` back
to whatever range the *first* image needed. See that module's docstring for
the full explanation. `_render_view_figure`/`_render_views_grid` (export/print
figures) independently re-derive similar axis setup for a throwaway
Figure/Axes and are flagged inline as a related risk if this bug class ever
resurfaces there.

### Overlays (all types)
All drawn/managed inside `MultiPreviewCanvas`
(`gui/canvases/detail_preview_canvas.py`) unless noted:
- **Scale bar** — `_add_scale_bar`/`_refresh_scale_bars`, drag via `_on_sb_*`,
  toggle Ctrl+4.
- **3D molecule overlay** — `Molecule` (`sxm_viewer/molecules/`, see
  "Molecular overlays" below) / `MoleculePropertiesDialog`
  (`gui/canvases/molecular_overlay.py`); `_draw_molecules` is a wrapper
  around `molecules.render.draw_molecules`, toggle Ctrl+3.
- **2D/SVG molecule overlay** — `SvgMoleculeOverlay`/`SvgAtom`/`SvgBond`
  (`gui/canvases/svg_molecule_overlay.py`); drawn by `_draw_svg_molecules`.
- **Profile line & HUD** — `_update_profile_artists`, saved profiles via
  `_add_saved_profile_from_pts`, toggle Ctrl+1.
- **Angle measurement** — `_update_angle_artists`, toggle Ctrl+2.
- **Crop rectangle/template + history** — `_render_template_overlay`,
  `_draw_fixed_crop_history`.
- **Acquisition/metadata text box** — `_draw_acquisition_overlay` (Ctrl+5),
  `_draw_image_size_overlay`, `_draw_filter_summary_overlay`.
- **Outline/contour** — `_draw_outlines` (Alt-drag to create).
- **Spectroscopy markers on thumbnails** — `gui/spectroscopy/overlays.py`:
  `_render_spectroscopy_overlays` draws matrix footprints, marker symbols
  (`_draw_marker_symbol`), and stack badges (`_draw_stack_badge`) onto
  thumbnail pixmaps.
- **Publication canvas** — molecule placement reuses `Molecule`/
  `SvgMoleculeOverlay` per-tile; alignment guides while dragging tiles are
  `AlignmentGuide` (`canvas_items.py`) shown/cleared via
  `CanvasGraphicsView.show_alignment_guides`/`clear_alignment_guides`.

### Publication canvas (figure composition workspace)
For series / zoom figures use the figure recipes instead (see "Figure
recipes" below); the canvas remains the free-layout tool, reachable from a
figure's ⋯ > "Open in canvas".

`ExperimentalCanvasWindow` (`gui/canvases/canvas_window.py`) is the top-level
window for composing a multi-panel publication figure from loaded scans;
launched lazily from `main_window.py` and cached on `self._canvas_window`.
Supporting modules: `canvas_view.py` (`CanvasGraphicsView`, a `QGraphicsView`),
`canvas_items.py` (`CanvasImageItem`, `RubberBandSelection`, `AlignmentGuide`
— `QGraphicsItem` subclasses), `canvas_rendering.py` (export-quality
matplotlib rendering), `canvas_layout.py` (2x2/1x3/3x1 layout presets),
`canvas_state.py` (undo/redo stack), `canvas_io.py` (`save_canvas`/
`load_canvas`/`export_image`, JSON with `{"version":1,"items":[...]}`).
A prior monolithic `experimental_canvas.py` predated this split (duplicate
definitions of `ExperimentalCanvasWindow`/`CanvasImageItem`/
`CanvasGraphicsView`/etc.), was never imported anywhere once superseded, and
has been removed.

### Spectroscopy dialogs & popups
`gui/dialogs/spectroscopy_dialogs.py` defines the dialog classes;
`gui/spectroscopy/popups.py` and `gui/controllers/spectro_compare.py`
(`SpectroCompareController`) are the actual orchestrators that decide which
dialog to open:
- `SpectroscopyPopup` — single-spectrum viewer/plotter.
- `MatrixSpectroViewer` — grid/matrix (CITS-style) spectroscopy viewer;
  click a pixel to plot/compare its spectrum.
- `SpectroscopyCompareDialog` — multi-trace comparison (stacks/sites of
  repeated spectra), background subtraction, waterfall/offset.
- `KPFMFitTrendDialog` — plots fit-derived metrics (e.g. contact potential)
  vs. Z/height across a stack of fits.
- `SpectroSummaryDialog` (`gui/spectroscopy/summary_dialog.py`) — browsing
  list of spectroscopy entries with thumbnail preview, hands off to the
  dialogs above.
- `gui/spectroscopy/controller.py` is **not** a UI controller — it's the
  spec-to-image assignment/grouping logic (`_assign_spectros_to_images`,
  `_build_spectro_sites`, `_annotate_xy_stacks`) that groups spectra into
  "sites"/"stacks" before `SpectroCompareController` opens a popup for them.
- `gui/spectroscopy/overlays.py` — marker/badge drawing (see Overlays above).
- **"Compare All Spectra on This Image (N)…"** —
  `SpectroCompareController.open_all_specs_popup(file_key)` gathers every
  non-matrix point spectrum assigned to one image (`all_specs_for_image`;
  matrix/grid points are excluded — they belong in the Grid Map Explorer) and
  opens them together in a `SpectroscopyCompareDialog`. It opens a comparison
  *window*, so it lives under the preview canvas' right-click **Analysis**
  submenu (wired via `MultiPreviewCanvas.set_spectra_compare_all_callback`),
  **not** the Overlays paint-toggle group, and is mirrored onto the thumbnail
  right-click menu (`_on_thumb_context_menu`) so it's reachable from the grid
  without opening the image. Shown only when the image has ≥2 point spectra.

### Position inset (universal, shared by every inset-bearing dialog)
The miniature acquisition image with per-spectrum position markers shown in
both `SpectroscopyPopup` and `SpectroscopyCompareDialog` is rendered by ONE
set of module-level functions in `spectroscopy_dialogs.py`, so the behaviour
can't drift between dialogs. Add an inset feature/fix once here and both get
it (do **not** re-implement per dialog):
- `build_position_inset(host, image, markers, *, draw_badges, image_key)` —
  the single renderer. Each dialog's `_update_position_inset` /
  `_update_position_inset_compare` only gathers the `image` + a normalized
  `markers` list (`{coords, color, is_current, spec, entry_index?}`) and
  delegates. It renders `set_aspect("equal", adjustable="box", anchor="NW")`
  so the display bbox hugs the image (true aspect kept, resize grip on the
  image corner, not out in the transparent allocation margin), hides the
  frame+patch (spectra show through the margin), draws halo'd markers, the
  optional per-marker labels + constant-height image-Z caption, the resize
  grip, and stores `host._inset_marker_specs = [(spec,col,row),...]` for
  click hit-testing.
- `populate_inset_settings_menu(host, menu)` — the shared options menu
  (show-other-points, marker symbol/size/fill/color, outline toggle+color,
  **Marker labels** submenu {Initial Z=`z_level_nm`, Time, Acquisition order,
  Image Z}, inset colormap). Reached from each dialog's **Inset** button and
  by **right-clicking the inset** (`inset_right_click_context_menu`). New
  options wire to module-level generic setters (`_set_inset_marker_outline`,
  `_toggle_inset_label_field`, ...) that mutate `host._inset_*` and call
  `_inset_refresh(host)` (dispatches to whichever `_update_position_inset*`
  the host has) — not per-dialog methods.
- **Click-to-select**: `_inset_marker_at_event(host, event)` hit-tests the
  stored marker coords via the inset's live `transData`; each dialog's press
  handler calls it *before* the resize/drag hit-tests and routes to
  `_select_spec_in_list` (compare → File list) / `_select_curve_for_spec`
  (single → curve list). Order labels/lookup use a **stable** spec key
  (`_inset_spec_key`, path+matrix_index) not `id(spec)`, which breaks when
  hydration re-creates the dict.
- **Resize**: `_begin_inset_resize` + `_resize_inset_bbox_scaled` scale the
  inset uniformly about its fixed top-left (aspect-locked corner drag);
  `_inset_drag_offset` anchors a move to the allocation bottom-left (not the
  now-tight bbox) so a rectangular inset doesn't jump at drag-start.
- **Image-Z caption** uses `thumbnail_ui.effective_tag(viewer, key)` (there is
  NO `viewer.effective_tag` method — a `getattr` for it silently always
  fails) and `_inset_image_abs_z_nm` (unfiltered topo median, cached on
  `viewer._inset_image_z_cache`), only when the background scan is tagged
  constant-height.
- **Preview-cmap → inset sync**: `viewer_preview._notify_spectro_insets` is
  called from `main_window._set_thumbnail_entry_cmap` (the choke point for
  preview/thumb/gallery cmap changes) and from `on_preview_cmap_changed`'s
  global fallthrough; it calls each open dialog's
  `refresh_inset_for_preview_cmap`.

### Coordinate frames & orientation: the one mental model (READ FIRST)
Almost every "the points are warped / mirrored / upside-down" report in this
app comes from confusing **three coordinate frames** for the same data. Hold
these three straight and the dense sections that follow (position mapping,
Grid Map Explorer, matrix fits, folder reports) all collapse into a single
rule.

**The three frames**
1. **Absolute nm** — the microscope's real stage coordinates. Every spec
   carries its true `(x, y)` here; a scan image carries a center +
   `Width`/`Height` + scan `Angle`. The only frame where different scans and
   specs are directly comparable. A *rotated* scan is a tilted rectangle in
   this frame.
2. **Anchor raster** — the pixel grid of one specific scan image (row 0 =
   top as displayed, col 0 = left). **This is what the user is looking at**:
   the thumbnail, the main preview, "Reference image" mode, and Nanonis's own
   viewer are ALL in this frame. Absolute-nm → this frame is exactly what
   `_map_spec_to_pixels` computes (rotate by the scan angle, then normalize).
3. **Grid-local** — a matrix/`.3ds` grid's own `[row, col]` index space (the
   order points were acquired/stored). Nanonis fills this in acquisition
   order, so for a rotated grid it can be flipped/mirrored versus the raster:
   index `[0, 0]` is **not** reliably "top-left on screen."

**The one rule**: *whenever you take grid-local data (a metric array, a fit
map, per-pixel values) and draw it, reorient it into the anchor raster frame
first.* The user compares it against the reference image, which is
raster-frame; skip the reorientation and it looks mirrored/upside-down even
though every number is correct.

**How reorientation works** (one idea, several ports):
`_grid_local_orientation(X, Y, rows, cols, angle_deg)` takes the grid's true
absolute `(x, y)`, rotates them by the anchor image's scan angle, and returns
`(row_flip, col_flip)` — whether to `np.flipud`/`np.fliplr` the array so grid
index order matches screen order. For an unrotated grid (`angle=0`) it's a
no-op — which is *why rotation bugs stay invisible until a real tilted scan*
(a `-141°` grid is a clean vertical flip; a `157°` grid is ~180° off).

**Two load-bearing details that cause silent wrongness (no crash, no obvious
misbehaviour — just plausible-looking wrong pixels):**
- **Rotate first, normalize per-axis second.** Normalizing each axis by its
  own span *before* rotating shears any non-square scan. Invisible on square
  scans (which is how it shipped).
- **Reorient every representation, or none agree.** One grid can be shown as
  raster / absolute / local, plus virtual copies, WSxM export, and fit maps.
  Each must apply the *same* flip or they disagree with each other and with
  the reference image.

**The recurring meta-trap** (why a real fix can look like it "didn't stick"
across clean reloads): derived per-spec fields (`grid_row`/`grid_col`,
`off_frame_*`, `assignment_*`) are computed once and must survive
`hydrate_spectro_entries` re-parsing — add any new one to
`_merge_payload_into_spec`'s preserved whitelist or it is silently wiped on
the next dialog open. Full story under "Grid Map Explorer" below.

**Every place that must obey the rule** (any one drifting = the bug class
returns): `_map_spec_to_pixels` (canonical, `main_window.py`),
`overlays.py::_spectros_near_thumb_pos` (the inverse, pixel→nm),
`controller.py::_spec_within_extent` / `_spec_frame_offset_info`,
`MatrixSpectroViewer._grid_local_orientation` + `_draw_image_layer` (metric /
relative / virtual-copy / WSxM), `matrix_fit.py`'s `_grid_local_orientation`
+ `MatrixFitDialog._local_flips` (the fit maps — most recent addition), and
`reporting/model.py::grid_local_orientation`. None share a common helper (by
choice, for locality/perf), so they are kept in sync by hand.

### Spectroscopy position mapping (rotation/frame gotchas)
This bug class took a long investigation to pin down, across three
distinct-but-related failure modes. All of it flows from one function:

**`_map_spec_to_pixels(self, spec, header, xpix, ypix, file_key=None,
thumb_crop=None)`** (`main_window.py`, ~9021) is the canonical transform
from a spec's absolute `(x, y)` in nm to a fractional/pixel `(col, row)`
inside a given image's raster, honoring the image's `Angle`/`ScanAngle`
header field. **The load-bearing invariant: rotate the spec's real-nm
offset by the header's scan angle FIRST, then normalize by width/height
SEPARATELY per axis.** Doing it in the other order (normalize each axis by
its own span, then rotate) is mathematically wrong whenever `Width != Height`
— normalizing by two different divisors ahead of a rotation mixes
non-uniformly-scaled components inside the rotation matrix, which **shears**
the result instead of rotating it. This is invisible on square scans (which
is why it shipped unnoticed) and only showed up as visibly smeared/distorted
point clouds on an elongated real scan (2.5nm × 7.5nm).

Three other places independently reimplement pieces of this same rotate-
then-normalize convention, on purpose (not shared via a common helper, for
locality/perf reasons) — if any of them drifts out of sync with
`_map_spec_to_pixels`, the exact shearing/mirroring bug class comes back:
- `gui/spectroscopy/overlays.py`'s `_spectros_near_thumb_pos` — the
  **inverse** transform (pixel click → nm) for thumbnail hit-testing; must
  un-rotate and un-normalize in the mirrored order.
- `gui/spectroscopy/controller.py`'s `_spec_within_extent` — a
  containment-only check (is a spec inside an image's rotated footprint,
  with a margin) reducing to a plain axis-aligned bbox check when angle=0.
- `gui/spectroscopy/controller.py`'s `_spec_frame_offset_info` — computes
  how far *outside* a footprint a spec sits (direction + distance in nm),
  used for the off-frame flagging described below.

**Nanonis row-order gotcha**: `.sxm` scans store rows in acquisition order,
and the header's `Direction: up` vs `Direction: down` changes which end of
the frame row 0 represents — `down` starts at the top (north), `up` starts
at the bottom (south), since the tip physically swept upward. `providers/
nanonis/adapter.py`'s `_extract_scan_channels` flips rows for `Direction:
up` scans so every converted image lands on the same "row 0 = north"
convention `_map_spec_to_pixels` assumes. A single real folder can (and did)
contain a genuine mix of `up`/`down` scans, so this must stay a per-image
conditional flip, never a blanket one.

**Thumbnail-crop-awareness gotcha**: `detect_valid_scan_region`
(`thumbnail_render.py`) trims blank/aborted rows off a thumbnail before
display. Anything that overlays a spec's mapped position onto a *rendered*
thumbnail (not the raw header-pixel grid) must thread that crop through
`_map_spec_to_pixels`'s `thumb_crop` param, then rescale into the final
on-screen pixmap using **separate** x/y scale factors (only rows get
cropped, not columns) — see `_render_spectroscopy_overlays`'s
`w_scale`/`h_scale` in `overlays.py` for the reference implementation.
Skipping this produces a plausible-*looking* but wrong marker position;
it doesn't crash or obviously misbehave, which is what let it slip into the
Spectrum popup's "Position" inset panel undetected initially.

**Off-frame specs** (real position outside every loaded image's footprint —
common for reference points deliberately acquired off to the side) still
need to render *somewhere*, and the established approach has two rules:
- **Don't** let them fall into the `_map_spec_by_spec_extent` "spec-cloud
  bounding box" fallback (meant for degenerate/zero header extents) — for a
  genuinely off-frame point on an otherwise-valid image, that fallback
  stretches a synthetic bounding box to include the outlier and places it
  at a plausible-*looking*-but-wrong interior position, making it
  indistinguishable from a normal in-frame point.
- **Do** clamp it to the true nearest edge/vertex and flag it explicitly:
  `spec['off_frame_direction']`/`spec['off_frame_distance_nm']`, computed
  once in `_assign_spectros_to_images` (`gui/spectroscopy/controller.py`)
  right after the *final* image is chosen — deliberately **not** tied to
  which assignment path/reason picked that image, because the common
  `causal_time` fallback path for `.dat` singles almost always wins before
  the geometric `xy_nearest_extent` check is ever reached, so relying on
  `assignment_reason == "xy_nearest_extent"` alone silently missed most
  real off-frame specs. Any decorative glyph drawn pointing further outward
  from an edge-clamped marker (e.g. the off-frame flag in `overlays.py`)
  also needs a rendering-side inset, or it draws past the pixmap's own
  boundary and gets silently clipped by Qt — invisible, not merely subtle.
  This computation now runs for matrix/grid points too, not just `.dat`
  singles (see "Grid Map Explorer" below for why that took a second pass to
  get right) — the singles-oriented "likely a reference point acquired
  off-frame" summary text is still singles-only, but the underlying
  `off_frame_direction`/`off_frame_distance_nm` fields are computed and
  meaningful for every spec type.

### Grid Map Explorer: anchoring, three rendering conventions, and the hydration-wipe trap
`MatrixSpectroViewer` (`gui/dialogs/spectroscopy_dialogs.py`, the "Grid map"
window for a `.3ds`/matrix file) went through a second, longer rotation/frame
investigation on top of the one above — warped points, a tilted slice view,
and mirrored virtual copies were each "fixed" across several
rounds before the real root causes were found. Every stage below is a
distinct, real bug; earlier fixes were genuine but kept getting silently
undone by a later one in the list, which is why this took so long to nail
down. Worth reading in full before touching anchoring, off-frame handling,
or rendering for matrix data again.

**1. Grid-to-image anchoring was rotation-blind and favored size over
specificity.** `_assign_matrix_reference` (`gui/viewer/loader.py`) — decides
which loaded image a matrix file's points/markers get drawn on — used a
plain axis-aligned bounding-box containment test (ignoring the candidate
image's own `Angle`) and picked whichever single candidate had the highest
*raw point count*. A big, older, unrelated overview scan's oversized
unrotated bbox geometrically "contained" more points from many different
grids than the small, correct, just-preceding zoom scan each grid was
actually acquired on — so on a real 14-grid folder, nearly every grid ended
up anchored to the same 2-3 images regardless of which scan it was really
taken on. Fixed to reuse the same rotate-then-normalize containment test as
`_spec_within_extent`, and to rank candidates by **coverage fraction**
(≥98% of the grid's points contained) before ever falling back to raw hit
count — mirrors the causal-time-first pattern `_choose_image_for_spec`
already uses for singles.

**2. Matrix points never got off-frame status at all.**
`_assign_spectros_to_images` used to skip `off_frame_direction`/
`off_frame_distance_nm` computation entirely for matrix points
(`is_matrix_point` gate), on the assumption a grid's anchored image always
contains it. Once anchoring was fixed (#1) to prefer the correct, small,
specific image, that assumption broke: a grid is routinely physically
larger than the particular scan it was centered on, so a real fraction of
its own boundary now legitimately falls outside its (correctly, tightly)
anchored image. Without off-frame status, those points fell into
`_map_spec_by_spec_extent`'s spec-cloud-bounding-box fallback instead of
clamping to the true edge — a warped/fanned cluster of points. Fixed by
computing off-frame status for matrix points too (the gate now only
excludes them from the singles-oriented summary text, not the computation).

**3. Truthy check vs. key-presence check in `_map_spec_to_pixels`'s
off-frame gate.** Even after #2, warping came back specifically for grids
whose extent was built to exactly touch their anchor image's edge (a
normal "grid over exactly this region" acquisition) — an entire boundary
row/column then sits at *exactly* the image edge, off by floating-point
noise (~1e-15) from the rotation trig. `_spec_frame_offset_info` correctly
reports these as "basically in bounds" (`direction=None`, `distance≈0`),
but the gate was `if not spec.get('off_frame_direction'):` — `True` for
both "confirmed in bounds" (`None`) *and* "never checked" (key absent), so
both got routed into the same bad fallback. Fixed by checking key
**presence** (`if 'off_frame_direction' not in spec:`) instead of
truthiness.

**4. The recurring bug: hydration silently wipes derived per-spec fields.**
This is the one that made every fix above look like it kept "not
sticking," across multiple clean-restart-and-reload cycles. `hydrate_spectro_entries`/`hydrate_spectro_file`
(`gui/viewer/loader.py`) re-parse a spec's source file to fill in
lazily-deferred payload (`channels`/`V`) — and this runs **every time** a
lazily-loaded grid dataset is opened in the Grid Map Explorer, not once
per session. The merge, `_merge_payload_into_spec`, does `target.clear();
target.update(payload)` and keeps only a small explicit "preserved" field
whitelist — and neither `grid_row`/`grid_col` (derived once at scan time
by `_ensure_grid_indices`, since Nanonis's own `.3ds` parser never sets
them) nor `off_frame_direction`/`off_frame_distance_nm`/`assignment_*`
(computed once by `_assign_spectros_to_images`) were in it. So simply
*opening* the Grid Map Explorer — the exact action used to test every fix
above — silently erased their own data on every open, making a genuinely
fixed bug look permanently unfixed. Fixed by adding both groups to the
preserved-fields whitelist; `off_frame_direction`/`off_frame_distance_nm`
needed special-case handling since `None` is itself a meaningful,
deliberately-computed value for them ("checked, confirmed not
off-frame"), not "not applicable yet" — the generic preserved-dict loop
skips `None` values (right for every other field), so these two are
captured before `target.clear()` and restored explicitly afterward,
unconditionally, whenever the key existed beforehand.
**Lesson**: when a bug seems to survive a fix across multiple clean-reload
cycles, suspect a per-spec-dict field getting silently dropped by a
re-merge path (check `_merge_payload_into_spec`'s whitelist first), not
stale caching — especially if the fix touches a field computed once
elsewhere rather than something present in every raw parse.

**5. Three rendering conventions that must visually agree.**
`_draw_image_layer` can show a grid's per-pixel data three ways, each in a
different frame:
- **"Reference image" mode** — the anchor image's own topo channel, drawn
  axis-aligned via `imshow(extent=_header_extent(header))`; markers are
  pre-rotated into that same local-pixel frame via `_map_spec_to_pixels`
  (same convention as the main preview/Position insets).
- **Absolute view** ("Slice at value" etc. with Relative axes off) —
  `pcolormesh(X, Y, metric)` at each pixel's true, rotated absolute nm
  position (`_grid_xy_coords`) — geometrically correct, matches "Reference
  image" mode's content and the true acquisition angle, and renders as a
  tilted parallelogram for any rotated grid (expected geometry, not a
  bug). Uses a standard upward ylim (physically north-up). Note the
  labeled-axis quirk: image views via `_header_extent` label their y axis
  increasing downward while this view labels it increasing upward — the
  *content* orientation agrees even though tick directions differ.
- **"Relative axes"** (the default) — a flattened, always axis-aligned
  view from the grid's own local pixel pitch (`_grid_local_extent`/
  `_grid_local_pitch`), deliberately discarding rotation for a "fill the
  box" look that's also the only representation `Virtual copy`/`Export
  WSxM XYZ` can round-trip into a normal image — they always use this
  local frame regardless of what's on-screen (`_create_virtual_copy_of_map`).

The local/relative flattening initially used a **fixed** "row 0 = top,
col 0 = left" rule, independent of the grid's actual geometry — fine for
small rotation angles, but for a grid whose acquisition angle happens to
swap which end is north/south or east/west (confirmed on a real
157°-rotated grid), that fixed rule disagreed with the absolute/reference-
image convention, so the local view (and everything derived from it -
virtual copies, WSxM export) rendered visibly mirrored relative to the
correct one. Fixed by `_grid_local_orientation`, which derives the needed
row/col flip from the grid's own true coordinates instead of assuming a
fixed direction - applied consistently to the metric array, marker
placement, virtual copy, and WSxM export.

**Which frame the flattening orients into matters as much as the flip
signs.** The first version of `_grid_local_orientation` oriented into
absolute north-up (matching the absolute/pcolormesh view) — internally
consistent, but for a 157° scan north-up is ~180° from the anchor image's
raster, so the relative view looked point-inverted ("mirrored") next to
the thumbnail/preview and Nanonis's own grid viewer, which are all
raster-frame. It now rotates the grid coordinates by the **anchor image's
scan angle** (`_anchor_scan_angle`, same +θ convention as
`_map_spec_to_pixels`) before the direction tests, flattening into the
anchor's raster frame — for angle 0 this reduces exactly to the old
behaviour, so unrotated data is unaffected. `sxm_viewer/reporting/
model.py`'s `grid_local_orientation` is a deliberate port and must stay
in sync. The decisive validation was quantitative and immune to
stale-cache traps: sample the anchor image at every grid point's
`_map_spec_to_pixels` position and correlate against the grid's own slice
values (|r| ≈ 0.98 on the real 30×30 case) — that pins the physically
correct display orientation regardless of what any cached thumbnail shows.

**Getting the flip's sign right requires checking against ground truth —
and make sure your ground truth isn't itself corrupted.** This bit three
times. (1) The flip's row condition was first validated by "does the
local view match its own virtual copy," which passed even with the sign
backwards since both shared the same code path. (2) A later change
"fixed" an apparent vertical mirror of both metric views against the
anchor thumbnail by inverting the absolute view's ylim and the flip
condition — but the thumbnail itself was the mirrored one: its `.sxm`
had been converted to the channel cache *before* the Direction=up row
flip was added to `providers/nanonis/adapter.py`, and that flip commit
never bumped `NANONIS_CACHE_VERSION`, so the stale unflipped array kept
being served. Comparing against Nanonis's own Scan Inspector (the only
truly independent reference) exposed it; the "fix" was reverted and the
cache version bumped (v5) instead. (3) The value readout
(`_sample_current_image`) samples the *unflipped* metric and must undo
the local display flips (`_current_local_flips`) or it reads the wrong
pixel in flipped local views. The sharpened lesson: for orientation
questions, validate against the vendor's own viewer on freshly-converted
data; and any change to conversion-time data-orientation semantics MUST
bump `NANONIS_CACHE_VERSION`, or the fix silently applies only to
never-before-converted files while every existing cache keeps the bug.

### Spectroscopy stale-path reconciliation (`hydrate_spectro_file`)
Specs persist **absolute** file paths, so data copied between machines or
locations (an old `C:\DATA\...` path after the folder moved under the user
profile; or an OneDrive-synced `.sxmviewer_spectro_cache` whose `meta.json`
recorded a relative `../DATA/...` path plus an empty `spec_count:0` result)
leaves specs pointing at dead files. Hydration then reads nothing and the
Spectrum popup shows "No channels" while Grid-map points go missing — and
the poisoned cache keeps serving the empty result until it's cleared.
`_reconcile_spectro_path` (`gui/viewer/loader.py`) guards against this: when
a spec's recorded path is absent, `hydrate_spectro_file` falls back to the
same *filename* inside a folder we currently know about (the loaded spec
folder, `last_dir`, or the parent of any already-resolved spectrum) and
rewrites the spec's `path` to the real location, so the disk cache heals on
the next store. The lookup only runs when the recorded file is actually
missing, so normal loads pay nothing. (The `.dat` parser itself always
records the true on-disk path — the stale paths come from persisted
caches/manifests, not the parse.)

### Spectroscopy browser (`gui/spectroscopy/browser.py`)
Follows the same module-function convention as `gui/viewer/*.py` (plain
functions taking `viewer` first, shimmed onto `SXMGridViewer` as one-liners)
rather than a real class. The floating "Spectro Browser" dock groups specs
per image into three distinct row kinds, each with its own click-to-open
route via `_open_browser_item`:
- **`matrix_dataset`** — one row per `.3ds`/matrix *file* (not one per grid
  point — a 32×32 grid is 1024 points, so per-point rows were never
  viable), grouped by `spec['matrix_dataset']` since one image can host
  several distinct grids at once. Opens `main_window.py`'s
  `_open_matrix_explorer_for_file(image_key, dataset_key=...)` — the
  `dataset_key` param exists specifically so the browser can pick *which*
  grid when an image hosts more than one (previously always silently used
  whichever grid was first in the list).
- **`spec`** (solo) — a single-spectrum position (the majority case).
  Name-first row (matches how users recognize files in Explorer); position/
  channel/assignment detail lives in expandable child rows
  (`_spec_detail_rows`) rather than a separate parent wrapper. Opens
  `gui/spectroscopy/popups.py`'s `_open_spectroscopy_popup` (the "Spectrum"
  trace window).
- **`site`** — a genuine multi-member group (Z-stack / grid-adjacent
  cluster sharing one physical position). Opens
  `_open_spectroscopy_compare_popup` (falls back to a single Spectrum popup
  if it turns out to hold only one spec after all).

**Lesson learned the hard way**: `_open_browser_item`'s spec-open path used
to call `viewer._show_spectro_popup(spec)` behind a `hasattr(viewer,
"_show_spectro_popup")` guard — that method is not defined *anywhere* in
the codebase, so the guard was always `False` and every "Open"/double-click
on an individual spectrum silently did nothing, for the browser's entire
lifetime. A `hasattr` guard passing is not evidence the referenced method
exists correctly — grep for the actual `def` before trusting one,
especially on a code path nothing exercises in normal manual testing.

Status flags (low-confidence assignment, off-frame) render as small
`QPainter`-drawn icons (`_browser_type_icon`/`_browser_status_icon`,
module-level cache) instead of bracketed text tags, reusing the same accent
colors as the on-thumbnail marker overlays for visual consistency.

### Fitting workflows
Both single-spectrum fitting (`SpectroscopyPopup._on_fit_clicked`) and batch
fitting (`_SpectroFitWorker.run`, feeding `KPFMFitTrendDialog`) call a shared
`fit_parabola_bias(V, data)` helper (parabolic bias-spectroscopy fit, e.g.
KPFM contact-potential). Matrix-wide per-pixel fitting is
`gui/dialogs/matrix_fit.py`'s `MatrixFitWorker`, which auto-picks a Δf/KPFM
channel and produces 2D fit-parameter maps, rendered/exported by
`MatrixFitDialog`. `MatrixFitWorker.run()` sizes the output grid from
`spec['grid_row']`/`grid_col'` when present, else falls back to guessing a
**square** grid from `matrix_index` (`side = sqrt(max_idx + 1)`) - Nanonis
`.3ds` entries never set the singular `grid_row`/`grid_col` fields (only
the plural `grid_rows`/`grid_cols` whole-grid dimensions), so every Nanonis
matrix fit used to hit that square-guess fallback, silently corrupting any
non-square grid (confirmed: an 8×32 grid got treated as ~16×16). Fixed by
preferring the plural `grid_rows`/`grid_cols` fields directly over the
square guess - invisible on square grids, which is why it shipped unnoticed
the same way the rotation-order bug did (see "Spectroscopy position
mapping" above).
Both worker classes follow the same
`QObject.moveToThread(QThread)` + `thread.started.connect(worker.run)`
pattern rather than `QRunnable`.

**Fit-map orientation** (an instance of "the one rule" — see "Coordinate
frames & orientation" above): the 2D parameter maps are indexed by raw
`grid_row`/`grid_col`, so, like every other grid-local render, they must be
reoriented into the anchor image's raster frame before display or they show
up mirrored/upside-down next to the reference image (a real `-141°` grid was
a clean vertical flip). `MatrixFitDialog._local_flips` reuses the same
`_grid_local_orientation` logic (ported into `matrix_fit.py`, since
`MatrixFitWorker` runs off-thread with only the raw specs) and applies the
flip consistently to the on-screen maps, double-click pop-outs, virtual
copies, and the `.npz` / WSxM XYZ exports. Two gotchas that made the first
pass look like it did nothing:
- The flip needs the *anchor image's scan angle*, which the off-thread worker
  can't reach — so it's computed in the dialog (`_anchor_scan_angle`) after
  the worker returns `grid_rows`/`grid_cols`/`zero_based` in its payload.
- The coordinate guard must tolerate **partial** NaN. `_grid_xy_coords`
  returns a full `rows×cols` array, so a grid with even one missing/aborted
  cell would short-circuit the whole orientation to unflipped (leaving the
  maps inverted), even though `_grid_local_orientation` averages via
  `nanmean` and copes with gaps fine. Only a *fully* empty coordinate grid
  now disables the flip. This is why the bug looked "still there" on one grid
  (partial coverage) but fine on another (complete) — the completed grid
  flipped correctly, the partial one silently bailed.

### Folder reports (PDF)
"Generate folder report (PDF)..." (toolbar
"Report" action) renders a multi-page PDF overview of the loaded folder.
Two-phase design: `ReportController.collect_payload`
(`gui/controllers/report.py`) snapshots viewer state on the GUI thread —
topo/signal arrays (downsampled), marker positions via
`_map_spec_to_pixels`, assignment metadata, user display prefs — into a
plain-data payload; `ReportWorker` (`gui/workers/report_worker.py`, a
`QRunnable` like `BatchExportWorker`) then runs `build_report_model` +
`render_report_pdf` from the Qt-free `sxm_viewer/reporting/` package
entirely off-thread. Never let the worker touch the viewer.

- **Sample "Regions"** (`model.build_session_regions`) — K-Means over
  (scan center x, y, acquisition time), elbow-selected k: scans from the
  same sample area in the same time window. This is the report's ordering
  unit ("Region 1..N"); the term replaced "chapters" as clearer for SPM use.
- **Image sequences (data-sets)** (`model.build_image_sequences`) — images
  re-scanning the same footprint (center/size/angle within tolerance,
  time-ordered) are grouped and classified by which acquisition parameter
  actually steps: bias series, height series (constant-height scans at a
  series of tip-sample distances), both, or plain repeats (kept only at
  ≥3 members). Per-image `z_level_nm` is the median of the **raw**
  topography channel (`_topo_z_stats_nm` in `gui/controllers/report.py` -
  unfiltered on purpose: background subtraction would erase the flat CH
  plane that *is* the tip height), and only counts as a controlled height
  when the frame is genuinely flat (`z_spread_nm` gate) so drifting
  topography can't fake a height series. `z_rel` is reported relative to
  the sequence's first image (image 0 = 0, positive = retracted).
  Nanonis-converted scans caption their Z channel "Z (Forward)" with bare
  unit "m", which `_find_topography_channel` deliberately doesn't match -
  `_z_channel_index` adds the report-local fallback. `pages_sequence`
  (pdf.py) renders annotated member panels, capped per sequence;
  `collect_payload` backfills panels for sequence members past the
  contact-sheet limit using the same Qt-free detector so the two always
  agree, rebuilding two-channel picks for members that only carried one.
  Members with both current and frequency-shift channels (constant-height)
  render in *paired rows* - per column one file, current on top in
  `Blues_r`, frequency shift below in `gray` (`_SEQ_CLASS_CMAPS`, a
  sequence-page-specific convention that deliberately differs from the
  general current -> Blues rule) - so both channels of the same file
  always correspond vertically.
- **Channel-selection heuristics** (`reporting/channels.py`, deliberate design choices):
  constant-height images show current + frequency-shift/lock-in panels
  instead of flat topo (with a dead-flat-Z retry for untagged CH scans);
  spectroscopy curves prefer freq shift > current > lock-in; z-spectroscopy
  keeps its Z sweep axis; grid slice defaults to the smallest |bias|.
- **Presentation conventions** (also deliberate, in `pdf.py`): current →
  Blues, frequency shift → gray, topography → Blues; KPFM maps bwr (LCPD) /
  gray (c, more negative = darker) / viridis (errors) — matching
  `MatrixFitDialog.PARAM_INFO`; a scale bar on every image panel; grid
  averages are never plotted (representative nearest-centroid member
  spectra instead).
- `reporting/model.py`'s grid helpers (`grid_dims`/`spec_grid_row_col`/
  `grid_local_orientation`/...) are deliberate ports of the
  `MatrixSpectroViewer` methods (which live on a QDialog and can't be
  imported Qt-free) — keep them in sync, especially the local-frame flip
  rules (see "Grid Map Explorer" above).

### Figure recipes: data-aware multi-panel figures (`figures/`, `gui/controllers/figures.py`, `gui/dialogs/figure_window.py`)
The successor to the publication canvas for the figures SPM users actually
make. Design goal (from the user): **one click from data to a finished,
publication-ready figure** - no dialog, no settings up front; every
adjustment made in place on the figure or in the always-visible Format bar.
The canvas stays reachable from a figure's ⋯ menu ("Open in canvas") and
the New Figure button's menu, marked legacy: it is frozen (no new
features) and planned for removal once custom figures cover it (next:
dropping thumbnails onto a figure, per-cell channel choice).

**Three recipes, one window.** Thumbnails carry clickable chips (painted in
`_decorate_thumbnail_pixmap`, hit-tested in `thumbnail_ui._handle_series_chip_click`
via the shared `series_chip_rect` / `zoom_chip_rect`); right-click offers
the same plus *Figure from selection* for anything else:
- **Height series** (green `z×N`): CH frames re-scanning one footprint at a
  ladder of tip heights; frames by z, `Δz = +20 pm` headers.
- **Bias series** (amber `V×N`): same footprint, only the bias steps (CC
  topography or CH); frames by bias, signed `V = +300 mV` headers, setpoint
  line when it changes.
- **Overview + zooms** (blue `⊡N`, may sit beside a series chip): a frame
  with smaller frames inside it; overview left, zoom grid right, each zoom
  outlined + lettered on the overview in its own colour.
N counts distinct acquisition conditions (`levels`), not files.

**Custom figures** (mode `"custom"`, replacing the canvas as the way to
start a free selection): the toolbar's **New Figure** split button
(`main_window_toolbar.new_figure_from_selection` / `populate_figure_menu`;
the old "Open Canvas" button, now only in its menu as legacy) and
*Figure from selection* call `FigureController.open_selection` -> a selection
that is exactly one series opens that series, anything else `open_custom`.
A custom figure has a stable id `custom-<hex>` and keeps `mode`, `keys`
(its frames **in the user's order** - `build_spec` mode `custom` never
sorts or collapses) and `name` in its state, so adding (`add_frames`),
dropping and reordering frames (`_move_frame`; drag a header,
`_frame_drop_target` draws the insertion line) never loses its edits.
`_rebuild` re-orders the cached `_members` from `state["keys"]` - they are
not re-read on a reorder. `recent_custom_figures` lists them (toolbar
menu); `open_saved` reopens by id, offering to load the folder if frames
are missing; `save_figure_file` / `open_figure_file` write / read
`*.sxmfig.json` (`{kind: "sxm_figure", fid, mode, keys, root_key, state}`;
the file wins over the remembered state). Series figures can be copied
into one (⋯ > "Edit as a custom figure"). The legacy `"sel-<hash>"` ids
are no longer created.

**Per-image processing** (`state["cell_edits"]`, keyed by
`clim_key(row_label, frame_key)` = one channel of one frame): optional
`filters` (a full step list, `[]` = none) and `zero`; an absent key follows
the viewer (per-file `thumbnail_filters` / `_is_image_zeroed`) - the viewer
itself is never changed. `FigureController.panel(key, idx, edits)` runs own
filters on `raw_array` (normalised, unfiltered); the window's panel cache is
keyed `(key, idx, repr(edits))`, so two rows showing the same channel stay
independent. A per-panel colour range in *shared* (`row`) mode detaches
that panel from the row (`own_scale` reason `"user"`, excluded from the ⓘ
"different baseline" list, no tag when colour bars are off). UI: the
panel's right-click section (`_cell_menu`: Filters… = `CustomFilterDialog`
live on that panel via `_live_edit`, quick tilt/plane2/flatten/line
toggles, Zero, range editor with `_HistogramStrip`, copy to row, reset),
keys on the hovered panel (`_cell_key_press`: A / Z / G / T, Shift removes,
a selected molecule keeps X/Y/Z), a screen-only `fx` badge
(`view.edit_boxes`) and ⓘ entries per edited panel.

**Layers** (same split as the folder report - Qt-free core, GUI gathers):
- `figures/height_series.py` (Qt-free; despite the name it hosts all
  recipes): detection (`find_height_series`, `find_bias_series`,
  `find_zoom_families`), `build_spec` (frames + channels + headers),
  `figure_layout` (deterministic inch boxes - the window hit-tests from
  these, never asking matplotlib), `render` (one matplotlib renderer for
  screen *and* every export), range helpers. `figures/svg_layers.py` turns
  the SVG into Inkscape layers by artist `gid`.
- `FigureController` (GUI thread): series index for the chips, inputs
  through the preview pipeline (`panel()`: filters, adjustments, quarter
  turns, Zero, aborted-row mask), `data_signature()` for the live link,
  remembered state/styles/export settings, Obsidian menu.
- `FigureWindow` + `_FormatPanel` (pure UI against that controller).

**Detection rules (each one came from a real folder):**
- Height/bias series reuse `reporting.model.build_image_sequences` so the
  report and the figure agree on what a series is. Height: only CH-tagged
  frames, `z_level_nm` = the **HUD** height (`ch_height.ch_abs_z_nm`, so
  captions can't contradict the preview), `z_spread_nm` = the CH tag's
  `rng_nm`. **A CH tag is not flatness**: real frames step Z mid-scan (rng =
  the step) or carry a z = 0 region (rng ~190 nm); the report's 20 pm gate
  drops their z and they come back as `excluded` ("Include anyway" in the
  ⓘ note).
- Bias: CH and CC/untagged frames go through the detector **separately** -
  mixed, a CC bias ladder and a CH height ladder of one spot become one
  "bias+height" group.
- Zooms: rotation-aware containment of all four corners (4 % overhang),
  parent = the *smallest* frame with >= 2.5x the area; one level deep (a
  zoom with zooms gets its own family). Geometry mirrors
  `geometry.spec_mapping` (`_local_to_abs` = `pixel_to_nm`, `_abs_to_local`
  = `map_spec_to_pixels`, rotate-then-normalise) - keep them in sync.
- Repeats collapse to the **latest** frame per `condition_key` (bias mV,
  z pm, setpoint 3 sig.; zooms: same footprint), `show_repeats` brings them
  back. Real series were often mostly repeats (13 frames = 3 heights).
- Stopped scans: `panel()` NaN-masks rows outside `detect_valid_scan_region`
  (a stopped scan's constant z = 0 rows beside -185 nm topography flattened
  whole rows). The detector's tolerance is relative to the frame's range
  (`min(1, span) x 1e-10`): it used to be absolute, so current in A
  (row ptp ~1e-11) read as "never scanned" - half-blank current panels and
  complete CH frames pruned as stopped. Frames < 25 % scanned are pruned (`excluded` reason
  `"aborted"`) - only knowable after reading the array, so
  `_idle_check_scans` checks chip members a few per idle tick after the
  index is built and repopulates thumbnails if a count changes; the scan
  fraction is part of the index signature; `_find_aborted` in the window
  is the backstop.
- `ensure_series_index()` runs at the top of every
  `populate_thumbnails_for_channel`, cached on (files, tag, abs_z_pm,
  rng_nm, scan fraction); `_z_nm` only asks the HUD for CH frames (others
  could sample a file each).

**Spec model:** `spec["columns"]` are always the *frames*, `spec["rows"]`
the *channels*; `style["transpose"]` / the zoom layout only change where
they go. Channels are identified by **caption** (`row["label"]`); state
keeps `rows` (ordered captions, None = defaults from `pick_image_channels`),
`row_cmaps`, `clims` keyed by `clim_key(label[, frame_key])`. Pseudo
channel `hs.BEST` ("Best channel per frame", `kind == "mixed"`) gives each
frame its own first pick - the zoom default when frames were taken in
different modes (CC overview topography + CH zoom current); mixed rows are
scaled per cell in their own unit. Unknown vendor captions fall back to the
unit (`_CLASS_OF_UNIT`: A -> current, Hz -> freq, m -> z).

**Ranges** (never pool across baselines): `baseline_outliers` flags frames
whose *median* sits > 3 typical spans (75th pct of frame spans) from the
row median ("offset", own scale) or that have no signal ("flat", dimmed,
tagged with the real level); `_flat_row_level` catches a flat channel or a
single flat frame (CH topography). Signal rows: `pooled_range` with **equal
weight per frame**; topography rows: `typical_range` (median of frames' own
2-98 %, so one tip crash saturates instead of darkening the rest) and
relative (each frame minus its p1). One display unit per row (the viewer's
`_scale_unit_for_display` picks per image - 4 pA next to 0.4 nA). Real
data that forced this: a Δf series with 5 PLL-unlocked frames at
-12 033 Hz and one mid-scan jump. Tick precision follows the *span*
(`_tick`), or 90.06…90.14 prints "90.1" twice.

**Style** - one dict, `STYLE_DEFAULTS` / `resolve_style` (bool
`scale_label` migrates). Keys in `RECIPE_KEYS` (ranges, colour bars, scale
bar, header, layout) have per-recipe defaults (`RECIPE_DEFAULTS`: zooms =
own ranges, no bars) and are remembered per recipe (config
`figure_style_recipe`); everything else globally (`figure_style_default`);
the header is never carried over (`"auto"` = Δz / V / size per recipe).
To add an option: `STYLE_DEFAULTS` + `render`/layout + a `_FormatPanel`
card row + the matching right-click menu. Notable ones:
- `page`: `PAGE_PRESETS` (Nature 89/183 mm, Science 55/120/183, ACS
  3.25/7 in, generic 85/120/178) set the width **and** `page_text_style`
  (text pt, letter pt/style). The panel width is solved from the page
  (every layout width is affine in it: two trial layouts give the answer);
  if panels would drop below `_MIN_PANEL_IN` the frames wrap into more rows.
  Fonts stay in absolute points; derived sizes go through `pt()` (>= 5 pt).
- `font_family` (Arial default, `available_fonts` journal fonts first):
  `render` wraps `_render` in `rc_context(font_rc(st))` - text props are
  fixed at artist creation, so this covers screen, PDF, SVG, PPT.
- Colours of text over images: `letter_color` / `scale_color` "auto" pick
  white or black **per letter / bar** from the pixels under it, through the
  live colormap + range (`auto_color` reuses the preview's Qt-free
  `gui/canvases/scale_bar_contrast.py` WCAG rule - one contrast rule app
  wide). Per-letter overrides live in figure state `letter_colors`
  (`letter_key(row_label, frame_key)`), set from a letter's right-click
  menu. Outlines take the opposite tone of whatever colour the text got.
- Letters: `letter_pos` corners or `out_tl`/`out_tr` (outside, own line
  above the header - `_letter_line_y`; the layout adds `_outside_letter_h`
  to the header band / row gaps); outline optional.
- Scale bar on every image by default; length label `first` = once per
  distinct length (zoom frames of different sizes all stay labelled),
  `all`, `none`; `scale_font`; thickness from the panel's **shorter** side
  (from the height, tall frames got slabs).
- Colour bars: `cbar_style` outside (ticks) / inside (values in the bar,
  contrast-coloured) / minmax (gradient labelled "min"/"max");
  `cbar_label_pos` inside = channel name centred in the bar; `cbar_scope`
  row = one bar per channel even with per-image ranges (min/max, not
  draggable). `cbar_dims` is shared by layout and renderer; end labels are
  aligned inside their bar so stacked bars never overlap.
- `rotate_tall`: portrait panels turned to landscape for the figure only
  (`rotate_tall_panels`: data rot90, width/height swapped, `turns` bumped so
  zoom outlines follow). `panel()`'s width/height are already the
  *displayed* (viewer-turned) size - never swap them again for `turns`.
- Headers can be `none`; `channel_labels` hides the row names (bands
  shrink); `ref_bold` (off by default).

**Molecular models** (Molecule card, right-click a model): one list of
`Molecule.to_dict()` in figure state `molecules`, positioned in **panel
nm** (each panel's axes are `0..w, 0..h` nm, y up), so one placement is the
same physical spot on every channel / frame of a series - the point is
comparing channels against the same model. Drawn with the shared renderer
(`_MoleculePainter` -> `molecules.render.draw_molecules`; palette / bond
mode / shadow / H from `controller.molecule_settings()` = the preview
canvas, except the palette when style `mol_palette` != "preview" -
`FigureWindow._mol_settings`) at `scale=panel_scale(ax) x mol_size`, because marks are sized in
points for the big preview. "From preview" (`controller.preview_molecules`) takes the
live preview / `viewer.molecule_overlays` and subtracts the frame's display-
extent lower-left corner (preview data coords are extent-based, y up).
Interaction cost: a drag caches the figure without molecules once
(`copy_from_bbox`) and each move redraws only the molecule artists
(`restore_region` + `draw_artist`): ~80 ms for 12 panels instead of a ~0.8 s
re-render. `render(draw_molecules=...)` is the Qt-free hook; `data_boxes`
in the layout (post-`apply_aspect` image boxes) map clicks to panel nm.
Molecules get their own Inkscape layer. Interactive rotation must match the
preview gesture for gesture (it was briefly lost - keep it): drag = move,
Shift+drag = spin about the model centre, Ctrl+Shift+drag or middle button
= free 3-D (0.5 deg/px: dx -> Y, dy -> X, y up), plus the X/Y/Z orientation
gizmo (`molecules.interaction.gizmo_vectors`; screen-only, painted by the
view top-left while a model is selected; inner disc = tumble 0.45 deg/px,
rim = spin) - all of it from `molecules.interaction`, so it cannot drift. All modes share the cached-background
redraw; keys X/Y/Z rotate 5 deg, arrows nudge.

**Window** (`FigureWindow`): figure on white "paper" (= the export) in a
scroll area, `_FormatPanel` under it as **one row of chips** (icon +
the category's current state in words, greyed when the feature is off, `•`
= differs from the recipe default set via `set_defaults`); a chip opens its
card as a `_CardPopover` (Qt popup, edits apply live) above it. History:
a right-hand panel (cramped, cropped text) and tabs (state hidden) were
rejected, then a strip of always-open cards needed sideways scrolling and
cropped rows - the chips keep every category and its state visible at no
height. Card grammar: whole-feature on/off = header switch (letters,
`_scale_on` <-> `scale_bar` "off", molecules; colour bars = bold section
check), 2-6 choices = `_Segmented` (QComboBox data API, so `_emit`/`sync`
use `_is_choice`), long lists = combo, numbers = spin with unit; ↺ per card
resets its keys (`_card_keys`). A left click on a letter / scale bar, and
"All ... options…" in the right-click menus, open the card at the cursor
(`FigureWindow.open_format_card`). `_sync_format_panel` checks
`isHidden()`, not `isVisible()` - the first sync runs before the window is
shown, and the panel used to stay blank until the first edit. Opens fitted to the window (actual size is a per-window view,
never remembered). Direct manipulation: Δz header = reference; hover × =
drop frame; channel label = switch/move/remove/recolour; "+ channel";
colour bar drag = contrast, click = Min/Max popover, double-click = auto;
right-click any element = its options; **Ctrl+wheel** = size of what is
under the cursor (text with letters in proportion / bar width / scale-bar
label / letters; Ctrl+Shift = panels), coalesced into one undo step. Live
link: `data_signature()` polled + on activation. Outputs: Copy (PNG+SVG /
PNG / SVG), PowerPoint (new/current slide, shapes/picture - always sent
with `width`/`height`/`image_size` = the figure's size in points; the
bridge's defaults fit everything into a 600 x 450 pt box, which is why it
was not at real scale), Obsidian (via
`viewer.export.export_to_obsidian(render=...)`), Inkscape (layered SVG,
opened from a temp folder), Save PDF / layered SVG / PNG / LZW TIFF / EPS,
300-1200 dpi, transparent background.
**Scan pixels are never resampled on export** (`_export_figure` sets
`interpolation="none"` on every scan image; `_vector_save` turns off
`image.composite_image`): SVG / PDF / EPS - and so PowerPoint shapes,
Inkscape and the clipboard SVG - embed each scan at its own pixel size,
tagged `image-rendering: pixelated` by matplotlib. With the old
`"nearest"` matplotlib resampled vector images to the figure's 100 dpi
(a 480 x 160 scan arrived as 110 x 37 px). Raster outputs (PNG, TIFF,
PNG clipboard, PowerPoint picture, Obsidian) raise the chosen dpi to
`native_dpi(fig)` (rounded up to 50) when it would drop scan rows or
columns, and say so in the status line. Verified pixel-exact: the
embedded PNG equals `im.to_rgba(data)`.

**Lessons (what made this take longer than it should have):**
- Validate against several *real* folders (Anfatec local, Nanonis-converted,
  network): every robustness rule above came from one of them, none from
  synthetic data.
- Matplotlib SVG puts an image's `gid` on the `<image>` element, not a
  `<g>` - match ids on any element (`svg_layers`).
- Anfatec headers: `Bias` is in `BiasPhysUnit` (mV) - `_header_bias_v`
  converts (report and figures both used to print -2000 V); tunnel current
  is `It_to_PC` / `It_ext` (unit A), which `classify_channel` used to call
  "other" (so CH report pages had no current).
- An exception escaping a Qt handler aborts the process when no
  `sys.excepthook` is installed (the app's `cli.py` installs one; headless
  test scripts don't) - `FigureController._open` catches and reports.

### Molecular overlays: one model, one renderer, one gesture map (`sxm_viewer/molecules/`)
A molecular model looks and behaves the same in the preview, pop-ups,
figure windows, preview exports and the publication canvas because each of
them is a thin host over one Qt-free package:
- `model.py` — `Molecule` (load .xyz/.pdb/..., `angles`/`offset`/`scale`,
  mirrors, `get_transformed_coordinates`, `to_dict`/`from_dict`,
  `reset_to_file_state`), palettes, radii, render-style profiles, colour
  resolution. `gui/canvases/molecular_overlay.py` re-exports all of it
  (old imports keep working) and holds only the Qt dialogs.
- `render.py` — `draw_molecules(ax, molecules, settings, *, scale=1.0)`,
  the only molecule renderer: `settings` = palette / bond colour mode /
  default bond colour / shadows / hydrogens (`DEFAULT_SETTINGS`), returns
  entries `{mol, ax, scatter, atom_underlay, atom_order, shadow,
  line_underlay, lines}` (`artists_of` flattens them). `scale` multiplies
  screen-sized marks (marker areas x scale^2, line/edge widths x scale);
  `panel_scale(ax, user)` derives it from a printed panel's width. Bonds are
  round-capped/joined in every style so they read as one continuous stick.
  Light atoms and bond halves (luminance > `LIGHT_LUMINANCE`: white H, the
  near-white default bond colour) get a dark rim/casing so they read on a
  white figure without sampling the background; same-colour split bonds are
  drawn as one segment (overlapping translucent round caps made dark dots).
  `BOND_STYLE_WIDTH` thin/default/thick = 0.45/1/2 - the depth cue alone
  varies widths up to 3x, so the styles must be further apart than that.
- `interaction.py` — what a gesture *does*: `drag_mode` /
  `drag_mode_for_key` (drag = translate, Shift = spin, Ctrl+Shift or middle
  = 3-D tumble 0.5 deg/px), `drag_update` (pose from press pose + pointer;
  pixel deltas and gizmo vectors y-up), `gizmo_hit_mode` (inner disc
  < 0.58 = tumble 0.45 deg/px, rim = spin), `orientation_matrix`,
  `axis_screen_vectors`, `gizmo_vectors`, `atoms_hit` (14 px),
  `key_rotation` (5 deg), `apply_transform`.
Hosts only translate their events (matplotlib data/pixel coords in
`MultiPreviewCanvas`; Qt positions -> panel nm in `FigureWindow`) and own
undo, selection and redraw strategy. A new look or gesture goes into the
package, never into a host. Verify a change by pixel-diffing preview
renders per style and replaying gestures through the real handlers (the
refactor that created the package was checked that way: identical pixels,
identical poses). Found by that replay: the preview's Ctrl+Shift tumble had
never worked - it tested `'control' in key`, matplotlib sends `ctrl+shift`.

### CH height label: z_abs or Δz to a reference (`gui/viewer/ch_height.py`)
The acquisition HUD's CH line is resolved at *draw* time, not baked into the
view: `MultiPreviewCanvas._ch_height_overlay_text` reads the view's
`meta["acquisition_z_abs_nm"]` and asks the owner via
`set_ch_height_label_callbacks` (wired by `ch_height.install_canvas_callbacks`
for the main preview and every pop-up). So switching mode/reference is a
redraw (`ch_height.refresh_labels`), never a view rebuild. Text formatting is
the Qt-free `utils.units.format_ch_height_label` (Δz > 0 = tip retracted,
same sign as the folder report's `z_rel_pm`). State: `ch_z_label_mode`
(`abs`/`rel`), `ch_z_reference` (`{file_key, z_abs_nm}` - the height is
stored so it works with the reference file unloaded), and
`ch_label_in_publication` (Publication preset hides the HUD; this opt-in
draws only the CH line, figure-style, via `_draw_publication_ch_label`).
Persisted in config and sessions; reference changes push `("chref", mode,
ref)` onto `_adjustment_undo_stack`. One shared menu builder,
`ch_height.populate_menu`, feeds Display > CH height label, the thumbnail
right-click menu and the preview's Display Options. The reference z comes
from `_build_acquisition_overlay_info`, i.e. the same raw source as the
displayed z_abs, so Δz never picks up filter/Zero offsets.
`reference_nm` re-reads a loaded reference's live height (the stored value
is only the unloaded-file fallback). Manual CH tags saved before `9a70df9`
carry a bogus `abs_z_pm` (~0, the displayed current channel's median);
`_heal_manual_ch_abs_z` recomputes them from the topography channel once
during `_auto_detect_tags_for_folder` and marks them `abs_z_src: "topo"`.
Look is user-editable via the menu's "Label style..." (`dialogs/ch_label_style.py`,
pure UI emitting `style_changed`): `viewer.ch_label_style` (config) holds only
keys the user set; `MultiPreviewCanvas._ch_label_style` overlays them on the
per-mode defaults (boxed HUD / outlined Publication text) and `_draw_ch_label`
is the single CH label renderer for both modes. CC HUD text is unaffected.

### Quarter-turn image orientation (R / Shift+R)
A per-file *display* state, `viewer.image_orientation` (`file_key ->`
clockwise quarter turns 1..3), distinct from Crop/Rotate below (which makes
resampled `[edit]` copies). Pure helpers live in
`sxm_viewer/geometry/orientation.py`. The turn is applied in one place,
`_apply_adjustments_for_channel` (array via `rotate_array`, extent via
`rotate_extent`, which swaps spans for odd turns), so preview, popups,
exports, PowerPoint and the publication canvas all get it; the folder report
deliberately doesn't. Raw arrays, headers and spec positions stay in the
acquisition raster frame — so this is another instance of "the one rule"
above, and every raster-frame overlay drawn on a turned image must be turned
too:
- **Thumbnails** call it with `orient=False` and turn *after*
  `detect_valid_scan_region`, so `thumb_crop` (r0/r1) stays raster-frame for
  `_map_spec_to_pixels`; the turn is part of the thumbnail data-cache key.
  `_render_spectroscopy_overlays(turns=...)` lays markers/footprints out in
  the unturned box, then turns positions (and off-frame compass letters,
  `_turn_compass`) onto the pixmap; `_spectros_near_thumb_pos` un-turns the
  click. Only `_decorate_thumbnail_pixmap` passes `turns` — other callers
  draw on unturned pixmaps.
- **Preview** markers: `preview._turn_spec_pixel` turns each
  `_map_spec_to_pixels` result before it goes into `spec_pixels`.
- Turning re-renders with `_suppress_profile_restore` and resets zoom
  (profile/angle lines are in the old frame's data coords). Per-file
  molecules and quick-crop history are *not* migrated.
- `_last_display_extent` (only read by Crop/Rotate, which works on the
  unturned `_last_base_array`) is kept in the acquired orientation.
- Virtual copies: channel copies (`create_virtual_channel_copies`) re-read
  the source's raw arrays, so they inherit the source's turn; view/crop
  copies are built from the displayed (already turned) array and must
  *not* get an orientation entry, or they'd be turned twice.
- Undo: `("orient", {key: old_turns})` on `_adjustment_undo_stack`; sessions
  persist `image_orientation`; the thumbnail chip is the teal `↻90°` badge
  next to ADJ.

### Crop/Rotate (Image menu) vs quick-crop — two tools, one output convention
Both crop tools are non-destructive and produce **virtual copies**
(`_create_virtual_view_copy` → `_processed_views`); the original file's
data is never altered.
- **Image > Crop/Rotate** (`gui/dialogs/image_adjust.py`,
  `MainWindow.on_adjust_image`): geometry (crop/rotate/flips) creates an
  adjacent `[edit]`-tagged copy on Apply, resampled in the rotated frame
  by `thumbnail_render.resample_geometry` (quick-crop's approach — output
  never contains NaN padding; the crop rect is clamped to
  `largest_inscribed_rect`). The dialog's workspace always shows the
  rotated image and the crop rect lives in that rotated frame; the result
  preview uses the same `resample_geometry`, so preview == result. Clip/
  gamma/cmap are **live display adjustments**: clip+gamma go to
  `viewer.image_adjustments` (tone-only now — geometry is never written
  there anymore), cmap to `per_file_channel_cmap`. Every
  `image_adjustments` mutation must run the invalidation trio
  (`_refresh_adjusted_channel`: per-file thumbnail cache invalidation +
  grid repopulate + preview re-render) or stale pixmaps linger — that was
  the tool's historical "persists after undo/reset" bug. Adjusted images
  show "• adjusted" in the preview title and an amber ADJ chip on the
  thumbnail; Image > "Reset display adjustments" (and the thumbnail
  context menu) clears the spec. All these actions push onto
  `viewer._adjustment_undo_stack`, consumed by `_undo_last_adjustment` in
  the global Ctrl+Z order (see "Global undo" under Keyboard shortcuts).
  **Legacy sessions**: `apply_adjustment_spec` keeps full geometry
  support forever so old geometry specs render unchanged; opening the
  dialog on one seeds its controls (migration-on-touch) and Accept
  converts it to a copy.
- **Quick-crop template** (Ctrl+Shift+C, `gui/controllers/quick_crop.py`)
  is unchanged: `[crop]`-tagged copies via `_on_preview_crop`.
- **`resample_geometry` frame conventions** (get these wrong and results
  are silently mirrored/rotated the wrong way): it works in the *display*
  frame — display array = `flipud(raw)`, `origin='lower'`, extent
  `(0, w, 0, h)`, y increasing up — with positive `rotate` meaning CCW in
  that frame; `crop_rect` is `(left, right, bottom, top)` in rotated
  display-frame pixels (or `None` → largest inscribed rect); flips apply
  *before* rotation. Trap when writing expectations for it: `np.rot90(a)`
  is CCW in matrix-print orientation, which is **CW** in the
  origin-lower display frame — an exact 90° CCW display rotation of a
  display array `d` equals `np.rot90(d, 1)` only after accounting for the
  flipud round-trip (`out = flipud(rot90(flipud_input...))` inside the
  function); validate against the function's own identity/90° cases, not
  a bare `rot90` guess.

### Central colormap registry (`sxm_viewer/cmap_registry.py`)
Qt-free single source of truth for colormaps. Registers the custom
`gui_amber_theme` cmap (stops synced by comment with `gui/theme.py`'s
AMBER tokens) and, when importable, the optional pratiman-91 `colormaps`
PyPI package (runtime-detected; never a hard dependency — the package
lazily registers its cmaps into matplotlib on attribute access, which
`_register_extra_package` accounts for). All pickers enumerate via
`all_cmap_names()`/`featured_cmap_names(context)` (curated per-context
shortlists live in `_FEATURED` here, not in widgets); name→Colormap
resolution goes through the never-raising `get_cmap(name, fallback)`
(shared-object cache — never mutate a returned Colormap). Renderers that
draw data imagery resolve through `effective_cmap`/`effective_cmap_name`,
which honor the "Full amber imagery" display-time override
(`set_forced_cmap`, driven by `main_window._sync_forced_cmap`); view
dicts, `per_file_channel_cmap`, sessions, and file exports always keep
the user's true cmap names. Display → "Extra colormaps..." shows the
optional-package status/install hint.

**Base-name / reversed-flag model.** `_FEATURED` entries are
`(base_name, is_reversed)` pairs — the `_r` suffix never appears in the
primary data lists. `split_cmap_name`/`join_cmap_name` translate to/from
the legacy `X_r` string form, which stays the persistence/interchange
format (`per_file_channel_cmap`, sessions, exports, combo texts);
`featured_cmap_names` is the unchanged legacy string view over
`featured_cmap_entries`. `get_colormap(name, is_reversed, fallback)`
resolves a pair with a *programmatic* `Colormap.reversed()` (an `_r`
suffix in `name` XOR-folds into the flag) and, on first reversal of a map
with no registered `_r` twin (extra-package maps), registers the variant
under its joined name so plain-string call sites keep resolving it.
`base_cmap_names()`/`grouped_base_cmap_names()` enumerate with `_r` twins
collapsed.

### Colormap gallery (`gui/colormap_manager.py`, `gui/colormap_gallery.py`, `cmap_sorting.py`)
The 🎨 button next to the preview cmap combo opens `ColormapGalleryDialog`.
`ColormapManager` (viewer-agnostic `QObject` at `gui/` root, like
`theme.py` — not a `gui/controllers/` class) holds `(selected base name,
is_reversed)` as two independently-settable axes, twice: a *pending*
selection (`pending_changed`, drives live preview) and an *applied* one
(`applied_changed`, emitted only by `apply()`; `revert()` rolls pending
back). `ColormapGallery` is a dumb `QListView` card grid over base names
(never `_r`); its delegate builds gradient pixmaps lazily on first paint,
so with the extras package (~970 maps) only the visible viewport renders.
It emits `name_selected`/`reverse_requested` independently and paints
selection purely from `set_selection` (NoSelection mode — the manager is
the single source of truth; a 🔄 click toggles the selected card or picks
a new card reversed). Ordering is delegated to
`cmap_sorting.ColormapSorter`, which returns `(section_label, [names])`
sections for a strategy (`functional`/`similarity`/`usage`/`alphabetical`)
— functional/color inject full-width section headers into the grid.
Colormap metadata (`cmap_sorting.classify`) is auto-derived: matplotlib's
documented category tables → name rules (colorcet class letter,
qualitative family stems) → LUT sampling (lightness profile +
discreteness), so it also classifies extra-package maps a static table
never could. Usage stats come from a provider callable
(`main_window._cmap_usage_stats`, persisted in config as
`colormap_usage`), keeping the sorter free of config/Qt.
Main-window wiring (`on_open_colormap_gallery`): pending changes recolor
the current preview via the cheap `set_cmap_for_current_views` path only;
Apply commits to the **thumbnail selection** (Ctrl+A / multi-select,
falling back to the highlighted/previewed image) through
`_set_thumbnail_entry_cmap` — the same targeting as the Thumb cmap combo
(`on_thumb_cmap_changed`) — so thumbnails and preview both update, and it
records usage + persists the sort strategy (`colormap_sort_strategy`).

### Favourite (default) colormaps
Small ★ buttons next to the thumbnail/preview colormap combos
(`main_window.py`) and the Grid map's colormap + color-cycle selectors
(`spectroscopy_dialogs.py`, mode-aware: saves `grid_metric` or
`grid_reference` depending on the active view) persist the current pick as
the user's default. Stored in config as `favorite_cmaps` (dict:
`preview`/`thumbnails`/`grid_metric`/`grid_reference`) and
`favorite_color_cycle`; helpers on `SXMGridViewer`
(`get_favorite_cmap`/`set_favorite_cmap`/`set_favorite_color_cycle`/
`clear_colormap_favorites`). Favourites win over merely-last-used values
at startup; Display → "Reset colormap defaults" clears them; the folder
report reads the same favourites via `payload["prefs"]`.

### PNG / Obsidian export (as displayed)
"Export PNGs" (toolbar + thumbnail menu) and "Export to Obsidian" share
`gui/viewer/export.py:render_targets_as_displayed`: each target (thumbnail
selection in grid order, else the previewed image) is loaded into the main
preview and `preview_canvas.fig` is saved, so the active display preset,
overlays (and the H-hidden state), per-file cmap/contrast/Zero/turn apply;
the previous preview is restored afterwards. One PNG per image (the whole
displayed figure, extra channel views included). Obsidian support is
file-based and Qt-free (`sxm_viewer/integrations/obsidian.py`): vaults from
`%APPDATA%/obsidian/obsidian.json`, the active note from
`<vault>/.obsidian/workspace.json`, link style/attachment folder from
`app.json`; embeds are appended to the end of the note. Options are
remembered in config `obsidian_export` (image folder per vault).

### Cross-cutting conventions

- Keep GUI code free of provider internals; providers (`providers/`) must not
  import GUI/Qt code, so format support can be tested/reused headlessly.
- Cache folders `.sxmviewer_nanonis/` are generated next to data and are
  gitignored, as are `.sxm`/`.dat` sample files themselves — sample/test data
  lives outside the repo (e.g. a local `data_local/`) or under `samples/` for
  curated, non-personal examples.
- Header cache and channel-data caches are bounded (`CHANNEL_DATA_CACHE_LIMIT`,
  `FILTERED_CACHE_LIMIT`) and versioned (`HEADER_CACHE_VERSION`) — bump the
  version constant when changing the cached data shape so stale caches are
  invalidated rather than misread.
- Sibling modules under `gui/spectroscopy/*.py` (`browser.py`,
  `controller.py`, `overlays.py`, `popups.py`, `summary_dialog.py`) import
  each other directly as plain Python modules (e.g. `from . import popups
  as spectro_popups`) rather than only going through a `viewer.*` shim —
  this is safe and already established (`gui/controllers/spectro_compare.py`
  does the same) since none of them import `main_window.py`, so there's no
  circular-import risk; prefer this over adding a new phantom
  `viewer._method` shim that nothing defines (see the browser's dead
  `_show_spectro_popup` lesson above).
- Assignment/position metadata computed once during spec-to-image
  assignment (`_assign_spectros_to_images` in `gui/spectroscopy/
  controller.py`) — `assignment_reason`, `assignment_confidence`,
  `assignment_summary`, `off_frame_direction`, `off_frame_distance_nm` —
  is written directly onto each spec dict and treated as authoritative
  everywhere downstream (overlays, the browser, tooltips, filters) rather
  than being recomputed per call site. When adding a new derived
  per-spec flag, prefer this same pattern (compute once at assignment
  time, read everywhere else) over recomputing it from geometry on every
  render. **This pattern has a real failure mode**: any such field also
  needs to be added to `_merge_payload_into_spec`'s (`gui/viewer/
  loader.py`) preserved-fields whitelist, or a later `hydrate_spectro_
  entries` re-parse (triggered just by opening certain dialogs on a
  lazily-loaded spec) will silently wipe it back to unset - confirmed as
  the root cause of a bug that looked "not fixed" across several clean-
  restart cycles because the fix's own data kept getting erased by the
  next dialog open, not by anything stale (see "Grid Map Explorer" above).
  `grid_row`/`grid_col` (derived once from `matrix_index` by
  `_ensure_grid_indices`) need the same protection for the same reason.
