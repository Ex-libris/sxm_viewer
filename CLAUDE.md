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

- Legacy launch shim: `python sxm_grid_viewer.py` (forwards to `python -m sxm_viewer`).
- There is no lint/format/test tooling configured for the main package (no
  `pyproject.toml`, `setup.cfg`, `pytest.ini`, or CI test workflow) — verify
  changes by running the app manually. The one exception is the vendored
  `nanonispy2` reader (see below), which ships its own `tests/test_read.py`.
- `scripts/install.py` / `scripts/install_sxm_viewer.bat` /
  `scripts/run_sxm_viewer.bat` are Windows installer/launcher helpers, not
  part of the app's runtime import path.

## Architecture

### Package layout (`sxm_viewer/`)

- `config.py`, `config_defaults.py`, `config_io.py` — user config file
  (`~/.sxm_viewer_config.json`), header cache (`~/.sxm_viewer_header_cache.json`,
  bump `HEADER_CACHE_VERSION` on breaking format changes), and in-memory
  cache size limits (`CHANNEL_DATA_CACHE_LIMIT`, `FILTERED_CACHE_LIMIT`).
- `_shared.py` — common Qt/matplotlib/numpy imports re-exported for the GUI
  layer, plus a process-wide patch of `FigureCanvasQTAgg.resizeEvent` to guard
  against non-finite/zero resize events. `matplotlib.use("Agg")` is set here.
- `app_meta.py` — app name/org constants and window-icon discovery
  (`samples/app_icon.*`).
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
- Right-click/context menus (not the toolbar) are separate methods:
  `_on_thumb_context_menu`, `_on_spectro_thumb_context_menu`,
  `_populate_marker_style_menu` in `main_window.py`.
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
  histogram auto-contrast, `M` opens molecule load, `R` resets zoom, X/Y/Z
  rotate the selected molecule, Enter/Esc apply/cancel crop-transform mode.
- The publication canvas has yet another independent key map
  (`CanvasGraphicsView.keyPressEvent` / `ExperimentalCanvasWindow._handle_canvas_key`
  in `gui/canvases/canvas_view.py` / `canvas_window.py`): arrows nudge
  selected items, Ctrl+A selects all, Escape clears selection,
  Delete/Backspace deletes, Ctrl+Z/Ctrl+Y undo/redo.
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
- **3D molecule overlay** — `Molecule`/`MoleculePropertiesDialog`
  (`gui/canvases/molecular_overlay.py`); drawn by `_draw_molecules`, toggle
  Ctrl+3.
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

### Spectroscopy position mapping (rotation/frame gotchas)
This bug class ate an entire session before being nailed down, across three
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
and mirrored virtual copies were reported and each "fixed" across several
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
sticking," across multiple clean-restart-and-reload cycles reported by the
user. `hydrate_spectro_entries`/`hydrate_spectro_file`
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
backwards since both shared the same code path. (2) A later session
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

### Folder reports (PDF)
"Generate folder report (PDF)..." (thumbnail context menu + toolbar
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
- **Channel-selection heuristics** (`reporting/channels.py`, user-specified):
  constant-height images show current + frequency-shift/lock-in panels
  instead of flat topo (with a dead-flat-Z retry for untagged CH scans);
  spectroscopy curves prefer freq shift > current > lock-in; z-spectroscopy
  keeps its Z sweep axis; grid slice defaults to the smallest |bias|.
- **Presentation conventions** (also user-specified, in `pdf.py`): current →
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
