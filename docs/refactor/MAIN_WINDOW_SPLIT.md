# Splitting `main_window.py`

`main_window.py` is one file holding essentially one class. This records
how it is being broken up, and - more importantly - **why the obvious
approach is the wrong one**.

## Why not mixin classes

The tempting move is to slice `SXMGridViewer` into mixin base classes
(`class SXMGridViewer(ThumbnailMixin, SpectroMixin, CanvasMixin, ...)`),
one per file. It shrinks the file immediately and needs no call-site
changes.

It is still the wrong tool here:

- **The coupling does not go away, it goes underground.** Mixins share one
  `self` namespace, so all ~700 attributes stay globally reachable from
  every mixin. Nothing becomes encapsulated; the mess is just spread over
  more files.
- **MRO makes it harder to trace, not easier.** A successor grepping for
  `self._thumb_cache` would now have to know which of eight base classes
  owns it, and Python will happily let two mixins collide on a name.
- **It rewards leaving the god object intact**, because adding "just one
  more mixin" is always easier than defining a real boundary.

The convention this repo already uses - plain module functions taking
`viewer` first - is better precisely because it is *honest*: the
dependency is a visible parameter, and the moment a function needs eight
viewer attributes you can see that it does.

## The approach

Move coherent domains into their own modules, in this order of preference:

1. **Qt-free module** (`sxm_viewer/geometry/`, `reporting/`) when the logic
   is pure computation. Best outcome: unit-testable without a display.
2. **Domain module taking `viewer`** (`gui/drift_animation.py`,
   `gui/virtual_copies.py`) when it needs viewer state but is one workflow.
3. **Owned collaborator class** (`gui/activity_log.py`, `gui/debounce.py`)
   when there is real *state* to encapsulate, not just behaviour.

`main_window.py` keeps: the class definition, `__init__`'s widget/wiring
phases, Qt event overrides, and thin delegating entry points.

## Done

| Module | Lines | What |
| --- | --- | --- |
| `main_window_state.py` | 487 | phase 1 of `__init__` - all non-widget state |
| `drift_animation.py` | 671 | drift correction, animation export, alignment preview |
| `virtual_copies.py` | 240 | non-destructive derived images |
| `geometry/spec_mapping.py` | ~250 | spec nm <-> raster pixel transforms (Qt-free) |
| `spectroscopy/loading.py` | ~260 | lazy scan, async autoload, manifest save |
| `spectroscopy/overrides.py` | ~210 | manual spec -> image assignment |
| `spectroscopy/details.py` | ~120 | Details-panel formatting (Qt-free) |
| `debounce.py` | 109 | `Debouncer` - replaced 4 hand-rolled copies |
| `activity_log.py` | 86 | batched activity-log writer |
| `qt_helpers.py` | 88 | `set_silent` - replaced 245 hand-rolled triads |

`main_window.py`: **~11,800 -> 9,823 lines.**
`SXMGridViewer`: 541 -> 536 methods, 814 -> 708 attributes,
~12,050 -> 10,152 class lines.

## Next, in priority order

1. **Retire shim groups** (`SHIM_CENSUS_SXMGridViewer.md`). 153 of 536
   methods are pure delegation. Point callers at the target module and
   delete the shim - this is the only thing that moves the *method* count.
   Biggest groups: `filter_controller` (29), `viewer_thumb_ui` (19),
   `main_window_spectro` (11), `viewer_measurement` (10).
2. **Canvas display/style** - `_apply_canvas_display_options` (129),
   `_apply_canvas_style_snapshot` (119), `_sync_view_cmaps_from_canvas`
   (77) -> `gui/canvas_display.py`.
3. **Window history / session restore** -
   `_restore_closed_window_payload` (120) + the `_remember_closed_*`
   family -> `gui/window_history.py`.
4. **Typography** - `set_plot_typography` (105) +
   `_apply_preview_workspace_theme` (75) -> extend `gui/plot_typography.py`.
5. **Thumbnail context menu** - `_on_thumb_context_menu` (228) is the
   largest single remaining method; it belongs with
   `gui/viewer/thumbnail_ui.py`.
6. **`__init__` phases 2 and 3** - widget construction (~730 lines) and
   wiring/layout (~380). Harder than phase 1 because widget construction
   builds locals used by later layout code; do it only with the same
   local-coupling check that phase 1 used.

## `MultiPreviewCanvas`

439 methods, **12,158 lines** - now larger than `SXMGridViewer`. Started,
but this file needs more care than the main window did: CLAUDE.md records
three separate bugs from one earlier attempt to share rendering setup
between the live path and a hand-rolled copy.

A domain scan (`scripts/analysis/_canvas_domains.py` in branch history)
measured each domain's size **and its coupling to render state**
(`_redraw`, `_images`, `_axes`, `set_extent`, `apply_aspect`, ...):

| Domain | Methods | Lines | Render-coupled |
| --- | --- | --- | --- |
| profile line | 79 | 1981 | 22 |
| molecule overlay | 66 | 1822 | 20 |
| crop/template | 52 | 1504 | 15 |
| SVG molecule overlay | ~40 | ~1200 | some |
| **export figures** | **3** | **286** | **0** ✅ |
| angle measure | 31 | 485 | 4 |
| outline | 14 | 453 | 4 |
| redraw core | 6 | 399 | 5 |

**Rule for this file: prefer the zero-render-coupled groups.** The export
figure trio (`_render_view_figure`, `_render_views_grid`,
`_style_export_figure`) qualified - they build a throwaway Figure and
never touch live artists - and are now
`gui/canvases/preview_export_figures.py`, covered by a smoke check that
renders both paths and asserts the live canvas is undisturbed.

Next candidates there, in increasing risk: the SVG molecule *export*
formats (XYZ/MOL/PDB text generation - pure formatting), then outline,
then angle. Leave `redraw core` alone.

**A trap worth recording**: these three methods are *not* contiguous -
nine unrelated methods sit between them. An extraction that lifts the
whole line span silently deletes those neighbours. Extract per method,
and assert the class's method count afterwards (`class_size.py`).

## Verification standard

Every extraction on this branch was checked by at least one of:

- **runtime equivalence** - construct the viewer before/after and diff a
  fingerprint of every `vars(viewer)` entry (used for the 451-line
  `__init__` move: 7 removed / 3 added / 1 changed, all intended);
- **output equivalence on real data** - e.g. 5,782 coordinate mappings and
  2,891 formatted spec dumps, both bit-identical;
- **a new smoke-test check** exercising the moved behaviour end to end
  (activity log, debouncer, override resolution, virtual copies).

Do not move code out of this class without one of those.
