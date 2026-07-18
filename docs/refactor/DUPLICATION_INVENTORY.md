# Duplication inventory (Phase B)

Triaged findings from the analysis toolkit (`scripts/analysis/`, run with
`python scripts/analysis/run_all.py`). Every number here is reproducible -
re-run the tools rather than trusting this snapshot.

**Baseline: 2026-07-18, 105 files, 79,502 lines (vendor excluded).**

Classification:

| Band | Meaning |
| --- | --- |
| 🟢 **Green** | Mechanical and behavior-preserving. Safe with a smoke test. |
| 🟡 **Amber** | Behavior-preserving but broad; migrate incrementally, verify each step. |
| 🔴 **Red** | Changes structure/ownership. Evidence recorded; **not** scheduled. |

---

## Headline numbers

| Signal | Count | Source |
| --- | --- | --- |
| Silent-widget (`blockSignals`) sites | **245** | A1 |
| `try/except Exception: pass` guards | **1,366** | A1 |
| Defensive `getattr(self/viewer, "x", None)` | **1,227** | A1 |
| `config[...] = x` immediately followed by `save_config` | **88** | A1 |
| Repeated header-field unpacking | **43** | A1 |
| Whole-function structural clones | **50 functions / 24 groups** | A2 |
| Block clones (≥4 statements, ≥3 instances) | **79 groups / ~2,798 statements** | A2 |
| Near-duplicate method clusters (≥80% similar) | **110 clusters / 247 methods / ~2,627 lines** | A3 |
| True phantom calls | **0** ✅ | A4 |
| Never-referenced definitions | **72** (58 with no dynamic-use hint) | A4 |
| Called-once definitions ≤4 lines | **82** | A4 |
| `SXMGridViewer` attributes | **814** distinct, **537** methods | A5 |

---

## 🟢 G1 - Silent widget update (`set_silent`)

**245 sites.** The single most repeated idiom in the codebase:

```python
try:
    widget.blockSignals(True)
    widget.setChecked(state)
    widget.blockSignals(False)
except Exception:
    pass
```

A **correct** helper already exists (`_set_combo_text_silent`,
`main_window.py`) and is used ~6 times; everything else is hand-rolled.
Crucially, the hand-rolled copies hardcode `blockSignals(False)` instead of
restoring the prior state - a latent bug: nesting silently unblocks signals
an outer caller deliberately blocked.

**Action:** `gui/qt_helpers.py::set_silent(widget, **props)`, null-safe and
state-restoring. Collapses ~6 lines to 1 and removes ~245 `except` guards
as a side effect.

**Estimated:** ~1,200 lines removed. **Risk:** very low (behavior-identical,
strictly safer).

## 🟢 G2 - Dead code removal

**72 never-referenced definitions**, 58 with no matching string literal
(so not dynamically dispatched). Pure deletion.

**Action:** review `docs/refactor/A4_USAGE.md` "Never referenced", delete
in one batch. Qt lifecycle names and signal handlers are already excluded
by the tool, but each still warrants a 10-second look - `getattr`-style
dispatch through a variable name the scanner can't see is possible.

**Risk:** low, but verify individually; deletion is trivially revertible.

## 🟢 G3 - Debounced-refresh pattern (`Debouncer`)

A5 exposed the same hand-rolled shape **four independent times**:
`_X_pending` + `_X_timer` + `_flush_X()`:

| Instance | Attributes |
| --- | --- |
| Activity log | `_activity_log_pending`, `_activity_log_flush_timer` |
| Preview request | `_pending_preview_request`, `_preview_request_timer`, `_preview_render_in_progress` |
| Thumbnail render state | `_thumbnail_render_state_pending_paths`, `_thumbnail_render_state_timer` |
| Compact histogram | `_pending_compact_histogram_clim`, `_compact_histogram_apply_timer`, `_suppress_compact_histogram_refresh` |

**Action:** one small `Debouncer` class (interval, pending payload, flush
callback). Removes ~12 attributes from `SXMGridViewer` and makes the
throttling policy inspectable in one place.

**Risk:** low-medium - timing behavior must be preserved exactly; migrate
one instance at a time.

## 🟡 A1 - Config-backed mirrored toggles (settings registry)

**88 config-write-then-save sites**, and A3 clusters the handlers directly:
`on_show_spectra_toggled`/`on_show_preview_spectra_toggled` (54 lines),
`on_show_matrix_markers_toggled`/`on_show_single_markers_toggled`/
`on_compact_markers_toggled` (39), `on_scale_bar_toggled`/
`on_unit_display_toggled` (36), `on_unit_relative_toggled`/
`on_preview_lock_toggled` (37), `on_show_spectro_miniatures_toggled`/
`on_detail_grid_toggled` (37).

Each setting is hand-wired in four places: `__init__` read, toggle handler,
config write + save, and a mirror loop syncing checkbox *and* menu action.

**Action:** declarative `Setting(name, default, mirrors, on_change)` table +
one generic handler.

**Estimated:** ~26 methods, ~400 lines. **Risk:** medium - touches user-
visible behavior; migrate **one setting per commit**, toggling it in the app
each time. Depends on G1.

## 🟡 A2 - Duplicated grid geometry helpers

A2 found these **structurally identical across three files**:

| Helper | Copies |
| --- | --- |
| `_spec_grid_row_col` | `matrix_fit.py`, `spectroscopy_dialogs.py`, `reporting/model.py` |
| `_grid_local_pitch` | same three |
| `_grid_dims` | `spectroscopy_dialogs.py`, `reporting/model.py` |

CLAUDE.md documents these as deliberate ports that "must stay in sync" - a
known, *managed* risk. The tools prove they are byte-identical in shape, so
the risk can be **eliminated** instead: extract to a Qt-free
`sxm_viewer/geometry/grid.py` that all three import. The original reason for
duplicating (the dialog versions live on a `QDialog` and `reporting/` must
stay Qt-free) is satisfied by a shared Qt-free module.

**Risk:** medium - this is the rotation/flip logic with a long bug history
(see CLAUDE.md "Grid Map Explorer"). High payoff, but do it with the
correlation check described there as verification.

## 🟡 A3 - Strategy-table method families

Clean parameterization candidates from A3:

- `_build_slice_metric` / `_build_peak_metric` / `_build_integral_metric` /
  `_build_stat_metric` - 4 methods, 73 lines (`MatrixSpectroViewer`)
- `_select_topo_axis` / `_select_z_axis` / `_select_bias_axis` /
  `_select_true_bias_axis` - 4 methods, 83 lines (`nanonis/adapter.py`)
- `set_show_shortcut_hint` / `set_show_profile_overlays` / … - **7** methods,
  63 lines (`MultiPreviewCanvas`) - identical show/hide setters
- `_remember_closed_*` - 3 methods, 36 lines

**Action:** one parameterized function + a table per family.
**Risk:** low-medium, isolated per family.

## 🟡 A4 - Target resolution (`resolve_target_files`)

The "selection → highlighted → previewed" fallback chain is re-derived at
call sites (`last_preview` read 136×, `thumb_multi_select` 76×,
`selected_file_for_thumbs` 39×). Divergence between copies has caused real
inconsistency between the thumbnail combo, gallery Apply, and export paths.

**Action:** one documented `resolve_target_files()` method.

## 🔴 R1 - `SXMGridViewer` decomposition

**537 methods, 814 distinct attributes.** A5 ranks candidate extractions by
isolation (share of touching methods that touch *nothing else*):

| Isolation | Group |
| --- | --- |
| 75% | activity log (`_activity_log_*`, `activity_log_box`) |
| 50% | session activity strip (4 attrs) |
| 50% | window history (3) |
| 50% | spectroscopy folder (3) |
| 33% | preview-request debounce (4) |
| 33% | thumbnail render state (3) |

Only the top few exceed 50% isolation; most of the class is genuinely
entangled. **Not scheduled.** Recorded as measured evidence so the eventual
split is argued from data, and so a successor can see which pieces are
detachable. The debouncer work (G3) already peels off two of these groups
without any structural commitment.

## 🔴 R2 - Defensive-guard sprawl

**1,366 `try/except Exception: pass` + 1,227 defensive `getattr`.** Do
**not** bulk-remove: some are load-bearing (Qt teardown ordering, optional
imports, matplotlib invalidation). They are a *symptom* of 814 attributes
with no guaranteed initialization order.

**Approach:** let them disappear as a side effect of G1/G3/A1. Re-measure
with A1 after each batch; a falling count is the progress metric.

## 🔵 Informational - injected callbacks

A4 found **50 names** called as `self.x(...)` with no `def`, but assigned
externally (`canvas._compare_menu_callback = fn`). Not bugs, but an
**implicit, untyped interface**: nothing declares them, so a rename fails
silently at runtime. Worth documenting as a real extension surface in the
handover docs rather than refactoring.

---

---

## Progress

| Item | Status | Notes |
| --- | --- | --- |
| Smoke test | ✅ done | `scripts/smoke_test.py`, 13/13 on a 224-scan folder |
| Regression counters | ✅ done | `scripts/analysis/check_regressions.py` + `baseline.json` |
| G1 `set_silent` | 🟨 in progress | helper landed + `main_window.py` batch: **245 → 215** sites |
| G2 dead code | ⬜ queued | 58 clean candidates in `A4_USAGE.md` |
| G3 `Debouncer` | ⬜ queued | 4 instances identified |
| A1 settings registry | ⬜ queued | depends on G1 |
| A2 grid geometry | ⬜ queued | needs the correlation check from CLAUDE.md |
| A3 strategy tables | ⬜ queued | 4 families |
| A4 target resolution | ⬜ queued | |

**G1 remaining call sites** (215), by file - the migration worklist:

Run `python scripts/analysis/find_idioms.py` and read the
`block_signals_triad` section of `A1_IDIOMS.md` for exact file:line
locations. Largest remaining concentrations are
`gui/dialogs/spectroscopy_dialogs.py`, `gui/canvases/canvas_window.py`,
`gui/controllers/quick_crop.py`, and `gui/canvases/detail_preview_canvas.py`.

Migrate in batches of ≤40 sites, running the smoke test and
`check_regressions.py` between batches, and merging each batch to `master`
before starting the next.

## Recommended sequence

1. **Smoke test first** (`scripts/smoke_test.py`) - prerequisite for all
   code changes.
2. G2 (dead code) - pure deletion, builds confidence.
3. G1 (`set_silent`) - biggest mechanical win.
4. A4 (`resolve_target_files`) - small, prevents a real bug class.
5. G3 (`Debouncer`) - removes attributes from the god class.
6. A3 (strategy tables) - per family.
7. A1 (settings registry) - one setting per commit.
8. A2 (grid geometry) - highest care, needs the correlation check.

Each step: branch off current `master`, ≤40 call sites, smoke test, re-run
A1, PR, merge within a day or two. Never accumulate on a long-lived branch -
`main_window.py` takes ~70 commits/month and will win any merge fight.
