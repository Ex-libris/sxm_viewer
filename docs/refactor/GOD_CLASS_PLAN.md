# Reducing `SXMGridViewer`'s god-class character

Working plan and progress log. Measurements come from
`scripts/analysis/attribute_cohesion.py` (A5) and `_domain_scan.py`;
re-run rather than trusting this snapshot.

## Where the mass actually is

`SXMGridViewer` at baseline: **541 methods, 814 `self.` attributes.**
Grouping by domain keyword shows it is not uniformly entangled - a few
subsystems dominate:

| Domain | Methods | Attributes | Existing home |
| --- | --- | --- | --- |
| **spectroscopy** | **99** | **146** | `gui/spectroscopy/` (5 modules) + `SpectroCompareController` |
| **thumbnail** | 63 | 88 | `gui/viewer/thumbnail_ui.py`, `thumbnail_controller.py` |
| popup/window mgmt | 34 | 42 | `controllers/preview_popup.py` |
| filter | 37 | 20 | `controllers/filter_controller.py` |
| preview | 29 | 52 | `gui/viewer/preview.py` |
| colormap/theme | 27 | 41 | `cmap_registry`, `gui/theme.py`, `colormap_manager.py` |
| session/recovery | 26 | 33 | `controllers/session.py` |
| collection | 24 | 30 | `controllers/collection.py` |
| canvas/publication | 22 | 22 | `gui/canvases/` |
| crop/adjust | 19 | 55 | `controllers/quick_crop.py`, `dialogs/image_adjust.py` |
| molecule | 20 | 33 | `gui/canvases/molecular_overlay.py` |
| *(unclaimed)* | 159 | 242 | — |

**The key insight: almost every domain already has a home.** This is not a
design problem needing new abstractions - it is a *migration that stalled*.
The 99 spectroscopy methods on the god class are the ones that never moved
into the `gui/spectroscopy/` package that exists for them.

Of those 99, **69 carry real logic (1,447 lines)**; the other 24 are already
thin delegating shims. So the work is bounded and mostly mechanical.

## Strategy

Two complementary moves, in this order:

1. **Move logic to the package that already owns the domain.** Follows the
   established `gui/viewer/*.py` module-function convention (plain function
   taking `viewer` first). Prefer having callers import the module directly
   over leaving a shim on the class - CLAUDE.md already prefers this for
   `gui/spectroscopy/*` siblings, and a shim keeps the method count up.

2. **Extract cohesive state into small owned classes** where A5 shows high
   isolation. These are smaller wins per unit, but they remove *state*, not
   just methods - and state is what makes the class hard to reason about.

Do **not** attempt a single grand decomposition. Each step must be
independently shippable and verified by `scripts/smoke_test.py`.

## Progress

Track with `python scripts/analysis/class_size.py`.

| Step | Status | Effect |
| --- | --- | --- |
| Remove shadowed/dead definitions | ✅ done | −4 methods, 3 latent `AttributeError`s removed |
| Extract `ActivityLog` | ✅ done | −3 attrs, −2 methods, new tested class |
| Extract `geometry/spec_mapping` | ✅ done | −4 attrs, ~190 lines of logic out; now Qt-free + testable |
| Extract `Debouncer` | ✅ done | −4 attrs, 3 hand-rolled copies unified |
| Extract `spectroscopy/overrides.py` | ✅ done | 7 methods → delegations, ~150 lines out |
| Extract `spectroscopy/details.py` | ✅ done | ~85 lines out, now pure/testable |
| Extract `spectroscopy/loading.py` | ✅ done | 9 methods → delegations, ~200 lines out |
| **Split `__init__` (1588 lines)** | ✅ done | **−451 lines**; state phase now `main_window_state.py` |
| Move remaining spectroscopy interaction handlers | ⬜ next | `_handle_spec_hover`, `_handle_spec_marker_click`, `_on_spectro_thumb_context_menu` (~200 lines) |
| Move thumbnail logic to `gui/viewer/` | ⬜ queued | up to −63 methods / −88 attrs |
| Delete shims; have callers import modules directly | ⬜ queued | this is what finally drops the **method** count |

**Current: 536 methods, 714 attributes, 11,037 class lines**
(from 541 / 814 / ~12,050 — **−1,013 lines, −100 attributes**).
`main_window.py` as a file: **10,674 lines** (was ~11,800).

### Reading these numbers honestly

- **Lines** is the real signal: ~1,000 lines of logic now live in focused
  modules instead of one file.
- **Attributes** dropped by 100, but ~85 of those simply moved to
  `main_window_state.py` - the object still *has* them at runtime. What
  changed is that they are now declared in one readable place instead of
  being scattered through a 1,588-line constructor.
- **Methods** has barely moved, and will not until the shims go. 153 of
  the 536 methods (29%) are now pure delegation - see
  `SHIM_CENSUS_SXMGridViewer.md`, which groups them by forwarding target
  so a whole group can be retired at once by pointing callers at the
  module directly.

### Verification for the `__init__` split

Moving 451 lines of constructor is the riskiest change on this branch, so
it was checked by **runtime attribute fingerprint**: construct the viewer
before and after, dump every `vars(viewer)` entry with its type/size, and
diff. Result - 7 attributes removed, 3 added, 1 changed, all of them the
intended debouncer/activity-log replacements, plus a timestamp. Nothing
unintended. (`_init_locals.py` in the branch history also proved the block
had no local-variable coupling to the rest of `__init__` before the move.)

### Completed detail

**Shadowed definitions** (found by `scripts/analysis/find_shadowed.py`,
written for this purpose). Five methods were defined twice in one class
body, so the first copy was dead:

- `SXMGridViewer._map_spec_to_pixels` (5584, 9472)
- `SXMGridViewer._matrix_bbox_pixels` (5587, 9718)
- `SXMGridViewer._fallback_spec_coords` (5590, 9669)
- `SXMGridViewer.on_dark_mode_toggled` (4756, 11607) - byte-identical copy
- `SpectroscopyCompareDialog._get_icon` (6425, 7009) - byte-identical copy

The first three were **also broken**: they delegated to
`viewer_preview._map_spec_to_pixels` etc., which `gui/viewer/preview.py`
does not define - that module only *calls* `viewer._map_spec_to_pixels`.
They would have raised `AttributeError` if reached, and were saved purely
by being shadowed 4,000 lines later. Same family as the
`_show_spectro_popup` bug in CLAUDE.md. Both counters are now
zero-tolerance in `check_regressions.py`.

**`ActivityLog`** (`gui/activity_log.py`) - A5's highest-isolation group at
75%. Owns the batched append/flush/clear of status messages. Verified by a
dedicated smoke-test check that exercises the full append → flush → clear
round trip, not just construction.

## Next: the spectroscopy migration

Largest single reduction available. Suggested order, smallest blast radius
first:

1. **Pure-geometry helpers** - `_map_spec_to_pixels` (105 lines),
   `_map_spec_by_spec_extent` (52), `_matrix_bbox_pixels` (47),
   `_map_spec_by_grid` (19). These take `(spec, header, pixels)` and return
   coordinates; no widget state. Moving them to a Qt-free module makes the
   most bug-prone logic in the repo unit-testable for the first time (see
   CLAUDE.md "Spectroscopy position mapping" for its history). Verify with
   the documented correlation check (|r| ≈ 0.98).
2. **Assignment/override plumbing** - `_resolve_spectro_override_targets`,
   `_spectro_override_signature`, `_refresh_spectro_assignment_overrides`,
   `_spec_identity_key`.
3. **Interaction handlers** - `_handle_spec_hover`, `_handle_spec_marker_click`,
   `_on_spectro_thumb_context_menu` → `gui/spectroscopy/`.
4. **Loading/lifecycle** - `ensure_spectros_loaded`,
   `_run_pending_spectro_load_async`, `_apply_spectro_scan_results`.

Each is a separate PR against `master`, smoke-tested, with the domain scan
re-run to confirm the counts actually drop.
