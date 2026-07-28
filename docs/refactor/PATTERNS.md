# Shared patterns: use these instead of hand-rolling

Short guide to the helpers that exist so common boilerplate is written once.

**Why this file exists.** The most-duplicated idiom in this codebase - the
`blockSignals` triad, 245 sites at baseline - had a *correct* shared helper
available the entire time. It was used 6 times. Availability alone does not
stop a pattern spreading, so each helper below is paired with an automated
counter (`scripts/analysis/check_regressions.py`) that fails when its idiom
count goes up.

---

## Updating a widget without firing its signal

**Use** `gui/qt_helpers.py::set_silent` / `set_many_silent`.

```python
from .qt_helpers import set_silent, set_many_silent

set_silent(self.scale_bar_cb, checked=True)
set_silent(self.zoom_slider, value=42)
set_silent(self.cmap_combo, current_text="viridis")

# One logical setting mirrored by a checkbox AND a menu action:
set_many_silent((
    getattr(self, "unit_display_cb", None),
    getattr(self, "display_units_si_act", None),
), checked=self.display_units_si)
```

**Do not write:**

```python
try:
    widget.blockSignals(True)
    widget.setChecked(state)
    widget.blockSignals(False)   # BUG
except Exception:
    pass
```

Beyond being six lines instead of one, the hand-rolled form is **wrong**:
`blockSignals(False)` unconditionally *unblocks*, so calling it inside an
outer block silently re-enables handlers the caller deliberately suppressed.
`set_silent` restores the previous state, and is null-safe - which is why
most of these sites needed the surrounding `try/except` in the first place.

Property names are snake_case (`checked`, `text`, `current_text`,
`current_index`, `value`, `enabled`, `visible`); anything else falls back to
`set<CamelCase>`. Tuples are splatted, so `set_silent(spin, range=(0, 50))`
works.

---

## Broad exception guards

`try/except Exception: pass` appears ~1,350 times. Many are load-bearing (Qt
teardown ordering, optional imports, matplotlib invalidation quirks) - **do
not bulk-remove them.**

But before adding a new one, check whether the thing you are guarding is
really the problem. Most existing guards protect a hand-rolled widget poke
that a null-safe helper makes unnecessary. Prefer:

1. a helper that is already null-safe (`set_silent`);
2. an explicit `if widget is None: return`;
3. a *narrow* except (`except AttributeError:`) with a comment saying what
   raises and why it is expected;
4. only then a broad guard, with a comment.

---

## Grid geometry helpers

`spec_grid_row_col`, `grid_local_pitch`, `grid_dims`, `grid_local_orientation`
currently exist in **three** copies (`gui/dialogs/matrix_fit.py`,
`gui/dialogs/spectroscopy_dialogs.py`, `reporting/model.py`), documented in
CLAUDE.md as ports that must be kept in sync manually.

If you touch any of them, change all three, and re-read the "Grid Map
Explorer" section of CLAUDE.md first - this is the rotation/flip logic with
a long history of mirrored-output bugs. Consolidating them into one Qt-free
module is tracked as item **A2** in `DUPLICATION_INVENTORY.md`.

---

## Adding a config-backed setting

Currently hand-wired in four places (init read, toggle handler, config write,
mirror-widget sync). A declarative registry is tracked as item **A1** in
`DUPLICATION_INVENTORY.md`. Until that lands, follow the existing shape and
use `set_many_silent` for the mirror sync.

---

## Checking your work

```bash
python scripts/smoke_test.py --folder "C:\DATA\some_folder" --report
python scripts/analysis/check_regressions.py
```

The first drives the real app offscreen end to end; the second fails if any
idiom counter increased. Both are fast and neither needs a display.

If a counter legitimately must increase, re-run with `--update` and explain
why in the commit message - the point is that it becomes a deliberate,
reviewable decision rather than silent drift.
