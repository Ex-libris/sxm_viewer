# Testing and verification

There is no pytest suite for the GUI, and that is a deliberate consequence of
the architecture rather than an oversight: `SXMGridViewer` is a ~540-method
class whose behaviour is entangled with live Qt widgets, so unit-testing it
in isolation would require the decomposition tracked in
`docs/refactor/DUPLICATION_INVENTORY.md` (item R1).

What exists instead, in increasing order of coverage:

## 1. Smoke test (the main safety net)

```powershell
python scripts\smoke_test.py --folder "C:\DATA\your_folder" --report
python scripts\smoke_test.py            # UI-only checks, no data needed
```

Drives the **real** viewer under Qt's offscreen platform: constructs the
window, loads a folder, parses headers, populates thumbnails, loads
spectroscopy, renders a preview, and generates a full PDF report. Exit code
0 = everything passed.

Run this before committing any non-trivial change. It catches the class of
breakage that matters most here - "the app still starts and can open data" -
in about a minute, with no display.

**It isolates config to a temp directory**, so it can never write to your
real `~/.sxm_viewer_config.json`. This matters: the offscreen viewer is the
actual application, and anything that persists settings would otherwise
modify your live preferences.

## 2. Visual regression fingerprint (required for render changes)

```powershell
# on the baseline revision
python scripts\render_fingerprint.py --folder "C:\DATA\your_folder" --out before
# after your change
python scripts\render_fingerprint.py --folder "C:\DATA\your_folder" --out after
python scripts\render_fingerprint.py --compare before after
```

Renders a deterministic set of canvas states to PNG and hashes them, so
two revisions can be compared pixel-for-pixel. Both rendering paths are
captured per image: the **live figure** (what the interactive canvas has
drawn - image, axes, colorbar, overlays) and the **export figure** (the
separate throwaway-Figure path).

**Why this exists, and why the smoke test is not enough.** Every bug this
repo has recorded from refactoring `detail_preview_canvas.py` was visual:
a smeared overlay, an axis dragged back to a stale range, a mirrored
grid. The smoke test proves the app runs; it says nothing about pixels.
That was measured, not assumed - injecting a 2% x-limit error into
`preview_axes_sync.sync_axes_to_view` gives:

| Check | Result |
| --- | --- |
| `smoke_test.py` | **17/17 pass** - completely blind to it |
| `render_fingerprint.py` | **5 of 5 live renders changed** - caught it |

(The export renders correctly stayed identical, since that injection only
affects the live path - so the tool also localises *which* path broke.)

**Run this before and after any change to `MultiPreviewCanvas` render
state.** The profile-line, molecule-overlay and crop-template domains are
all render-coupled and cannot be verified any other way.

Two determinism traps were fixed while building it, both worth knowing if
you extend it: captures must settle after `show_file_channel` (an early
grab caught a half-drawn second view), and the capture uses
`figure.savefig` rather than a Qt widget grab - the widget's bottom rows
pick up neighbouring chrome and differed by 2,416 pixels between two
*identical* revisions. A visual check with false positives is worse than
none; it trains you to ignore it.

## 3. Regression counters

```powershell
python scripts\analysis\check_regressions.py
```

Fails if a known-boilerplate counter increases, or if any *phantom call*
appears (a call to a method that exists nowhere - dead code hiding behind an
always-False `hasattr` guard). The phantom counter is at 0 and should stay
there. See `docs/refactor/PATTERNS.md`.

## 4. Analysis toolkit

```powershell
python scripts\analysis\run_all.py
```

Regenerates the duplication reports in `docs/refactor/`. Read-only; useful
when planning a refactor rather than as a per-commit check.

## 5. Vendored reader tests

`sxm_viewer/providers/nanonis/vendor/` ships upstream `nanonispy2` with its
own `tests/test_read.py`. Do not edit that directory - it mirrors an
external package.

## 6. Qt-free unit testing (where it is easy)

`reporting/`, `providers/`, `data/`, `processing/`, `utils/`, and
`cmap_registry` import no Qt and can be tested directly with plain Python -
no display, no `QApplication`. This is verified: `check_coverage` in the
analysis toolkit and the layering rule in CLAUDE.md both depend on it.

New logic should be born on this side of the line wherever possible. The
folder-report feature is the reference example: all of its modelling and
rendering lives in the Qt-free `reporting/` package behind a plain-data
payload, so it can be exercised in seconds against synthetic inputs, while
the Qt-touching part is a thin controller.

## Offscreen gotchas (each cost real debugging time)

- `QT_QPA_PLATFORM=offscreen` renders **no text** unless `QT_QPA_FONTDIR` is
  also set - PyQt5 ships no fonts. Screenshots then look like a theme
  regression that is not there.
- `_maybe_offer_recovery_session` opens a modal dialog from a startup timer
  when a recovery snapshot exists, which hard-crashes offscreen
  (`0xC0000005`) at a point that looks random. `smoke_test.py` stubs it.
- End scripts with `os._exit()` to avoid Qt teardown corrupting the exit
  code - but flush stdout first, or buffered output vanishes.
- The offscreen screen is 800x600, and control-dense dialogs have a layout
  `minimumSizeHint()` that overrides `resize()`; compare against
  `max(your_clamp, dlg.minimumSizeHint())` when asserting sizes.
