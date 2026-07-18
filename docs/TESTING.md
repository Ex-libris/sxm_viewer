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

## 2. Regression counters

```powershell
python scripts\analysis\check_regressions.py
```

Fails if a known-boilerplate counter increases, or if any *phantom call*
appears (a call to a method that exists nowhere - dead code hiding behind an
always-False `hasattr` guard). The phantom counter is at 0 and should stay
there. See `docs/refactor/PATTERNS.md`.

## 3. Analysis toolkit

```powershell
python scripts\analysis\run_all.py
```

Regenerates the duplication reports in `docs/refactor/`. Read-only; useful
when planning a refactor rather than as a per-commit check.

## 4. Vendored reader tests

`sxm_viewer/providers/nanonis/vendor/` ships upstream `nanonispy2` with its
own `tests/test_read.py`. Do not edit that directory - it mirrors an
external package.

## 5. Qt-free unit testing (where it is easy)

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
