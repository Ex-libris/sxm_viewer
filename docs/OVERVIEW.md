# Project layout

- `sxm_viewer/` – application code
  - `gui/` – Qt UI and widgets
  - `data/` – header/spectroscopy parsing helpers
  - `processing/` – image/channel processing utilities (native/Omicron path)
  - `providers/` – format-specific adapters; Nanonis lives here
    - `nanonis/adapter.py` – converts `.sxm` scans to Omicron-style headers
    - `nanonis/vendor/` – bundled `nanonispy2` reader (upstream code, frozen)
- `scripts/` – CLI helpers and launch scripts (if present)
- `docs/` – documentation like this overview
- `tests/` – add format fixtures and regression tests here

## Entry points
- GUI: `python -m sxm_viewer`
- Provider API: `sxm_viewer.providers` (`convert_nanonis`, `parse_nanonis_spectroscopy`)

## Notes for contributors
- Do not edit files under `sxm_viewer/providers/nanonis/vendor/`; they mirror upstream.
- Keep GUI code free of provider internals; providers should not import GUI/Qt.
- Cache folders `.sxmviewer_nanonis/` are generated alongside data and are ignored.
