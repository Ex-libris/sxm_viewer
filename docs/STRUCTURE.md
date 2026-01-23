# Repository Structure

- `sxm_viewer/` – application package (GUI, data loaders, providers, utilities)
- `scripts/` – installers, helper launchers, legacy utilities
- `docs/` – overview, deep dives, and workflow notes
- `screenshots/` – gallery referenced in README
- `samples/` – optional curated datasets (keep personal data elsewhere)

Local experiments and cached conversions belong outside the repo (e.g., `data_local/`), and `.gitignore` already blocks `.sxm`, `.dat`, and `.sxmviewer_nanonis` caches.
