# Samples

This directory is reserved for intentionally curated datasets (small, anonymised examples used in docs or tests).

Personal `.sxm`/`.dat` files should stay outside the repository (e.g., `data_local/`) so the repo stays lightweight.

## Optional app icon drop-in

The viewer now checks this folder at startup for a branding asset and will use
it for the application/window icon when available.

Supported filenames:

- `app_icon.ico`
- `app_icon.icns`
- `app_icon.png`
- `app_icon.svg`
- `sxm_viewer_icon.ico`
- `sxm_viewer_icon.icns`
- `sxm_viewer_icon.png`
- `sxm_viewer_icon.svg`
- `sxmviewer_icon.ico`
- `sxmviewer_icon.icns`
- `sxmviewer_icon.png`
- `sxmviewer_icon.svg`

Recommended:

- Windows: `app_icon.ico`
- macOS bundle workflows: `app_icon.icns`
- Cross-platform development fallback: `app_icon.png`
