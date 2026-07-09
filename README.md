# SXM Viewer

SXM Viewer is a Python-based desktop application for scientific SPM (Scanning Probe Microscopy) data analysis and visualization, designed for Anfatec/Omicron systems. But also Nanonis. Maybe in the future Matrix. We will see.

---


## Documentation

Full documentation is available at:

https://ex-libris.github.io/sxm_viewer/

Key pages:
- Installation: https://ex-libris.github.io/sxm_viewer/getting-started/installation/
- First Steps: https://ex-libris.github.io/sxm_viewer/getting-started/first-steps/
- Profiles and Measurements: https://ex-libris.github.io/sxm_viewer/image-analysis/profiles/

---

## Overview

SXM Viewer provides an integrated environment for:

- Fast browsing of large SPM datasets
- Image analysis (profiles, angles, cropping, filtering)
- Spectroscopy visualization (traces, matrix scans, KPFM)
- Overlay tools (molecules, metadata, scale bars)
- Publication-ready figure composition (canvas)
- Session and collection management



![Main interface](screenshots/main_menu.png)

## Quick start

```powershell
git clone https://github.com/Ex-libris/sxm_viewer.git
cd sxm_viewer
conda create -n sxmviewer python=3.11
conda activate sxmviewer
cd .\scripts
python -m pip install -r .\requirements.txt
cd ..
python -m sxm_viewer
```

See the full installation guide in the MkDocs site for the Windows installer helper and troubleshooting notes.





### 2026-07-10 UPDATES (spectroscopy position accuracy & browser overhaul)
A deep pass on where spectroscopy points actually land and how you find
them — this was the biggest correctness fix to spectroscopy in a while,
plus a full redesign of the Spectro Browser:

**Position accuracy fixes:**
- Spectroscopy markers on rotated and non-square (elongated) scans could
  land visibly off from their true acquired position, sometimes badly
  smeared across the image — fixed a shearing bug in the position-mapping
  math that only showed up on non-square scans
- Some Nanonis scans were displaying upside-down/mirrored relative to their
  real acquisition direction (and therefore so were their spectroscopy
  markers) — fixed the row-order handling for the `Direction: up` scan
  convention, which a real folder can mix with `Direction: down` scans
  file-by-file
- The Spectrum window's "Position" inset (the small reference-image
  thumbnail with the marker) could show the marker in a completely wrong
  spot, or even outside the visible thumbnail entirely, on any scan that
  had blank/aborted rows trimmed off for display — fixed

**Off-frame spectroscopy points**, common for reference spectra
deliberately acquired off to the side of the scanned area:
- these now render with a distinct orange flag pointing toward their real
  direction, clamped to the nearest edge of the image, instead of being
  invisibly misplaced somewhere inside the frame
- new "Off-frame" filter in the Spectro Browser, with a live count, and a
  Spectroscopy → "Review off-frame spectroscopies" menu action that jumps
  straight to them

**Position inset (Spectrum window) improvements:**
- now a faithful color copy of whatever channel/cmap the main Preview is
  showing, instead of always desaturating to gray
- resizable by dragging its corner, and bigger by default
- new "Inset settings" menu: show/hide the other plotted points, and
  change the position marker's symbol/size/color

**Spectro Browser redesign** — it was a dense wall of small text before;
now:
- small colored icons per row (single point / grid / Z-series, plus
  low-confidence and off-frame flags) instead of bracketed text tags
- large position groups (a 512-point grid, a big Z-series) collapse by
  default instead of dumping hundreds of rows open at once
- rows lead with the file name first (matching how you'd recognize files
  in a folder explorer), with position/channel/assignment info tucked into
  an expandable detail underneath, instead of a redundant "Position X/Y nm"
  wrapper around every single spectrum
- a grid/matrix scan (which can be 1000+ individual points) now collapses
  to one row for the whole grid, showing its dimensions, point count,
  channel count, and acquisition time — double-click or right-click → Open
  jumps straight into the Grid Map Explorer instead of drilling through
  individual points
- fixed "Open" on an individual spectrum doing nothing at all (a leftover
  broken hook from an earlier version of the browser) — it now reliably
  opens the Spectrum trace window
- opening a Z-series/cluster now opens a proper multi-trace comparison
  window instead of an older, disconnected summary dialog
- live counts on every filter checkbox, and a "Showing X of Y spectra"
  line so filtering never silently empties the list with no explanation
- the preview panel now shows an actual colored image of the spectrum's
  source scan with its position marked, not just text; the tree and
  preview panel are now in a resizable split so you can drag the preview
  bigger

### 2026-07-07 UPDATES (spectroscopy UX pass)
A round of workflow polish for spectroscopic data (single spectra, Z series at one position, and grid/CITS maps), aimed at making the spectra ↔ image navigation symmetric and the click behavior predictable:

- **Show on image**: every spectroscopy window (single spectrum, comparison, grid map explorer) now has a "Show on image" button (also in its right-click menu) that focuses the main preview on the image the spectrum was acquired on, scrolls its thumbnail into view, and pulses its marker
- **Selection tray**: Shift+clicking markers no longer auto-opens a comparison window at the second selection; instead a small tray appears under the thumbnails ("N spectra selected - Compare | Clear") so you decide when to compare
- **"With spectroscopy" view filter** (`Ctrl+Alt+P`, also in Display → Show only and the Filter dropdown): show only the images that have linked spectroscopy points; press again to show all
- **Marker legend**: Spectroscopy toolbar → "What do these markers mean?" opens a legend rendered with your actual marker style/colors: single vs grid-map markers, the dashed low-confidence ring, the highlight glow, repeat/Z-series badges, and the per-image totals badge
- **Always-on presence banner**: as soon as a folder is scanned, a pill above the thumbnails announces what spectroscopy exists ("⚡ 1655 spectra · 14 grid maps · click to browse") - no toggles or menus needed to know the data is there; clicking it opens the Spectro Browser, which also lists spectra that aren't linked to any image
- **Scientist vocabulary**: window titles, tooltips, and summaries now say "Position 12.3/45.6 nm" instead of "site", "Z series" instead of "stack", and "Grid map" instead of "matrix" - e.g. "Z series (x5) - Position 12.3/45.6 nm" for a tip-height series

### 2026-07-07 UPDATES (starred favourites and quick view filters)
You can now star your favourite images and browse only those:

- star/unstar thumbnails from the right-click menu ("★ Star" / "Remove star") or by pressing `S` while the thumbnail area has focus - `S` toggles, so pressing it again removes the star; starred images show a gold star badge and a brief star animation plays when you star one
- an in-app guide explains all of this: `Display → Show only → How favourites & filters work...`, and the shortcuts panel lists the new keys
- stars are remembered across sessions, so reopening a folder later still shows which images you starred
- new `Display → Show only` menu filters the thumbnail grid to Starred (`Ctrl+Alt+F`), Constant height (`Ctrl+Alt+H`), or Constant current (`Ctrl+Alt+C`) images - press the same shortcut again to show everything (the thumbnail Filter dropdown reflects and controls the same state)
- virtual copies follow their source image under these filters: a copy of a CH/CC image inherits the tag (and border badge), and a copy of a starred image starts out starred, so copies created while a filter is active stay visible instead of vanishing until you switch back to "All"

### 2026-07-05 UPDATES (auto-enhance and periodic-noise removal)
Two new tools in the preview panel's Filters, aimed at quick cleanup of common scan problems:

- a "Let the robot" button that diagnoses the current image (tilt, incomplete scans, glitched scan lines, isolated spike pixels, high-frequency noise, poor contrast) and applies whichever fixes actually apply, as ordinary steps in the same filter pipeline you'd build by hand - fully visible and editable afterward, not a hidden transform
- a "Remove periodic noise..." dialog for scan-line banding and pump/mains-frequency vibration: shows the image's frequency spectrum, flags likely noise regions with a plain-language reason, and lets you draw your own regions (rectangle or ellipse) to remove or protect, since genuine surface structure can look similar to noise in frequency space and this always stays a manual, reviewed step
- Nanonis scan-timing metadata (used to match noise against known pump/mains frequencies) is now read correctly - a bug meant it was previously never being picked up at all

### 2026-07-04 UPDATES (performance pass)
A round of work aimed at the two things that show up most during everyday use: browsing thumbnails/previews and loading a folder for the first time.

- clicking between thumbnails and previews is noticeably snappier - removed several redundant full-canvas re-renders that used to happen on every click
- fixed a couple of display glitches introduced while chasing that speed-up (an accumulating scale bar, and image/axis misalignment when switching between very differently-sized scans)
- loading a folder you've never opened before is roughly 4-5x faster for Nanonis `.sxm` conversion, now parallelized across CPU cores with verified byte-identical output
- fixed a few accidentally-quadratic slow paths in spectroscopy loading (matching spectra to the right image, building per-spectrum metadata) - up to ~60x faster in the worst cases
- the on-disk header cache no longer grows without bound across every folder you've ever opened

### 2026-05-07 UPDATES (talking with Kelvin's group)
Nanonis support has been updated with a focus on faster scan loading and reloads:

- converted `.sxm` channel caches now use binary NumPy `.npy` files instead of ASCII text exports
- this reduces conversion I/O overhead and speeds up subsequent channel reads
- automatic CH/CC tag detection now reuses cached results when the header and topography source have not changed
- warm folder reloads therefore avoid unnecessary topography re-reads when auto-tagging is enabled


