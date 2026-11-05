# Changelog

Notable changes to SXM Viewer. Versions follow [Semantic Versioning](https://semver.org/).

## Unreleased

- **Figures replace the canvas as the way to build a figure**: the toolbar's
  `New Figure` turns the selected images into a figure in grid order (the
  canvas stays reachable from its menu, marked legacy). Custom figures keep
  their identity and name when frames are added, removed or dragged into a
  new order, are listed under Recent figures and can be saved / opened as
  `*.sxmfig.json` files. A series figure can be copied into a custom one.
- **Per-image processing in figures**: each image of a figure (one channel
  of one frame) can have its own filters, Zero and colour range (with a
  histogram) without touching the other images, the other channels of the
  same frame or the viewer. Keys on the image under the cursor: A, Z, G, T.

## 0.1.0 - 2026-11-05

First public release, archived on Zenodo.

Release preparation (no change to how images, spectra or figures look - the
preview renders were checked pixel for pixel against the previous version on
Anfatec and Nanonis data):

- **Fix**: pop-up windows that had a profile line drawn on them stayed in
  memory after being closed (figure and image data included).
- **Fix**: pressing Esc on the recent-molecules menu opened a file dialog.
- **Installer**: `scripts/install.py` no longer downloads or installs without
  HTTPS certificate verification; it needs Python 3.10-3.13, as the package
  does.
- **Packaging**: `pip install .` gives a complete installation, application
  icon included; the wheel no longer picks up compiled `.pyc` files.
- **Cleanup**: about 1,900 lines of unused imports, unreachable code and
  leftovers of earlier UI versions removed; deprecated calls updated for
  current Python.
- **Continuous integration**: every push runs the unit tests, an import of
  every module and a headless start of the application on Windows (Python
  3.10 and 3.12). The documentation site is built from `master` only.

## Development history (pre-release)

### 2026-11-04 (one molecular-model mechanism)

- **Molecular models are one shared mechanism**: the preview, pop-ups,
  figure windows, exports and the publication canvas now load, draw and
  manipulate models through one package (`sxm_viewer/molecules`), so a look
  or gesture changed in one place changes everywhere.
- **Publication canvas** draws models with the same styles as the preview
  (it had its own simpler circles-and-lines drawing).
- **Fix**: Ctrl+Shift+drag on a model in the preview now tumbles it in 3-D
  as documented (it only spun it in plane).
- **Hydrogens visible on white**: white atoms and white bond halves get a
  thin dark rim, so hydrogens no longer vanish on figures or bright images.
- **Thin / Thick bonds** are now clearly different (0.45x / 2x of default).

### 2026-11-03 (overview + zooms, publication formatting)

- **Overview + zooms figure**: a frame with smaller frames scanned inside
  it gets a blue ⊡N chip; the figure shows the overview next to its zooms,
  each outlined and lettered on the overview in its own colour. Frames
  taken in different modes show their own best channel (topography of a CC
  overview, current of its CH zooms).
- **Journal pages**: Nature, Science, ACS and generic 1 / 1.5 / 2 column
  widths set the figure width *and* every text size (letters, headers,
  channel names, colour-bar values and names, scale-bar labels; nothing
  below 5 pt). Too many frames for the page wrap into more rows. *Show at
  actual size* draws the figure at its printed size with a width ruler.
- **Scale bars** on every image by default; the length label once per
  distinct length (or on every bar / none); Ctrl+wheel on a bar changes the
  label size (the bar stays put); thickness follows the shorter panel side.
- **Colour bars**: channel name centred inside the bar (vertical or
  horizontal), values outside / inside / min-max text, rotated or
  horizontal text, optional frame.
- **Outlines** on letters, scale bars and colour bars are each optional.
- **Tall images** can be turned to landscape for the figure (clockwise or
  counter-clockwise); scale bars and zoom outlines follow.
- **Ctrl+wheel** sizes what is under the cursor (text, colour bar, scale-bar
  label, letters; Ctrl+Shift = panels), one undo step per gesture.
- **Exports**: Copy as PNG / SVG / both; PowerPoint to a new or the current
  slide, editable or as a picture; Obsidian (active note / choose); Inkscape
  (opens a layered SVG: images, outlines, scale bars, letters, colour bars,
  labels); Save as PDF / layered SVG / PNG / TIFF / EPS at 300-1200 dpi,
  optional transparent background.
- Fix: images quarter-turned in the viewer had their width/height swapped
  twice in figures.
- **Format panel** opens with the figure, in icon tabs (Page & layout,
  Channels, Titles, Colour, Panel letters, Scale bar, Zoom outlines) - no
  scrolling; closing it is remembered. Figures always open fitted to the
  window (actual size is a per-window view).
- **Format bar under the figure**: every option at once, as cards with an
  icon, a title and a one-line hint (Page & layout, Channels, Text,
  Colour, Panel letters, Scale bar, Zoom outlines); drag the splitter to
  give it more or less room. Drop-downs keep a fixed width and show full
  names in their lists, so nothing is cut off.
- **Font**: any installed font for the whole figure (Arial by default;
  journal-friendly fonts listed first), carried into PDF/SVG/PowerPoint.
- **Panel letters outside the image**: "Outside, above top left/right",
  on their own line above the header, in dark text on the page.
- **min / max colour bars**: the bar stays, its ends read "min" and "max";
  and **one bar per channel** even when each image has its own range.
- **Molecular model on every panel**: load a .pdb / .xyz / .mol (or take
  the model already placed in the preview) and it appears at the same spot
  on every panel - drag to move, Shift+drag to spin, Ctrl+Shift+drag or
  middle button to rotate in 3-D, or use the X/Y/Z orientation gizmo (as in
  the preview), X / Y / Z keys and arrows - and every panel follows live, so channels and heights compare against one model. Same
  renderer and styles as the preview (Appearance... opens the preview's
  dialog); atoms sized for the printed panel (Atom size to tune); on every
  panel or the first channel row; own Inkscape layer.
- **Automatic letter and scale-bar colour**: white or black per letter,
  judged on the image under it (grey rows get black letters, dark rows
  white); right-click a letter to give just that one its own colour.
- **PowerPoint at real size**: the figure lands at its printed size (an
  85 mm figure is 85 mm on the slide), like PDF and Inkscape.
- **Titles**: frame headers can be *None*; channel names beside the rows can
  be hidden (the layout reclaims the space); the Δz = 0 reference is bold
  only if you ask for it.

### 2026-11-03 (bias-series figures)

**The same one-click figure for bias series.** A frame re-scanned while the
sample bias steps (constant-current topography vs V, or a constant-height
frame at a ladder of biases) carries an amber **V×N** chip (N = distinct
biases); click it or right-click → *Bias-series figure*:
- frames ordered by bias, headers `V = -1.50 V` / `V = +300 mV` (always
  signed), the setpoint as a second line when it changes between frames;
- everything from the height-series figure works the same: any channel,
  transpose, Format panel / right-click options, ranges, Copy / PowerPoint /
  PDF, live link; headers also offer *Bias*, and *Reverse the frame order*
  is new for both recipes. The header choice is per recipe (Δz for height,
  V for bias), not carried from one to the other;
- CC and CH frames are grouped separately, so a CC bias ladder and a CH
  height ladder of the same spot stay two series.

Also for both recipes:
- **repeats collapse**: re-scans under identical conditions (bias, tip
  height, setpoint) show once - the latest - with "N repeats hidden ·
  Show"; chips count conditions, and a group that is only repeats gets no
  chip;
- **stopped scans**: rows a scan never reached are masked (they were a
  constant, e.g. z = 0 next to -185 nm real topography, and flattened the
  whole row), and frames with < 25 % of their rows scanned are left out
  ("scan stopped early", *Include anyway*). Series frames are checked in
  idle time after a folder load, so chip counts match the figure;
- topography rows share a height scale fitted to the *typical* frame, so
  a tip crash in one frame saturates there instead of darkening the rest;
- opening a figure can no longer take the application down: an error is
  logged and shown instead.

### 2026-11-03 (height-series figures)

**One click from a constant-height height series to a finished figure.**
Every thumbnail that belongs to a height series (same frame re-scanned at a
ladder of tip heights) now carries a green **z×N** chip. Click it (or
right-click → *Height-series figure*) and a figure opens with no dialog:
- columns = the frames, ordered by tip height, captioned **Δz** (closest
  approach = 0 pm unless your CH reference is in the series); rows = current
  (Blues_r) and frequency shift (gray); panel letters; one scale bar;
- **one colour range per row**, pooled with every frame weighted equally, in
  one display unit per row (no more 4 pA next to 0.4 nA);
- frames whose Z changes within the scan are left out (no single Δz) and
  listed in the strip with *Include anyway*; frames on a different Δf
  baseline (PLL lost lock, reference changed) are drawn on their own scale
  and tagged, dead frames dimmed as *flat signal* - they never wreck the
  shared range;
- long series wrap into balanced blocks of at most 8 columns;
- **any simultaneously acquired channel**: click a channel label to switch
  it (fwd/bwd, It_to_PC/It_ext, Δf/lock-in, ...), move or remove it, or
  add another one ("+ channel", Format → Channels); a channel with no
  contrast (Z of CH frames) shows its level instead of a blank block;
- **Transpose**: frames down and channels across, or the other way round;
- **Format** panel (and right-click on any element): panel letters (a / A /
  (a) / a) / a., one per frame or every panel, corner, colour, size, bold),
  scale bar (one panel / every panel / off, automatic or fixed length,
  corner, colour, label), colour bars (one range per channel or each image
  its own range with its own colourbar, 2/3/5 ticks, label), headers (Δz /
  absolute z / file name, plus the file name line), per-channel colormap,
  panel width and text size. Your format is remembered for new figures;
- colour ranges: drag a colourbar (either end, Shift = both), **click it to
  type exact Min/Max**, double-click for automatic - per channel, or per
  image when each image has its own range;
- edits happen on the figure: click a Δz header to make it the reference,
  hover a header and click × to drop a frame, double-click a panel to show
  it in the viewer, Ctrl+Z/Ctrl+Y;
- **Copy** (PNG + SVG on the clipboard), **PowerPoint** (editable vector
  shapes on a new slide) and **Save PDF/SVG/PNG** all come from the same
  renderer as the window;
- live: changing a filter, Zero, turn or display adjustment of a member
  updates the figure; edits are remembered per series (config + sessions).
Multi-selections that aren't a series get the same window via *Figure from
selection*. Fixes (also in the folder report): bias values honor the header
unit (Anfatec `Bias: -2000` with `BiasPhysUnit: mV` used to read as
-2000 V); Anfatec's tunnel current `It_to_PC` / `It_ext` is now recognised
as current (it was "other", so constant-height pages showed Δf only), with
`It_to_PC` preferred; `QPlusAmpl` / `Zeit` classify as amplitude / time.

### 2026-07-18 (folder-report data-sets & colormap gallery fixes)

**Folder reports now recognize "data-sets" (image sequences).** If a folder
contains the same frame scanned over and over while one acquisition
parameter is stepped, the report groups those images into a sequence and
gives it its own section instead of scattering them through the report:
- **Height (tip-sample distance) series** — constant-height current and
  frequency-shift images taken at a ladder of tip heights. Each image is
  labelled with how far the tip was retracted/approached relative to the
  first one (image 0 = 0, then +20 pm, +40 pm, …), read from the raw Z
  plane so a background-subtraction filter can't hide it. Drift alone can't
  fake a series — a frame only counts as a controlled height when it's
  genuinely flat.
- **Bias series** — the same spot scanned at a sequence of biases (STM
  current/topography vs. voltage).
- For constant-height data-sets the two channels are shown in **paired
  rows**, one file per column: current on top (in reversed Blues),
  frequency shift directly below (in gray), so the two channels of the same
  image always line up vertically.
- Validated on a real 224-scan folder: a 65-image height series came out
  correctly laddered in exact 10 pm steps.

**Report button works without selecting anything.** The toolbar **Report**
button used to stay greyed out until you clicked a thumbnail to preview an
image; it's now clickable as soon as a folder is loaded — the report
describes the whole folder, so no selection is needed.

**Colormap gallery reversal fix.** Because the default colormap is a
reversed map (`Blues_r`), opening the colormap gallery and clicking a plain
card used to silently give you the *reversed* version of that colormap
(click "Accent", get "Accent_r"), which could quietly become your session's
preview default. A plain card click now selects the colormap the normal way;
the little 🔄 on a card is still how you ask for a reversed version.

**Optional extra colormaps — how to add them if you already have SXM Viewer
installed.** SXM Viewer can use a big library of extra scientific colormaps
(the [`colormaps`](https://pratiman-91.github.io/colormaps/) package by
pratiman-91, ~970 maps) if it's present, but it's completely optional and
not required for anything. If you'd like them in the gallery and dropdowns:

1. Open a terminal (Anaconda Prompt on Windows) and activate the same
   environment you run SXM Viewer in — if you followed the Quick start
   above that's:
   ```powershell
   conda activate sxmviewer
   ```
2. Install the package with pip:
   ```powershell
   pip install colormaps
   ```
3. Restart SXM Viewer. The extra maps show up automatically in the colormap
   gallery (🎨) and the thumbnail/preview colormap dropdowns. You can check
   status any time under **Display → "Extra colormaps..."**, which also shows
   this install hint if the package isn't found.

If you use a different environment name, or a plain `python -m venv`
environment instead of conda, activate that one first; the only actual step
is `pip install colormaps` into whatever environment SXM Viewer runs in.

### 2026-07-21 (compare-all-spectra reachability)
- **"Compare All Spectra on This Image"** — the action that overlays every
  point spectrum acquired on an image in one comparison window — moved from
  the preview's *Overlays* submenu to its **Analysis** submenu (it opens a
  window, so it never really belonged among the paint-on-image toggles), and
  is now **also on the thumbnail right-click menu**, so you can launch it
  straight from the grid without opening the image first. It appears whenever
  an image has 2+ point spectra; grid/CITS points stay in the Grid Map
  Explorer.

### 2026-07-10 (spectroscopy position accuracy & browser overhaul)
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

### 2026-07-07 (spectroscopy UX pass)
A round of workflow polish for spectroscopic data (single spectra, Z series at one position, and grid/CITS maps), aimed at making the spectra ↔ image navigation symmetric and the click behavior predictable:

- **Show on image**: every spectroscopy window (single spectrum, comparison, grid map explorer) now has a "Show on image" button (also in its right-click menu) that focuses the main preview on the image the spectrum was acquired on, scrolls its thumbnail into view, and pulses its marker
- **Selection tray**: Shift+clicking markers no longer auto-opens a comparison window at the second selection; instead a small tray appears under the thumbnails ("N spectra selected - Compare | Clear") so you decide when to compare
- **"With spectroscopy" view filter** (`Ctrl+Alt+P`, also in Display → Show only and the Filter dropdown): show only the images that have linked spectroscopy points; press again to show all
- **Marker legend**: Spectroscopy toolbar → "What do these markers mean?" opens a legend rendered with your actual marker style/colors: single vs grid-map markers, the dashed low-confidence ring, the highlight glow, repeat/Z-series badges, and the per-image totals badge
- **Always-on presence banner**: as soon as a folder is scanned, a pill above the thumbnails announces what spectroscopy exists ("⚡ 1655 spectra · 14 grid maps · click to browse") - no toggles or menus needed to know the data is there; clicking it opens the Spectro Browser, which also lists spectra that aren't linked to any image
- **Scientist vocabulary**: window titles, tooltips, and summaries now say "Position 12.3/45.6 nm" instead of "site", "Z series" instead of "stack", and "Grid map" instead of "matrix" - e.g. "Z series (x5) - Position 12.3/45.6 nm" for a tip-height series

### 2026-07-07 (starred favourites and quick view filters)
You can now star your favourite images and browse only those:

- star/unstar thumbnails from the right-click menu ("★ Star" / "Remove star") or by pressing `S` while the thumbnail area has focus - `S` toggles, so pressing it again removes the star; starred images show a gold star badge and a brief star animation plays when you star one
- an in-app guide explains all of this: `Display → Show only → How favourites & filters work...`, and the shortcuts panel lists the new keys
- stars are remembered across sessions, so reopening a folder later still shows which images you starred
- new `Display → Show only` menu filters the thumbnail grid to Starred (`Ctrl+Alt+F`), Constant height (`Ctrl+Alt+H`), or Constant current (`Ctrl+Alt+C`) images - press the same shortcut again to show everything (the thumbnail Filter dropdown reflects and controls the same state)
- virtual copies follow their source image under these filters: a copy of a CH/CC image inherits the tag (and border badge), and a copy of a starred image starts out starred, so copies created while a filter is active stay visible instead of vanishing until you switch back to "All"

### 2026-07-05 (auto-enhance and periodic-noise removal)
Two new tools in the preview panel's Filters, aimed at quick cleanup of common scan problems:

- a "Let the robot" button that diagnoses the current image (tilt, incomplete scans, glitched scan lines, isolated spike pixels, high-frequency noise, poor contrast) and applies whichever fixes actually apply, as ordinary steps in the same filter pipeline you'd build by hand - fully visible and editable afterward, not a hidden transform
- a "Remove periodic noise..." dialog for scan-line banding and pump/mains-frequency vibration: shows the image's frequency spectrum, flags likely noise regions with a plain-language reason, and lets you draw your own regions (rectangle or ellipse) to remove or protect, since genuine surface structure can look similar to noise in frequency space and this always stays a manual, reviewed step
- Nanonis scan-timing metadata (used to match noise against known pump/mains frequencies) is now read correctly - a bug meant it was previously never being picked up at all

### 2026-07-04 (performance pass)
A round of work aimed at the two things that show up most during everyday use: browsing thumbnails/previews and loading a folder for the first time.

- clicking between thumbnails and previews is noticeably snappier - removed several redundant full-canvas re-renders that used to happen on every click
- fixed a couple of display glitches introduced while chasing that speed-up (an accumulating scale bar, and image/axis misalignment when switching between very differently-sized scans)
- loading a folder you've never opened before is roughly 4-5x faster for Nanonis `.sxm` conversion, now parallelized across CPU cores with verified byte-identical output
- fixed a few accidentally-quadratic slow paths in spectroscopy loading (matching spectra to the right image, building per-spectrum metadata) - up to ~60x faster in the worst cases
- the on-disk header cache no longer grows without bound across every folder you've ever opened

### 2026-05-07 (Nanonis loading performance)
Nanonis support has been updated with a focus on faster scan loading and reloads:

- converted `.sxm` channel caches now use binary NumPy `.npy` files instead of ASCII text exports
- this reduces conversion I/O overhead and speeds up subsequent channel reads
- automatic CH/CC tag detection now reuses cached results when the header and topography source have not changed
- warm folder reloads therefore avoid unnecessary topography re-reads when auto-tagging is enabled


