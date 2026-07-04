# Thumbnail Grid

The thumbnail grid is the main browsing surface in SXM Viewer. It is designed for fast navigation through large folders without forcing a full preview render for every item.

---

## Basic interactions

| Gesture | Effect |
|---|---|
| Click | Select thumbnail and load it into the preview |
| Double-click | Open the image as a pop-out |
| Shift+Click | Range-select |
| Ctrl+Click | Add or remove single items from selection |
| Drag | Rubber-band multi-selection |

These same patterns are also used for spectroscopy miniatures where possible.

---

## What thumbnails show

Depending on the file and current workflow, thumbnails may show:

- image previews for scan files
- CH / CC tags and related metadata
- spectroscopy markers or miniature spectroscopy plots
- crop or virtual-copy variants inserted near the source image
- selection highlights for batch operations

The current preview is tied directly to thumbnail selection.

### Source-file actions

Right-clicking an image thumbnail or a spectroscopy miniature exposes a **Source file** submenu so you can:

- show the file in the operating-system file manager
- open the raw file in the default text editor for the current OS
- copy the full file path

---

## Sorting and filtering

Two drop-downs above the grid control how thumbnails are arranged:

**Sort**:

- **Name (A-Z)** - natural sort (handles embedded numbers sensibly, e.g. `scan2` before `scan10`)
- **Date (new-old)** / **Date (old-new)** - by acquisition date/time from the file header
- **Tag (CH-CC-U)** - groups constant-height scans, then constant-current, then untagged

**Filter**:

- **All**
- **Constant height**
- **Constant current**
- **Untagged**
- **Matrix datasets**

These are tag/type-only filters - there is no free-text search box. Sorting and filtering apply to the currently loaded folder; they don't change which files are loaded, only how they're arranged and which ones are shown.

---

## Multi-selection

Selecting several thumbnails lets you:

- batch-copy images
- batch-apply filters
- export multiple items
- curate a subset before creating a collection

Press ++ctrl+a++ to select all visible thumbnails.

---

## Virtual copies and processed entries

The grid can contain more than raw source files. It can also include:

- crop snapshots
- processed virtual copies
- popup-derived copies inserted near the source image

These entries retain their own display and analysis state and can be reopened later like normal thumbnails.

See [Cropping](../image-analysis/cropping.md) and [Sessions & Collections](sessions-and-collections.md).

---

## Drag and drop

You can drag files or folders into the application window to load them. Explicit file drops append to the current thumbnail session instead of replacing it, which is useful when curating images from different locations.

A **Clear thumbnails** button lets you wipe the current list and start again.

---

## Spectroscopy miniatures

Associated spectroscopy entries appear inside the same browsing workspace. Single click selects them, double-click opens a spectroscopy popup, and Shift/Ctrl selection can build multi-spectrum plots.

For more, see [Spectroscopy Overview](../spectroscopy/overview.md) and [Spectroscopy Browser](../spectroscopy/browser.md).
