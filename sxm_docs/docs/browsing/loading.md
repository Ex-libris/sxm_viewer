# Loading Data

## Opening a folder

Click **Open folder** in the toolbar, or drag and drop a folder directly onto the main window. The viewer loads all recognised files into the thumbnail grid and automatically detects associated spectroscopy files.

A **Recent folders** drop-down on the Open folder button remembers up to 30 previously used paths. Use **Clear recent folders** at the bottom of that menu to reset it.

!!! note
    Opening a new folder always starts from a clean workspace: any open pop-outs and tool windows are closed first.

---

## Opening individual files

You can drag and drop one or more individual files onto the main window. Dropped files are **appended** to the current thumbnail list rather than replacing it, so you can curate a set of images from different folders.

!!! tip
    Dropping a folder replaces the current session; dropping individual files extends it.

A **Clear thumbnails** button in the thumbnails toolbar wipes the current list so you can start fresh.

---

## Supported file formats

| Format | Extension | Notes |
|---|---|---|
| Anfatec SXM | `.sxm` | Primary format; full header and multi-channel support |
| MATRIX | `.mtrx` | Vendor import; see [MATRIX scans](../spectroscopy/matrix.md) |
| Nanonis | various | Single-file and grid import |
| WSxM XYZ | `.xyz` | Import and export |

See [Supported File Formats](../reference/file-formats.md) for full details.

---

## Constant-height frame detection

When loading a folder, the viewer automatically detects **constant-height (CH)** and **constant-current (CC)** frames and tags each file accordingly. The `dz` offset is preserved per file so you can distinguish CH frames at a glance in the thumbnail grid.

---

## Sessions and collections

If you want to return to a curated workspace later, use [Sessions](sessions-and-collections.md) (folder-oriented) or [Collections](sessions-and-collections.md#collections) (cross-folder curated sets).