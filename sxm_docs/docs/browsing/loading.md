# Loading Data

## Opening a folder

Click **Open folder** in the toolbar, or drag and drop a folder directly onto the main window. The viewer loads all recognised files into the thumbnail grid and automatically detects associated spectroscopy files.

A **Recent folders** drop-down on the Open folder button remembers up to 30 previously used paths. Use **Clear recent folders** at the bottom of that menu to reset it.

!!! note
    Opening a new folder always starts from a clean workspace: any open pop-outs and tool windows are closed first.

!!! tip
    SXM Viewer remembers the last folder you had open and reloads it automatically on startup (as long as that folder still contains data), so on relaunch you may already see your previous session's folder rather than an empty workspace.

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
| Omicron / Anfatec | `.txt` header + binary channel files | Native format; full header and multi-channel support |
| Nanonis | `.sxm` | Converted on load into an Omicron-style header/channel cache (see below); full multi-channel support |

`.xyz` files are handled two different ways depending on context: they can be **imported** as a 2D or 3D molecule-model overlay to place on top of a scan (see [Molecule Overlays](../image-analysis/molecules.md)), and a scan's own channel data can be **exported** as an XYZ point cloud (see [CSV & Data Export](../export-and-sharing/data.md)) - there is no XYZ import path for scan/image data itself.

See [Supported File Formats](../reference/file-formats.md) for full details.

!!! note
    The first time a Nanonis folder is loaded, SXM Viewer converts each `.sxm` scan into a cached header/channel representation stored in a `.sxmviewer_nanonis/` folder next to the data. This folder is safe to delete (it will be rebuilt), and speeds up subsequent loads of the same folder.

---

## Constant-height frame detection

When loading a folder, the viewer automatically detects **constant-height (CH)** and **constant-current (CC)** frames and tags each file accordingly. The `dz` offset is preserved per file so you can distinguish CH frames at a glance in the thumbnail grid.

---

## Sessions and collections

If you want to return to a curated workspace later, use [Sessions](sessions-and-collections.md) (folder-oriented) or [Collections](sessions-and-collections.md#collections) (cross-folder curated sets).