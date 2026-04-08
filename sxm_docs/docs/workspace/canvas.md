# Publication Canvas

![Canvas export workflow](../assets/screenshots/canvas-export-flow.gif){ width="900" }

The **Canvas** is a multi-image figure composer for building publication-ready layouts from your SPM data.

Open it with the **Canvas** toolbar button or **File → Open Canvas**.

---

## Adding images

Drag thumbnails from the main grid onto the canvas to add them as tiles. Each tile is an independent image with its own colormap, contrast, overlays, and display state.

---

## Layout

The canvas supports flexible tile arrangements:

- **Stack layout** — tiles arranged vertically
- **Column layout** — tiles arranged in columns
- Drag tiles to reorder

---

## Per-tile overlay chips

When a tile is selected, overlay chips appear directly on it for instant toggling:

| Chip | Toggles |
|---|---|
| T | Tile title |
| S | Scale bar |
| C | Colorbar |
| M | Metadata bar |
| U | Unit badge |
| F | Filename badge |

Each chip updates the tile immediately, saves to undo history, and keeps the inspector in sync.

---

## Display presets

The canvas toolbar offers three global presets:

| Preset | Description |
|---|---|
| Clean | Minimal overlays, no tick labels |
| Analysis | Colorbars, scale bar, acquisition metadata |
| Publication | Journal-ready: scale bar only, no ticks, no title clutter |

Making any manual display change marks the state as **Custom**.

---

## Right-click canvas menu

The shared canvas right-click menu exposes:

**Range**
: Auto-range selected tiles, copy range to selected, sync ranges across tiles.

**Colormap**
: Copy colormap to selected, common colormap presets, sync colors by channel across tiles.

**Display**
: Show/hide metadata bar, unit badge, title, colorbar, colorbar ticks, scale bar, colorbar position.

**Alignment, overlay, view, and layout** actions are also available.

---

## Molecule overlays on canvas

The canvas left rail includes molecule controls: **Show**, **Load onto selected**, and **Clear from selected**. Canvas tiles carry molecule overlay state and render it directly.

---

## Export

Export the full canvas figure as PNG, SVG, or PDF via **File → Export** or the right-click menu. See [Export to SVG & PDF](../export-and-sharing/vector.md).

PowerPoint export preserves the source aspect ratio of each tile — square scans stay square. See [Export to PowerPoint](../export-and-sharing/powerpoint.md).
