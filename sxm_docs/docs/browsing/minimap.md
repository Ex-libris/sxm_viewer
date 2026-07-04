# Minimap

The minimap ("Folder layout") shows where each scan in the current folder was physically taken, laid out spatially - not a scaled-down copy of the thumbnail grid's layout.

---

## What it shows

Each scan frame is drawn as a small rectangle at its real acquisition position, color-coded by tag:

- **Green** - constant-height (CH) frames
- **Blue** - constant-current (CC) frames
- **Gray** - untagged frames

Hovering over a frame shows its filename in a tooltip.

Toggle **Show real view** to switch from flat colored rectangles to actual rendered channel thumbnails at each position - useful when you want to recognize a specific scan by its appearance rather than by color/tag alone.

---

## Navigating

- **Click** a frame to load it into the main preview.
- **Shift+click** a frame to hide it from the layout (useful for decluttering a busy area); use **Show all frames** to bring back everything you've hidden.
- **Mouse wheel** zooms in/out, centered on the cursor position. A separate **zoom slider** and **Reset** button next to the panel do the same thing without needing the cursor over the map.
- **Middle-drag** or **right-drag** pans the view.

---

## Typical workflow

Use the thumbnail grid for precise selection, and the minimap for fast spatial navigation - especially useful when scans overlap or were taken at slightly different positions across the same sample area, since the minimap is the only view that shows those positions to scale.

A common pattern is:

1. Use the minimap to spot the physical region you're interested in (helped by the CH/CC color coding).
2. Click the relevant frame to load it into the main preview.
3. Use the thumbnail grid or spectroscopy markers for more precise selection from there.

---

## Relationship to the grid

The minimap does not replace the main thumbnail grid - it's a complementary spatial view. Selections, preview loading, and pop-out actions still happen in the main grid; the minimap is for figuring out *where* to look next.

---

## Tips

!!! tip
    If you are working with spectroscopy-heavy folders, use the minimap to navigate to the relevant image region, then use the main grid or spectroscopy markers for precise selection.

!!! tip
    Turn on **Show real view** when you need to visually recognize a scan rather than rely on its CH/CC color alone.
