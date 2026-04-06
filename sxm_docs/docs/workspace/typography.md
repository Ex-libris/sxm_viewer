# Dark Mode & Typography

SXM Viewer supports shared styling for both the application theme and plot typography.

---

## Dark and light mode

The GUI supports both light and dark presentation modes. Theme changes propagate through the workspace so preview canvases, pop-outs, and canvas-related controls remain visually consistent.

---

## Shared typography model

Preview canvases, pop-outs, profile dialogs, and spectroscopy windows use the same typography model.

Available styling includes:

- font family
- bold
- italic
- underline

These settings are propagated from the main window and persisted in configuration so they remain consistent across different plotting surfaces.

---

## Where typography is applied

Typography changes affect more than just titles. The project history explicitly notes styling for:

- plot titles
- tick labels
- colorbar labels
- acquisition HUD text
- shortcut hint text
- scale bars

That means a typography change can noticeably alter the whole visual language of the workspace.

---

## Popup and plot scaling

In several plotting windows, holding ++ctrl++ while using the scroll wheel adjusts font scale for the current view. This is especially useful in pop-outs and spectroscopy plots when preparing screenshots or presentation figures.

---

## Related pages

- [Preview & Popups](../image-analysis/preview-and-popups.md)
- [Publication Canvas](canvas.md)
- [Display Presets](presets.md)
