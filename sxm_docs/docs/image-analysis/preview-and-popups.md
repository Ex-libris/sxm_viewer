# Preview & Pop-out Windows

## The main preview

Clicking a thumbnail loads it into the **main preview** pane. The preview shows the full image with colourmap, colorbar, scale bar, and any active overlays. A metadata panel alongside it shows acquisition parameters in correctly scaled SI units (e.g. SetPoint in pA or fA, not 0 A).

### Switching channels

Use the **channel selector** at the top of the preview area, or the **previous/next channel** arrow buttons, to switch between simultaneously acquired channels (topography, current, KPFM, etc.). Each channel has its own independent contrast and colormap state.

### Relative-zero display

Press ++0++ (or use **Display → Values relative to zero**) to shift the colorbar so it starts at zero. This is useful for constant-height images where absolute z values matter. The colorbar lower limit is clamped to zero in this mode; auto-contrast and range resets respect this constraint.

### Relative axes

Enable **Axes → Relative coordinates** to display image axes starting from zero (in physical units) instead of absolute piezo coordinates.

### Acquisition HUD

Enable **Display → Show acquisition overlay** to add a top-right HUD showing:

- **CC images**: Bias (V) and SetPoint (current)
- **CH images**: absolute z position (nm)

---

## Pop-out windows

**Double-click** any thumbnail to open it as a floating pop-out window. Pop-outs are independent: each has its own channel selector, contrast, overlays, and analysis state.

### Creating pop-outs from the preview

Right-click the preview canvas → **Pop out** to open the current view as a floating window.

### What pop-outs support

- Independent channel switching (prev/next arrows per popup)
- All overlay types: profiles, angles, molecules, scale bar, acquisition HUD
- Local relative-zero toggle (independent from main preview)
- Crop and quick-crop workflows
- Profile measurement dialogs (detached, independent of main window stacking)
- **Apply this style to all pop-ups** — copies font scale, typography, and display layout from the active popup to all others

### Managing pop-outs

| Action | How |
|---|---|
| Bring all to front | ++ctrl+shift+p++ or toolbar **Pop-ups** menu |
| Minimize all | ++ctrl+shift+m++ |
| Arrange / tile | **Pop-ups** toolbar split-button → Arrange |
| Close all | **Pop-ups** toolbar menu → Close all |
| Reopen last closed | ++ctrl+z++ (with nothing to undo) |

The **Pop-ups** toolbar button is a split button: primary click recalls open pop-outs; the menu gives per-window focus, arrange, minimize, and restore actions.

### Popup size badge

Pop-outs show the physical scan dimensions (or pixel dimensions as fallback) in the window title bar.

---

## Display options that sync across windows

Many display settings made in the right-click menu propagate automatically to the main preview and all open pop-outs:

- Ticks, colorbar visibility and orientation
- Title, acquisition HUD, shortcut hint
- Profile and angle overlay visibility
- Molecules, scale bar
- Frame fill mode
- Relative axes override
- Layout mode