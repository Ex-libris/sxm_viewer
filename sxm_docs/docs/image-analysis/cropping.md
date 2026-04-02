# Cropping

SXM Viewer has two complementary cropping workflows.

---

## Quick crop (Shift+Click)

Hold **Shift** and click anywhere on the preview or a pop-out to crop a fixed-size square region centred on the click point. The cropped region opens immediately as a new pop-out.

Shift+click crops respect the current display orientation so the extracted data is consistent whether relative or absolute axes are active.

### Crop history

Every quick crop is recorded in the **crop history panel**. Each row has a checkbox that controls whether that crop outline is drawn on the source image. You can:

- Show/hide individual crop outlines via their row checkboxes
- Open any past crop as a pop-out again
- Add selected crop history entries to a [Collection](../browsing/sessions-and-collections.md#collections)

---

## Crop frame editor

The crop frame editor lets you draw, move, resize, and rotate an arbitrary crop region before committing.

**To enter editor mode**: right-click → **Quick tools → Edit crop frame**, or press ++ctrl+e++.

### Interactions

| Gesture | Effect |
|---|---|
| Drag corner handles | Resize the frame |
| Drag frame body | Move the frame |
| Drag R handle | Rotate the frame |
| ++ctrl++ + drag body | Rotate (alternative) |
| ++enter++ | Apply the crop |
| ++ctrl+e++ | Exit editor without cropping |

The crop is extracted with rotated resampling so the full output frame is filled without edge gaps. The top of the output corresponds to the side of the frame opposite the rotate handle, consistently regardless of display mode.

After applying a template crop, editor mode exits automatically and the frame is hidden.

### Crop frame on pop-outs

The crop frame editor works identically on pop-out windows. Cropped results from pop-outs go through the same crop pipeline as main-preview crops and appear in the thumbnail grid as virtual copies inserted next to the source image.

---

## Virtual copies

Any crop result (quick crop or frame editor) can be saved as a **virtual copy** in the thumbnail grid:

- Right-click a pop-out or preview → **Create virtual copy in thumbnails**
- Drag a pop-out window onto the thumbnail area to insert a snapshot at the drop position

Virtual copies carry their analysis state (overlays, display settings) and stay ordered relative to their source image even after grid refreshes.

---

## Undo

++ctrl+z++ undoes the most recent canvas edit, including crop operations. The undo stack covers filters, profile/angle overlays, molecule edits, and contrast changes, not just crops.