# Sessions & Collections

SXM Viewer has two complementary ways to save and restore your work.

---

## Sessions

A **session** is folder-oriented: it saves the full state of a working folder, including the current preview, open pop-outs, profile measurements, crop history, display options, and spectroscopy selection.

### Saving a session

- Press ++ctrl+s++ to save to the current session file. If none exists yet, you are prompted once for a location.
- Use **File -> Save Session** from the toolbar.

**Autosave** is enabled by default every 5 minutes. Recovery controls live in the top-level **Tools** menu, where you can enable or disable autosave, change the interval, recover the latest autosave, or discard it.

### Loading a session

- Use **File -> Load Session** from the toolbar.
- A **Recent sessions** drop-down remembers previously used session files and folders.

### What is restored

Sessions restore:

- cached headers and processed views for fast first paint
- thumbnail-grid state and the active preview
- all open pop-out windows with their geometry and analysis state
- profile-measurement window positions
- spectroscopy selection and browser state
- display options such as colormap, overlays, typography, and dark/light mode

!!! tip
    Pop-outs from a previous session are loaded as **deferred pop-outs** to avoid opening a storm of windows on startup. A **Pop-ups (N)** toolbar menu lets you restore them one by one or all at once.

---

## Collections

A **collection** is cross-folder: it saves a curated set of selected previews, pop-outs, and crop snapshots gathered from multiple folders or sessions. Collections are independent of any single folder.

### When to use collections

Use a collection when you want to assemble a set of "hero" images from different experiments without committing to saving a full folder session for each one.

### Saving a collection

Entry points:

- toolbar collection actions
- preview or pop-out **Collection** submenus
- right-click on any preview or popup -> **Collection**

Available actions:

| Action | What is saved |
|---|---|
| Add Current Preview | The main preview at its current state |
| Add Active Pop-up | The focused pop-out with all its analysis overlays |
| Add All Open Pop-ups | Every currently open pop-out |
| Add Selected Crop History | Chosen entries from the crop history panel |

### Linked vs portable

When saving, you choose between two modes:

**Linked** saves references to the original source files. Smaller file size, but requires access to the original data path when reopening.

**Portable** caches all image arrays into the collection file. Larger file, but safe to move to another machine or share with a colleague.

### Restoring a collection

Opening a collection clears the current workspace and rebuilds it as a curated virtual set. Items flagged as pop-outs are reopened as pop-out windows. All per-item analysis state such as profiles, angles, molecules, and scale-bar settings is reapplied.

The collection file format is `.sxmcoll.json`.
