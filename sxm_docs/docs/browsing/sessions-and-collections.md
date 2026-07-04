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

A **collection** is cross-folder: it is a curated, named list of scans and spectroscopy gathered while browsing several different folders, that you can reopen later like a virtual folder. Unlike a session, a collection is not tied to any single folder or working directory.

Collection items are stored as **lightweight references** to the real file, channel, and any associated spectroscopy - not a rendered copy. Opening a collection reads the real files directly, so all channels, measurements, and filters work exactly as they would from a normal folder load. The one exception is pop-up and crop-history items, which carry their own overlay/crop state and are saved as a snapshot instead (see [Adding items](#adding-items) below).

!!! tip
    Use a collection when you want to sort images from one or more folders into a curated set - or several different sets at once - without saving a full folder session for each one.

### The current collection

At any time, one collection can be the app's **current collection** - the default target that plain "Add to Collection" actions append to. It is shown in the toolbar **Collections** button (e.g. *Collection: myset (12 items)*) and persists across app restarts.

- **Create a Collection...** starts a brand-new collection and makes it current. Nothing is written to disk until you actually add an item to it.
- **Open a Collection...** opens the collection browser: a list of your recent collections (plus a **Browse for Another Collection...** option) with a live preview - item count, last-modified time, and a thumbnail strip - for whichever one is selected. From there you choose one of two explicit actions:
    - **Set as Current Collection** - only changes the append target; your loaded folder and thumbnails are untouched.
    - **Open (Load Into Workspace)** - fully replaces your current workspace with this collection's contents, like opening a folder.
- **Recent Collections** (in the same toolbar menu) lists collections you've used recently for one-click reopening - clicking an entry always loads it (same as "Open (Load Into Workspace)" above).
- **Clear Current Collection Target** forgets the current collection for this session without deleting anything.

Opening a different collection, or picking one from Recent Collections, never overwrites or alters the collection you were previously working in - each collection file is only ever appended to when you explicitly add items to it.

### Adding items

Entry points:

- toolbar **Collections** menu
- right-click a thumbnail selection -> **Collections**
- preview or pop-out **Collections** submenus

| Action | What is saved | Storage |
|---|---|---|
| Add Selected Thumbnails | The selected thumbnail(s), with their associated spectroscopy | Lightweight reference |
| Add Current Preview | The main preview's file and channel | Lightweight reference (falls back to a snapshot only if the view is a crop or otherwise has no real file behind it) |
| Add Active Pop-up | The focused pop-out, with all its analysis overlays | Snapshot |
| Add All Open Pop-ups | Every currently open pop-out | Snapshot |
| Add Selected Crop History | Chosen entries from the crop history panel | Snapshot |

Pop-ups and crop-history items are saved as a snapshot because they carry derived state - crop region, filter pipeline, profile/angle/molecule overlays - that a plain file+channel reference cannot represent. They are heavier and less robust if the original file is later moved, but faithfully restore that exact analysis state when reopened.

#### Adding to a specific collection without switching

Use **Add Selected Thumbnails to...** to route a batch of thumbnails to a collection you pick right now - from your recent list, browsed for, or created fresh via an explicit **+ New Collection...** option - without changing the current collection. This is the way to sort different selections from the same folder session into several different collections in one sitting: select a few thumbnails, route them to collection A, select a different few, route them to collection B, and so on, with no need to switch the current-collection target back and forth in between.

### Folder awareness

Reopening a folder that already has files sorted into one or more collections shows a one-line notice naming them, so you don't lose track of previous sorting across sessions.

### File format and compatibility

The collection file format is `.sxmcoll.json`. Collections created by older versions of SXM Viewer (which saved every item as a rendered snapshot) still open correctly - items whose source file can still be found are read as live references with full functionality; items whose source file is missing fall back to their original saved snapshot.
