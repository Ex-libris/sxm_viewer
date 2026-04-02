# Configuration

SXM Viewer persists a range of user-interface and display preferences so the workspace behaves consistently across sessions.

---

## What is typically persisted

Examples described in the project history include:

- recent folders and recent session locations
- typography choices
- display state and overlay preferences
- filter parameter defaults such as recent Laplacian settings
- autosave / recovery behavior

---

## Why configuration matters

Persistent configuration helps SXM Viewer feel like a stable workspace rather than a stateless file viewer. It is especially useful for:

- restoring familiar display behavior
- keeping recent paths close at hand
- preserving figure and plotting style choices
- reducing repeated setup for common analysis tasks

---

## Session state vs configuration

It helps to distinguish two layers:

- **Configuration** stores general preferences and recent history.
- **Sessions** store the detailed state of a specific workspace.

If you want to reopen the exact working set of images, pop-outs, and overlays, use a session. If you just want the app to remember recent folders or style choices, that is configuration.

See [Sessions & Collections](../browsing/sessions-and-collections.md).
