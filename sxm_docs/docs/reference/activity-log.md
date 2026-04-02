# Activity Log

The Activity Log is an in-app record of status messages emitted during loading, export, and other operations.

---

## What it is for

Use the Activity Log when you want to see what the application is doing without watching the terminal.

Typical examples include:

- folder-load progress
- export progress
- parser or file-rejection messages
- general status updates from longer operations

---

## Behavior

The log lives inside the main window as a collapsible panel. Project history notes also describe:

- timestamps
- auto-scrolling
- a clear button
- a bounded line count to keep the widget manageable

---

## Why it is useful

The Activity Log is especially helpful when a folder contains a mix of valid and invalid files, or when you want to understand why a long-running load produced the resulting workspace.

---

## Related pages

- [Loading Data](../browsing/loading.md)
- [Supported File Formats](file-formats.md)
- [Configuration](configuration.md)
