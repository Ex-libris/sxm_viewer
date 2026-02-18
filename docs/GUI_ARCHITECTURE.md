# GUI Architecture

The GUI layer is organised around a thin `MainWindow` coordinator and a set of
feature–specific controllers. Widgets and dialogs live in their own modules so
each subsystem can evolve independently.

```
gui/
├── main_window.py
├── controllers/
│   ├── preview_popup.py        # pop-out spawning, toolbar wiring
│   ├── thumbnail_controller.py # thumbnail selection & drag interactions
│   ├── histogram.py            # histogram & live preview management
│   ├── profile.py              # profile dialog + measurement plumbing
│   ├── spectro_compare.py      # spectroscopy comparison workflows
│   └── quick_crop.py           # quick crop overlays & history
├── viewer/                     # reusable widgets/canvases
├── dialogs/                    # modal dialogs (histogram, profiles, filters…)
└── …
```

## Main window responsibilities

`main_window.py` now focuses on:

* creating the shared widgets (thumbnail list, preview canvas, toolbars)
* instantiating controllers and passing them the widgets they need
* orchestrating high-level application events (folder load, theme switch)

Whenever a subsystem grows complex, it is peeled into a controller inside
`gui/controllers/`. Controllers expose a compact API (e.g. `show_histogram()`,
`handle_thumbnail_event()`) and emit callbacks for the main window to react to.

## Controllers

| Controller            | Responsibilities                                                                 |
| --------------------- | --------------------------------------------------------------------------------- |
| `preview_popup.py`    | builds/updates pop-out windows, syncs toolbars and state                          |
| `thumbnail_controller.py` | handles thumbnail clicks, drag-to-canvas, keyboard navigation, CH/CC tagging |
| `histogram.py`        | manages the Histogram & Range dialog, auto CLIM, live preview updates             |
| `profile.py`          | wires canvases to the `ProfileDialog`, keeps measurements and overlays in sync    |
| `spectro_compare.py`  | gathers spectroscopy selections, opens comparison windows, trace export, minima   |
| `quick_crop.py`       | manages the quick crop panel, overlays, and pop-out history                       |
| `session.py`          | serialises/deserialises the full viewer state to session JSON files               |

Each controller receives only the widgets or callbacks it needs, keeping
dependencies explicit and testable.

## Widgets & dialogs

Custom canvases (`viewer/`), dialogs (`dialogs/`), and layout helpers remain
pure UI components. They emit Qt signals but avoid referencing the main window
directly. Controllers subscribe to their signals and issue updates.

## Extending the GUI

When a new feature touches several widgets:

1. Prototype inside `main_window.py`.
2. Once the workflow solidifies, move it into a controller module with a clear
   public surface.
3. Document the controller at the top of the file (expected widgets, signals).
4. Wire it up in the main window constructor via composition.

This keeps `main_window.py` readable and makes it obvious where to look when a
bug is reported for the thumbnails, histogram, spectroscopy comparison, etc.
