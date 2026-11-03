"""Qt-free figure recipes (data-aware multi-panel figures).

Like ``reporting/`` and ``providers/``, nothing here imports Qt or the GUI:
recipes take plain per-image dicts and arrays gathered on the GUI thread
(``gui/figures/``) and turn them into a spec + a matplotlib Figure, so the
same renderer serves the on-screen window and the vector export, and the
logic is testable headlessly.
"""
