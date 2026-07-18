"""Small Qt boilerplate helpers shared across the GUI layer.

Kept deliberately tiny and dependency-free (only ``_shared``) so any GUI
module can import it without circularity.

See ``docs/refactor/PATTERNS.md`` - these exist to replace hand-rolled
idioms. Prefer them over re-implementing the pattern inline; the analysis
toolkit (``scripts/analysis/find_idioms.py``) reports regressions.
"""
from __future__ import annotations

# Property name -> Qt setter, for the shapes actually used in this codebase.
_SETTERS = {
    "checked": "setChecked",
    "text": "setText",
    "current_text": "setCurrentText",
    "current_index": "setCurrentIndex",
    "value": "setValue",
    "enabled": "setEnabled",
    "visible": "setVisible",
    "range": "setRange",
    "maximum": "setMaximum",
    "minimum": "setMinimum",
}


def set_silent(widget, **props):
    """Set widget properties without emitting their change signals.

    Replaces the hand-rolled triad that appears ~245 times in this
    codebase::

        try:
            w.blockSignals(True)
            w.setChecked(state)
            w.blockSignals(False)     # <-- bug
        except Exception:
            pass

    Two things this fixes beyond brevity:

    * **Restores the previous block state** instead of hardcoding
      ``False``. The hand-rolled form unblocks signals unconditionally, so
      calling it inside an outer block silently re-enables handlers the
      caller deliberately suppressed. That is a real latent bug, not
      merely verbosity.
    * **Null-safe**: a missing widget is a no-op, which is why most of
      those call sites needed a surrounding ``try/except`` at all.

    Property names are the snake_case forms in ``_SETTERS``; an unknown
    name falls back to ``set<CamelCase>``. Unknown/absent setters are
    ignored rather than raising - this is UI-sync glue, never worth
    crashing the app over.

    Returns True when the widget existed.
    """
    if widget is None:
        return False
    prev = None
    blocked = False
    try:
        prev = widget.blockSignals(True)
        blocked = True
        for name, value in props.items():
            setter_name = _SETTERS.get(name)
            if setter_name is None:
                setter_name = "set" + "".join(
                    part.capitalize() for part in name.split("_"))
            setter = getattr(widget, setter_name, None)
            if setter is None:
                continue
            try:
                if isinstance(value, tuple):
                    setter(*value)
                else:
                    setter(value)
            except Exception:
                continue
    except Exception:
        return False
    finally:
        if blocked:
            try:
                widget.blockSignals(prev)
            except Exception:
                pass
    return True


def set_many_silent(widgets, **props):
    """``set_silent`` over an iterable, skipping ``None`` entries.

    The mirrored-setting case: one logical value shown by a checkbox *and*
    a menu action (and sometimes a toolbar button), all of which must be
    updated without re-triggering the handler that is currently running.
    """
    count = 0
    for widget in widgets or ():
        if set_silent(widget, **props):
            count += 1
    return count


__all__ = ["set_silent", "set_many_silent"]
