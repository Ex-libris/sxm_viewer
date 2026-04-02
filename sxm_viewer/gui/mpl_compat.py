"""Small Matplotlib compatibility helpers for GUI plotting code."""
from __future__ import annotations

try:
    from mpl_toolkits.axes_grid1.inset_locator import InsetPosition as InsetPosition
except ImportError:
    # Matplotlib 3.8+/3.9+ removed InsetPosition from inset_locator.
    # Keep the same callable locator API expected by set_axes_locator().
    from matplotlib.transforms import Bbox, BboxTransform, TransformedBbox

    class InsetPosition:  # type: ignore[no-redef]
        """Compatibility replacement for axes_grid1.inset_locator.InsetPosition."""

        def __init__(self, parent, lbwh):
            self._parent = parent
            self._lbwh = lbwh

        def __call__(self, ax, renderer):
            bbox_parent = self._parent.get_position(original=False)
            trans = BboxTransform(Bbox.unit(), bbox_parent)
            bbox_inset = Bbox.from_bounds(*self._lbwh)
            return TransformedBbox(bbox_inset, trans)
