"""Shared state container for SXMGridViewer caches."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List


@dataclass
class ViewerState:
    """Collect mutable caches used across viewer helper modules."""

    files: List[str]
    headers: Dict[str, object]
    header_cache: Dict[str, object]
    thumb_cache: Dict[str, object]
    thumb_data_cache: Dict[str, object]
    topo_stats_cache: Dict[str, object]
    channel_data_cache: Dict[str, object]
    filtered_channel_cache: Dict[str, object]
    thumb_labels: Dict[str, object]
    frame_entry_pixmaps: Dict[str, object]
    frame_real_pixmap_cache: Dict[str, object]
    spectro_cache: Dict[str, object]
    spectro_hist_cache: Dict[str, object]
    matrix_datasets: Dict[str, object]

    @classmethod
    def from_viewer(cls, viewer) -> "ViewerState":
        """Build a state container with references to viewer cache objects."""
        return cls(
            files=viewer.files,
            headers=viewer.headers,
            header_cache=viewer.header_cache,
            thumb_cache=viewer.thumb_cache,
            thumb_data_cache=viewer._thumb_data_cache,
            topo_stats_cache=viewer._topo_stats_cache,
            channel_data_cache=viewer._channel_data_cache,
            filtered_channel_cache=viewer._filtered_channel_cache,
            thumb_labels=viewer._thumb_labels,
            frame_entry_pixmaps=viewer.frame_entry_pixmaps,
            frame_real_pixmap_cache=viewer._frame_real_pixmap_cache,
            spectro_cache=viewer._spectro_cache,
            spectro_hist_cache=viewer._spectro_hist_cache,
            matrix_datasets=viewer.matrix_datasets,
        )


__all__ = ["ViewerState"]



