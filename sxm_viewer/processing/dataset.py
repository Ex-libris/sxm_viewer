"""High-level services for loading SXM folders."""
from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Dict, Optional

import numpy as np

from sxm_viewer.data.io import parse_header, read_channel_file, normalize_unit_and_data
from sxm_viewer.utils.logging import log, log_progress


@dataclass
class ChannelDescriptor:
    caption: str
    file_name: str
    phys_unit: str
    scale: float = 1.0
    offset: float = 0.0


@dataclass
class SXMFile:
    header_path: Path
    header: dict
    channels: List[ChannelDescriptor]


@dataclass
class SXMFolder:
    files: List[SXMFile] = field(default_factory=list)
    headers_by_path: Dict[str, SXMFile] = field(default_factory=dict)

    def load_folder(self, folder):
        folder = Path(folder)
        if not folder.exists():
            raise FileNotFoundError(folder)
        log(f"Loading folder {folder}")
        txts = sorted(folder.glob('*.txt'))
        self.files.clear(); self.headers_by_path.clear()
        total = len(txts)
        for idx, txt in enumerate(txts, 1):
            try:
                header, fds = parse_header(txt)
            except Exception as exc:
                log(f"Skipping {txt.name}: {exc}")
                continue
            channels = []
            for fd in fds:
                channels.append(ChannelDescriptor(
                    caption=fd.get('Caption', fd.get('FileName','')),
                    file_name=fd.get('FileName'),
                    phys_unit=fd.get('PhysUnit',''),
                    scale=float(fd.get('Scale',1.0)),
                    offset=float(fd.get('Offset',0.0))
                ))
            sxm_file = SXMFile(header_path=txt, header=header, channels=channels)
            self.files.append(sxm_file)
            self.headers_by_path[str(txt)] = sxm_file
            if idx % max(1, total//10 or 1) == 0 or idx == total:
                log_progress('Parsing headers', idx, total)
        log(f"Loaded {len(self.files)} descriptor(s)")

    def list_channel_labels(self) -> List[str]:
        if not self.files:
            return []
        first = self.files[0]
        labels = []
        for idx, ch in enumerate(first.channels):
            labels.append(f"{idx}: {ch.caption or ch.file_name or f'chan{idx}'}")
        return labels

    def load_channel_array(self, path: str, channel_index: int):
        sxm = self.headers_by_path.get(path)
        if not sxm:
            raise KeyError(path)
        if channel_index < 0 or channel_index >= len(sxm.channels):
            raise IndexError(channel_index)
        ch = sxm.channels[channel_index]
        header = sxm.header
        xpix = int(header.get('xPixel', 128))
        ypix = int(header.get('yPixel', xpix))
        arr = read_channel_file(sxm.header_path.parent / ch.file_name, xpix, ypix,
                                scale=ch.scale, offset=ch.offset)
        unit, arr = normalize_unit_and_data(arr, ch.phys_unit)
        return arr, unit

    def channel_extent(self, path: str):
        sxm = self.headers_by_path.get(path)
        if not sxm:
            return None
        header = sxm.header
        xr = header.get('XScanRange'); yr = header.get('YScanRange')
        if xr and yr:
            return (0.0, float(xr), float(yr), 0.0)
        return None
