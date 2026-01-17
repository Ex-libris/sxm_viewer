"""Nanonis SXM file adapter using the local nanonispy2 library."""
from __future__ import annotations

import sys
import functools
from pathlib import Path
import numpy as np

# --- Step 1: Add local nanonispy2 to Python's path ---
# This allows us to import it without installing it system-wide.
try:
    # Navigate from this file (sxm_viewer/data/nanonis.py) up to the project root
    project_root = Path(__file__).resolve().parents[2]
    
    # The user's file tree shows a nested structure, so we target the inner folder
    # which contains the 'nanonispy2' package.
    nanonispy_lib_path = project_root / "nanonispy2-1.2.0" / "nanonispy2-1.2.0"
    
    if not nanonispy_lib_path.exists():
        # Fallback to the parent folder just in case
        nanonispy_lib_path = project_root / "nanonispy2-1.2.0"

    if nanonispy_lib_path.exists() and str(nanonispy_lib_path) not in sys.path:
        # Add the library path to the *beginning* of sys.path to give it priority
        sys.path.insert(0, str(nanonispy_lib_path))
        print(f"[sxm_viewer.nanonis] Added local nanonispy2 to path: {nanonispy_lib_path}")

except Exception as e:
    print(f"[sxm_viewer.nanonis] Could not add local nanonispy2 to path: {e}")


# --- Step 2: Apply compatibility patches for modern NumPy ---
# The nanonispy2 library uses `np.float`, `np.int`, `np.bool` which were removed.
if not hasattr(np, 'float'):
    np.float = float
if not hasattr(np, 'int'):
    np.int = int
if not hasattr(np, 'bool'):
    np.bool = bool


# --- Step 3: Import the library ---
nap = None
try:
    # We specifically import nanonispy2 as requested
    import nanonispy2 as nap
    print(f"[sxm_viewer.nanonis] Successfully imported local nanonispy2 from: {nap.__file__}")
except ImportError:
    nap = None
    print("[sxm_viewer.nanonis] WARNING: Failed to import local nanonispy2 library.")


# --- Adapter Functions ---

@functools.lru_cache(maxsize=32)
def get_scan_object(path: str):
    """
    Cached reader for Nanonis Scan objects. This avoids re-reading the entire
    .sxm file from disk every time a different channel is requested.
    """
    if nap is None:
        return None
    try:
        # Let nanonispy2 handle the file parsing
        return nap.read.Scan(path)
    except Exception as e:
        print(f"Error creating Nanonis scan object for {path}: {e}")
        return None


def is_nanonis_file(path):
    """Checks if a file has the .sxm extension."""
    return Path(path).suffix.lower() == '.sxm'


def parse_nanonis_header(path: Path):
    """
    Interface function: Parses an .sxm header using nanonispy2 and translates
    it into the dictionary format expected by the sxm_viewer GUI.
    """
    scan = get_scan_object(str(path))
    if not scan:
        return {}, []
    
    h = scan.header
    gui_header = {}
    
    # Map dimensions, range, center, angle, etc.
    if 'scan_pixels' in h:
        gui_header['xPixel'] = int(h['scan_pixels'][0])
        gui_header['yPixel'] = int(h['scan_pixels'][1])
    
    if 'scan_range' in h:
        gui_header['XScanRange'] = float(h['scan_range'][0])
        gui_header['YScanRange'] = float(h['scan_range'][1])
        gui_header['XPhysUnit'] = 'm'
        gui_header['YPhysUnit'] = 'm'
        
    if 'scan_offset' in h:
        gui_header['xCenter'] = float(h['scan_offset'][0])
        gui_header['yCenter'] = float(h['scan_offset'][1])
        
    if 'scan_angle' in h:
        gui_header['Angle'] = float(h['scan_angle'])
        
    if 'bias' in h:
        gui_header['Bias'] = float(h['bias'])
        gui_header['BiasPhysUnit'] = 'V'
        
    if 'rec_date' in h:
        gui_header['Date'] = str(h['rec_date'])
    if 'rec_time' in h:
        gui_header['Time'] = str(h['rec_time'])
        
    # Create the File Descriptor (fd) list for the GUI's channel dropdown
    fds = []
    for channel_name in scan.signals:
        unit = ""
        lower_name = channel_name.lower()
        if 'current' in lower_name: unit = 'A'
        elif 'z' in lower_name or 'height' in lower_name: unit = 'm'
        elif 'freq' in lower_name or 'shift' in lower_name: unit = 'Hz'
        elif 'bias' in lower_name: unit = 'V'
        elif 'cond' in lower_name or 'li' in lower_name: unit = 'S'
        
        fd = {
            'FileName': str(path.name),
            'Caption': channel_name,
            'PhysUnit': unit,
            'Scale': 1.0,
            'Offset': 0.0,
            'channel_key': channel_name  # Key for retrieving data later
        }
        fds.append(fd)
        
    return gui_header, fds


def read_nanonis_data(path: str, channel_key: str):
    """
    Interface function: Reads the 2D image data for a specific channel from an
    .sxm file using the cached nanonispy2 object and ensures it matches the
    structure expected by the GUI.
    """
    print(f"[nanonis.py] read_nanonis_data called for path='{path}', channel='{channel_key}'")
    scan = get_scan_object(path)
    if not scan:
        print(f"[nanonis.py] FAILED: could not get scan object for '{path}'")
        return np.zeros((128, 128))
    
    if channel_key in scan.signals:
        signal_data = scan.signals[channel_key]
        
        # Nanonis files can have forward and backward scans. We prefer forward.
        data = None
        if hasattr(signal_data, 'forward') and signal_data.forward is not None:
            data = signal_data.forward
        elif hasattr(signal_data, 'backward') and signal_data.backward is not None:
            data = signal_data.backward
        elif isinstance(signal_data, np.ndarray):
            data = signal_data
        
        if data is not None:
            print(f"[nanonis.py] SUCCESS: Got data for channel '{channel_key}'. Original dtype: {data.dtype}, shape: {data.shape}, max value: {np.max(np.nan_to_num(data))}")

            # The library should handle endianness, but if it fails, values are huge.
            # This is a safety net. Real SPM data is never this large.
            if data.dtype.kind == 'f' and np.max(np.abs(np.nan_to_num(data))) > 1e20:
                print(f"[nanonis.py] WARNING: Absurdly large values detected. Attempting to fix by swapping byte order (endianness).")
                data = data.byteswap().newbyteorder()

            # Convert to a standard float type for our GUI. This also ensures
            # the byte order is native to the system.
            data = data.astype(np.float64)
            
            # Nanonis format is typically (width, height). Our GUI needs (height, width).
            # So we must transpose.
            if 'scan_pixels' in scan.header:
                nx = int(scan.header['scan_pixels'][0])
                ny = int(scan.header['scan_pixels'][1])
                # If shape is (width, height), transpose it.
                if data.shape == (nx, ny):
                    data = data.T
            
            print(f"[nanonis.py] FINAL data shape: {data.shape}, max value: {np.max(np.nan_to_num(data))}")
            return np.nan_to_num(data) # Final cleanup of any remaining NaNs
            
    print(f"[nanonis.py] FAILED: Channel '{channel_key}' not found in signals for '{path}'")
    return np.zeros((128, 128))