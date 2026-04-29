"""Vendored RHK SM4 reader.

This is sourced from the `spym` project (MIT license). We vendor it so the
viewer can import `.sm4` files without requiring an external dependency.
"""

from .spym_sm4 import RHKsm4  # noqa: F401

__all__ = ["RHKsm4"]

