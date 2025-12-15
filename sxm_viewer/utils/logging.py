"""Shared logging helpers for the refactored viewer."""
import sys
from contextlib import contextmanager


def log(message: str):
    sys.stdout.write(f"{message}\n")
    sys.stdout.flush()


def log_progress(prefix: str, current: int, total: int):
    pct = (current/total*100) if total else 0
    log(f"{prefix} [{current}/{total} | {pct:4.0f}%]")


@contextmanager
def stage(message: str):
    log(f"--> {message} ...")
    try:
        yield
        log(f"<-- {message} done")
    except Exception:
        log(f"<-- {message} FAILED")
        raise
