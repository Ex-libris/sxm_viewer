"""Persistent header cache backed by SQLite (Qt-free).

Parsing an image header costs ~4-6 ms per file, so caching parsed
(header, fds) pairs across sessions saves about a second per 200-scan folder.
The cache used to be one JSON file holding every header ever parsed: read in
full at every startup and rewritten in full on every save, so it grew without
bound (10k entries / 50 MB / ~310 ms startup parse on a real install) and
stalled the GUI thread on each write.

Here each header is one row, so a folder load reads only that folder's rows
and a save writes only new/changed rows - cost scales with the folder, not
with every folder ever opened. Rows are validated against the file's mtime on
lookup, exactly like the JSON cache was.

Every operation swallows errors and degrades to "cache miss": the cache must
never break folder loading. The database is purely derived data, so a corrupt
file is simply deleted and rebuilt.
"""
from __future__ import annotations

import json
import os
import sqlite3
import threading
from pathlib import Path

_SQL_CHUNK = 500  # stay under SQLite's host-parameter limit on older builds

_SCHEMA = """
CREATE TABLE IF NOT EXISTS headers (
    path   TEXT PRIMARY KEY,
    folder TEXT NOT NULL,
    mtime  REAL NOT NULL,
    header TEXT NOT NULL,
    fds    TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS headers_folder ON headers(folder);
"""


def db_path_for(json_path):
    """Database location derived from the legacy JSON path, so anything that
    redirects config_io.HEADER_CACHE_PATH (smoke tests) redirects both."""
    return Path(json_path).with_suffix(".sqlite3")


def _dumps(value):
    return json.dumps(value, separators=(",", ":"))


class HeaderCacheStore:
    """Path -> (mtime, header, fds) store. Thread-safe via one lock; the
    connection is shared across threads (check_same_thread=False)."""

    def __init__(self, db_path, version, legacy_json_path=None):
        self._db_path = Path(db_path)
        self._version = int(version)
        self._legacy_json_path = Path(legacy_json_path) if legacy_json_path else None
        self._lock = threading.Lock()
        self._conn = None
        self.migrated_entries = 0
        try:
            self._conn = self._connect()
        except Exception:
            # Corrupt/unreadable database: it's only a cache, rebuild it.
            try:
                self._db_path.unlink()
                self._conn = self._connect()
            except Exception:
                self._conn = None

    def _connect(self):
        conn = sqlite3.connect(str(self._db_path), timeout=5.0, check_same_thread=False)
        try:
            return self._init_connection(conn)
        except Exception:
            # Release the file handle, or Windows refuses the rebuild's unlink.
            conn.close()
            raise

    def _init_connection(self, conn):
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.executescript(_SCHEMA)
        stored_version = conn.execute("PRAGMA user_version").fetchone()[0]
        if stored_version != self._version:
            fresh = stored_version == 0
            conn.execute("DELETE FROM headers")
            if fresh:
                self._import_legacy_json(conn)
            conn.execute(f"PRAGMA user_version={self._version}")
            conn.commit()
        return conn

    def _import_legacy_json(self, conn):
        """One-time import of the old JSON cache so the first run after the
        switch is still warm. The JSON file is left untouched."""
        path = self._legacy_json_path
        if path is None or not path.exists():
            return
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return
        if not isinstance(data, dict) or data.get("_version") != self._version:
            return
        rows = []
        for key, entry in (data.get("entries") or {}).items():
            if not isinstance(entry, dict):
                continue
            header, fds = entry.get("header"), entry.get("fds")
            if header is None or fds is None:
                continue
            rows.append((key, str(Path(key).parent), float(entry.get("mtime", 0.0)),
                         _dumps(header), _dumps(fds)))
        conn.executemany("INSERT OR REPLACE INTO headers VALUES (?,?,?,?,?)", rows)
        self.migrated_entries = len(rows)

    def __len__(self):
        if self._conn is None:
            return 0
        try:
            with self._lock:
                return self._conn.execute("SELECT COUNT(*) FROM headers").fetchone()[0]
        except Exception:
            return 0

    def lookup_many(self, paths):
        """Return {path_str: (header, fds)} for the given paths whose cached
        mtime still matches the file on disk. Missing/stale paths are absent."""
        keys = [str(p) for p in paths]
        if self._conn is None or not keys:
            return {}
        rows = []
        try:
            with self._lock:
                for i in range(0, len(keys), _SQL_CHUNK):
                    chunk = keys[i:i + _SQL_CHUNK]
                    marks = ",".join("?" * len(chunk))
                    rows.extend(self._conn.execute(
                        f"SELECT path, mtime, header, fds FROM headers WHERE path IN ({marks})",
                        chunk).fetchall())
        except Exception:
            return {}
        result = {}
        for key, mtime, header, fds in rows:
            try:
                if abs(os.stat(key).st_mtime - mtime) > 1e-6:
                    continue
                result[key] = (json.loads(header), json.loads(fds))
            except Exception:
                continue
        return result

    def store_many(self, items):
        """Insert/replace entries. ``items`` is an iterable of
        (path, mtime, header, fds); one transaction for the whole batch."""
        if self._conn is None:
            return
        rows = []
        for path, mtime, header, fds in items:
            try:
                rows.append((str(path), str(Path(path).parent), float(mtime),
                             _dumps(header), _dumps(fds)))
            except Exception:
                continue
        if not rows:
            return
        try:
            with self._lock, self._conn:
                self._conn.executemany("INSERT OR REPLACE INTO headers VALUES (?,?,?,?,?)", rows)
        except Exception:
            pass

    def prune_folders(self, folders, keep_paths=()):
        """Drop rows in ``folders`` whose file no longer exists. Paths in
        ``keep_paths`` (just seen on disk) skip the stat. Scoped to the folders
        being loaded, so it never touches - or stats - unrelated data (e.g.
        an unplugged drive's folders keep their entries)."""
        if self._conn is None:
            return
        keep = {str(p) for p in keep_paths}
        try:
            with self._lock:
                stale = []
                for folder in {str(f) for f in folders}:
                    for (key,) in self._conn.execute(
                            "SELECT path FROM headers WHERE folder = ?", (folder,)):
                        if key not in keep and not os.path.exists(key):
                            stale.append((key,))
                if stale:
                    with self._conn:
                        self._conn.executemany("DELETE FROM headers WHERE path = ?", stale)
        except Exception:
            pass

    def close(self):
        if self._conn is None:
            return
        try:
            with self._lock:
                self._conn.close()
        except Exception:
            pass
        self._conn = None


__all__ = ["HeaderCacheStore", "db_path_for"]
