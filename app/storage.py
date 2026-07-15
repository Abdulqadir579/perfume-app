"""
Server-side storage for the shop's inventory sheet.

Deliberately simple: one inventory file on disk plus a small JSON metadata file.
No database, no login — matches current scope (single shop owner).

The storage directory is configurable via the DATA_DIR environment variable so
this moves cleanly to a container/cloud volume later.
"""
import json
import os
import shutil
from datetime import datetime, timezone
from io import BytesIO

import pandas as pd

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "..", "data"))
DATA_DIR = os.path.abspath(DATA_DIR)
INVENTORY_PATH = os.path.join(DATA_DIR, "inventory.xlsx")
META_PATH = os.path.join(DATA_DIR, "inventory_meta.json")


def _ensure_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def save_inventory(data: bytes, filename: str, row_count: int) -> dict:
    """Persist the inventory file, replacing any existing one. Returns metadata."""
    _ensure_dir()
    # Write to a temp file first, then move into place, so a failed write
    # never leaves a half-written inventory behind.
    tmp = INVENTORY_PATH + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    shutil.move(tmp, INVENTORY_PATH)

    meta = {
        "filename": filename,
        "uploaded_at": datetime.now(timezone.utc).isoformat(),
        "row_count": row_count,
        "size_bytes": len(data),
    }
    with open(META_PATH, "w") as f:
        json.dump(meta, f)
    return meta


def load_inventory() -> bytes | None:
    """Return the stored inventory bytes, or None if nothing is stored."""
    if not os.path.exists(INVENTORY_PATH):
        return None
    with open(INVENTORY_PATH, "rb") as f:
        return f.read()


def get_meta() -> dict | None:
    """Return metadata about the stored inventory, or None."""
    if not os.path.exists(INVENTORY_PATH) or not os.path.exists(META_PATH):
        return None
    try:
        with open(META_PATH) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def clear_inventory() -> bool:
    """Delete the stored inventory. Returns True if something was removed."""
    removed = False
    for p in (INVENTORY_PATH, META_PATH):
        if os.path.exists(p):
            os.remove(p)
            removed = True
    return removed


def count_rows(data: bytes) -> int:
    """Best-effort row count for display purposes."""
    try:
        return len(pd.read_excel(BytesIO(data)))
    except Exception:
        return 0
