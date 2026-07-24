"""
Server-side storage for the shop's inventory.

The inventory can be split across MULTIPLE files (e.g. one per brand). Each file
is stored individually under DATA_DIR/inventory/, and a manifest tracks them.
At comparison time they're merged into one catalog (see compare.merge_inventories).

No database — a directory of files plus a JSON manifest. Moves cleanly to a
cloud volume via the DATA_DIR env var.
"""
import json
import os
import re
import shutil
import uuid
from datetime import datetime, timezone
from io import BytesIO

import pandas as pd

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(__file__), "..", "data"))
DATA_DIR = os.path.abspath(DATA_DIR)
INV_DIR = os.path.join(DATA_DIR, "inventory")
MANIFEST_PATH = os.path.join(DATA_DIR, "inventory_manifest.json")


def _ensure_dir():
    os.makedirs(INV_DIR, exist_ok=True)


def _read_manifest() -> list:
    if not os.path.exists(MANIFEST_PATH):
        return []
    try:
        with open(MANIFEST_PATH) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def _write_manifest(entries: list):
    tmp = MANIFEST_PATH + ".tmp"
    with open(tmp, "w") as f:
        json.dump(entries, f)
    shutil.move(tmp, MANIFEST_PATH)


def add_inventory_file(data: bytes, filename: str, row_count: int) -> dict:
    """Store one inventory file (added to any already present). Returns its entry."""
    _ensure_dir()
    file_id = uuid.uuid4().hex
    stored_name = file_id + ".xlsx"
    path = os.path.join(INV_DIR, stored_name)

    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    shutil.move(tmp, path)

    entry = {
        "id": file_id,
        "stored_name": stored_name,
        "filename": filename,
        "uploaded_at": datetime.now(timezone.utc).isoformat(),
        "row_count": row_count,
        "size_bytes": len(data),
    }
    entries = _read_manifest()
    entries.append(entry)
    _write_manifest(entries)

    # Snapshot this upload so inventory price history accrues from now on.
    # (Past uploads can't be reconstructed — nothing was kept before this.)
    try:
        _snapshot_inventory(data, filename, entry["uploaded_at"])
    except Exception:
        pass  # history is a nice-to-have; never fail an upload because of it

    return entry


SNAPSHOT_DIR = os.path.join(DATA_DIR, "inventory_snapshots")


def _snapshot_inventory(data: bytes, filename: str, uploaded_at: str):
    """Record Barcode -> (Last Cost, Selling Price) for this upload."""
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    df = pd.read_excel(BytesIO(data))
    df.columns = [str(c).strip() for c in df.columns]
    lower = {c.lower(): c for c in df.columns}
    bc = lower.get("barcode")
    lc = lower.get("last cost")
    sp = lower.get("market/selling price")
    if not bc:
        return
    keep = {"barcode": bc}
    if lc:
        keep["last_cost"] = lc
    if sp:
        keep["selling_price"] = sp

    rows = []
    for _, r in df.iterrows():
        _b = str(r[bc]).strip()
        if _b.endswith(".0"):
            _b = _b[:-2]
        rec = {"barcode": _b}
        if lc and pd.notna(r[lc]):
            rec["last_cost"] = float(r[lc])
        if sp and pd.notna(r[sp]):
            rec["selling_price"] = float(r[sp])
        rows.append(rec)

    snap = {
        "uploaded_at": uploaded_at,
        "filename": filename,
        "rows": rows,
    }
    name = uuid.uuid4().hex + ".json"
    with open(os.path.join(SNAPSHOT_DIR, name), "w") as f:
        json.dump(snap, f)


def list_inventory_snapshots() -> list:
    """All inventory snapshots, oldest first."""
    if not os.path.isdir(SNAPSHOT_DIR):
        return []
    out = []
    for fn in os.listdir(SNAPSHOT_DIR):
        if not fn.endswith(".json"):
            continue
        try:
            with open(os.path.join(SNAPSHOT_DIR, fn)) as f:
                out.append(json.load(f))
        except (json.JSONDecodeError, OSError):
            continue
    return sorted(out, key=lambda s: s.get("uploaded_at", ""))


def list_inventory_files() -> list:
    """Metadata for every stored inventory file, oldest first."""
    return _read_manifest()


def load_inventory_files() -> list:
    """Return [(filename, bytes), ...] for all stored files, in upload order."""
    out = []
    for entry in _read_manifest():
        path = os.path.join(INV_DIR, entry["stored_name"])
        if os.path.exists(path):
            with open(path, "rb") as f:
                out.append((entry["filename"], f.read()))
    return out


def remove_inventory_file(file_id: str) -> bool:
    """Remove a single inventory file by id."""
    entries = _read_manifest()
    kept, removed = [], False
    for e in entries:
        if e["id"] == file_id:
            path = os.path.join(INV_DIR, e["stored_name"])
            if os.path.exists(path):
                os.remove(path)
            removed = True
        else:
            kept.append(e)
    if removed:
        _write_manifest(kept)
    return removed


def clear_inventory() -> bool:
    """Remove ALL inventory files."""
    removed = False
    if os.path.isdir(INV_DIR):
        shutil.rmtree(INV_DIR)
        removed = True
    if os.path.exists(MANIFEST_PATH):
        os.remove(MANIFEST_PATH)
        removed = True
    return removed


def has_inventory() -> bool:
    return len(_read_manifest()) > 0


def count_rows(data: bytes) -> int:
    """Best-effort row count for display purposes."""
    try:
        return len(pd.read_excel(BytesIO(data)))
    except Exception:
        return 0
