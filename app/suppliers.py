"""
Supplier registry and sheet archive.

Each supplier has a short CODE (e.g. PT for Perfume Trading) and a name.
Every uploaded sheet is archived under a reference like "PT-001-2026-07-23":
  <CODE>-<sequence for that supplier>-<upload date>

The newest sheet for a supplier automatically becomes their "current" one —
that's what comparisons use. Older sheets are kept and can be re-downloaded.

Storage layout under DATA_DIR:
  suppliers.json              registry: code -> name
  supplier_sheets.json        manifest of every archived sheet
  supplier_sheets/<uuid>.xlsx the files themselves
"""
import json
import os
import re
import shutil
import uuid
from datetime import datetime, timezone

from .storage import DATA_DIR  # reuse the same configurable data directory

REGISTRY_PATH = os.path.join(DATA_DIR, "suppliers.json")
SHEETS_MANIFEST = os.path.join(DATA_DIR, "supplier_sheets.json")
SHEETS_DIR = os.path.join(DATA_DIR, "supplier_sheets")

CODE_RE = re.compile(r"^[A-Z][A-Z0-9]{1,5}$")


class SupplierError(Exception):
    """User-fixable problem (bad code, duplicate, unknown supplier)."""


def _ensure_dir():
    os.makedirs(SHEETS_DIR, exist_ok=True)


def _read_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return default


def _write_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(data, f)
    shutil.move(tmp, path)


# ---------- registry ----------

def list_suppliers() -> list:
    """[{code, name, sheet_count, current_ref, current_uploaded_at}, ...]"""
    reg = _read_json(REGISTRY_PATH, {})
    sheets = _read_json(SHEETS_MANIFEST, [])
    out = []
    for code, name in sorted(reg.items()):
        mine = [s for s in sheets if s["code"] == code]
        current = mine[-1] if mine else None
        out.append({
            "code": code,
            "name": name,
            "sheet_count": len(mine),
            "current_ref": current["ref"] if current else None,
            "current_uploaded_at": current["uploaded_at"] if current else None,
        })
    return out


def add_supplier(code: str, name: str) -> dict:
    code = (code or "").strip().upper()
    name = (name or "").strip()
    if not CODE_RE.match(code):
        raise SupplierError(
            "Supplier code must be 2-6 characters, letters/numbers, starting with a letter "
            "(e.g. PT, FT, MT)."
        )
    if not name:
        raise SupplierError("Please give the supplier a name.")
    reg = _read_json(REGISTRY_PATH, {})
    if code in reg:
        raise SupplierError(f"Supplier code '{code}' already exists ({reg[code]}).")
    reg[code] = name
    _write_json(REGISTRY_PATH, reg)
    return {"code": code, "name": name}


def rename_supplier(code: str, name: str) -> dict:
    code = (code or "").strip().upper()
    name = (name or "").strip()
    if not name:
        raise SupplierError("Please give the supplier a name.")
    reg = _read_json(REGISTRY_PATH, {})
    if code not in reg:
        raise SupplierError(f"No supplier with code '{code}'.")
    reg[code] = name
    _write_json(REGISTRY_PATH, reg)
    return {"code": code, "name": name}


def delete_supplier(code: str) -> bool:
    """Remove a supplier and all their archived sheets."""
    code = (code or "").strip().upper()
    reg = _read_json(REGISTRY_PATH, {})
    if code not in reg:
        return False
    del reg[code]
    _write_json(REGISTRY_PATH, reg)

    sheets = _read_json(SHEETS_MANIFEST, [])
    kept = []
    for s in sheets:
        if s["code"] == code:
            p = os.path.join(SHEETS_DIR, s["stored_name"])
            if os.path.exists(p):
                os.remove(p)
        else:
            kept.append(s)
    _write_json(SHEETS_MANIFEST, kept)
    return True


def supplier_name(code: str) -> str:
    return _read_json(REGISTRY_PATH, {}).get(code, code)


# ---------- sheets ----------

def add_sheet(code: str, data: bytes, filename: str, row_count: int) -> dict:
    """Archive a supplier sheet. Becomes that supplier's current sheet."""
    code = (code or "").strip().upper()
    reg = _read_json(REGISTRY_PATH, {})
    if code not in reg:
        raise SupplierError(f"No supplier with code '{code}'. Add the supplier first.")

    _ensure_dir()
    sheets = _read_json(SHEETS_MANIFEST, [])
    seq = sum(1 for s in sheets if s["code"] == code) + 1
    now = datetime.now(timezone.utc)
    ref = f"{code}-{seq:03d}-{now.strftime('%Y-%m-%d')}"

    stored_name = uuid.uuid4().hex + ".xlsx"
    path = os.path.join(SHEETS_DIR, stored_name)
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(data)
    shutil.move(tmp, path)

    entry = {
        "ref": ref,
        "code": code,
        "seq": seq,
        "stored_name": stored_name,
        "filename": filename,
        "uploaded_at": now.isoformat(),
        "row_count": row_count,
        "size_bytes": len(data),
    }
    sheets.append(entry)
    _write_json(SHEETS_MANIFEST, sheets)
    return entry


def list_sheets(code: str | None = None) -> list:
    """Archived sheets, newest first. Optionally filtered to one supplier."""
    sheets = _read_json(SHEETS_MANIFEST, [])
    if code:
        code = code.strip().upper()
        sheets = [s for s in sheets if s["code"] == code]
    return sorted(sheets, key=lambda s: s["uploaded_at"], reverse=True)


def current_sheets() -> list:
    """The latest sheet for each supplier that has one. [(code, name, ref, bytes)]"""
    sheets = _read_json(SHEETS_MANIFEST, [])
    reg = _read_json(REGISTRY_PATH, {})
    latest = {}
    for s in sheets:  # manifest is append-order, so last wins
        latest[s["code"]] = s
    out = []
    for code in sorted(latest):
        s = latest[code]
        p = os.path.join(SHEETS_DIR, s["stored_name"])
        if os.path.exists(p):
            with open(p, "rb") as f:
                out.append((code, reg.get(code, code), s["ref"], f.read()))
    return out


def get_sheet_bytes(ref: str):
    """Return (filename, bytes) for an archived sheet, or None."""
    for s in _read_json(SHEETS_MANIFEST, []):
        if s["ref"] == ref:
            p = os.path.join(SHEETS_DIR, s["stored_name"])
            if os.path.exists(p):
                with open(p, "rb") as f:
                    return s["filename"], f.read()
    return None


def delete_sheet(ref: str) -> bool:
    sheets = _read_json(SHEETS_MANIFEST, [])
    kept, removed = [], False
    for s in sheets:
        if s["ref"] == ref:
            p = os.path.join(SHEETS_DIR, s["stored_name"])
            if os.path.exists(p):
                os.remove(p)
            removed = True
        else:
            kept.append(s)
    if removed:
        _write_json(SHEETS_MANIFEST, kept)
    return removed
