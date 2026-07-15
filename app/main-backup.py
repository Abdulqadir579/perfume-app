"""
FastAPI app: compare a supplier sheet against the shop's inventory.

The inventory is uploaded once and stored server-side (it changes weekly, so it
can be replaced at any time). Supplier sheets are never stored — they're
processed in memory for the request only.
"""
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
import os

from .compare import build_comparison, CompareError
from . import storage

MAX_BYTES = 15 * 1024 * 1024  # 15 MB per file
ALLOWED_EXT = (".xlsx", ".xlsm")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="Perfume Cost & Margin Comparator", version="2.0")

STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")


@app.get("/health")
def health():
    return {"status": "ok"}


async def _read_validated(upload: UploadFile, label: str) -> bytes:
    if not upload.filename.lower().endswith(ALLOWED_EXT):
        raise HTTPException(400, f"The {label} must be an .xlsx file (got '{upload.filename}').")
    data = await upload.read()
    if len(data) == 0:
        raise HTTPException(400, f"The {label} file is empty.")
    if len(data) > MAX_BYTES:
        raise HTTPException(400, f"The {label} file is too large (max 15 MB).")
    return data


# ---------- Inventory management ----------

@app.get("/api/inventory")
def inventory_status():
    """What inventory is currently stored (if any)?"""
    meta = storage.get_meta()
    if not meta:
        return {"stored": False}
    return {"stored": True, **meta}


@app.post("/api/inventory")
async def upload_inventory(inventory: UploadFile = File(...)):
    """Upload or replace the stored inventory sheet."""
    data = await _read_validated(inventory, "inventory file")

    # Validate it's actually usable BEFORE storing it, so a bad file can never
    # become the stored inventory.
    try:
        from .compare import validate_master
        validate_master(data)
    except CompareError as e:
        raise HTTPException(400, str(e))

    rows = storage.count_rows(data)
    meta = storage.save_inventory(data, inventory.filename, rows)
    return {"stored": True, **meta}


@app.delete("/api/inventory")
def delete_inventory():
    removed = storage.clear_inventory()
    return {"stored": False, "removed": removed}


# ---------- Comparison ----------

def _stream_result(result, prefix="Perfume_Cost_Margin_Comparison"):
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"{prefix}_{stamp}.xlsx"
    return StreamingResponse(
        result,
        media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.post("/api/compare-supplier")
async def compare_supplier(supplier: UploadFile = File(...)):
    """Compare a supplier sheet against the STORED inventory."""
    master_bytes = storage.load_inventory()
    if master_bytes is None:
        raise HTTPException(
            400, "No inventory is stored yet. Upload your inventory sheet first."
        )
    supplier_bytes = await _read_validated(supplier, "supplier file")
    try:
        result = build_comparison(master_bytes, supplier_bytes)
    except CompareError as e:
        raise HTTPException(400, str(e))
    except Exception:
        raise HTTPException(500, "Something went wrong while building the comparison. "
                                 "Please check the file is a valid Excel export and try again.")
    return _stream_result(result)


@app.post("/api/compare")
async def compare(master: UploadFile = File(...), supplier: UploadFile = File(...)):
    """Original two-file endpoint — kept so existing use keeps working."""
    master_bytes = await _read_validated(master, "shop master file")
    supplier_bytes = await _read_validated(supplier, "supplier file")
    try:
        result = build_comparison(master_bytes, supplier_bytes)
    except CompareError as e:
        raise HTTPException(400, str(e))
    except Exception:
        raise HTTPException(500, "Something went wrong while building the comparison. "
                                 "Please check both files are valid Excel exports and try again.")
    return _stream_result(result)


# Serve the single-page frontend at "/"
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
