"""
FastAPI app: compare a supplier sheet against the shop's inventory.

The inventory is uploaded once and stored server-side (it changes weekly, so it
can be replaced at any time). Supplier sheets are never stored — they're
processed in memory for the request only.

Protected by a single shared password (see auth.py).
"""
import os
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException, Depends, Request, Form
from fastapi.responses import StreamingResponse, FileResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import auth, storage
from .compare import build_comparison, validate_master, CompareError

MAX_BYTES = 15 * 1024 * 1024  # 15 MB per file
ALLOWED_EXT = (".xlsx", ".xlsm")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="Perfume Cost & Margin Comparator", version="3.0")

STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")

# Cookies are marked Secure unless explicitly disabled for plain-HTTP local dev.
COOKIE_SECURE = os.environ.get("COOKIE_SECURE", "1") != "0"


@app.get("/health")
def health():
    return {"status": "ok"}


# ---------- Auth ----------

@app.get("/login")
def login_page():
    return FileResponse(os.path.join(STATIC_DIR, "login.html"))


@app.post("/api/login")
def login(request: Request, password: str = Form(...)):
    if not auth.auth_configured():
        raise HTTPException(503, "Login is not configured on the server (APP_PASSWORD is not set).")
    if not auth.check_password(password):
        raise HTTPException(401, "Incorrect password.")
    resp = JSONResponse({"ok": True})
    resp.set_cookie(
        auth.COOKIE_NAME,
        auth.make_session(),
        max_age=auth.SESSION_MAX_AGE,
        httponly=True,
        secure=COOKIE_SECURE,
        samesite="lax",
    )
    return resp


@app.post("/api/logout")
def logout():
    resp = JSONResponse({"ok": True})
    resp.delete_cookie(auth.COOKIE_NAME)
    return resp


# ---------- Helpers ----------

async def _read_validated(upload: UploadFile, label: str) -> bytes:
    if not upload.filename.lower().endswith(ALLOWED_EXT):
        raise HTTPException(400, f"The {label} must be an .xlsx file (got '{upload.filename}').")
    data = await upload.read()
    if len(data) == 0:
        raise HTTPException(400, f"The {label} file is empty.")
    if len(data) > MAX_BYTES:
        raise HTTPException(400, f"The {label} file is too large (max 15 MB).")
    return data


def _stream_result(result, prefix="Perfume_Cost_Margin_Comparison"):
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"{prefix}_{stamp}.xlsx"
    return StreamingResponse(
        result,
        media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


# ---------- Inventory management (protected) ----------

@app.get("/api/inventory", dependencies=[Depends(auth.require_auth)])
def inventory_status():
    """What inventory is currently stored (if any)?"""
    meta = storage.get_meta()
    if not meta:
        return {"stored": False}
    return {"stored": True, **meta}


@app.post("/api/inventory", dependencies=[Depends(auth.require_auth)])
async def upload_inventory(inventory: UploadFile = File(...)):
    """Upload or replace the stored inventory sheet."""
    data = await _read_validated(inventory, "inventory file")

    # Validate BEFORE storing, so a bad file can never replace a good one.
    try:
        validate_master(data)
    except CompareError as e:
        raise HTTPException(400, str(e))

    rows = storage.count_rows(data)
    meta = storage.save_inventory(data, inventory.filename, rows)
    return {"stored": True, **meta}


@app.delete("/api/inventory", dependencies=[Depends(auth.require_auth)])
def delete_inventory():
    removed = storage.clear_inventory()
    return {"stored": False, "removed": removed}


# ---------- Comparison (protected) ----------

@app.post("/api/compare-supplier", dependencies=[Depends(auth.require_auth)])
async def compare_supplier(supplier: UploadFile = File(...)):
    """Compare a supplier sheet against the STORED inventory."""
    master_bytes = storage.load_inventory()
    if master_bytes is None:
        raise HTTPException(400, "No inventory is stored yet. Upload your inventory sheet first.")
    supplier_bytes = await _read_validated(supplier, "supplier file")
    try:
        result = build_comparison(master_bytes, supplier_bytes)
    except CompareError as e:
        raise HTTPException(400, str(e))
    except Exception:
        raise HTTPException(500, "Something went wrong while building the comparison. "
                                 "Please check the file is a valid Excel export and try again.")
    return _stream_result(result)


@app.post("/api/compare", dependencies=[Depends(auth.require_auth)])
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


# ---------- Frontend (protected) ----------

@app.get("/")
def index(request: Request):
    """The app itself. Redirect to /login when not signed in."""
    if not auth.valid_session(request.cookies.get(auth.COOKIE_NAME)):
        return RedirectResponse("/login", status_code=302)
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


# Static assets only (login.html is served explicitly above; index.html is gated by "/").
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
