"""
FastAPI app: upload shop master + supplier sheet -> download comparison workbook.
Stateless. Nothing is stored on disk; files live in memory only for the request.
"""
from datetime import datetime
from io import BytesIO

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
import os

from .compare import build_comparison, CompareError

MAX_BYTES = 15 * 1024 * 1024  # 15 MB per file
ALLOWED_EXT = (".xlsx", ".xlsm")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="Perfume Cost & Margin Comparator", version="1.0")

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


@app.post("/api/compare")
async def compare(master: UploadFile = File(...), supplier: UploadFile = File(...)):
    master_bytes = await _read_validated(master, "shop master file")
    supplier_bytes = await _read_validated(supplier, "supplier file")
    try:
        result = build_comparison(master_bytes, supplier_bytes)
    except CompareError as e:
        # User-fixable problem -> 400 with a clear message the frontend shows
        raise HTTPException(400, str(e))
    except Exception:
        raise HTTPException(500, "Something went wrong while building the comparison. "
                                 "Please check both files are valid Excel exports and try again.")

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"Perfume_Cost_Margin_Comparison_{stamp}.xlsx"
    return StreamingResponse(
        result,
        media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


# Serve the single-page frontend at "/"
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
