"""
Core comparison engine.
Matches a shop master sheet against a supplier sheet by Barcode and produces
a formatted comparison workbook with cost-difference, margin, and flag columns.

This is the SAME logic validated in the Excel demo. No simulation here —
it reads the supplier's real Current Cost.
"""
from io import BytesIO
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule

FONT = "Arial"
LOW_MARGIN_THRESHOLD = 0.15  # below this = "Low Margin"
DELTA = "\u0394"  # Greek delta, kept out of f-string expressions for py3.9 compatibility

# Columns we require to exist (case-insensitive, trimmed)
MASTER_REQUIRED = ["Barcode", "Last Cost", "Market/Selling Price"]
SUPPLIER_REQUIRED = ["Barcode", "Current Cost"]

OUTPUT_COLS = [
    "Barcode", "SKU", "Brand", "Perfume Name", "Description", "Size (ml)",
    "Gender", "Category", "Stock Qty", "Country of Origin",
    "Last Cost", "Supplier Current Cost", "Market/Selling Price", "Match Status",
]


class CompareError(Exception):
    """Raised for user-fixable problems (missing columns, unreadable file)."""


def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _check_columns(df: pd.DataFrame, required, label):
    have = {c.lower(): c for c in df.columns}
    missing = [c for c in required if c.lower() not in have]
    if missing:
        raise CompareError(
            f"The {label} file is missing required column(s): {', '.join(missing)}. "
            f"Found columns: {', '.join(df.columns)}"
        )
    # return actual-cased names
    return {c: have[c.lower()] for c in required}


def build_comparison(master_bytes: bytes, supplier_bytes: bytes) -> BytesIO:
    """Take two uploaded xlsx files (as bytes), return a formatted xlsx in a BytesIO."""
    try:
        master = _norm_cols(pd.read_excel(BytesIO(master_bytes)))
    except Exception as e:
        raise CompareError(f"Could not read the shop master file as Excel. ({e})")
    try:
        supplier = _norm_cols(pd.read_excel(BytesIO(supplier_bytes)))
    except Exception as e:
        raise CompareError(f"Could not read the supplier file as Excel.
