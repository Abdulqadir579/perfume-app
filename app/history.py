"""
Price history.

Two sources:
  1. Supplier sheet archives (PT-001, PT-002 ...) — every sheet ever uploaded is
     kept, so supplier cost history is fully available, including retroactively.
  2. Inventory snapshots — recorded from the moment snapshotting was added, so
     inventory history builds up going forward and is empty for earlier uploads.

Product names/descriptions live in the inventory, not in supplier sheets (those
carry only Barcode + Current Cost), so search resolves names from the current
inventory and then looks up that barcode's history everywhere.
"""
from datetime import datetime, timedelta, timezone
from io import BytesIO

import pandas as pd

from . import storage, suppliers as sup

PERIODS = {
    "yesterday": 1,
    "last_week": 7,
    "last_month": 31,
    "last_year": 365,
    "all": None,
}


def _bc(value) -> str:
    """Normalise a barcode to a plain string.

    Excel often reads barcodes as floats, so 890000000001 arrives as
    890000000001.0 and string comparisons silently fail. Strip the trailing
    '.0' so every source keys the same way.
    """
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _cutoff(period: str):
    days = PERIODS.get(period, None)
    if days is None:
        return None
    return datetime.now(timezone.utc) - timedelta(days=days)


def _within(iso_ts: str, cutoff) -> bool:
    if cutoff is None:
        return True
    try:
        ts = datetime.fromisoformat(iso_ts)
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=timezone.utc)
        return ts >= cutoff
    except (ValueError, TypeError):
        return True


def _inventory_catalog() -> pd.DataFrame:
    """Barcode -> descriptive fields, from the current stored inventory."""
    files = storage.load_inventory_files()
    frames = []
    for _fn, data in files:
        try:
            df = pd.read_excel(BytesIO(data))
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        lower = {c.lower(): c for c in df.columns}
        if "barcode" not in lower:
            continue
        ren = {lower["barcode"]: "Barcode"}
        for want in ["perfume name", "description", "brand", "size (ml)",
                     "category", "last cost", "market/selling price"]:
            if want in lower:
                ren[lower[want]] = want.title() if want != "size (ml)" else "Size (ml)"
        df = df.rename(columns=ren)
        keep = [c for c in ["Barcode", "Perfume Name", "Description", "Brand",
                            "Size (ml)", "Category", "Last Cost", "Market/Selling Price"]
                if c in df.columns]
        frames.append(df[keep])
    if not frames:
        return pd.DataFrame(columns=["Barcode"])
    cat = pd.concat(frames, ignore_index=True)
    cat["Barcode"] = cat["Barcode"].map(_bc)
    return cat.drop_duplicates(subset="Barcode", keep="last")


def dashboard() -> dict:
    """Headline figures for the landing view.

    Everything here is derived from data already stored — no extra uploads
    needed — so the dashboard is useful the moment a supplier sheet exists.
    """
    cat = _inventory_catalog()
    product_count = 0 if cat.empty else len(cat)

    suppliers = sup.list_suppliers()
    current = sup.current_sheets()

    # Best cost per product across current sheets
    best = {}          # barcode -> (cost, code)
    per_code_wins = {}
    coverage = {}      # code -> how many products they carry
    for code, _name, _ref, data in current:
        try:
            df = pd.read_excel(BytesIO(data))
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        lower = {c.lower(): c for c in df.columns}
        if "barcode" not in lower or "current cost" not in lower:
            continue
        bc_col, cost_col = lower["barcode"], lower["current cost"]
        sub = df[[bc_col, cost_col]].dropna()
        coverage[code] = len(sub)
        for _, r in sub.iterrows():
            bc = _bc(r[bc_col])
            try:
                cost = float(r[cost_col])
            except (TypeError, ValueError):
                continue
            if bc not in best or cost < best[bc][0]:
                best[bc] = (cost, code)

    for _bcode, (_cost, code) in best.items():
        per_code_wins[code] = per_code_wins.get(code, 0) + 1

    # Margin health against best cost
    at_loss = low_margin = healthy = 0
    import_count = export_count = 0
    total_margin_pct = 0.0
    counted = 0
    if not cat.empty and "Market/Selling Price" in cat.columns:
        for _, r in cat.iterrows():
            bc = _bc(r["Barcode"])
            if bc not in best:
                continue
            try:
                sell = float(r["Market/Selling Price"])
            except (TypeError, ValueError):
                continue
            if not sell:
                continue
            cost = best[bc][0]
            m = (sell - cost) / sell
            counted += 1
            total_margin_pct += m
            if m < 0:
                at_loss += 1
            elif m < 0.15:
                low_margin += 1
            else:
                healthy += 1
            if sell > cost:
                import_count += 1
            elif sell < cost:
                export_count += 1

    supplier_rows = []
    for s in suppliers:
        code = s["code"]
        supplier_rows.append({
            **s,
            "wins": per_code_wins.get(code, 0),
            "covers": coverage.get(code, 0),
        })
    supplier_rows.sort(key=lambda x: x["wins"], reverse=True)

    return {
        "products": product_count,
        "inventory_files": len(storage.list_inventory_files()),
        "supplier_count": len(suppliers),
        "sheets_total": len(sup.list_sheets()),
        "priced_products": len(best),
        "at_loss": at_loss,
        "low_margin": low_margin,
        "healthy": healthy,
        "avg_margin_pct": round(total_margin_pct / counted * 100, 1) if counted else None,
        "good_for_import": import_count,
        "good_for_export": export_count,
        "suppliers": supplier_rows,
    }


def _current_costs_for(barcodes: set) -> dict:
    """barcode -> {"best_cost": float, "best_code": str, "per_supplier": {code: cost}}

    Reads each supplier's CURRENT sheet only, and only for the barcodes asked
    for, so this stays cheap even with a large catalog.
    """
    if not barcodes:
        return {}
    out = {}
    for code, _name, _ref, data in sup.current_sheets():
        try:
            df = pd.read_excel(BytesIO(data))
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        lower = {c.lower(): c for c in df.columns}
        if "barcode" not in lower or "current cost" not in lower:
            continue
        bc_col, cost_col = lower["barcode"], lower["current cost"]
        sub = df[[bc_col, cost_col]].dropna()
        for _, r in sub.iterrows():
            bc = _bc(r[bc_col])
            if bc not in barcodes:
                continue
            try:
                cost = float(r[cost_col])
            except (TypeError, ValueError):
                continue
            rec = out.setdefault(bc, {"per_supplier": {}})
            rec["per_supplier"][code] = round(cost, 2)

    for bc, rec in out.items():
        per = rec["per_supplier"]
        if per:
            best_code = min(per, key=per.get)
            rec["best_cost"] = per[best_code]
            rec["best_code"] = best_code
        else:
            rec["best_cost"] = None
            rec["best_code"] = None
    return out


def search_products(query: str, limit: int = 25) -> list:
    """Find products by Perfume Name, Description or Barcode.

    Results include the shop's Last Cost and Selling Price, plus the current
    Best Cost across suppliers and the margin that implies.
    """
    q = (query or "").strip().lower()
    if not q:
        return []
    cat = _inventory_catalog()
    if cat.empty:
        return []

    name_col = "Perfume Name" if "Perfume Name" in cat.columns else None
    desc_col = "Description" if "Description" in cat.columns else None

    mask = pd.Series(False, index=cat.index)
    if name_col:
        mask |= cat[name_col].astype(str).str.lower().str.contains(q, na=False, regex=False)
    if desc_col:
        mask |= cat[desc_col].astype(str).str.lower().str.contains(q, na=False, regex=False)
    mask |= cat["Barcode"].str.contains(q, na=False, regex=False)

    hits = cat[mask].head(limit)
    barcodes = {_bc(b) for b in hits["Barcode"]}
    costs = _current_costs_for(barcodes)

    def _num(row, col):
        if col not in cat.columns:
            return None
        v = row.get(col)
        try:
            return round(float(v), 2) if pd.notna(v) else None
        except (TypeError, ValueError):
            return None

    out = []
    for _, r in hits.iterrows():
        bc = _bc(r["Barcode"])
        c = costs.get(bc, {})
        best = c.get("best_cost")
        sell = _num(r, "Market/Selling Price")
        margin = round(sell - best, 2) if (best is not None and sell is not None) else None
        margin_pct = round((sell - best) / sell * 100, 1) if (
            best is not None and sell not in (None, 0)
        ) else None
        out.append({
            "barcode": bc,
            "name": str(r.get(name_col, "")) if name_col else "",
            "description": str(r.get(desc_col, "")) if desc_col else "",
            "brand": str(r.get("Brand", "")) if "Brand" in cat.columns else "",
            "size": str(r.get("Size (ml)", "")) if "Size (ml)" in cat.columns else "",
            "last_cost": _num(r, "Last Cost"),
            "selling_price": sell,
            "best_cost": best,
            "best_code": c.get("best_code"),
            "per_supplier": c.get("per_supplier", {}),
            "margin": margin,
            "margin_pct": margin_pct,
        })
    return out


def product_history(barcode: str, period: str = "all") -> dict:
    """Full price timeline for one product."""
    barcode = _bc(barcode)
    cutoff = _cutoff(period)

    cat = _inventory_catalog()
    info = {}
    if not cat.empty:
        row = cat[cat["Barcode"] == barcode]
        if not row.empty:
            r = row.iloc[0]
            info = {
                "barcode": barcode,
                "name": str(r.get("Perfume Name", "")),
                "description": str(r.get("Description", "")),
                "brand": str(r.get("Brand", "")),
                "size": str(r.get("Size (ml)", "")),
            }
    if not info:
        info = {"barcode": barcode, "name": "", "description": "", "brand": "", "size": ""}

    # Current figures: shop's own costs plus the best supplier price right now
    if not cat.empty:
        row = cat[cat["Barcode"] == barcode]
        if not row.empty:
            r = row.iloc[0]
            for key, col in [("last_cost", "Last Cost"), ("selling_price", "Market/Selling Price")]:
                if col in cat.columns:
                    try:
                        v = r.get(col)
                        info[key] = round(float(v), 2) if pd.notna(v) else None
                    except (TypeError, ValueError):
                        info[key] = None
    costs = _current_costs_for({barcode}).get(barcode, {})
    info["best_cost"] = costs.get("best_cost")
    info["best_code"] = costs.get("best_code")
    info["per_supplier"] = costs.get("per_supplier", {})
    sell = info.get("selling_price")
    best = info.get("best_cost")
    info["margin"] = round(sell - best, 2) if (sell is not None and best is not None) else None
    info["margin_pct"] = round((sell - best) / sell * 100, 1) if (
        best is not None and sell not in (None, 0)
    ) else None

    # --- supplier cost history (from every archived sheet) ---
    supplier_points = []
    for sheet in sup.list_sheets():          # newest first
        if not _within(sheet["uploaded_at"], cutoff):
            continue
        got = sup.get_sheet_bytes(sheet["ref"])
        if not got:
            continue
        _fn, data = got
        try:
            df = pd.read_excel(BytesIO(data))
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        lower = {c.lower(): c for c in df.columns}
        if "barcode" not in lower or "current cost" not in lower:
            continue
        bc_col, cost_col = lower["barcode"], lower["current cost"]
        match = df[df[bc_col].map(_bc) == barcode]
        if match.empty:
            continue
        val = match.iloc[-1][cost_col]
        if pd.isna(val):
            continue
        supplier_points.append({
            "ref": sheet["ref"],
            "code": sheet["code"],
            "supplier": sup.supplier_name(sheet["code"]),
            "date": sheet["uploaded_at"],
            "cost": round(float(val), 2),
        })

    supplier_points.sort(key=lambda p: p["date"])

    # --- inventory history (from snapshots; empty before snapshotting existed) ---
    inventory_points = []
    for snap in storage.list_inventory_snapshots():
        if not _within(snap.get("uploaded_at", ""), cutoff):
            continue
        for rec in snap.get("rows", []):
            if _bc(rec.get("barcode")) == barcode:
                pt = {"date": snap["uploaded_at"], "filename": snap.get("filename", "")}
                if "last_cost" in rec:
                    pt["last_cost"] = round(rec["last_cost"], 2)
                if "selling_price" in rec:
                    pt["selling_price"] = round(rec["selling_price"], 2)
                inventory_points.append(pt)
                break

    # Per-supplier change summary
    by_supplier = {}
    for p in supplier_points:
        by_supplier.setdefault(p["code"], []).append(p)
    summary = []
    for code, pts in sorted(by_supplier.items()):
        first, last = pts[0], pts[-1]
        change = last["cost"] - first["cost"]
        pct = (change / first["cost"] * 100) if first["cost"] else None
        summary.append({
            "code": code,
            "supplier": pts[0]["supplier"],
            "points": len(pts),
            "first_cost": first["cost"],
            "latest_cost": last["cost"],
            "change": round(change, 2),
            "change_pct": round(pct, 1) if pct is not None else None,
        })

    return {
        "product": info,
        "supplier_history": supplier_points,
        "inventory_history": inventory_points,
        "summary": summary,
        "period": period,
    }


def changes_table(period: str = "last_month", limit: int = 300) -> dict:
    """Every product whose supplier cost changed within the period."""
    cutoff = _cutoff(period)
    cat = _inventory_catalog()
    names = {}
    if not cat.empty:
        for _, r in cat.iterrows():
            names[_bc(r["Barcode"])] = {
                "name": str(r.get("Perfume Name", "")),
                "brand": str(r.get("Brand", "")),
            }

    # barcode -> code -> [(date, cost)]
    series = {}
    for sheet in sup.list_sheets():
        if not _within(sheet["uploaded_at"], cutoff):
            continue
        got = sup.get_sheet_bytes(sheet["ref"])
        if not got:
            continue
        _fn, data = got
        try:
            df = pd.read_excel(BytesIO(data))
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        lower = {c.lower(): c for c in df.columns}
        if "barcode" not in lower or "current cost" not in lower:
            continue
        bc_col, cost_col = lower["barcode"], lower["current cost"]
        sub = df[[bc_col, cost_col]].dropna()
        for _, r in sub.iterrows():
            bc = _bc(r[bc_col])
            series.setdefault(bc, {}).setdefault(sheet["code"], []).append(
                (sheet["uploaded_at"], float(r[cost_col]))
            )

    rows = []
    for bc, per_code in series.items():
        for code, pts in per_code.items():
            if len(pts) < 2:
                continue
            pts.sort(key=lambda t: t[0])
            first_cost, last_cost = pts[0][1], pts[-1][1]
            if abs(last_cost - first_cost) < 0.005:
                continue
            meta = names.get(bc, {})
            rows.append({
                "barcode": bc,
                "name": meta.get("name", ""),
                "brand": meta.get("brand", ""),
                "code": code,
                "supplier": sup.supplier_name(code),
                "from_cost": round(first_cost, 2),
                "to_cost": round(last_cost, 2),
                "change": round(last_cost - first_cost, 2),
                "change_pct": round((last_cost - first_cost) / first_cost * 100, 1) if first_cost else None,
                "points": len(pts),
            })

    rows.sort(key=lambda r: abs(r["change_pct"] or 0), reverse=True)
    return {"period": period, "count": len(rows), "rows": rows[:limit]}
