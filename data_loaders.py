# -*- coding: utf-8 -*-
"""Load MS004, SO003, MP008, RD004 master, and forecast for balance engine."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DIR = BASE_DIR / "Sample Balance"
STOCK_DIR = BASE_DIR / "Stock"
FORECAST_DIR = BASE_DIR / "Froecast"
OUTPUT_DIR = BASE_DIR / "Output"

SAMPLE_PATH = SAMPLE_DIR / "All Customer review Jun '2026 review 20.06.2026.xlsx"
STOCK_PATH = STOCK_DIR / "MS004-260619.xls"

TARGET_MONTHS = ["2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
MONTH_TO_YM = {"Jun": "2026-06", "Jul": "2026-07", "Aug": "2026-08", "Sep": "2026-09", "Oct": "2026-10"}
FORECAST_KG_COLS = {"Jun": 43, "Jul": 45, "Aug": 47, "Sep": 49, "Oct": 51}


def _num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, str):
        v = v.strip().replace(",", "")
        if v in ("", "-"):
            return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _text(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v).strip())


def discover_file(patterns: list[str], search_dirs: list[Path] | None = None) -> Path | None:
    """Return newest matching file across search directories."""
    search_dirs = search_dirs or [SAMPLE_DIR, FORECAST_DIR, BASE_DIR]
    matches: list[Path] = []
    for directory in search_dirs:
        if not directory.exists():
            continue
        for pattern in patterns:
            matches.extend(directory.glob(pattern))
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def resolve_sample_path() -> Path:
    found = discover_file(["All Customer review*.xlsx", "*Customer review*.xlsx"], [SAMPLE_DIR])
    return found or SAMPLE_PATH


def resolve_stock_path() -> Path:
    found = discover_file(["MS004*.xls", "MS004*.xlsx"], [STOCK_DIR, SAMPLE_DIR])
    return found or STOCK_PATH


def _detect_header_row(raw: pd.DataFrame, keywords: tuple[str, ...]) -> int:
    for i in range(min(20, len(raw))):
        row_text = " ".join(_text(v).lower() for v in raw.iloc[i].tolist())
        if all(k.lower() in row_text for k in keywords[:2]) and any(
            k.lower() in row_text for k in keywords[2:]
        ):
            return i
    return 5


def _parse_so003_raw(raw: pd.DataFrame, source: str) -> pd.DataFrame:
    hdr = _detect_header_row(raw, ("customer", "fg code", "bal kg", "delivery"))
    col_map: dict[str, int] = {}
    for i, h in enumerate(raw.iloc[hdr].tolist()):
        key = _text(h).lower().replace("\n", " ")
        if key == "customer":
            col_map["Customer"] = i
        elif key == "po no.":
            col_map["PO No."] = i
        elif key in ("due date", "new due date"):
            col_map.setdefault("Due Date", i)
        elif key == "fg code":
            col_map["FG Code"] = i
        elif key == "mat spec":
            col_map["Mat Spec"] = i
        elif key == "t":
            col_map["T"] = i
        elif key == "w":
            col_map["W"] = i
        elif "delivery" in key and "kg" in key:
            col_map["Delivery Kg"] = i
        elif key == "bal kg":
            col_map["Bal Kg"] = i
        elif key == "basic code":
            col_map["Basic Code"] = i
        elif key == "delivery place":
            col_map["Delivery Place"] = i

    rows = []
    for i in range(hdr + 1, len(raw)):
        fg = raw.iloc[i, col_map.get("FG Code", 7)] if "FG Code" in col_map else None
        if pd.isna(fg) or not _text(fg):
            continue
        rows.append({
            "Customer": _text(raw.iloc[i, col_map.get("Customer", 0)]),
            "PO No.": _text(raw.iloc[i, col_map.get("PO No.", 2)]),
            "Due Date": raw.iloc[i, col_map.get("Due Date", 5)],
            "FG Code": _text(fg),
            "Delivery Place": _text(raw.iloc[i, col_map.get("Delivery Place", 8)]),
            "Mat Spec": _text(raw.iloc[i, col_map.get("Mat Spec", 9)]),
            "T": _num(raw.iloc[i, col_map.get("T", 10)]),
            "W": _num(raw.iloc[i, col_map.get("W", 11)]),
            "Delivery Kg": _num(raw.iloc[i, col_map.get("Delivery Kg", 33)]),
            "Bal Kg": _num(raw.iloc[i, col_map.get("Bal Kg", 35)]),
            "Basic Code": _text(raw.iloc[i, col_map.get("Basic Code", 28)]),
            "Source": source,
        })
    return pd.DataFrame(rows)


def load_rd004_master(sample_path: Path | None = None) -> pd.DataFrame:
    """Build material master from Balance sheet + Act order."""
    sample_path = sample_path or resolve_sample_path()
    bal = pd.read_excel(sample_path, sheet_name="Balance sheet (Update)", header=None)
    act = pd.read_excel(sample_path, sheet_name="Act order", header=None)

    fg_map: dict[str, set[str]] = {}
    spec_map: dict[str, dict] = {}
    for _, row in act.iloc[4:].iterrows():
        code = row[0]
        if pd.isna(code):
            continue
        code = _text(code)
        fg = _text(row[3])
        if fg:
            fg_map.setdefault(code, set()).add(fg)
        spec_map[code] = {
            "Spec": _text(row[1]),
            "Thickness": _num(row[8]),
            "Width": _num(row[9]),
            "Kind": _text(row[6]),
        }

    records = []
    for i in range(6, len(bal)):
        code = bal.iloc[i, 1]
        if pd.isna(code):
            continue
        code = _text(code)
        info = spec_map.get(code, {})
        t = _num(bal.iloc[i, 3]) or info.get("Thickness", 0)
        w = info.get("Width", 1219)
        fg_codes = fg_map.get(code, set())
        dim_fg = f"{t:g}x{w:g}x1219".replace(".0", "")
        fg_primary = dim_fg if dim_fg else (next(iter(fg_codes)) if fg_codes else code)
        records.append({
            "Material_Code": code,
            "Common_Group": info.get("Spec", code),
            "Spec": info.get("Spec", ""),
            "Thickness": t,
            "Width": w,
            "FG_Code": fg_primary,
            "FG_Codes_All": ",".join(sorted(fg_codes)) if fg_codes else fg_primary,
            "Kind": info.get("Kind", _text(bal.iloc[i, 2])),
            "Main_Customer": _text(bal.iloc[i, 4]),
            "MOQ": 0.0,
        })
    return pd.DataFrame(records)


def load_ms004(stock_path: Path | None = None) -> pd.DataFrame:
    stock_path = stock_path or resolve_stock_path()
    df = pd.read_excel(stock_path)
    owner_col = "OWNER" if "OWNER" in df.columns else next(
        (c for c in df.columns if str(c).upper() == "OWNER"), None
    )
    if owner_col:
        out = df[df[owner_col].astype(str).str.upper() == "PTT"].copy()
    else:
        out = df.copy()
    spec_col = "MAT SPEC" if "MAT SPEC" in out.columns else "Mat Spec"
    out["Spec"] = out[spec_col].map(_text)
    out["Thickness"] = pd.to_numeric(out["T"], errors="coerce")
    out["Width"] = pd.to_numeric(out["W"], errors="coerce")
    p_wt_col = next((c for c in out.columns if str(c).upper().replace("  ", " ") == "P REMAIN WT"), None)
    rem_wt_col = next((c for c in out.columns if str(c).upper() == "REMAIN WT"), None)
    if p_wt_col:
        out["P REMAIN WT"] = pd.to_numeric(out[p_wt_col], errors="coerce").fillna(0)
    if rem_wt_col:
        out["REMAIN WT"] = pd.to_numeric(out[rem_wt_col], errors="coerce").fillna(0)
    wt_kg = out["P REMAIN WT"] if "P REMAIN WT" in out.columns else out.get("REMAIN WT", 0)
    out["Quantity"] = pd.to_numeric(wt_kg, errors="coerce").fillna(0) / 1000.0
    cust_col = next((c for c in out.columns if str(c).upper().replace(" ", "") == "CUSTCODE"), None)
    if cust_col:
        out["CUST CODE"] = out[cust_col].map(_text)
    maker_col = next((c for c in out.columns if str(c).upper().replace(" ", "") == "MAKERCODE"), None)
    if maker_col:
        out["MAKER CODE"] = out[maker_col].map(_text)
    out["Source_File"] = stock_path.name
    return out


def load_so003(sample_path: Path | None = None) -> pd.DataFrame:
    """Prefer SO003-Carrier file in Sample Balance; fallback to workbook SO003 sheet."""
    carrier_path = discover_file(
        ["*SO003*Carrier*", "*SO003-Carrier*", "*SO003 Carrier*"],
        [SAMPLE_DIR],
    )
    if carrier_path:
        if carrier_path.suffix.lower() in (".xlsx", ".xls"):
            xl = pd.ExcelFile(carrier_path)
            sheet = next((s for s in xl.sheet_names if "SO003" in s.upper() or "CARRIER" in s.upper()), xl.sheet_names[0])
            raw = pd.read_excel(carrier_path, sheet_name=sheet, header=None)
            df = _parse_so003_raw(raw, source=carrier_path.name)
            if not df.empty:
                return df

    sample_path = sample_path or resolve_sample_path()
    raw = pd.read_excel(sample_path, sheet_name="SO003", header=None)
    return _parse_so003_raw(raw, source=sample_path.name)


def load_mp008(sample_path: Path | None = None) -> pd.DataFrame:
    """Load open PO from standalone MP008 export or Sample Balance workbook."""
    mp008_path = discover_file(["*MP008*", "MP008*.xlsx", "MP008*.xls"], [SAMPLE_DIR, BASE_DIR])
    if mp008_path and mp008_path.suffix.lower() in (".xlsx", ".xls"):
        xl = pd.ExcelFile(mp008_path)
        sheet = next((s for s in xl.sheet_names if "MP008" in s.upper()), xl.sheet_names[0])
        raw = pd.read_excel(mp008_path, sheet_name=sheet, header=None)
        return _parse_mp008_raw(raw, source=mp008_path.name)

    sample_path = sample_path or resolve_sample_path()
    raw = pd.read_excel(sample_path, sheet_name="MP008 xxxxxxxx", header=None)
    return _parse_mp008_raw(raw, source=sample_path.name)


def _parse_mp008_raw(raw: pd.DataFrame, source: str) -> pd.DataFrame:
    hdr = _detect_header_row(raw, ("po no.", "mat spec", "eta date", "close flag"))
    col_map: dict[str, int] = {}
    for i, h in enumerate(raw.iloc[hdr].tolist()):
        key = _text(h).lower()
        if key == "po no.":
            col_map["Po No."] = i
        elif key == "mat spec":
            col_map["Mat Spec"] = i
        elif key == "t":
            col_map["T"] = i
        elif key == "w":
            col_map["W"] = i
        elif "eta date" in key:
            col_map["ETA"] = i
        elif key == "close flag":
            col_map["Close Flag"] = i
        elif "po balance wt" in key or key == "po balance wt":
            col_map["PO Balance WT"] = i
        elif key == "customer":
            col_map["Customer"] = i

    rows = []
    for i in range(hdr + 1, len(raw)):
        if col_map.get("Close Flag") is not None and raw.iloc[i, col_map["Close Flag"]] is True:
            continue
        eta = raw.iloc[i, col_map.get("ETA", 17)] if "ETA" in col_map else None
        if pd.isna(eta):
            continue
        eta_dt = pd.to_datetime(eta, errors="coerce")
        if pd.isna(eta_dt):
            continue
        bal_wt = _num(raw.iloc[i, col_map.get("PO Balance WT", 24)])
        if bal_wt <= 0:
            continue
        close_flag = raw.iloc[i, col_map["Close Flag"]] if "Close Flag" in col_map else False
        rows.append({
            "Po No.": _text(raw.iloc[i, col_map.get("Po No.", 0)]),
            "Customer": _text(raw.iloc[i, col_map.get("Customer", 4)]),
            "Mat Spec": _text(raw.iloc[i, col_map.get("Mat Spec", 9)]),
            "T": _num(raw.iloc[i, col_map.get("T", 11)]),
            "W": _num(raw.iloc[i, col_map.get("W", 12)]),
            "ETA": eta_dt,
            "ETA_Month": eta_dt.strftime("%Y-%m"),
            "PO Balance WT": bal_wt,
            "Close Flag": close_flag,
            "Status": "Open",
            "Source": source,
        })
    return pd.DataFrame(rows)


def load_sample_balance_forecast(sample_path: Path | None = None) -> pd.DataFrame:
    """Forecast from Sample Balance Forecast sheet (kg/month by Material Code)."""
    sample_path = sample_path or resolve_sample_path()
    fc = pd.read_excel(sample_path, sheet_name="Forecast", header=None)
    records = []
    for r in range(4, len(fc)):
        code = fc.iloc[r, 0]
        if pd.isna(code):
            continue
        rec = {
            "Material_Code": _text(code),
            "FG_Code": _text(fc.iloc[r, 2]),
            "Customer": _text(fc.iloc[r, 3]),
            "Spec": _text(fc.iloc[r, 9]),
        }
        for month, col in FORECAST_KG_COLS.items():
            rec[MONTH_TO_YM[month]] = _num(fc.iloc[r, col]) / 1000.0
        records.append(rec)
    if not records:
        return pd.DataFrame(columns=["Material_Code", *TARGET_MONTHS])
    return pd.DataFrame(records).groupby("Material_Code", as_index=False).sum(numeric_only=True)


def load_penta_schedule_forecast() -> pd.DataFrame:
    """
    Load future order forecast from PENTA Schedule wk 25 (preferred).
    Falls back to Sample Balance Forecast + Froecast folder merge.
    """
    schedule_path = discover_file(
        [
            "*PENTA*Schedule*wk*25*",
            "*PENTA Schedule*wk 25*",
            "*PENTA*Schedule*WK*25*",
            "*Schedule*wk*25*",
        ],
        [SAMPLE_DIR, FORECAST_DIR, BASE_DIR],
    )

    frames: list[pd.DataFrame] = []
    if schedule_path:
        frames.append(_parse_penta_schedule(schedule_path))

    sample_fc = load_sample_balance_forecast()
    if not sample_fc.empty:
        sample_fc = sample_fc.copy()
        sample_fc["Source"] = "Sample Balance Forecast"
        frames.append(sample_fc)

    vendor_fc = load_client_forecast_from_froecast()
    if not vendor_fc.empty:
        vendor_fc = vendor_fc.copy()
        vendor_fc["Source"] = "Froecast"
        frames.append(vendor_fc)

    if not frames:
        return pd.DataFrame(columns=["Material_Code", *TARGET_MONTHS])

    combined = pd.concat(frames, ignore_index=True)
    month_cols = [m for m in TARGET_MONTHS if m in combined.columns]
    grouped = combined.groupby("Material_Code", as_index=False)[month_cols].sum()
    grouped["Forecast_Source"] = schedule_path.name if schedule_path else "Sample Balance + Froecast"
    return grouped


def _parse_penta_schedule(path: Path) -> pd.DataFrame:
    """Parse PENTA Schedule weekly file; auto-detect material + month columns."""
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    raw = pd.read_excel(path, sheet_name=sheet, header=None)

    header_row = None
    for i in range(min(25, len(raw))):
        row_vals = [_text(v) for v in raw.iloc[i].tolist()]
        joined = " ".join(row_vals).lower()
        if ("material" in joined or "fg code" in joined or "ptt material" in joined) and any(
            m in joined for m in ("jun", "jul", "aug", "2026", "forecast")
        ):
            header_row = i
            break

    if header_row is None:
        for i in range(min(15, len(raw))):
            if _text(raw.iloc[i, 0]).lower() in ("item", "material", "material code", "ptt material code"):
                header_row = i
                break

    if header_row is None:
        header_row = 2

    headers = [_text(v) for v in raw.iloc[header_row].tolist()]
    code_col = next(
        (i for i, h in enumerate(headers) if h and any(k in h.lower() for k in ("material code", "ptt material", "fg code", "material"))),
        0,
    )

    month_col_map: dict[str, int] = {}
    for i, h in enumerate(headers):
        hl = h.lower().replace("'", "")
        for ym, names in {
            "2026-06": ("jun", "06.2026", "2026-06", "6/2026"),
            "2026-07": ("jul", "07.2026", "2026-07", "7/2026"),
            "2026-08": ("aug", "08.2026", "2026-08", "8/2026"),
            "2026-09": ("sep", "09.2026", "2026-09", "9/2026"),
            "2026-10": ("oct", "10.2026", "2026-10", "10/2026"),
        }.items():
            if any(n in hl for n in names) and "kg" in hl or hl in names or hl.endswith("2026"):
                month_col_map.setdefault(ym, i)
            elif hl in ("jun", "jul", "aug", "sep", "oct") and i + 1 < len(headers) and "kg" in headers[i + 1].lower():
                month_col_map.setdefault(ym, i + 1)

    # Sample Balance Forecast layout fallback inside schedule-like files
    if not month_col_map:
        month_col_map = {f"2026-{k}": v for k, v in {"06": 43, "07": 45, "08": 47, "09": 49, "10": 51}.items()}

    records = []
    for r in range(header_row + 1, len(raw)):
        code = raw.iloc[r, code_col]
        if pd.isna(code) or not _text(code):
            continue
        rec = {"Material_Code": _text(code)}
        for ym, col in month_col_map.items():
            if col < raw.shape[1]:
                rec[ym] = _num(raw.iloc[r, col]) / 1000.0 if _num(raw.iloc[r, col]) > 500 else _num(raw.iloc[r, col])
        if any(rec.get(m, 0) > 0 for m in TARGET_MONTHS):
            records.append(rec)

    df = pd.DataFrame(records)
    if df.empty:
        return df
    month_cols = [m for m in TARGET_MONTHS if m in df.columns]
    out = df.groupby("Material_Code", as_index=False)[month_cols].sum()
    out["Forecast_Source"] = path.name
    return out


def load_client_forecast_from_froecast(target_months: list[str] | None = None) -> pd.DataFrame:
    import contextlib
    import io

    from material_balance import load_forecast

    target_months = target_months or TARGET_MONTHS
    with contextlib.redirect_stdout(io.StringIO()):
        fc = load_forecast(FORECAST_DIR)
    if fc.empty:
        return pd.DataFrame(columns=["Material_Code", *target_months])

    records = []
    for _, row in fc.iterrows():
        rec = {"Material_Code": _text(row["material_code"])}
        for ym in target_months:
            short = {"06": "Jun", "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct"}[ym.split("-")[1]]
            rec[ym] = _num(row.get(short, 0)) / 1000.0
        records.append(rec)
    df = pd.DataFrame(records)
    return df.groupby("Material_Code", as_index=False).sum(numeric_only=True)


def load_client_forecast(target_months: list[str] | None = None) -> pd.DataFrame:
    """Unified forecast: PENTA Schedule wk25 + Sample Balance + Froecast."""
    return load_penta_schedule_forecast()


def load_order_history(sample_path: Path | None = None) -> pd.DataFrame:
    sample_path = sample_path or resolve_sample_path()
    act = pd.read_excel(sample_path, sheet_name="Act order", header=None)
    rows = []
    for _, row in act.iloc[4:].iterrows():
        fg = row[3]
        if pd.isna(fg):
            continue
        rows.append({
            "FG_Code": _text(fg),
            "Material_Code": _text(row[0]),
            "M-1": _num(row[25]),
            "M-2": _num(row[24]),
            "M-3": _num(row[23]),
        })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.groupby("FG_Code", as_index=False).agg({
        "Material_Code": "first",
        "M-1": "sum",
        "M-2": "sum",
        "M-3": "sum",
    })


def attach_material_codes(df: pd.DataFrame, rd004: pd.DataFrame) -> pd.DataFrame:
    """Assign Material_Code to open MP008 via U-Stock Mat Spec rules + customer priority."""
    from stock_matching import allocate_mp008_inbound

    return allocate_mp008_inbound(df, rd004)


def get_data_source_summary() -> dict[str, str]:
    """Report which source files were resolved."""
    from stock_matching import resolve_rules_path

    carrier = discover_file(["*SO003*Carrier*", "*SO003-Carrier*"], [SAMPLE_DIR])
    schedule = discover_file(["*PENTA*Schedule*wk*25*", "*Schedule*wk*25*"], [SAMPLE_DIR, FORECAST_DIR])
    rules_path = resolve_rules_path()
    return {
        "sample_balance": str(resolve_sample_path().name),
        "so003": carrier.name if carrier else f"{resolve_sample_path().name} (SO003 sheet)",
        "ms004": resolve_stock_path().name,
        "mp008": (discover_file(["*MP008*"], [SAMPLE_DIR]) or resolve_sample_path()).name,
        "forecast": schedule.name if schedule else "Sample Balance Forecast + Froecast",
        "stock_matching_rules": rules_path.name if rules_path.exists() else f"{rules_path.name} (embedded defaults)",
    }
