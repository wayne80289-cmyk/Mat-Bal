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


def is_carrier_label(*values) -> bool:
    """True when customer/cust/source text refers to Carrier (excluded from this engine)."""
    for v in values:
        t = _text(v).upper().replace(" ", "")
        if t and "CARRIER" in t:
            return True
    return False


def is_carrier_path(path: Path) -> bool:
    return "CARRIER" in path.name.upper()


def exclude_carrier_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Drop rows tied to Carrier customers/sources."""
    if df.empty:
        return df
    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        key = str(col).upper().replace(" ", "").replace("_", "")
        if key in ("CUSTOMER", "CUSTCODE", "CUST", "MAINCUSTOMER", "DELIVERYPLACE"):
            mask |= df[col].map(lambda v: is_carrier_label(v))
    if "Source" in df.columns:
        mask |= df["Source"].map(lambda v: is_carrier_label(v))
    if "Source_File" in df.columns:
        mask |= df["Source_File"].map(lambda v: is_carrier_label(v))
    return df.loc[~mask].copy()


def discover_file(
    patterns: list[str],
    search_dirs: list[Path] | None = None,
    *,
    exclude_carrier: bool = True,
) -> Path | None:
    """Return newest matching file across search directories."""
    search_dirs = search_dirs or [SAMPLE_DIR, FORECAST_DIR, BASE_DIR]
    matches: list[Path] = []
    for directory in search_dirs:
        if not directory.exists():
            continue
        for pattern in patterns:
            matches.extend(directory.glob(pattern))
    if exclude_carrier:
        matches = [p for p in matches if not is_carrier_path(p)]
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def resolve_sample_path() -> Path:
    found = discover_file(
        ["All Customer review*.xlsx", "*Customer review*.xlsx"],
        [SAMPLE_DIR, BASE_DIR],
        exclude_carrier=True,
    )
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


def _load_forecast_fg_map(sample_path: Path) -> dict[str, str]:
    """Material_Code → FG_Code from Sample Balance Forecast sheet."""
    try:
        fc = pd.read_excel(sample_path, sheet_name="Forecast", header=None)
    except Exception:
        return {}
    fg_map: dict[str, str] = {}
    for r in range(4, len(fc)):
        code = fc.iloc[r, 0]
        fg = fc.iloc[r, 2]
        if pd.isna(code):
            continue
        code_t = _text(code)
        fg_t = _text(fg)
        if code_t and fg_t:
            fg_map[code_t] = fg_t
    return fg_map


def _parse_fg_from_main_customer(main_customer: str) -> str:
    """e.g. TSGP 1.0x1326xCoil → 1.0x1326xCoil"""
    text = _text(main_customer)
    m = re.search(r"([\d.]+)\s*[xX]\s*([\d.]+)\s*[xX]\s*(\w+)", text)
    if not m:
        return ""
    return f"{m.group(1)}x{m.group(2)}x{m.group(3)}".replace(".0", "")


def _resolve_material_fg_and_width(
    code: str,
    thickness: float,
    act_info: dict,
    fg_codes: set[str],
    forecast_fg: str,
    main_customer: str,
) -> tuple[str, str, float]:
    """Resolve FG_Code / FG_Codes_All / finished Width for RD004 master."""
    from stock_matching import parse_product_width_from_material_code

    product_w = parse_product_width_from_material_code(code, thickness)
    act_w = _num(act_info.get("Width"))
    if product_w and product_w > 0 and round(product_w, 0) != 1219:
        width = product_w
    elif act_w > 0:
        width = act_w
    else:
        width = 1219.0

    if fg_codes:
        fg_primary = sorted(fg_codes)[0]
        fg_all = ",".join(sorted(fg_codes))
    elif forecast_fg:
        fg_primary = forecast_fg
        fg_all = forecast_fg
    elif product_w and product_w > 0 and round(product_w, 0) != 1219:
        fg_primary = f"{thickness:g}x{product_w:g}xCoil".replace(".0", "")
        fg_all = fg_primary
    else:
        mc_fg = _parse_fg_from_main_customer(main_customer)
        if mc_fg:
            fg_primary = mc_fg
            fg_all = mc_fg
        elif act_w > 0 and round(act_w, 0) != 1219:
            fg_primary = f"{thickness:g}x{act_w:g}xCoil".replace(".0", "")
            fg_all = fg_primary
        else:
            fg_primary = code
            fg_all = code

    return fg_primary, fg_all, width


def load_rd004_master(sample_path: Path | None = None) -> pd.DataFrame:
    """Build material master from Balance sheet + Act order."""
    sample_path = sample_path or resolve_sample_path()
    bal = pd.read_excel(sample_path, sheet_name="Balance sheet (Update)", header=None)
    act = pd.read_excel(sample_path, sheet_name="Act order", header=None)
    forecast_fg_map = _load_forecast_fg_map(sample_path)

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
        main_customer = _text(bal.iloc[i, 4])
        fg_codes = fg_map.get(code, set())
        fg_primary, fg_all, width = _resolve_material_fg_and_width(
            code,
            t,
            info,
            fg_codes,
            forecast_fg_map.get(code, ""),
            main_customer,
        )
        records.append({
            "Material_Code": code,
            "Common_Group": info.get("Spec", code),
            "Spec": info.get("Spec", ""),
            "Thickness": t,
            "Width": width,
            "FG_Code": fg_primary,
            "FG_Codes_All": fg_all,
            "Kind": info.get("Kind", _text(bal.iloc[i, 2])),
            "Main_Customer": main_customer,
            "MOQ": 0.0,
        })
    df = pd.DataFrame(records)
    df = df[~df["Main_Customer"].map(lambda v: is_carrier_label(v))]
    from stock_matching import supplement_rd004_pairing_materials

    return supplement_rd004_pairing_materials(df)


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
    return exclude_carrier_rows(out)


def load_so003(sample_path: Path | None = None) -> pd.DataFrame:
    """Load SO003 from Sample Balance workbook (Carrier sources excluded)."""
    sample_path = sample_path or resolve_sample_path()
    raw = pd.read_excel(sample_path, sheet_name="SO003", header=None)
    return exclude_carrier_rows(_parse_so003_raw(raw, source=sample_path.name))


def _resolve_mp008_sheet(workbook: Path) -> str:
    xl = pd.ExcelFile(workbook)
    for name in xl.sheet_names:
        if name.strip().upper().startswith("MP008"):
            return name
    raise ValueError(f"No MP008 sheet found in {workbook.name}")


def load_mp008(sample_path: Path | None = None) -> pd.DataFrame:
    """Load open PO from standalone MP008 export or Sample Balance workbook."""
    mp008_path = discover_file(["MP008*.xlsx", "MP008*.xls"], [SAMPLE_DIR, BASE_DIR])
    if mp008_path and mp008_path.suffix.lower() in (".xlsx", ".xls"):
        sheet = _resolve_mp008_sheet(mp008_path)
        raw = pd.read_excel(mp008_path, sheet_name=sheet, header=None)
        return exclude_carrier_rows(_parse_mp008_raw(raw, source=mp008_path.name))

    sample_path = sample_path or resolve_sample_path()
    sheet = _resolve_mp008_sheet(sample_path)
    raw = pd.read_excel(sample_path, sheet_name=sheet, header=None)
    return exclude_carrier_rows(_parse_mp008_raw(raw, source=sample_path.name))


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
    df = pd.DataFrame(records)
    df = exclude_carrier_rows(df)
    return df.groupby("Material_Code", as_index=False).sum(numeric_only=True)


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
    combined = exclude_carrier_rows(combined)
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
    """U7 fallback: Act order M-1/M-2/M-3 (kg) by FG_Code."""
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


def _resolve_sa_sales_sheet(workbook: Path) -> str:
    xl = pd.ExcelFile(workbook)
    for name in xl.sheet_names:
        if name.strip().upper() == "SA006":
            return name
    for name in xl.sheet_names:
        if name.strip().upper() == "SA007":
            return name
    raise ValueError(f"No SA006/SA007 sheet in {workbook.name}")


def _parse_sa007_sales(raw: pd.DataFrame, source: str) -> pd.DataFrame:
    """Parse SA007 pivot: Weight cols for relative months 3–5 = last 3 months (kg)."""
    hdr = 4
    wt_cols = {3: 14, 4: 15, 5: 16}  # M-3, M-2, M-1 (oldest → newest)
    rows = []
    for i in range(hdr + 1, len(raw)):
        fg = raw.iloc[i, 2]
        if pd.isna(fg) or not _text(fg):
            continue
        m3 = _num(raw.iloc[i, wt_cols[3]])
        m2 = _num(raw.iloc[i, wt_cols[4]])
        m1 = _num(raw.iloc[i, wt_cols[5]])
        if m1 + m2 + m3 <= 0:
            continue
        rows.append({
            "Customer": _text(raw.iloc[i, 0]),
            "Delivery_Place": _text(raw.iloc[i, 1]),
            "FG_Code": _text(fg),
            "Spec": _text(raw.iloc[i, 3]),
            "Thickness": _num(raw.iloc[i, 4]),
            "Width": _num(raw.iloc[i, 5]),
            "M-3_kg": m3,
            "M-2_kg": m2,
            "M-1_kg": m1,
            "Total_3mo_kg": m1 + m2 + m3,
            "Source_Sheet": "SA007",
            "Source_File": source,
        })
    return pd.DataFrame(rows)


def _parse_sa006_sheet(raw: pd.DataFrame, source: str) -> pd.DataFrame:
    """Parse SA006 when layout matches SA007 (Weight months 3–5)."""
    return _parse_sa007_sales(raw, source).assign(Source_Sheet="SA006")


def load_sa006_sales(sample_path: Path | None = None) -> pd.DataFrame:
    """
    Load last-3-month actual sales (kg) for plan universe.
    Prefers SA006 sheet; falls back to SA007 (same layout in Sample Balance).
    Carrier customers excluded.
    """
    sample_path = sample_path or resolve_sample_path()
    sheet = _resolve_sa_sales_sheet(sample_path)
    raw = pd.read_excel(sample_path, sheet_name=sheet, header=None)
    if sheet.strip().upper() == "SA006":
        df = _parse_sa006_sheet(raw, sample_path.name)
    else:
        df = _parse_sa007_sales(raw, sample_path.name)
    return exclude_carrier_rows(df)


def build_fg_to_material_map(rd004: pd.DataFrame) -> dict[str, str]:
    fg_to_mat: dict[str, str] = {}
    for _, m in rd004.iterrows():
        fg_to_mat[_text(m["FG_Code"])] = _text(m["Material_Code"])
        for fg in str(m.get("FG_Codes_All", "")).split(","):
            if fg.strip():
                fg_to_mat[fg.strip()] = _text(m["Material_Code"])
    return fg_to_mat


def aggregate_sa006_by_material(sa006_df: pd.DataFrame, rd004: pd.DataFrame) -> pd.DataFrame:
    """Roll SA006/SA007 FG sales up to Material_Code with U-Stock matching fallback."""
    from stock_matching import find_rd004_matches, load_matching_rules, resolve_customer_spec_pairing

    if sa006_df.empty:
        return pd.DataFrame(columns=[
            "Material_Code", "FG_Code", "Main_Customer", "Customers",
            "M-3_kg", "M-2_kg", "M-1_kg", "Total_3mo_kg", "Avg_Monthly_Ton", "Source_Sheet",
        ])

    rules = load_matching_rules()
    fg_to_mat = build_fg_to_material_map(rd004)
    df = sa006_df.copy()
    mat_codes: list[str] = []
    for _, row in df.iterrows():
        fg = _text(row["FG_Code"])
        code = fg_to_mat.get(fg, "")
        if not code:
            code = resolve_customer_spec_pairing(
                row.get("Spec", ""),
                row.get("Thickness", 0),
                row.get("Width", 0),
                rd004,
                fg,
                row.get("Customer", ""),
            ) or ""
        if not code:
            hits = find_rd004_matches(
                row.get("Spec", ""),
                row.get("Thickness", 0),
                row.get("Width", 0),
                rd004,
                rules,
            )
            if not hits.empty:
                code = _text(hits.iloc[0]["Material_Code"])
        if not code:
            code = fg
        mat_codes.append(code)
    df["Material_Code"] = mat_codes
    df = df[df["Material_Code"] != ""].copy()
    if df.empty:
        return pd.DataFrame(columns=[
            "Material_Code", "FG_Code", "Main_Customer", "Customers",
            "M-3_kg", "M-2_kg", "M-1_kg", "Total_3mo_kg", "Avg_Monthly_Ton", "Source_Sheet",
        ])

    main_cust = (
        rd004.set_index("Material_Code")["Main_Customer"].to_dict()
        if "Main_Customer" in rd004.columns
        else {}
    )
    fg_primary = (
        rd004.set_index("Material_Code")["FG_Code"].to_dict()
        if "FG_Code" in rd004.columns
        else {}
    )

    grouped = df.groupby("Material_Code", as_index=False).agg({
        "M-3_kg": "sum",
        "M-2_kg": "sum",
        "M-1_kg": "sum",
        "Total_3mo_kg": "sum",
        "Customer": lambda s: ", ".join(sorted({c for c in s if c})[:5]),
        "Source_Sheet": "first",
    })
    grouped.rename(columns={"Customer": "Customers"}, inplace=True)
    grouped["FG_Code"] = grouped["Material_Code"].map(lambda m: fg_primary.get(m, m))
    grouped["Main_Customer"] = grouped["Material_Code"].map(lambda m: main_cust.get(m, ""))
    grouped["Avg_Monthly_Ton"] = (grouped["Total_3mo_kg"] / 3.0 / 1000.0).round(3)
    grouped = grouped[grouped["Total_3mo_kg"] > 0].copy()
    return grouped.sort_values("Total_3mo_kg", ascending=False).reset_index(drop=True)


def attach_material_codes(df: pd.DataFrame, rd004: pd.DataFrame) -> pd.DataFrame:
    """Assign Material_Code to open MP008 via U-Stock Mat Spec rules + customer priority."""
    from stock_matching import allocate_mp008_inbound

    return allocate_mp008_inbound(df, rd004)


def get_data_source_summary() -> dict[str, str]:
    """Report which source files were resolved (Carrier sources excluded)."""
    from stock_matching import resolve_rules_path

    schedule = discover_file(["*PENTA*Schedule*wk*25*", "*Schedule*wk*25*"], [SAMPLE_DIR, FORECAST_DIR])
    rules_path = resolve_rules_path()
    return {
        "sample_balance": str(resolve_sample_path().name),
        "so003": f"{resolve_sample_path().name} (SO003 sheet; Carrier excluded)",
        "ms004": f"{resolve_stock_path().name} (Carrier excluded)",
        "mp008": (
            f"{(discover_file(['MP008*.xlsx', 'MP008*.xls'], [SAMPLE_DIR, BASE_DIR]) or resolve_sample_path()).name}"
            " (Carrier excluded)"
        ),
        "forecast": (
            f"{schedule.name} + Sample Balance + Froecast (Carrier excluded)"
            if schedule
            else "Sample Balance Forecast + Froecast (Carrier excluded)"
        ),
        "sa006_sales": (
            f"{resolve_sample_path().name} (SA006/SA007 近3月銷售; Carrier excluded)"
        ),
        "stock_matching_rules": rules_path.name if rules_path.exists() else f"{rules_path.name} (embedded defaults)",
        "scope": "不含 Carrier 客戶（本系統僅分析其他客戶訂單）",
    }
