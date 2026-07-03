# -*- coding: utf-8 -*-
"""Load MS004, SO003, MP008, RD004 master, and forecast for balance engine."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DIR = BASE_DIR / "Sample Balance"
HISTORY_BALANCE_DIR = BASE_DIR / "History Balance"
STOCK_DIR = BASE_DIR / "Stock"
SA007_DIR = BASE_DIR / "SA007"
SA007_HISTORY_LABEL = "SA007歷史銷售紀錄"
# Legacy Sample Balance workbook sheet name (fallback when SA007/ is empty)
SAMPLE_BALANCE_SA007_SHEET = "Act order"
SO003_DIR = BASE_DIR / "SO003"
MP008_DIR = BASE_DIR / "MP008"
RD004_DIR = BASE_DIR / "RD004"
FORECAST_DIR = BASE_DIR / "Forecast"
FORECAST_DIR_LEGACY = BASE_DIR / "Froecast"
# U1: Forecast/ 僅含「有提供預估表」的客戶，非全部客戶總需求
FORECAST_SCOPE_LABEL = "有提供預估表之客戶（非全部客戶需求）"
# History Balance / Sample Balance：僅報告樣本取樣參考，不作 Supply Plan 運行資料來源
TEMPLATE_REFERENCE_LABEL = "僅報告樣本取樣參考（非運行資料來源）"
OUTPUT_DIR = BASE_DIR / "Output"

SAMPLE_PATH = SAMPLE_DIR / "All Customer review Jun '2026 review 20.06.2026.xlsx"
STOCK_PATH = STOCK_DIR / "MS004-260619.xls"

TARGET_MONTHS = ["2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
MONTH_TO_YM = {"Jun": "2026-06", "Jul": "2026-07", "Aug": "2026-08", "Sep": "2026-09", "Oct": "2026-10"}
FORECAST_KG_COLS = {"Jun": 43, "Jul": 45, "Aug": 47, "Sep": 49, "Oct": 51}

BASE_RD004_COLUMNS = [
    "Material_Code",
    "Common_Group",
    "Spec",
    "Thickness",
    "Width",
    "FG_Code",
    "FG_Codes_All",
    "Kind",
    "Main_Customer",
    "MOQ",
]


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


def resolve_forecast_dir() -> Path:
    """Primary: Forecast/ folder; fallback Froecast/ for legacy layouts."""
    if FORECAST_DIR.is_dir() and any(FORECAST_DIR.iterdir()):
        return FORECAST_DIR
    return FORECAST_DIR_LEGACY


def resolve_sa007_paths() -> list[Path]:
    """All SA007 sales workbooks under SA007/ (newest last)."""
    if not SA007_DIR.is_dir():
        return []
    paths: list[Path] = []
    for pattern in ("*.xlsx", "*.xls"):
        paths.extend(SA007_DIR.glob(pattern))
    paths = [p for p in paths if not is_carrier_path(p)]
    return sorted(paths, key=lambda p: p.stat().st_mtime)


def resolve_so003_paths() -> list[Path]:
    """All SO003 order exports under SO003/ (newest last)."""
    if not SO003_DIR.is_dir():
        return []
    paths: list[Path] = []
    for pattern in ("SO003*.xlsx", "SO003*.xls", "*.xlsx", "*.xls"):
        paths.extend(SO003_DIR.glob(pattern))
    paths = [p for p in paths if not is_carrier_path(p)]
    # de-dup while preserving order
    seen: set[Path] = set()
    unique: list[Path] = []
    for p in sorted(paths, key=lambda x: x.stat().st_mtime):
        if p not in seen:
            seen.add(p)
            unique.append(p)
    return unique


def resolve_rd004_master_path() -> Path | None:
    """Newest Material Master export under RD004/."""
    if not RD004_DIR.is_dir():
        return None
    matches: list[Path] = []
    for pattern in ("Material Master*.xlsx", "*Material*Master*.xlsx", "RD004*.xlsx"):
        matches.extend(RD004_DIR.glob(pattern))
    matches = [
        p for p in matches
        if p.suffix.lower() == ".xlsx"
        and not is_carrier_path(p)
        and "matching-rules" not in p.name.lower()
        and "stock-material" not in p.name.lower()
    ]
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def resolve_mp008_paths() -> list[Path]:
    """All MP008 open-PO exports under MP008/ (newest last)."""
    if not MP008_DIR.is_dir():
        return []
    paths: list[Path] = []
    for pattern in ("MP008*.xlsx", "MP008*.xls", "*.xlsx", "*.xls"):
        paths.extend(MP008_DIR.glob(pattern))
    paths = [p for p in paths if not is_carrier_path(p)]
    seen: set[Path] = set()
    unique: list[Path] = []
    for p in sorted(paths, key=lambda x: x.stat().st_mtime):
        if p not in seen:
            seen.add(p)
            unique.append(p)
    return unique


def discover_file(
    patterns: list[str],
    search_dirs: list[Path] | None = None,
    *,
    exclude_carrier: bool = True,
) -> Path | None:
    """Return newest matching file across search directories."""
    search_dirs = search_dirs or [SAMPLE_DIR, resolve_forecast_dir(), SA007_DIR, SO003_DIR, MP008_DIR, RD004_DIR, BASE_DIR]
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


def _load_forecast_fg_map_from_folder() -> dict[str, str]:
    """Material_Code → FG_Code from Forecast/ vendor xlsx (best-effort)."""
    fg_map: dict[str, str] = {}
    fc_dir = resolve_forecast_dir()
    if not fc_dir.is_dir():
        return fg_map
    for path in sorted(fc_dir.glob("*.xlsx")):
        if is_carrier_path(path):
            continue
        try:
            raw = pd.read_excel(path, sheet_name=0, header=None)
            for r in range(min(30, len(raw)), len(raw)):
                code = raw.iloc[r, 0]
                fg = raw.iloc[r, 2] if raw.shape[1] > 2 else None
                if pd.isna(code):
                    continue
                code_t, fg_t = _text(code), _text(fg)
                if code_t and fg_t:
                    fg_map[code_t] = fg_t
        except Exception:
            continue
    return fg_map


def _normalize_rd004_master_df(df: pd.DataFrame, forecast_fg_map: dict[str, str] | None = None) -> pd.DataFrame:
    """Keep base RD004 columns; fix nulls and FG/Width from Material Code where needed."""
    from stock_matching import parse_product_width_from_material_code

    forecast_fg_map = forecast_fg_map or {}
    out = df.copy()
    for col in BASE_RD004_COLUMNS:
        if col not in out.columns:
            out[col] = "" if col not in ("Thickness", "Width", "MOQ") else 0.0

    records = []
    for _, row in out.iterrows():
        code = _text(row["Material_Code"])
        if not code:
            continue
        t = _num(row.get("Thickness"))
        spec = _text(row.get("Spec"))
        common = _text(row.get("Common_Group")) or spec or code
        pw = parse_product_width_from_material_code(code, t)
        width = pw if pw and pw > 0 else _num(row.get("Width")) or 1219.0
        fg = _text(row.get("FG_Code"))
        fg_all = _text(row.get("FG_Codes_All")) or fg
        if forecast_fg_map.get(code):
            fg = forecast_fg_map[code]
            fg_all = fg
        elif pw and pw > 0 and round(pw, 0) != 1219:
            if not fg or "1219x1219" in fg.replace(" ", ""):
                fg = f"{t:g}x{pw:g}xCoil".replace(".0", "")
                fg_all = fg
        records.append({
            "Material_Code": code,
            "Common_Group": common,
            "Spec": spec,
            "Thickness": t,
            "Width": width,
            "FG_Code": fg or code,
            "FG_Codes_All": fg_all or fg or code,
            "Kind": _text(row.get("Kind")),
            "Main_Customer": _text(row.get("Main_Customer")),
            "MOQ": _num(row.get("MOQ")),
        })
    return pd.DataFrame(records)


def _load_rd004_from_folder() -> pd.DataFrame:
    path = resolve_rd004_master_path()
    if not path:
        return pd.DataFrame(columns=BASE_RD004_COLUMNS)
    xl = pd.ExcelFile(path)
    sheet = "Material Master" if "Material Master" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    fg_map = _load_forecast_fg_map_from_folder()
    if not fg_map:
        fg_map = _load_forecast_fg_map(resolve_sample_path())
    return _normalize_rd004_master_df(df, fg_map)


def _load_rd004_from_sample_balance(sample_path: Path) -> pd.DataFrame:
    """Legacy: build RD004 from Sample Balance Balance sheet + SA007歷史銷售紀錄（Sample Balance fallback）。"""
    bal = pd.read_excel(sample_path, sheet_name="Balance sheet (Update)", header=None)
    sa007_legacy = pd.read_excel(sample_path, sheet_name=SAMPLE_BALANCE_SA007_SHEET, header=None)
    forecast_fg_map = _load_forecast_fg_map(sample_path)

    fg_map: dict[str, set[str]] = {}
    spec_map: dict[str, dict] = {}
    for _, row in sa007_legacy.iloc[4:].iterrows():
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
    return pd.DataFrame(records)


def load_rd004_master(sample_path: Path | None = None) -> pd.DataFrame:
    """Material master + pairing base from RD004/ folder (fallback: Sample Balance)."""
    df = _load_rd004_from_folder()
    if df.empty:
        sample_path = sample_path or resolve_sample_path()
        df = _load_rd004_from_sample_balance(sample_path)
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


def _read_so003_workbook(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    for name in xl.sheet_names:
        upper = name.strip().upper()
        if upper == "SO003" or upper.startswith("SO003"):
            sheet = name
            break
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    return _parse_so003_raw(raw, source=path.name)


def _load_so003_from_sample_fallback(sample_path: Path | None = None) -> pd.DataFrame:
    sample_path = sample_path or resolve_sample_path()
    raw = pd.read_excel(sample_path, sheet_name="SO003", header=None)
    return _parse_so003_raw(raw, source=sample_path.name)


def load_so003(sample_path: Path | None = None) -> pd.DataFrame:
    """Load SO003 open orders from SO003/ folder (Carrier excluded)."""
    frames: list[pd.DataFrame] = []
    for path in resolve_so003_paths():
        try:
            part = _read_so003_workbook(path)
            if not part.empty:
                frames.append(part)
        except Exception:
            continue
    if frames:
        df = pd.concat(frames, ignore_index=True)
    else:
        df = _load_so003_from_sample_fallback(sample_path)
    return exclude_carrier_rows(df)


def _resolve_mp008_sheet(workbook: Path) -> str:
    xl = pd.ExcelFile(workbook)
    for name in xl.sheet_names:
        if name.strip().upper().startswith("MP008"):
            return name
    raise ValueError(f"No MP008 sheet found in {workbook.name}")


def _read_mp008_workbook(path: Path) -> pd.DataFrame:
    sheet = _resolve_mp008_sheet(path)
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    return _parse_mp008_raw(raw, source=path.name)


def _load_mp008_from_sample_fallback(sample_path: Path | None = None) -> pd.DataFrame:
    sample_path = sample_path or resolve_sample_path()
    sheet = _resolve_mp008_sheet(sample_path)
    raw = pd.read_excel(sample_path, sheet_name=sheet, header=None)
    return _parse_mp008_raw(raw, source=sample_path.name)


def load_mp008(sample_path: Path | None = None) -> pd.DataFrame:
    """Load open PO from MP008/ folder (Carrier excluded)."""
    frames: list[pd.DataFrame] = []
    for path in resolve_mp008_paths():
        try:
            part = _read_mp008_workbook(path)
            if not part.empty:
                frames.append(part)
        except Exception:
            continue
    if frames:
        df = pd.concat(frames, ignore_index=True)
    else:
        df = _load_mp008_from_sample_fallback(sample_path)
    return exclude_carrier_rows(df)


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
        [SAMPLE_DIR, resolve_forecast_dir(), BASE_DIR],
    )

    frames: list[pd.DataFrame] = []
    if schedule_path:
        frames.append(_parse_penta_schedule(schedule_path))

    sample_fc = load_sample_balance_forecast()
    if not sample_fc.empty:
        sample_fc = sample_fc.copy()
        sample_fc["Source"] = "Sample Balance Forecast"
        frames.append(sample_fc)

    vendor_fc = load_client_forecast()
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


def load_client_forecast(target_months: list[str] | None = None) -> pd.DataFrame:
    """有提供預估表之客戶 Forecast（Forecast/ 資料夾；非全部客戶需求，見 U1）。"""
    import contextlib
    import io

    from material_balance import load_forecast

    target_months = target_months or TARGET_MONTHS
    fc_dir = resolve_forecast_dir()
    with contextlib.redirect_stdout(io.StringIO()):
        fc = load_forecast(fc_dir)
    if fc.empty:
        return pd.DataFrame(columns=["Material_Code", *target_months, "Forecast_Source", "Forecast_Scope"])

    records = []
    for _, row in fc.iterrows():
        rec = {"Material_Code": _text(row["material_code"])}
        for ym in target_months:
            short = {"06": "Jun", "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct"}[ym.split("-")[1]]
            rec[ym] = _num(row.get(short, 0)) / 1000.0
        records.append(rec)
    df = pd.DataFrame(records)
    if df.empty:
        return pd.DataFrame(columns=["Material_Code", *target_months, "Forecast_Source", "Forecast_Scope"])
    month_cols = [m for m in target_months if m in df.columns]
    grouped = df.groupby("Material_Code", as_index=False)[month_cols].sum()
    grouped["Forecast_Scope"] = FORECAST_SCOPE_LABEL
    grouped["Forecast_Source"] = fc_dir.name
    grouped["Source"] = "Forecast"
    return grouped


def load_balance_forecast(
    rd004: pd.DataFrame,
    target_months: list[str] | None = None,
) -> pd.DataFrame:
    """
    Map Forecast/ vendor rows onto RD004 Material_Code (spec + T + W).
    This is what Balance Forecast_Ton uses; vendor codes rarely match RD004 exactly.
    """
    import contextlib
    import io

    from material_balance import MONTHS_BALANCE, load_forecast, match_forecast_to_material

    target_months = target_months or TARGET_MONTHS
    month_short = {"06": "Jun", "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct"}
    fc_dir = resolve_forecast_dir()
    with contextlib.redirect_stdout(io.StringIO()):
        vendor_fc = load_forecast(fc_dir)

    cols = ["Material_Code", "FG_Code", *target_months, "Forecast_Scope", "Forecast_Source"]
    if vendor_fc.empty or rd004.empty:
        return pd.DataFrame(columns=cols)

    records = []
    for _, row in rd004.iterrows():
        code = _text(row["Material_Code"])
        if not code:
            continue
        matched_kg = match_forecast_to_material(code, vendor_fc)
        rec = {
            "Material_Code": code,
            "FG_Code": _text(row.get("FG_Code")),
            "Forecast_Scope": FORECAST_SCOPE_LABEL,
            "Forecast_Source": fc_dir.name,
        }
        total_ton = 0.0
        for ym in target_months:
            short = month_short[ym.split("-")[1]]
            ton = _num(matched_kg.get(short, 0)) / 1000.0
            rec[ym] = round(ton, 6)
            total_ton += ton
        if total_ton > 0:
            records.append(rec)

    if not records:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(records)


_MONTH_SHORT = {"06": "Jun", "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct"}


def build_forecast_integrated_report(target_months: list[str] | None = None) -> dict[str, pd.DataFrame]:
    """
    Integrate Forecast/ vendor files into one report.
    Scope: customers who submitted forecasts only (U1); not all customer demand.
    """
    import contextlib
    import io

    from material_balance import MONTHS_BALANCE, load_forecast_vendor_detail

    target_months = target_months or TARGET_MONTHS
    fc_dir = resolve_forecast_dir()
    with contextlib.redirect_stdout(io.StringIO()):
        detail = load_forecast_vendor_detail(fc_dir)

    info_rows = [
        {"項目": "資料範圍", "內容": FORECAST_SCOPE_LABEL},
        {"項目": "說明", "內容": "未提供預估表之客戶需求由 SA007歷史銷售紀錄平均估算（U7），不在此報告內"},
        {"項目": "資料夾", "內容": str(fc_dir)},
        {"項目": "各家檔案數", "內容": detail["Source_File"].nunique() if not detail.empty and "Source_File" in detail.columns else 0},
        {"項目": "明細列數", "內容": len(detail)},
    ]
    if not detail.empty and "Source_File" in detail.columns:
        for fname, cnt in detail["Source_File"].value_counts().sort_index().items():
            info_rows.append({"項目": f"檔案: {fname}", "內容": f"{int(cnt)} 列"})

    if detail.empty:
        empty_cols = ["Source_File", "material_code", "spec", "t", "w", *target_months]
        return {
            "Forecast_說明": pd.DataFrame(info_rows),
            "Forecast_各家明細": pd.DataFrame(columns=empty_cols),
            "Forecast_整合彙總": pd.DataFrame(columns=["Material_Code", *target_months, "Vendor_Count"]),
        }

    vendor_rows = []
    for _, row in detail.iterrows():
        rec = {
            "Source_File": _text(row.get("Source_File")),
            "Parser": _text(row.get("source")),
            "Material_Code": _text(row.get("material_code")),
            "Spec": _text(row.get("spec")),
            "T": _num(row.get("t")),
            "W": _num(row.get("w")),
        }
        for ym in target_months:
            short = _MONTH_SHORT[ym.split("-")[1]]
            rec[f"{ym}_kg"] = _num(row.get(short, 0))
            rec[f"{ym}_Ton"] = round(_num(row.get(short, 0)) / 1000.0, 3)
        vendor_rows.append(rec)
    vendor_df = pd.DataFrame(vendor_rows)

    summary_records = []
    ton_cols = [f"{ym}_Ton" for ym in target_months]
    for mat, grp in vendor_df.groupby("Material_Code", dropna=False):
        if not _text(mat):
            continue
        rec = {"Material_Code": _text(mat), "Vendor_Count": grp["Source_File"].nunique()}
        for ym in target_months:
            rec[ym] = round(grp[f"{ym}_Ton"].sum(), 3)
        summary_records.append(rec)
    summary_df = pd.DataFrame(summary_records)

    return {
        "Forecast_說明": pd.DataFrame(info_rows),
        "Forecast_各家明細": vendor_df,
        "Forecast_整合彙總": summary_df,
    }


def _detect_sa007_workbook_format(raw: pd.DataFrame) -> str:
    """material_code_layout = SA007歷史銷售紀錄 Material Code 月欄位；pivot = Customer/FG weight pivot。"""
    for ri in (2, 4):
        if ri >= len(raw):
            continue
        line = " ".join(_text(v).lower() for v in raw.iloc[ri].tolist()[:10])
        if "material code" in line:
            return "material_code_layout"
        if "customer" in line and "fg code" in line:
            return "pivot"
    return "material_code_layout"


def _parse_sa007_history_sales_style(
    raw: pd.DataFrame,
    source_file: str,
    source_sheet: str = SA007_HISTORY_LABEL,
) -> pd.DataFrame:
    """
    SA007歷史銷售紀錄格式（Material Code + FG + 近三月 kg）。
    資料來源：SA007/ 資料夾；Sample Balance 工作表僅 fallback。
    """
    rows = []
    for i in range(4, len(raw)):
        row = raw.iloc[i]
        fg = row[3]
        if pd.isna(fg) or not _text(fg):
            continue
        rows.append({
            "Material_Code": _text(row[0]),
            "FG_Code": _text(fg),
            "Customer": _text(row[4]),
            "Spec": _text(row[1]),
            "Mat_Spec": _text(row[7]) if len(row) > 7 else _text(row[1]),
            "Thickness": _num(row[8]) if len(row) > 8 else 0.0,
            "Width": _num(row[9]) if len(row) > 9 else 0.0,
            "M-1": _num(row[25]) if len(row) > 25 else 0.0,
            "M-2": _num(row[24]) if len(row) > 24 else 0.0,
            "M-3": _num(row[23]) if len(row) > 23 else 0.0,
            "Source_File": source_file,
            "Source_Sheet": source_sheet,
            "Source_Format": "sa007_history",
        })
    return pd.DataFrame(rows)


def _read_sa007_workbook(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    sheet = xl.sheet_names[0]
    for name in xl.sheet_names:
        upper = name.strip().upper()
        if upper.startswith("SA007") or upper.startswith("SA006"):
            sheet = name
            break
    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    fmt = _detect_sa007_workbook_format(raw)
    if fmt == "pivot":
        return _parse_sa007_sales(raw, path.name)
    return _parse_sa007_history_sales_style(raw, path.name, SA007_HISTORY_LABEL)


def _load_sa007_from_sample_fallback() -> pd.DataFrame:
    """Fallback when SA007/ folder is empty: Sample Balance SA007 pivot or SA007歷史銷售紀錄工作表。"""
    sample_path = resolve_sample_path()
    try:
        sheet = _resolve_sa_sales_sheet(sample_path)
        raw = pd.read_excel(sample_path, sheet_name=sheet, header=None)
        if sheet.strip().upper() == "SA006":
            return _parse_sa006_sheet(raw, sample_path.name)
        return _parse_sa007_sales(raw, sample_path.name)
    except ValueError:
        pass
    try:
        sa007_legacy = pd.read_excel(sample_path, sheet_name=SAMPLE_BALANCE_SA007_SHEET, header=None)
        return _parse_sa007_history_sales_style(sa007_legacy, sample_path.name, SA007_HISTORY_LABEL)
    except Exception:
        return pd.DataFrame()


def _normalize_sa007_sales_to_pivot(df: pd.DataFrame) -> pd.DataFrame:
    """Unify SA007歷史銷售紀錄列為 pivot 欄位以便彙總。"""
    if df.empty:
        return df
    if "M-3_kg" in df.columns:
        return df
    rows = []
    for _, row in df.iterrows():
        m1 = _num(row.get("M-1"))
        m2 = _num(row.get("M-2"))
        m3 = _num(row.get("M-3"))
        if m1 + m2 + m3 <= 0:
            continue
        spec = _text(row.get("Mat_Spec")) or _text(row.get("Spec"))
        rows.append({
            "Customer": _text(row.get("Customer")),
            "Delivery_Place": _text(row.get("Customer")),
            "FG_Code": _text(row.get("FG_Code")),
            "Spec": spec,
            "Thickness": _num(row.get("Thickness")),
            "Width": _num(row.get("Width")),
            "M-3_kg": m3,
            "M-2_kg": m2,
            "M-1_kg": m1,
            "Total_3mo_kg": m1 + m2 + m3,
            "Source_Sheet": _text(row.get("Source_Sheet", SA007_HISTORY_LABEL)),
            "Source_File": _text(row.get("Source_File")),
        })
    return pd.DataFrame(rows)


def load_sa007_history_detail() -> pd.DataFrame:
    """SA007歷史銷售紀錄明細（近三月 kg），資料來源 SA007/ 資料夾。"""
    frames: list[pd.DataFrame] = []
    for path in resolve_sa007_paths():
        try:
            xl = pd.ExcelFile(path)
            sheet = xl.sheet_names[0]
            for name in xl.sheet_names:
                if "SA007" in name.upper() or "SA006" in name.upper():
                    sheet = name
                    break
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            if _detect_sa007_workbook_format(raw) == "material_code_layout":
                part = _parse_sa007_history_sales_style(raw, path.name, SA007_HISTORY_LABEL)
                if not part.empty:
                    frames.append(part)
        except Exception:
            continue
    if frames:
        return exclude_carrier_rows(pd.concat(frames, ignore_index=True))
    sample_path = resolve_sample_path()
    try:
        sa007_legacy = pd.read_excel(sample_path, sheet_name=SAMPLE_BALANCE_SA007_SHEET, header=None)
        return exclude_carrier_rows(
            _parse_sa007_history_sales_style(sa007_legacy, sample_path.name, SA007_HISTORY_LABEL)
        )
    except Exception:
        return pd.DataFrame()


def load_sa007_act_order_detail() -> pd.DataFrame:
    """Deprecated alias for load_sa007_history_detail()."""
    return load_sa007_history_detail()


def load_order_history(sample_path: Path | None = None) -> pd.DataFrame:
    """U7: mean(M-1,M-2,M-3)/1000 ton from SA007/ SA007歷史銷售紀錄。"""
    del sample_path  # legacy arg; SA007 folder is canonical
    detail = load_sa007_history_detail()
    if detail.empty:
        return pd.DataFrame(columns=["FG_Code", "Material_Code", "M-1", "M-2", "M-3"])
    return detail.groupby("FG_Code", as_index=False).agg({
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
            "Source_Sheet": SA007_HISTORY_LABEL,
            "Source_File": source,
        })
    return pd.DataFrame(rows)


def _parse_sa006_sheet(raw: pd.DataFrame, source: str) -> pd.DataFrame:
    """Parse SA006 when layout matches SA007 (Weight months 3–5)."""
    return _parse_sa007_sales(raw, source).assign(Source_Sheet="SA006")


def load_sa007_sales(sa007_dir: Path | None = None) -> pd.DataFrame:
    """
    SA007歷史銷售紀錄（近 3 月）彙總，資料來源 SA007/ 資料夾。
    Sample Balance 工作表僅在 SA007/ 為空時 fallback。
    """
    frames: list[pd.DataFrame] = []
    paths = resolve_sa007_paths() if sa007_dir is None else sorted(
        [p for p in sa007_dir.glob("*.xlsx")] + [p for p in sa007_dir.glob("*.xls")],
        key=lambda p: p.stat().st_mtime,
    )
    for path in paths:
        if is_carrier_path(path):
            continue
        try:
            part = _read_sa007_workbook(path)
            if not part.empty:
                frames.append(part)
        except Exception:
            continue
    if not frames:
        combined = _load_sa007_from_sample_fallback()
    else:
        combined = pd.concat(frames, ignore_index=True)
    combined = _normalize_sa007_sales_to_pivot(combined)
    return exclude_carrier_rows(combined)


def load_sa006_sales(sample_path: Path | None = None) -> pd.DataFrame:
    """Alias: actual sales from SA007/ folder."""
    del sample_path
    return load_sa007_sales()


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
    """Report Mat Bal folder sources used by Supply Plan (Carrier excluded)."""
    from stock_matching import resolve_rules_path

    rules_path = resolve_rules_path()
    rd004_master = resolve_rd004_master_path()
    sa007_paths = resolve_sa007_paths()
    so003_paths = resolve_so003_paths()
    mp008_paths = resolve_mp008_paths()
    fc_dir = resolve_forecast_dir()

    def _folder_files(label: str, folder: Path, paths: list[Path]) -> str:
        if paths:
            return f"{label}/: " + ", ".join(p.name for p in paths)
        return f"{label}/: (empty — 請放入匯出檔)"

    rd004_master_label = (
        f"RD004/{rd004_master.name}"
        if rd004_master
        else "RD004/: (empty — 請放入 Material Master 匯出檔)"
    )
    rules_label = (
        f"RD004/{rules_path.name}"
        if rules_path.exists() and rules_path.parent.resolve() == RD004_DIR.resolve()
        else (rules_path.name if rules_path.exists() else "內建 U-Stock 預設規則")
    )
    return {
        "資料根目錄": f"Mat Bal/ — 各資料夾檔案為 Supply Plan 分析主來源",
        "SA007": f"{_folder_files('SA007', SA007_DIR, sa007_paths)} ({SA007_HISTORY_LABEL})",
        "Forecast": f"{fc_dir.name}/ ({FORECAST_SCOPE_LABEL})",
        "SO003": _folder_files("SO003", SO003_DIR, so003_paths),
        "MP008": _folder_files("MP008", MP008_DIR, mp008_paths),
        "RD004": f"{rd004_master_label} + {rules_label}",
        "Stock": f"Stock/{resolve_stock_path().name}",
        "樣本參考": f"History Balance/、Sample Balance/ — {TEMPLATE_REFERENCE_LABEL}",
        "rd004_diff": (
            "RD004_差異摘要 / RD004_差異_主檔 / RD004_差異_配對規則"
            if rd004_master
            else "N/A (RD004/ 為空)"
        ),
        "scope": "不含 Carrier 客戶（本系統僅分析其他客戶訂單）",
    }
