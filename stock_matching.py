# -*- coding: utf-8 -*-
"""Stock Material Code matching (U-Stock-01..15) for MS004 and MP008 allocation."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
RULES_CANDIDATES = [
    BASE_DIR / "History Balance" / "Stock-Material-Code-Matching-Rules.xlsx",
    BASE_DIR / "Stock-Material-Code-Matching-Rules.xlsx",
]


def resolve_rules_path() -> Path:
    for path in RULES_CANDIDATES:
        if path.exists():
            return path
    return RULES_CANDIDATES[0]

# Embedded defaults when rules workbook is absent (from export_matching_rules_excel.py)
DEFAULT_SPEC_MAP: dict[str, str] = {
    "SPCC": "SPCC-SD",
    "SPCC-M": "SPCC-SD",
    "SGCC-Z08": "SGCC-Z22",
    "SGCD2-Z08": "SGCC-Z22",
    "SGCD1-Z18": "SGCC-Z18",
    "SAPH440-P/O": "SAPH440-PO",
    "SAPH400-P/O": "SAPH440-P/O",
    "JSH590R-P/O": "SPHC-PO",
    "FC440": "JSC440W",
    "SPFC440": "JSC440W",
    "SECC-AF,E16/E16": "SECC-16/16",
    "SPCD": "SPCC-SD",
    "SPCEN-SD": "SPCC-SD",
}

DEFAULT_PROJECT_CUST_CODES = frozenset({"MING TAI", "MINGTAI"})


def _text(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v).strip())


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


def _norm_spec(value: str) -> str:
    return _text(value).upper().replace("?", "")


def _compact_spec(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", _norm_spec(value))


def load_matching_rules(rules_path: Path | None = None) -> dict:
    """Load MAT SPEC map and exclusion rules from Stock-Material-Code-Matching-Rules.xlsx."""
    rules_path = rules_path or RULES_PATH
    spec_map = dict(DEFAULT_SPEC_MAP)
    project_cust = set(DEFAULT_PROJECT_CUST_CODES)

    if rules_path.exists():
        try:
            xl = pd.ExcelFile(rules_path)
            if "MAT SPEC對照" in xl.sheet_names:
                df = pd.read_excel(rules_path, sheet_name="MAT SPEC對照")
                for _, row in df.iterrows():
                    ms = _norm_spec(row.get("MS004_MAT_SPEC", ""))
                    code = _norm_spec(row.get("配對_Code前綴", ""))
                    if ms and code:
                        spec_map[ms] = code.split(" 或 ")[0].strip()
        except Exception:
            pass

    return {"spec_map": spec_map, "project_cust_codes": project_cust}


def map_ms004_spec(mat_spec: str, rules: dict | None = None) -> str:
    rules = rules or load_matching_rules()
    raw = _norm_spec(mat_spec)
    mapped = rules["spec_map"].get(raw, raw)
    return _compact_spec(mapped)


def thickness_matches(stock_t: float, target_t: float) -> bool:
    """U-Stock-08: round(T,1) tolerance."""
    if stock_t <= 0 or target_t <= 0:
        return False
    return round(stock_t, 1) == round(target_t, 1)


def width_matches(stock_w: float, target_w: float) -> bool:
    """U-Stock-09: coil common width 1219 group ignores actual W variation."""
    if target_w <= 0:
        return True
    if round(target_w, 0) == 1219:
        return True
    if stock_w <= 0:
        return True
    return round(stock_w, 0) == round(target_w, 0)


def spec_group_matches(stock_spec: str, rd004_spec: str, rules: dict | None = None) -> bool:
    """U-Stock-03/04: Common Group via MAT SPEC crosswalk + prefix compare."""
    stock_key = map_ms004_spec(stock_spec, rules)
    target_key = _compact_spec(rd004_spec)
    if not stock_key or not target_key:
        return False
    if stock_key == target_key:
        return True
    return stock_key.startswith(target_key[:6]) or target_key.startswith(stock_key[:6])


def customer_matches(a: str, b: str) -> bool:
    a_c = _compact_spec(a)
    b_c = _compact_spec(b)
    if not a_c or not b_c:
        return False
    return a_c == b_c or a_c in b_c or b_c in a_c


def is_customer_exclusive(cust_code: str, rules: dict | None = None) -> bool:
    """U-Stock-10: project-specific CUST CODE rows are not allocatable."""
    rules = rules or load_matching_rules()
    code = _norm_spec(cust_code)
    if not code:
        return False
    for proj in rules["project_cust_codes"]:
        if proj in code or code in proj:
            return True
    return False


def is_common_material_code(material_code: str) -> bool:
    return "_COMMON" in _norm_spec(material_code).upper()


def find_rd004_matches(
    mat_spec: str,
    thickness: float,
    width: float,
    rd004: pd.DataFrame,
    rules: dict | None = None,
) -> pd.DataFrame:
    rules = rules or load_matching_rules()
    hits = []
    for _, m in rd004.iterrows():
        if not spec_group_matches(mat_spec, m.get("Spec", m.get("Common_Group", "")), rules):
            continue
        if not thickness_matches(thickness, _num(m.get("Thickness"))):
            continue
        if not width_matches(width, _num(m.get("Width"))):
            continue
        hits.append(m)
    if not hits:
        return pd.DataFrame()
    return pd.DataFrame(hits)


def is_allocatable_ms004_row(row: pd.Series, rules: dict | None = None) -> bool:
    """U-Stock-01/10/14: PTT, weight>0, not customer-exclusive."""
    rules = rules or load_matching_rules()
    owner = _norm_spec(row.get("OWNER", row.get("Owner", "PTT")))
    if owner and owner != "PTT":
        return False
    qty_kg = _num(row.get("P_REMAIN_WT_kg", row.get("P REMAIN WT", 0)))
    if qty_kg <= 0:
        qty_kg = _num(row.get("REMAIN WT", row.get("Quantity", 0))) * (
            1000.0 if _num(row.get("Quantity", 0)) < 500 else 1.0
        )
    if qty_kg <= 0:
        return False
    cust = row.get("CUST CODE", row.get("Cust Code", row.get("CUST_CODE", "")))
    if is_customer_exclusive(cust, rules):
        return False
    return True


def _ms004_row_weight_ton(row: pd.Series) -> float:
    p_wt = _num(row.get("P_REMAIN_WT_kg", row.get("P REMAIN WT", 0)))
    if p_wt > 0:
        return p_wt / 1000.0
    return _num(row.get("Quantity", 0))


def allocate_ms004_stock(ms004_df: pd.DataFrame, rd004: pd.DataFrame, rules: dict | None = None) -> pd.DataFrame:
    """
    Build allocatable spot stock per U-Stock rules.
    Returns detail rows + aggregated Quantity per Material_Code.
    """
    rules = rules or load_matching_rules()
    if ms004_df.empty:
        return pd.DataFrame(columns=[
            "Material_Code", "Common_Group", "Quantity", "Mat_Spec", "T", "W",
            "CUST_CODE", "Allocatable", "Match_Rule",
        ])

    records = []
    for _, row in ms004_df.iterrows():
        if not is_allocatable_ms004_row(row, rules):
            continue

        spec = _text(row.get("Spec", row.get("MAT SPEC", "")))
        t = _num(row.get("Thickness", row.get("T")))
        w = _num(row.get("Width", row.get("W")))
        qty_ton = _ms004_row_weight_ton(row)
        if qty_ton <= 0:
            continue

        matches = find_rd004_matches(spec, t, w, rd004, rules)
        if matches.empty:
            continue

        mat_code = _pick_material_for_stock(
            matches,
            row.get("CUST CODE", row.get("Cust Code", "")),
            row.get("MAKER CODE", row.get("Maker Code", "")),
        )
        if not mat_code:
            continue

        hit = matches[matches["Material_Code"] == mat_code].iloc[0]
        records.append({
            "Material_Code": mat_code,
            "Common_Group": hit.get("Common_Group", hit.get("Spec", "")),
            "Quantity": qty_ton,
            "Mat_Spec": spec,
            "T": t,
            "W": w,
            "CUST_CODE": _text(row.get("CUST CODE", row.get("Cust Code", ""))),
            "MAKER_CODE": _text(row.get("MAKER CODE", row.get("Maker Code", ""))),
            "Allocatable": True,
            "Match_Rule": "U-Stock-03",
        })

    if not records:
        return pd.DataFrame(columns=[
            "Material_Code", "Common_Group", "Quantity", "Mat_Spec", "T", "W",
            "CUST_CODE", "Allocatable", "Match_Rule",
        ])

    detail = pd.DataFrame(records)
    # U-Stock-12: same Material Code sum weight
    summary = (
        detail.groupby(["Material_Code", "Common_Group"], as_index=False)["Quantity"]
        .sum()
        .round(6)
    )
    summary["Allocatable"] = True
    summary["Match_Rule"] = "U-Stock-12"
    return summary


def filter_open_mp008(mp008_df: pd.DataFrame) -> pd.DataFrame:
    """U4: Close Flag=False / not received."""
    if mp008_df.empty:
        return mp008_df
    out = mp008_df.copy()
    if "Close Flag" in out.columns:
        out = out[out["Close Flag"].astype(str).str.upper().isin(["FALSE", "0", "", "NAN"]) | out["Close Flag"].isna()]
    if "Status" in out.columns:
        out = out[out["Status"].astype(str).str.upper() != "RECEIVED"]
    if "PO Balance WT" in out.columns:
        out = out[_num_series(out["PO Balance WT"]) > 0]
    return out


def _num_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _pick_material_for_stock(
    candidates: pd.DataFrame,
    cust_code: str,
    maker_code: str,
) -> str:
    """One MS004 row → one Material Code (manual allocate behaviour)."""
    if candidates.empty:
        return ""
    cust = _text(cust_code)
    if cust:
        for _, m in candidates.iterrows():
            if customer_matches(cust, m.get("Main_Customer", "")):
                return m["Material_Code"]
    common = candidates[candidates["Material_Code"].map(is_common_material_code)]
    if not common.empty:
        return common.iloc[0]["Material_Code"]
    maker = _norm_spec(maker_code)
    if maker:
        for _, m in candidates.iterrows():
            if maker in _norm_spec(m["Material_Code"]):
                return m["Material_Code"]
    return candidates.iloc[0]["Material_Code"]


def _pick_material_for_po(candidates: pd.DataFrame, po_customer: str) -> str:
    """Customer-first allocation (same customer name), then Common pool."""
    if candidates.empty:
        return ""
    cust = _text(po_customer)
    if cust:
        for _, m in candidates.iterrows():
            main_c = _text(m.get("Main_Customer", ""))
            if customer_matches(cust, main_c):
                return m["Material_Code"]
    common = candidates[candidates["Material_Code"].map(is_common_material_code)]
    if not common.empty:
        return common.iloc[0]["Material_Code"]
    return candidates.iloc[0]["Material_Code"]


def allocate_mp008_inbound(
    mp008_df: pd.DataFrame,
    rd004: pd.DataFrame,
    rules: dict | None = None,
) -> pd.DataFrame:
    """
    Assign open MP008 lines to Material_Code by Mat Spec + T + W.
    Same-customer PO lines are prioritized to matching Main_Customer materials.
    """
    rules = rules or load_matching_rules()
    open_po = filter_open_mp008(mp008_df)
    if open_po.empty:
        return pd.DataFrame(columns=[
            "Po No.", "Customer", "Mat Spec", "T", "W", "ETA", "ETA_Month",
            "Material_Code", "Quantity", "Alloc_Priority", "Close Flag",
        ])

    rows = []
    for _, row in open_po.iterrows():
        spec = _text(row.get("Mat Spec", ""))
        t = _num(row.get("T"))
        w = _num(row.get("W"))
        qty_ton = _num(row.get("Quantity"))
        if qty_ton <= 0 and "PO Balance WT" in row.index:
            qty_ton = _num(row["PO Balance WT"]) / 1000.0
        if qty_ton <= 0:
            continue

        matches = find_rd004_matches(spec, t, w, rd004, rules)
        mat_code = _text(row.get("Material_Code", ""))
        po_customer = _text(row.get("Customer", ""))
        priority = "Pre-matched"
        if not mat_code:
            if matches.empty:
                continue
            mat_code = _pick_material_for_po(matches, po_customer)
            main_c = ""
            hit = matches[matches["Material_Code"] == mat_code]
            if not hit.empty:
                main_c = _text(hit.iloc[0].get("Main_Customer", ""))
            if customer_matches(po_customer, main_c):
                priority = "Customer-Match"
            elif is_common_material_code(mat_code):
                priority = "Common-Pool"
            else:
                priority = "Spec-Match"

        if not mat_code:
            continue

        rows.append({
            "Po No.": _text(row.get("Po No.", "")),
            "Customer": _text(row.get("Customer", "")),
            "Mat Spec": spec,
            "T": t,
            "W": w,
            "ETA": row.get("ETA"),
            "ETA_Month": row.get("ETA_Month", ""),
            "Material_Code": mat_code,
            "Quantity": qty_ton,
            "Alloc_Priority": priority,
            "Close Flag": row.get("Close Flag", False),
            "PO Balance WT": _num(row.get("PO Balance WT", qty_ton * 1000)),
        })

    if not rows:
        return pd.DataFrame(columns=[
            "Po No.", "Customer", "Mat Spec", "T", "W", "ETA", "ETA_Month",
            "Material_Code", "Quantity", "Alloc_Priority", "Close Flag",
        ])
    return pd.DataFrame(rows)


def inbound_by_material_month(allocated_po: pd.DataFrame) -> dict[tuple[str, str], float]:
    """{(Material_Code, YYYY-MM): ton} for balance engine."""
    pool: dict[tuple[str, str], float] = {}
    if allocated_po.empty:
        return pool
    for _, row in allocated_po.iterrows():
        key = (_text(row["Material_Code"]), _text(row.get("ETA_Month", "")))
        if not key[0] or not key[1]:
            continue
        pool[key] = pool.get(key, 0.0) + _num(row["Quantity"])
    return pool
