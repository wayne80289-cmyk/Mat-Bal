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
    BASE_DIR / "stock-material-code-matching-rules.xlsx",
]


def resolve_rules_path() -> Path:
    for path in RULES_CANDIDATES:
        if path.exists():
            return path
    search_dirs = [BASE_DIR, BASE_DIR / "History Balance"]
    for directory in search_dirs:
        if not directory.exists():
            continue
        for path in directory.iterdir():
            if path.suffix.lower() != ".xlsx":
                continue
            name = path.name.lower()
            if "stock" in name and "matching" in name and "rules" in name:
                return path
    return RULES_CANDIDATES[0]


RULES_REFERENCE_SHEETS = (
    "配對規則清單",
    "MAT SPEC對照",
    "Code命名格式",
    "排除情況",
    "配對流程",
)

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

# Customer + Mat Spec + raw coil width → single finished Material Code (no multi-size split)
DEFAULT_CUSTOMER_SPEC_PAIRINGS: list[dict] = [
    {
        "mat_spec": "DC05",
        "thickness": 0.7,
        "raw_widths": (1165.0, 1105.0),
        "customer_keys": ("SSK",),
        "material_code": "DC05_0.7_580_C",
        "product_width": 580.0,
        "fg_code": "0.7x580xC",
        "main_customer": "SSK KOLAKARN",
        "spec": "DC05",
        "common_group": "DC05",
        "note": "SSK 原卷1165/1105僅配成品DC05 0.7x580",
    },
    {
        "mat_spec": "DC05",
        "thickness": 0.7,
        "raw_widths": (1250.0,),
        "customer_keys": ("THAISUMMITGOL", "THAISUMMIT"),
        "material_code": "DC05_0.7_1250_BAO",
        "product_width": 415.0,
        "fg_code": "0.7x415x505",
        "main_customer": "THAI SUMMIT GOL",
        "spec": "DC05",
        "common_group": "DC05",
        "note": "Thai Summit GOL 原卷1250配0.7x415成品",
    },
]


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


def _build_reverse_spec_index(spec_crosswalk: list[dict]) -> dict[str, list[dict]]:
    """Code prefix → MS004 MAT SPEC rows from rules workbook."""
    index: dict[str, list[dict]] = {}
    for row in spec_crosswalk:
        prefixes = re.split(r"\s*或\s*", _norm_spec(row.get("配對_Code前綴", "")))
        for prefix in prefixes:
            key = _compact_spec(prefix)
            if key:
                index.setdefault(key, []).append(row)
    return index


def _prefix_matches_group(prefix_compact: str, spec: str, common_group: str) -> bool:
    for raw in (spec, common_group):
        key = _compact_spec(raw)
        if not key:
            continue
        if key == prefix_compact:
            return True
        if len(prefix_compact) >= 6 and len(key) >= 6:
            if key.startswith(prefix_compact[:6]) or prefix_compact.startswith(key[:6]):
                return True
    return False


def load_matching_rules(rules_path: Path | None = None) -> dict:
    """Load rules from Stock-Material-Code-Matching-Rules.xlsx."""
    rules_path = rules_path or resolve_rules_path()
    spec_map = dict(DEFAULT_SPEC_MAP)
    spec_crosswalk: list[dict] = []
    project_cust = set(DEFAULT_PROJECT_CUST_CODES)
    rules_path_exists = rules_path.exists()

    if rules_path_exists:
        try:
            xl = pd.ExcelFile(rules_path)
            if "MAT SPEC對照" in xl.sheet_names:
                df = pd.read_excel(rules_path, sheet_name="MAT SPEC對照")
                for _, row in df.iterrows():
                    ms = _norm_spec(row.get("MS004_MAT_SPEC", ""))
                    code = _norm_spec(row.get("配對_Code前綴", ""))
                    note = _text(row.get("備註", ""))
                    if ms and code:
                        spec_map[ms] = code.split(" 或 ")[0].strip()
                        spec_crosswalk.append({
                            "MS004_MAT_SPEC": ms,
                            "配對_Code前綴": code,
                            "備註": note,
                        })
            if "排除情況" in xl.sheet_names:
                excl = pd.read_excel(rules_path, sheet_name="排除情況")
                for _, row in excl.iterrows():
                    situation = _text(row.get("情況", ""))
                    m = re.search(r"如\s+([A-Z0-9 /]+)", situation, re.I)
                    if m:
                        for token in re.split(r"[/、,]", m.group(1)):
                            token = _norm_spec(token)
                            if token:
                                project_cust.add(token)
        except Exception:
            pass

    if not spec_crosswalk:
        spec_crosswalk = [
            {"MS004_MAT_SPEC": k, "配對_Code前綴": v, "備註": ""}
            for k, v in DEFAULT_SPEC_MAP.items()
        ]

    return {
        "spec_map": spec_map,
        "spec_crosswalk": spec_crosswalk,
        "reverse_spec_index": _build_reverse_spec_index(spec_crosswalk),
        "project_cust_codes": project_cust,
        "rules_path": rules_path,
        "rules_path_exists": rules_path_exists,
    }


def load_rules_workbook_sheets(rules_path: Path | None = None) -> dict[str, pd.DataFrame]:
    """Load reference sheets from matching rules workbook for report export."""
    rules_path = rules_path or resolve_rules_path()
    sheets: dict[str, pd.DataFrame] = {}
    if not rules_path.exists():
        return sheets
    try:
        xl = pd.ExcelFile(rules_path)
        for name in RULES_REFERENCE_SHEETS:
            if name in xl.sheet_names:
                sheets[name] = pd.read_excel(rules_path, sheet_name=name)
    except Exception:
        pass
    return sheets


def parse_material_code_parts(material_code: str) -> dict:
    """Parse Code命名格式 A/B from Material_Code string."""
    code = _text(material_code)
    if not code:
        return {
            "Code_Format": "",
            "Code_Suffix": "",
            "Parsed_Common_Group": "",
            "Parsed_T": None,
            "Parsed_W": None,
        }

    fmt_b = re.match(r"^(.+)_([\d.]+)_([\d.]+)_0_(.+)$", code, re.I)
    if fmt_b:
        return {
            "Code_Format": "格式B",
            "Code_Suffix": _text(fmt_b.group(4)),
            "Parsed_Common_Group": _text(fmt_b.group(1)),
            "Parsed_T": _num(fmt_b.group(2)),
            "Parsed_W": _num(fmt_b.group(3)),
        }

    fmt_a = re.match(r"^(.+)_([\d.]+)_([\d.]+)_(.+)$", code, re.I)
    if fmt_a:
        return {
            "Code_Format": "格式A",
            "Code_Suffix": _text(fmt_a.group(4)),
            "Parsed_Common_Group": _text(fmt_a.group(1)),
            "Parsed_T": _num(fmt_a.group(2)),
            "Parsed_W": _num(fmt_a.group(3)),
        }

    return {
        "Code_Format": "其他",
        "Code_Suffix": "",
        "Parsed_Common_Group": code,
        "Parsed_T": None,
        "Parsed_W": None,
    }


def classify_pool_type(code_suffix: str) -> tuple[str, str]:
    """Return (Pool_Type, Stock_Rule_Ref) from code tail."""
    suffix = _norm_spec(code_suffix).upper().replace(" ", "")
    if suffix == "COMMON" or suffix.endswith("COMMON"):
        return "共通池", "U-Stock-05"
    if "CASH" in suffix:
        return "現金採購", "U-Stock-07"
    if suffix:
        return "Maker", "U-Stock-06"
    return "其他", "U-Stock-03"


def ms004_specs_for_material(spec: str, common_group: str, rules: dict) -> tuple[str, str]:
    """MS004 MAT SPECs and notes that map to this material group."""
    hits: list[str] = []
    notes: list[str] = []
    for prefix, rows in rules.get("reverse_spec_index", {}).items():
        if not _prefix_matches_group(prefix, spec, common_group):
            continue
        for row in rows:
            ms = row.get("MS004_MAT_SPEC", "")
            if ms:
                hits.append(ms)
            note = row.get("備註", "")
            if note:
                notes.append(note)
    return ", ".join(sorted(set(hits))), "; ".join(sorted(set(notes)))


def width_match_policy(width: float) -> str:
    if round(width, 0) == 1219:
        return "1219群組忽略寬度 (U-Stock-09)"
    return "精確比對寬度"


def mp008_alloc_hint(pool_type: str, main_customer: str) -> str:
    if pool_type == "共通池":
        return "MP008共通池調撥"
    if main_customer:
        return "MP008同客戶優先"
    return "MP008規格配對"


def enrich_material_master(
    rd004_df: pd.DataFrame,
    rules: dict | None = None,
    allocatable_stock: pd.DataFrame | None = None,
    allocated_po: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """
    Merge RD004 master with Stock-Material-Code-Matching-Rules fields
    and current-run allocatable stock / inbound totals.
    """
    if rd004_df.empty:
        return rd004_df

    rules = rules or load_matching_rules()
    out = rd004_df.copy()

    stock_map: dict[str, float] = {}
    if allocatable_stock is not None and not allocatable_stock.empty:
        stock_map = (
            allocatable_stock.groupby("Material_Code")["Quantity"]
            .sum()
            .round(3)
            .to_dict()
        )

    po_total_map: dict[str, float] = {}
    po_customer_map: dict[str, float] = {}
    if allocated_po is not None and not allocated_po.empty:
        for _, row in allocated_po.iterrows():
            mat = _text(row.get("Material_Code", ""))
            if not mat:
                continue
            qty = _num(row.get("Quantity", 0))
            po_total_map[mat] = po_total_map.get(mat, 0.0) + qty
            if row.get("Alloc_Priority") == "Customer-Match":
                po_customer_map[mat] = po_customer_map.get(mat, 0.0) + qty

    rules_file = rules["rules_path"].name if rules.get("rules_path_exists") else "(embedded defaults)"

    enriched_rows = []
    for _, row in out.iterrows():
        rec = row.to_dict()
        code = _text(rec.get("Material_Code", ""))
        spec = _text(rec.get("Spec", rec.get("Common_Group", "")))
        common_group = _text(rec.get("Common_Group", spec))
        t = _num(rec.get("Thickness"))
        w = _num(rec.get("Width")) or 1219.0
        main_customer = _text(rec.get("Main_Customer", ""))

        parts = parse_material_code_parts(code)
        pool_type, suffix_rule = classify_pool_type(parts["Code_Suffix"])
        ms004_specs, ms004_notes = ms004_specs_for_material(spec, common_group, rules)

        rule_refs = ["U-Stock-02", "U-Stock-03", "U-Stock-08", suffix_rule, "U-Stock-12"]
        if pool_type == "共通池":
            rule_refs.append("U-Stock-09")

        rec.update({
            "Rules_Source": rules_file,
            "Code_Format": parts["Code_Format"],
            "Code_Suffix": parts["Code_Suffix"],
            "Pool_Type": pool_type,
            "Match_Key": f"{common_group}|{round(t, 1):g}|{round(w, 0):.0f}",
            "MS004_MAT_SPECS": ms004_specs,
            "MS004_SPEC_備註": ms004_notes,
            "Width_Match_Policy": width_match_policy(w),
            "Thickness_Tolerance": "round(T,1)",
            "Stock_Rule_Refs": ";".join(dict.fromkeys(rule_refs)),
            "MP008_Alloc_Hint": mp008_alloc_hint(pool_type, main_customer),
            "Allocatable_Stock_Ton": round(stock_map.get(code, 0.0), 3),
            "Inbound_PO_Total_Ton": round(po_total_map.get(code, 0.0), 3),
            "Inbound_PO_同客戶_Ton": round(po_customer_map.get(code, 0.0), 3),
        })
        enriched_rows.append(rec)

    master = pd.DataFrame(enriched_rows)

    # Column order: RD004 base fields first, then rules integration
    base_cols = [c for c in rd004_df.columns if c in master.columns]
    rule_cols = [
        "Rules_Source", "Code_Format", "Code_Suffix", "Pool_Type", "Match_Key",
        "MS004_MAT_SPECS", "MS004_SPEC_備註", "Width_Match_Policy", "Thickness_Tolerance",
        "Stock_Rule_Refs", "MP008_Alloc_Hint",
        "Allocatable_Stock_Ton", "Inbound_PO_Total_Ton", "Inbound_PO_同客戶_Ton",
    ]
    tail_cols = [c for c in master.columns if c not in base_cols and c not in rule_cols]
    return master[base_cols + rule_cols + tail_cols]


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


def customer_keys_match(customer: str, cust_code: str, keys: tuple[str, ...]) -> bool:
    combined = _compact_spec(f"{customer} {cust_code}")
    for key in keys:
        k = _compact_spec(key)
        if k and k in combined:
            return True
    return False


def load_customer_spec_pairings(rules_path: Path | None = None) -> list[dict]:
    """Load exclusive customer/spec pairings from rules workbook or defaults."""
    pairings = [dict(p) for p in DEFAULT_CUSTOMER_SPEC_PAIRINGS]
    rules_path = rules_path or resolve_rules_path()
    if not rules_path.exists():
        return pairings
    try:
        xl = pd.ExcelFile(rules_path)
        sheet = next((s for s in xl.sheet_names if "客戶專屬" in s or "專屬配對" in s), None)
        if not sheet:
            return pairings
        df = pd.read_excel(rules_path, sheet_name=sheet)
        loaded: list[dict] = []
        for _, row in df.iterrows():
            spec = _text(row.get("Mat_Spec", row.get("MAT_SPEC", "")))
            code = _text(row.get("Material_Code", ""))
            if not spec or not code:
                continue
            raw_ws = row.get("原料寬度_W", row.get("Raw_Widths", ""))
            if isinstance(raw_ws, str):
                widths = tuple(_num(x) for x in re.split(r"[,;/\s]+", raw_ws) if _text(x))
            elif isinstance(raw_ws, (list, tuple)):
                widths = tuple(_num(x) for x in raw_ws)
            else:
                widths = (float(raw_ws),) if pd.notna(raw_ws) and _num(raw_ws) > 0 else ()
            cust_keys = _text(row.get("客戶關鍵字", row.get("Customer_Keys", "")))
            keys = tuple(k.strip() for k in re.split(r"[,;/]", cust_keys) if k.strip())
            loaded.append({
                "mat_spec": spec,
                "thickness": _num(row.get("厚度_T", row.get("Thickness", 0))),
                "raw_widths": widths,
                "customer_keys": keys,
                "material_code": code,
                "product_width": _num(row.get("成品寬_W", row.get("Product_W", 0))),
                "fg_code": _text(row.get("FG_Code", "")),
                "main_customer": _text(row.get("Main_Customer", "")),
                "spec": spec,
                "common_group": _text(row.get("Common_Group", spec)),
                "note": _text(row.get("備註", "")),
            })
        return loaded if loaded else pairings
    except Exception:
        return pairings


def match_exclusive_customer_spec_pairing(
    mat_spec: str,
    thickness: float,
    raw_width: float,
    customer: str = "",
    cust_code: str = "",
    pairings: list[dict] | None = None,
) -> dict | None:
    """
    Exclusive 1:1 pairing: e.g. SSK DC05 0.7x1165 → DC05_0.7_580_C only.
    """
    pairings = pairings or load_customer_spec_pairings()
    for rule in pairings:
        if _norm_spec(mat_spec) != _norm_spec(rule.get("mat_spec", "")):
            continue
        if not thickness_matches(thickness, _num(rule.get("thickness"))):
            continue
        raw_ws = rule.get("raw_widths") or ()
        if raw_ws and not any(round(raw_width, 0) == round(rw, 0) for rw in raw_ws):
            continue
        if not customer_keys_match(customer, cust_code, tuple(rule.get("customer_keys", ()))):
            continue
        return rule
    return None


def customer_matches(a: str, b: str) -> bool:
    a_c = _compact_spec(a)
    b_c = _compact_spec(b)
    if not a_c or not b_c:
        return False
    return a_c == b_c or a_c in b_c or b_c in a_c


def supplement_rd004_pairing_materials(rd004: pd.DataFrame, pairings: list[dict] | None = None) -> pd.DataFrame:
    """Add Material Codes from customer-spec pairing rules if missing from RD004."""
    pairings = pairings or load_customer_spec_pairings()
    if rd004.empty:
        existing: set[str] = set()
    else:
        existing = set(rd004["Material_Code"].astype(str))
    rows = []
    for rule in pairings:
        code = _text(rule.get("material_code", ""))
        if not code or code in existing:
            continue
        pw = _num(rule.get("product_width"))
        rows.append({
            "Material_Code": code,
            "Common_Group": rule.get("common_group", rule.get("spec", "")),
            "Spec": rule.get("spec", ""),
            "Thickness": _num(rule.get("thickness")),
            "Width": pw if pw > 0 else 1219.0,
            "FG_Code": rule.get("fg_code", code),
            "FG_Codes_All": rule.get("fg_code", code),
            "Kind": "CR",
            "Main_Customer": rule.get("main_customer", ""),
            "MOQ": 0.0,
        })
        existing.add(code)
    if not rows:
        return rd004
    return pd.concat([rd004, pd.DataFrame(rows)], ignore_index=True)


def pairing_blocks_coil_split(
    mat_spec: str,
    thickness: float,
    raw_width: float,
    customer: str = "",
    cust_code: str = "",
    pairings: list[dict] | None = None,
) -> bool:
    """If exclusive pairing applies, do not split across other finished widths."""
    return match_exclusive_customer_spec_pairing(
        mat_spec, thickness, raw_width, customer, cust_code, pairings
    ) is not None

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

        cust_code = _text(row.get("CUST CODE", row.get("Cust Code", "")))
        exclusive = match_exclusive_customer_spec_pairing(spec, t, w, "", cust_code)
        if exclusive:
            records.append({
                "Material_Code": exclusive["material_code"],
                "Common_Group": exclusive.get("common_group", spec),
                "Quantity": qty_ton,
                "Mat_Spec": spec,
                "T": t,
                "W": w,
                "CUST_CODE": cust_code,
                "MAKER_CODE": _text(row.get("MAKER CODE", row.get("Maker Code", ""))),
                "Allocatable": True,
                "Match_Rule": "Customer-Spec-Pair",
                "Product_W": exclusive.get("product_width"),
            })
            continue

        matches = find_rd004_matches(spec, t, w, rd004, rules)
        if not matches.empty:
            mat_code = _pick_material_for_stock(
                matches,
                row.get("CUST CODE", row.get("Cust Code", "")),
                row.get("MAKER CODE", row.get("Maker Code", "")),
            )
            if mat_code:
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
            continue

        split_targets = find_coil_split_targets(spec, t, w, rd004, rules, cust_code=cust_code)
        if split_targets:
            for m, product_w in split_targets:
                ratio = product_w / w
                records.append({
                    "Material_Code": m["Material_Code"],
                    "Common_Group": m.get("Common_Group", m.get("Spec", "")),
                    "Quantity": round(qty_ton * ratio, 6),
                    "Mat_Spec": spec,
                    "T": t,
                    "W": w,
                    "CUST_CODE": _text(row.get("CUST CODE", row.get("Cust Code", ""))),
                    "MAKER_CODE": _text(row.get("MAKER CODE", row.get("Maker Code", ""))),
                    "Allocatable": True,
                    "Match_Rule": "Coil-Split",
                    "Raw_Coil_W": w,
                    "Product_W": product_w,
                    "Split_Ratio": round(ratio, 6),
                })
            continue

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


def material_row_matches_spec(mat_spec: str, m: pd.Series, rules: dict | None = None) -> bool:
    """Match MS004/MP008 spec against RD004 Spec, Common_Group, or Material_Code."""
    rules = rules or load_matching_rules()
    for field in ("Spec", "Common_Group", "Material_Code"):
        val = _text(m.get(field, ""))
        if val and spec_group_matches(mat_spec, val, rules):
            return True
    return False


def parse_product_width_from_fg_code(fg_code: str, thickness: float | None = None) -> float | None:
    """
    Finished width from FG Code dimensions (when Material Code has no product width).
    e.g. 1.6x61x1219 → 61, 0.7x580xC → 580
    """
    fg = _text(fg_code).replace(" ", "")
    if not fg:
        return None

    m = re.match(r"^([\d.]+)[tT]?[xX]([\d.]+)(?:[xX](.+))?$", fg, re.I)
    if not m:
        return None

    t_val = _num(m.group(1))
    w_val = _num(m.group(2))
    if thickness is not None and thickness > 0:
        if round(t_val, 1) != round(thickness, 1):
            return None
    if w_val <= 0 or round(w_val, 0) == 1219:
        return None
    return w_val


def resolve_product_width(m: pd.Series, thickness: float) -> float | None:
    """
    Resolve finished product width for coil-split from Material Code, FG Code, or RD004 Width.
    Used when FG is dimension-based or Material Code omits finished width.
    """
    mat_code = _text(m.get("Material_Code", ""))

    pw = parse_product_width_from_material_code(mat_code, thickness)
    if pw is not None and pw > 0:
        return pw

    for fg_field in ("FG_Code",):
        pw = parse_product_width_from_fg_code(m.get(fg_field, ""), thickness)
        if pw is not None and pw > 0:
            return pw

    for fg in _text(m.get("FG_Codes_All", "")).split(","):
        pw = parse_product_width_from_fg_code(fg.strip(), thickness)
        if pw is not None and pw > 0:
            return pw

    rd_w = _num(m.get("Width"))
    if rd_w > 0 and round(rd_w, 0) != 1219:
        return rd_w

    return None


def parse_product_width_from_material_code(material_code: str, thickness: float | None = None) -> float | None:
    """
    Finished product width from Material Code.
    - Format A underscore: SPHC-P/O_2.6_1086_CASH → 1086
    - Embedded TxW: BUSDE+Z-CSG+0 0/50_1.0x440_____ → 440
    Skips nominal coil width 1219.
    """
    parts = parse_material_code_parts(material_code)
    if parts["Code_Format"] == "格式A":
        pw = parts.get("Parsed_W")
        if pw is not None and pw > 0 and round(pw, 0) != 1219:
            if thickness is None or thickness <= 0 or parts.get("Parsed_T") is None:
                return pw
            if round(parts["Parsed_T"], 1) == round(thickness, 1):
                return pw

    code = _text(material_code)
    for m in re.finditer(r"([\d.]+)[xX]([\d.]+)", code):
        t_val = _num(m.group(1))
        w_val = _num(m.group(2))
        if w_val <= 0 or round(w_val, 0) == 1219:
            continue
        if thickness is not None and thickness > 0:
            if round(t_val, 1) != round(thickness, 1):
                continue
        return w_val
    return None


def find_coil_split_targets(
    mat_spec: str,
    thickness: float,
    raw_coil_width: float,
    rd004: pd.DataFrame,
    rules: dict | None = None,
    customer: str = "",
    cust_code: str = "",
) -> list[tuple[pd.Series, float]]:
    """
    Proportional split when:
    - MS004/MP008 Mat Spec matches RD004 spec group (even if FG Code is non-dimensional in Material Code)
    - Finished widths resolved from FG Code dimensions and/or RD004 Width field
    - Raw coil width > sum of multiple finished product widths (at least 2)
    Weight per code = (product_W / raw_coil_W) * coil weight
    """
    if raw_coil_width <= 0:
        return []
    if pairing_blocks_coil_split(mat_spec, thickness, raw_coil_width, customer, cust_code):
        return []

    rules = rules or load_matching_rules()
    targets: list[tuple[pd.Series, float]] = []
    seen_codes: set[str] = set()

    for _, m in rd004.iterrows():
        if not material_row_matches_spec(mat_spec, m, rules):
            continue
        if not thickness_matches(thickness, _num(m.get("Thickness"))):
            continue

        mat_code = _text(m.get("Material_Code", ""))
        if not mat_code or mat_code in seen_codes:
            continue

        product_w = resolve_product_width(m, thickness)
        if product_w is None or product_w <= 0:
            continue
        if round(product_w, 0) == round(raw_coil_width, 0):
            continue
        if product_w >= raw_coil_width:
            continue

        seen_codes.add(mat_code)
        targets.append((m, product_w))

    if len(targets) < 2:
        return []

    sum_product_w = sum(pw for _, pw in targets)
    if raw_coil_width <= sum_product_w:
        return []

    targets.sort(key=lambda x: x[1])
    return targets


def _mp008_alloc_row(row: pd.Series, mat_code: str, qty_ton: float, priority: str, **extra) -> dict:
    base = {
        "Po No.": _text(row.get("Po No.", "")),
        "Customer": _text(row.get("Customer", "")),
        "Mat Spec": _text(row.get("Mat Spec", "")),
        "T": _num(row.get("T")),
        "W": _num(row.get("W")),
        "ETA": row.get("ETA"),
        "ETA_Month": row.get("ETA_Month", ""),
        "Material_Code": mat_code,
        "Quantity": round(qty_ton, 6),
        "Alloc_Priority": priority,
        "Close Flag": row.get("Close Flag", False),
        "PO Balance WT": _num(row.get("PO Balance WT", qty_ton * 1000)),
    }
    base.update(extra)
    return base


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

        po_customer = _text(row.get("Customer", ""))
        exclusive = match_exclusive_customer_spec_pairing(spec, t, w, po_customer, "")
        if exclusive:
            rows.append(_mp008_alloc_row(
                row,
                exclusive["material_code"],
                qty_ton,
                "Customer-Spec-Pair",
                Product_W=exclusive.get("product_width"),
            ))
            continue

        matches = find_rd004_matches(spec, t, w, rd004, rules)
        mat_code = _text(row.get("Material_Code", ""))
        priority = "Pre-matched"

        if not mat_code and not matches.empty:
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

        if mat_code:
            rows.append(_mp008_alloc_row(row, mat_code, qty_ton, priority))
            continue

        split_targets = find_coil_split_targets(spec, t, w, rd004, rules, customer=po_customer)
        if split_targets:
            for m, product_w in split_targets:
                ratio = product_w / w
                rows.append(_mp008_alloc_row(
                    row,
                    m["Material_Code"],
                    qty_ton * ratio,
                    "Coil-Split",
                    Raw_Coil_W=w,
                    Product_W=product_w,
                    Split_Ratio=round(ratio, 6),
                    Product_W_Sum=sum(pw for _, pw in split_targets),
                ))
            continue

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
