# -*- coding: utf-8 -*-
"""Stock Material Code matching (U-Stock-01..15) for MS004 and MP008 allocation."""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
RD004_DIR = BASE_DIR / "RD004"
RULES_CANDIDATES = [
    RD004_DIR / "Stock-Material-Code-Matching-Rules.xlsx",
    BASE_DIR / "Stock-Material-Code-Matching-Rules.xlsx",
    BASE_DIR / "stock-material-code-matching-rules.xlsx",
]


def resolve_rules_path() -> Path:
    if RD004_DIR.is_dir():
        rd004_rules = RD004_DIR / "Stock-Material-Code-Matching-Rules.xlsx"
        if rd004_rules.exists():
            return rd004_rules
    for path in RULES_CANDIDATES:
        if path.exists():
            return path
    search_dirs = [RD004_DIR, BASE_DIR]
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


def resolve_baseline_rules_path() -> Path | None:
    """Original pairing rules workbook outside RD004/ (pre-folder layout)."""
    for path in RULES_CANDIDATES:
        if path.parent.resolve() == RD004_DIR.resolve():
            continue
        if path.exists():
            return path
    for directory in [BASE_DIR]:
        if not directory.exists():
            continue
        for path in directory.iterdir():
            if path.suffix.lower() != ".xlsx":
                continue
            name = path.name.lower()
            if "stock" in name and "matching" in name and "rules" in name:
                return path
    return None


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
    "SPCC-4D": "SPCC-SD",
    "SPCC-SD CQ1 GP": "SPCC-SD",
    "SPCD": "SPCC-SD",
    "SPCD-SD": "SPCC-SD",
    "SPCEN-SD": "SPCC-SD",
    "SGCC-Z08": "SGCC-Z22",
    "SGCD2-Z08": "SGCC-Z22",
    "SGCD1-Z18": "SGCC-Z18",
    "SGCD1-F08": "SGCD",
    "SAPH440-P/O": "SAPH440-PO",
    "SAPH400-P/O": "SAPH440-P/O",
    "JSH590R-P/O": "SPHC-PO",
    "FC440": "JSC440W",
    "SPFC440": "JSC440W",
    "SECC-AF,E16/E16": "SECC-16/16",
}

DEFAULT_PROJECT_CUST_CODES = frozenset({"MING TAI", "MINGTAI"})

# Customer + Mat Spec + raw coil width → single finished Material Code (no multi-size split)
DEFAULT_CUSTOMER_SPEC_PAIRINGS: list[dict] = [
    {
        "mat_spec": "DC05",
        "thickness": 0.7,
        "raw_widths": (1165.0,),
        "customer_keys": ("SSK",),
        "material_code": "DC05_0.7_580_C",
        "product_width": 580.0,
        "fg_code": "0.7x580xC",
        "main_customer": "SSK KOLAKARN",
        "spec": "DC05",
        "common_group": "DC05",
        "note": "SSK 原卷1165僅配成品DC05 0.7x580",
    },
    {
        "mat_spec": "DC05",
        "thickness": 0.7,
        "raw_widths": (1105.0,),
        "customer_keys": ("SSK",),
        "material_code": "DC05_0.7_550_C",
        "product_width": 550.0,
        "fg_code": "0.7x550xC",
        "main_customer": "SSK KOLAKARN",
        "spec": "DC05",
        "common_group": "DC05",
        "note": "SSK 原卷1105僅配成品DC05 0.7x550（現貨/期貨）",
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
    {
        "mat_spec": "JSH270C-P/O",
        "thickness": 0,
        "raw_widths": (),
        "customer_keys": ("THAISUMMITGOL", "TSGP"),
        "target_spec_contains": "SPHC-P/O",
        "material_code": "",
        "main_customer": "THAI SUMMIT GOL",
        "spec": "SPHC-P/O",
        "common_group": "SPHC-P/O",
        "note": "TSGP JSH270C-P/O 對應 FG/Material Code 內 SPHC-P/O",
    },
    {
        "mat_spec": "JSH440W-P/O",
        "thickness": 0,
        "raw_widths": (),
        "customer_keys": ("THAISUMMITGOL", "TSGP"),
        "target_spec_contains": "SAPH440-P/O",
        "material_code": "",
        "main_customer": "THAI SUMMIT GOL",
        "spec": "SAPH440-P/O",
        "common_group": "SAPH440-P/O",
        "note": "TSGP JSH440W-P/O 對應 SAPH440-P/O",
    },
]

# Exclusive mother-coil → finished product splits (rolls × width ratio).
# Weight per Material = coil_weight × (product_W / raw_W) × rolls
DEFAULT_EXCLUSIVE_COIL_SPLITS: list[dict] = [
    {
        "mat_spec": "BUSDE+Z-CSG+0 0/50",
        "thickness": 1.0,
        "raw_width": 926.0,
        "note": "1.0x926xC → 1.0x460×2卷",
        "splits": [
            {
                "product_width": 460.0,
                "rolls": 2,
                "material_code": "BUSDE+Z-CSG+0 0/50_1.0x460_____",
                "fg_code": "50026363",
            },
        ],
    },
    {
        "mat_spec": "BUSDE+Z-CSG+0 0/50",
        "thickness": 1.0,
        "raw_width": 1326.0,
        "note": "1.0x1326xC → 1.0x440×3C",
        "splits": [
            {
                "product_width": 440.0,
                "rolls": 3,
                "material_code": "BUSDE+Z-CSG+0 0/50_1.0x440_____",
                "fg_code": "50025836",
            },
        ],
    },
    {
        "mat_spec": "BUSDE+Z-CSG+0 0/50",
        "thickness": 1.0,
        "raw_width": 906.0,
        "note": "1.0x906xC → 1.0x440×1卷 + 1.0x460×1卷",
        "splits": [
            {
                "product_width": 440.0,
                "rolls": 1,
                "material_code": "BUSDE+Z-CSG+0 0/50_1.0x440_____",
                "fg_code": "50025836",
            },
            {
                "product_width": 460.0,
                "rolls": 1,
                "material_code": "BUSDE+Z-CSG+0 0/50_1.0x460_____",
                "fg_code": "50026363",
            },
        ],
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


def load_embedded_rules_reference_sheets() -> dict[str, pd.DataFrame]:
    """Built-in U-Stock reference sheets (baseline when no external rules workbook)."""
    rules = [
        ("U-Stock-01", "OWNER 篩選", "必要", "僅 OWNER=PTT 納入配對與 Balance；非 PTT 不填 Code、不計平衡"),
        ("U-Stock-02", "主檔 RD004", "必要", "Material Code 來自 RD004（SA007歷史銷售紀錄 + Balance sheet）；一 Code 對一 Common Group"),
        ("U-Stock-03", "配對主鍵", "必要", "1.Common Group 2.厚度T 3.寬度W；MS004 無原生 Code，需 allocate 貼上"),
        ("U-Stock-04", "MAT SPEC 跨規格", "對照", "實際 MAT SPEC 可與 Code 前綴不同，同 RD004 群組即可（見 MAT SPEC對照表）"),
        ("U-Stock-05", "尾碼 _Common", "尾碼", "跨客戶共通備貨池；同 Code 多列加總進 Balance"),
        ("U-Stock-06", "尾碼 Maker", "尾碼", "依 MAKER CODE 區分，例：_CHINA STEEL、_SSI、_BAOSHAN"),
        ("U-Stock-07", "尾碼 Customer", "尾碼", "特訂寬度母捲：尾碼用 Customer code（非標準1219），例：SPHC-P/O_2.6_1086_CASH、DC05_0.7_580_C"),
        ("U-Stock-08", "厚度容差", "容差", "0.55→0.5、3.02→3、0.75→0.7；建議 round(T,1) 或依 RD004"),
        ("U-Stock-09", "寬度容差", "容差", "Coil 共通料 Code 常固定 1219；實際 W 1270/1225/1200/1195/1260 可同碼"),
        ("U-Stock-10", "客戶專料", "排除", "專案專購 CUST CODE 特定客戶 → Material Code 留空（約85~90% PTT列）"),
        ("U-Stock-11", "Allocate 方式", "作業", "程式自動配碼（模擬原 Excel 手動貼上）：依客戶專配→Spec/T/W→尾碼優先序選 Material Code；同 Code 再加總（U-Stock-12）"),
        ("U-Stock-12", "同 Code 合併", "加總", "相同 Material Code 之 P REMAIN WT 加總 → Balance 期初現貨（噸）"),
        ("U-Stock-13", "Balance 對應", "對應", "群組型 Code 與 Balance 重疊約88~100%；另有尺寸型命名"),
        ("U-Stock-14", "不納入配對", "排除", "OWNER≠PTT、Code空白、Code=0、REMAIN WT=0 等（見排除清單）"),
        ("U-Stock-15", "CARRIER 例外", "例外", "無 Material Code 欄；改用 Master list G00料號+Spec+T+W+L"),
    ]
    return {
        "配對規則清單": pd.DataFrame(rules, columns=["規則編號", "規則名稱", "類別", "規則內容"]),
        "MAT SPEC對照": pd.DataFrame([
            {"MS004_MAT_SPEC": k, "配對_Code前綴": v, "備註": "同 Common Group" if k == "SPCC" else ""}
            for k, v in DEFAULT_SPEC_MAP.items()
        ]),
        "Code命名格式": pd.DataFrame([
            {"格式類型": "格式 A（現行，2026-02起）", "結構": "{Common_Group}_{厚度}_{寬度}_{尾碼}", "範例": "SPCC-SD_0.6_1219_Common"},
            {"格式類型": "格式 B（2026-01 Original material code）", "結構": "{MAT SPEC}_{T}_{W}_0_{Maker或PTT}", "範例": "CR_2.2_1219_0_PTT"},
        ]),
        "排除情況": pd.DataFrame([
            {"情況": "OWNER ≠ PTT", "處理": "排除，不填 Material Code"},
            {"情況": "Material Code 空白", "處理": "不計入 Balance（專料或待處理）"},
            {"情況": "Material Code = 0", "處理": "視為未配對（01月常見）"},
            {"情況": "P REMAIN WT = 0", "處理": "可保留列但重量不計"},
            {"情況": "CUST CODE 為專案客戶（如 MING TAI）", "處理": "通常不配碼"},
        ]),
        "配對流程": pd.DataFrame([
            {"步驟": 1, "動作": "MS004 匯出 → Stock 頁"},
            {"步驟": 2, "動作": "檢查 OWNER = PTT？否 → 跳過"},
            {"步驟": 3, "動作": "查 RD004 Common Group（Spec + T [+ W]）"},
            {"步驟": 4, "動作": "是否專案專料（CUST CODE）？是 → Material Code 留空"},
            {"步驟": 5, "動作": "決定尾碼（Common / Maker / 特訂寬度用 Customer code）"},
            {"步驟": 6, "動作": "填入 A 欄 Material Code（手動 allocate）"},
            {"步驟": 7, "動作": "同 Code 之 P REMAIN WT 加總 → Balance 表期初現貨"},
        ]),
    }


_MATCHING_RULES_CACHE: dict | None = None


def load_matching_rules(rules_path: Path | None = None) -> dict:
    """Load rules from Stock-Material-Code-Matching-Rules.xlsx (cached)."""
    global _MATCHING_RULES_CACHE
    rules_path = rules_path or resolve_rules_path()
    if _MATCHING_RULES_CACHE is not None and _MATCHING_RULES_CACHE.get("rules_path") == rules_path:
        return _MATCHING_RULES_CACHE

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

    _MATCHING_RULES_CACHE = {
        "spec_map": spec_map,
        "spec_crosswalk": spec_crosswalk,
        "reverse_spec_index": _build_reverse_spec_index(spec_crosswalk),
        "project_cust_codes": project_cust,
        "rules_path": rules_path,
        "rules_path_exists": rules_path_exists,
    }
    return _MATCHING_RULES_CACHE


def load_rules_workbook_sheets(rules_path: Path | None = None) -> dict[str, pd.DataFrame]:
    """Load pairing reference sheets from RD004/ (Material Master export Rules_* tabs, then rules workbook)."""
    sheets: dict[str, pd.DataFrame] = {}

    from data_loaders import resolve_rd004_master_path

    master_path = resolve_rd004_master_path()
    if master_path and master_path.exists():
        try:
            xl = pd.ExcelFile(master_path)
            for name in xl.sheet_names:
                if name.startswith("Rules_"):
                    key = name[len("Rules_"):]
                    sheets[key] = pd.read_excel(master_path, sheet_name=name)
        except Exception:
            pass

    rules_path = rules_path or resolve_rules_path()
    if rules_path.exists():
        try:
            xl = pd.ExcelFile(rules_path)
            for name in RULES_REFERENCE_SHEETS:
                if name not in sheets and name in xl.sheet_names:
                    sheets[name] = pd.read_excel(rules_path, sheet_name=name)
        except Exception:
            pass
    return sheets


def build_mat_spec_crosswalk_sheet(
    ms004_df: pd.DataFrame | None = None,
    mp008_df: pd.DataFrame | None = None,
    unmatched_stock: pd.DataFrame | None = None,
    existing_sheet: pd.DataFrame | None = None,
    rd004: pd.DataFrame | None = None,
    rules: dict | None = None,
) -> pd.DataFrame:
    """
    MAT SPEC對照 for Supply Plan: known mappings + SPECs with no Code yet
    (empty 配對_Code前綴) so the user can fill them in manually.
    """
    rules = rules or load_matching_rules()
    raw_map: dict[str, str] = dict(rules.get("spec_map") or DEFAULT_SPEC_MAP)
    # Normalized key → Code前綴
    spec_map: dict[str, str] = {_norm_spec(k): _text(v) for k, v in raw_map.items() if _norm_spec(k)}

    # RD004 identity keys (Spec / Common_Group / Material_Code prefix)
    rd004_keys: dict[str, str] = {}
    if rd004 is not None and not rd004.empty:
        for _, m in rd004.iterrows():
            for field in ("Spec", "Common_Group", "Material_Code"):
                val = _text(m.get(field, ""))
                if not val:
                    continue
                nk = _norm_spec(val)
                if nk and nk not in rd004_keys:
                    rd004_keys[nk] = val
                # Format-A group prefix before first _digit
                parts = parse_material_code_parts(val)
                grp = _text(parts.get("Parsed_Common_Group", ""))
                if grp:
                    gk = _norm_spec(grp)
                    if gk and gk not in rd004_keys:
                        rd004_keys[gk] = grp

    def _auto_code_for(spec_key: str, display_spec: str) -> tuple[str, str]:
        """Return (code_prefix, note) when auto-resolvable."""
        if spec_key in spec_map and spec_map[spec_key]:
            return spec_map[spec_key], ""
        if spec_key in rd004_keys:
            return rd004_keys[spec_key], "同 RD004 規格／群組"
        # mapped target lands on RD004
        mapped = spec_map.get(spec_key, "")
        if mapped:
            mk = _norm_spec(mapped)
            if mk in rd004_keys:
                return mapped, ""
        # fuzzy: compact compare against RD004 keys
        compact = _compact_spec(display_spec)
        if compact:
            for rk, rv in rd004_keys.items():
                if _compact_spec(rk) == compact or _compact_spec(rv) == compact:
                    return rv, "同 RD004（正規化）"
        return "", ""

    # Seed from existing Rules sheet / defaults
    rows_by_key: dict[str, dict] = {}
    sources: list[pd.DataFrame] = []
    if existing_sheet is not None and not existing_sheet.empty:
        sources.append(existing_sheet)
    else:
        sources.append(pd.DataFrame([
            {"MS004_MAT_SPEC": k, "配對_Code前綴": v, "備註": ""}
            for k, v in DEFAULT_SPEC_MAP.items()
        ]))

    for src in sources:
        for _, r in src.iterrows():
            ms = _text(r.get("MS004_MAT_SPEC", r.get("MAT SPEC", "")))
            if not ms:
                continue
            key = _norm_spec(ms)
            code = _text(r.get("配對_Code前綴", r.get("Code前綴", "")))
            note = _text(r.get("備註", ""))
            if not code:
                code, auto_note = _auto_code_for(key, ms)
                if auto_note and not note:
                    note = auto_note
            if key not in rows_by_key:
                rows_by_key[key] = {
                    "MS004_MAT_SPEC": ms,
                    "配對_Code前綴": code,
                    "狀態": "已對照" if code else "待手動輸入",
                    "備註": note,
                    "PTT列數": 0,
                    "PTT噸": 0.0,
                    "MP008列數": 0,
                    "來源": "規則表",
                }
            else:
                if code and not rows_by_key[key]["配對_Code前綴"]:
                    rows_by_key[key]["配對_Code前綴"] = code
                    rows_by_key[key]["狀態"] = "已對照"
                if note and not rows_by_key[key]["備註"]:
                    rows_by_key[key]["備註"] = note

    # Ensure every DEFAULT / rules spec_map entry appears
    for ms_key, code in spec_map.items():
        if not ms_key:
            continue
        if ms_key not in rows_by_key:
            rows_by_key[ms_key] = {
                "MS004_MAT_SPEC": ms_key,
                "配對_Code前綴": code,
                "狀態": "已對照" if code else "待手動輸入",
                "備註": "",
                "PTT列數": 0,
                "PTT噸": 0.0,
                "MP008列數": 0,
                "來源": "規則表",
            }
        elif code and not rows_by_key[ms_key]["配對_Code前綴"]:
            rows_by_key[ms_key]["配對_Code前綴"] = code
            rows_by_key[ms_key]["狀態"] = "已對照"

    def _ensure_spec(spec: str, source: str) -> dict | None:
        ms = _text(spec)
        if not ms:
            return None
        key = _norm_spec(ms)
        if key not in rows_by_key:
            code, note = _auto_code_for(key, ms)
            rows_by_key[key] = {
                "MS004_MAT_SPEC": ms,
                "配對_Code前綴": code,
                "狀態": "已對照" if code else "待手動輸入",
                "備註": note if code else "請手動填入配對_Code前綴",
                "PTT列數": 0,
                "PTT噸": 0.0,
                "MP008列數": 0,
                "來源": source,
            }
        else:
            # fill blank code from RD004 identity if possible
            if not rows_by_key[key]["配對_Code前綴"]:
                code, note = _auto_code_for(key, ms)
                if code:
                    rows_by_key[key]["配對_Code前綴"] = code
                    rows_by_key[key]["狀態"] = "已對照"
                    if note and not rows_by_key[key]["備註"]:
                        rows_by_key[key]["備註"] = note
            src = rows_by_key[key]["來源"]
            if source and source not in str(src):
                rows_by_key[key]["來源"] = f"{src}+{source}" if src else source
        return rows_by_key[key]

    # MS004 PTT specs
    if ms004_df is not None and not ms004_df.empty:
        spec_col = "Spec" if "Spec" in ms004_df.columns else (
            "MAT SPEC" if "MAT SPEC" in ms004_df.columns else None
        )
        if spec_col:
            for _, row in ms004_df.iterrows():
                rec = _ensure_spec(row.get(spec_col, ""), "MS004")
                if not rec:
                    continue
                rec["PTT列數"] = int(rec["PTT列數"]) + 1
                qty = _ms004_row_weight_ton(row) if (
                    "P REMAIN WT" in row.index or "Quantity" in row.index
                ) else _num(row.get("Quantity", 0))
                rec["PTT噸"] = round(float(rec["PTT噸"]) + qty, 6)

    # MP008 open specs
    if mp008_df is not None and not mp008_df.empty:
        spec_col = "Mat Spec" if "Mat Spec" in mp008_df.columns else (
            "MAT SPEC" if "MAT SPEC" in mp008_df.columns else None
        )
        if spec_col:
            for _, row in mp008_df.iterrows():
                rec = _ensure_spec(row.get(spec_col, ""), "MP008")
                if not rec:
                    continue
                rec["MP008列數"] = int(rec["MP008列數"]) + 1

    # Unmatched stock — keep as 待手動輸入 when still no code
    if unmatched_stock is not None and not unmatched_stock.empty:
        for _, row in unmatched_stock.iterrows():
            rec = _ensure_spec(row.get("Mat_Spec", row.get("Spec", "")), "Unmatched")
            if not rec:
                continue
            if not rec["配對_Code前綴"]:
                rec["狀態"] = "待手動輸入"
                if not rec["備註"] or "同 RD004" in str(rec["備註"]):
                    rec["備註"] = "本期 MS004 未配到 Material Code，請填 Code 前綴"
            qty = _num(row.get("Quantity", 0))
            if "Unmatched" in str(rec["來源"]) and float(rec["PTT噸"]) == 0:
                rec["PTT噸"] = round(qty, 6)

    out_rows = list(rows_by_key.values())
    for rec in out_rows:
        if not rec["配對_Code前綴"]:
            rec["狀態"] = "待手動輸入"
            if not rec["備註"]:
                rec["備註"] = "請手動填入配對_Code前綴"
        rec["PTT噸"] = round(float(rec["PTT噸"]), 3)

    out = pd.DataFrame(out_rows)
    if out.empty:
        return pd.DataFrame(columns=[
            "MS004_MAT_SPEC", "配對_Code前綴", "狀態", "備註",
            "PTT列數", "PTT噸", "MP008列數", "來源",
        ])

    out["_pending"] = (out["狀態"] != "已對照").astype(int)
    out = out.sort_values(
        by=["_pending", "PTT噸", "MS004_MAT_SPEC"],
        ascending=[False, False, True],
    ).drop(columns=["_pending"])
    return out.reset_index(drop=True)


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


def classify_pool_type(code_suffix: str, parsed_width: float | None = None) -> tuple[str, str]:
    """Return (Pool_Type, Stock_Rule_Ref) from code tail (+ optional width)."""
    suffix = _norm_spec(code_suffix).upper().replace(" ", "")
    if suffix == "COMMON" or suffix.endswith("COMMON"):
        return "共通池", "U-Stock-05"
    # U-Stock-07: special-width mother coil → suffix is Customer code (incl. CASH)
    special_w = parsed_width is not None and parsed_width > 0 and round(parsed_width, 0) != 1219
    if special_w and suffix:
        return "特訂寬度(Customer)", "U-Stock-07"
    if "CASH" in suffix:
        return "特訂寬度(Customer)", "U-Stock-07"
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
        pool_type, suffix_rule = classify_pool_type(
            parts["Code_Suffix"], parts.get("Parsed_W")
        )
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


def load_exclusive_coil_splits(rules_path: Path | None = None) -> list[dict]:
    """Load fixed mother-coil split rules (BUSDE+Z etc.)."""
    return [dict(r) for r in DEFAULT_EXCLUSIVE_COIL_SPLITS]


def _matches_coil_rule_spec(mat_spec: str, rule_spec: str) -> bool:
    a = _compact_spec(mat_spec)
    b = _compact_spec(rule_spec)
    if not a or not b:
        return False
    if a == b or b in a or a in b:
        return True
    return a.startswith("BUSDEZ") and b.startswith("BUSDEZ")


def match_exclusive_coil_split(
    mat_spec: str,
    thickness: float,
    raw_width: float,
    splits: list[dict] | None = None,
) -> dict | None:
    """Return exclusive coil split rule for spec/T/raw_W if defined."""
    splits = splits or load_exclusive_coil_splits()
    for rule in splits:
        if not _matches_coil_rule_spec(mat_spec, rule.get("mat_spec", "")):
            continue
        rule_t = _num(rule.get("thickness"))
        if rule_t > 0 and not thickness_matches(thickness, rule_t):
            continue
        if round(_num(rule.get("raw_width")), 0) != round(raw_width, 0):
            continue
        return rule
    return None


def find_exclusive_coil_split_targets(
    mat_spec: str,
    thickness: float,
    raw_coil_width: float,
    rd004: pd.DataFrame,
    splits: list[dict] | None = None,
) -> list[tuple[pd.Series, float, int]]:
    """
    Fixed coil splits with roll count.
    Returns (rd004_row, product_width, rolls).
    """
    rule = match_exclusive_coil_split(mat_spec, thickness, raw_coil_width, splits)
    if not rule:
        return []

    targets: list[tuple[pd.Series, float, int]] = []
    for split in rule.get("splits", []):
        code = _text(split.get("material_code", ""))
        if not code:
            continue
        hit = rd004[rd004["Material_Code"].astype(str) == code]
        if hit.empty:
            pw = _num(split.get("product_width"))
            hit = rd004[rd004["Width"].map(lambda w: round(_num(w), 0) == round(pw, 0))]
            if not hit.empty:
                hit = hit[hit["Material_Code"].astype(str).str.contains("BUSDE", case=False, na=False)]
        if hit.empty:
            continue
        rolls = max(int(_num(split.get("rolls", 1))), 1)
        targets.append((hit.iloc[0], _num(split.get("product_width")), rolls))
    return targets


def _append_coil_split_records(
    records: list[dict],
    row: pd.Series,
    spec: str,
    t: float,
    w: float,
    qty_ton: float,
    split_targets: list[tuple[pd.Series, float, int]],
    *,
    match_rule: str,
    cust_code: str = "",
    maker_code: str = "",
    extra_fields: dict | None = None,
) -> None:
    for m, product_w, rolls in split_targets:
        ratio = (product_w / w) * rolls if w > 0 else 0.0
        rec = {
            "Material_Code": m["Material_Code"],
            "Common_Group": m.get("Common_Group", m.get("Spec", "")),
            "Quantity": round(qty_ton * ratio, 6),
            "Mat_Spec": spec,
            "T": t,
            "W": w,
            "CUST_CODE": cust_code,
            "MAKER_CODE": maker_code,
            "Allocatable": True,
            "Match_Rule": match_rule,
            "Raw_Coil_W": w,
            "Product_W": product_w,
            "Rolls": rolls,
            "Split_Ratio": round(ratio, 6),
        }
        if extra_fields:
            rec.update(extra_fields)
        records.append(rec)


_CUSTOMER_SPEC_PAIRINGS_CACHE: dict | None = None


def load_customer_spec_pairings(rules_path: Path | None = None) -> list[dict]:
    """Load exclusive customer/spec pairings from rules workbook or defaults."""
    global _CUSTOMER_SPEC_PAIRINGS_CACHE
    rules_path = rules_path or resolve_rules_path()
    if (
        _CUSTOMER_SPEC_PAIRINGS_CACHE is not None
        and _CUSTOMER_SPEC_PAIRINGS_CACHE.get("rules_path") == rules_path
    ):
        return _CUSTOMER_SPEC_PAIRINGS_CACHE["pairings"]

    pairings = [dict(p) for p in DEFAULT_CUSTOMER_SPEC_PAIRINGS]
    if not rules_path.exists():
        _CUSTOMER_SPEC_PAIRINGS_CACHE = {"rules_path": rules_path, "pairings": pairings}
        return pairings
    try:
        xl = pd.ExcelFile(rules_path)
        sheet = next((s for s in xl.sheet_names if "客戶專屬" in s or "專屬配對" in s), None)
        if not sheet:
            _CUSTOMER_SPEC_PAIRINGS_CACHE = {"rules_path": rules_path, "pairings": pairings}
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
        result = loaded if loaded else pairings
        _CUSTOMER_SPEC_PAIRINGS_CACHE = {"rules_path": rules_path, "pairings": result}
        return result
    except Exception:
        _CUSTOMER_SPEC_PAIRINGS_CACHE = {"rules_path": rules_path, "pairings": pairings}
        return pairings


def _row_matches_target_spec(m: pd.Series, target_spec: str, rules: dict | None = None) -> bool:
    """True if RD004 row Material Code / Spec / FG_Code corresponds to target spec group."""
    target = _norm_spec(target_spec)
    if not target:
        return False
    tk = _compact_spec(target)
    for field in ("Material_Code", "Spec", "Common_Group", "FG_Code"):
        val = _text(m.get(field, ""))
        if not val:
            continue
        if spec_group_matches(target, val, rules):
            return True
        if tk in _compact_spec(val):
            return True
    fg_all = _text(m.get("FG_Codes_All", ""))
    if target.replace("-", "") in fg_all.upper().replace("-", ""):
        return True
    return False


def resolve_cross_spec_material_code(
    rule: dict,
    thickness: float,
    rd004: pd.DataFrame,
    customer: str = "",
    rules: dict | None = None,
) -> str:
    """Map inbound/source spec to a concrete Material_Code in RD004 by target spec group + T."""
    target = _text(rule.get("target_spec_contains", ""))
    if not target or rd004.empty:
        return ""

    hits: list[pd.Series] = []
    for _, m in rd004.iterrows():
        if not _row_matches_target_spec(m, target, rules):
            continue
        if not thickness_matches(thickness, _num(m.get("Thickness"))):
            continue
        hits.append(m)

    if not hits:
        return ""

    for m in hits:
        main_c = _text(m.get("Main_Customer", ""))
        if customer_matches("TSGP", main_c) or customer_matches(customer, main_c):
            return _text(m["Material_Code"])

    common = [m for m in hits if is_common_material_code(m["Material_Code"])]
    if common:
        return _text(common[0]["Material_Code"])

    return _text(hits[0]["Material_Code"])


def match_exclusive_customer_spec_pairing(
    mat_spec: str,
    thickness: float,
    raw_width: float,
    customer: str = "",
    cust_code: str = "",
    pairings: list[dict] | None = None,
) -> dict | None:
    """Match customer/spec pairing rule (fixed Material Code or cross-spec target)."""
    pairings = pairings or load_customer_spec_pairings()
    for rule in pairings:
        if _norm_spec(mat_spec) != _norm_spec(rule.get("mat_spec", "")):
            continue
        rule_t = _num(rule.get("thickness"))
        if rule_t > 0 and not thickness_matches(thickness, rule_t):
            continue
        raw_ws = rule.get("raw_widths") or ()
        if raw_ws and not any(round(raw_width, 0) == round(rw, 0) for rw in raw_ws):
            continue
        if not customer_keys_match(customer, cust_code, tuple(rule.get("customer_keys", ()))):
            continue
        return rule
    return None


def resolve_customer_spec_pairing(
    mat_spec: str,
    thickness: float,
    raw_width: float,
    rd004: pd.DataFrame,
    customer: str = "",
    cust_code: str = "",
    pairings: list[dict] | None = None,
    rules: dict | None = None,
) -> dict | None:
    """Resolve pairing rule to a concrete Material_Code (fixed or cross-spec lookup)."""
    rule = match_exclusive_customer_spec_pairing(
        mat_spec, thickness, raw_width, customer, cust_code, pairings
    )
    if not rule:
        return None

    code = _text(rule.get("material_code", ""))
    if not code:
        code = resolve_cross_spec_material_code(rule, thickness, rd004, customer, rules)

    if not code:
        return None

    out = dict(rule)
    out["material_code"] = code
    return out


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


def rd004_match_width(m: pd.Series) -> float:
    """
    Width used for U-Stock-09 stock/PO matching.
    Format-A Material Codes with Parsed_W=1219 are mother-coil commons: match on 1219
    even when the Excel Width cell holds a finished-product strip width.
    """
    code = _text(m.get("Material_Code", ""))
    parts = parse_material_code_parts(code)
    parsed_w = parts.get("Parsed_W")
    if parsed_w is not None and round(float(parsed_w), 0) == 1219:
        return 1219.0
    return _num(m.get("Width"))


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
        # Spec column is often blank; match Spec / Common_Group / Material_Code
        if not material_row_matches_spec(mat_spec, m, rules):
            continue
        if not thickness_matches(thickness, _num(m.get("Thickness"))):
            continue
        if not width_matches(width, rd004_match_width(m)):
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


def allocate_ms004_stock(
    ms004_df: pd.DataFrame,
    rd004: pd.DataFrame,
    rules: dict | None = None,
    *,
    return_unmatched: bool = False,
):
    """
    Build allocatable spot stock per U-Stock rules.
    Returns aggregated Quantity per Material_Code.
    If return_unmatched=True, returns (summary, unmatched_df).
    """
    rules = rules or load_matching_rules()
    empty_summary = pd.DataFrame(columns=[
        "Material_Code", "Common_Group", "Quantity", "Mat_Spec", "T", "W",
        "CUST_CODE", "Allocatable", "Match_Rule",
    ])
    empty_unmatched = pd.DataFrame(columns=[
        "Mat_Spec", "T", "W", "Quantity", "CUST_CODE", "MAKER_CODE", "Reason",
    ])
    if ms004_df.empty:
        return (empty_summary, empty_unmatched) if return_unmatched else empty_summary

    pairings = load_customer_spec_pairings()
    work_rows = []
    for _, row in ms004_df.iterrows():
        if not is_allocatable_ms004_row(row, rules):
            continue
        qty_ton = _ms004_row_weight_ton(row)
        if qty_ton <= 0:
            continue
        work_rows.append({
            "Spec": _text(row.get("Spec", row.get("MAT SPEC", ""))),
            "Thickness": _num(row.get("Thickness", row.get("T"))),
            "Width": _num(row.get("Width", row.get("W"))),
            "Quantity": qty_ton,
            "CUST_CODE": _text(row.get("CUST CODE", row.get("Cust Code", ""))),
            "MAKER_CODE": _text(row.get("MAKER CODE", row.get("Maker Code", ""))),
        })
    if not work_rows:
        return (empty_summary, empty_unmatched) if return_unmatched else empty_summary

    work_df = (
        pd.DataFrame(work_rows)
        .groupby(["Spec", "Thickness", "Width", "CUST_CODE", "MAKER_CODE"], as_index=False, dropna=False)["Quantity"]
        .sum()
    )

    records = []
    unmatched = []
    for _, row in work_df.iterrows():
        spec = _text(row["Spec"])
        t = _num(row["Thickness"])
        w = _num(row["Width"])
        qty_ton = _num(row["Quantity"])
        cust_code = _text(row["CUST_CODE"])
        maker_code = _text(row["MAKER_CODE"])
        exclusive = resolve_customer_spec_pairing(
            spec,
            t,
            w,
            rd004,
            "",
            cust_code,
            pairings=pairings,
            rules=rules,
        )
        if exclusive:
            records.append({
                "Material_Code": exclusive["material_code"],
                "Common_Group": exclusive.get("common_group", spec),
                "Quantity": qty_ton,
                "Mat_Spec": spec,
                "T": t,
                "W": w,
                "CUST_CODE": cust_code,
                "MAKER_CODE": maker_code,
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
                stock_width=w,
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
                    "CUST_CODE": cust_code,
                    "MAKER_CODE": maker_code,
                    "Allocatable": True,
                    "Match_Rule": "U-Stock-07" if (
                        w > 0 and round(w, 0) != 1219
                        and not is_common_material_code(mat_code)
                    ) else "U-Stock-03",
                })
            continue

        split_targets = find_coil_split_targets(spec, t, w, rd004, rules, cust_code=cust_code)
        if split_targets:
            split_rule = (
                "Coil-Split-Exclusive"
                if match_exclusive_coil_split(spec, t, w)
                else "Coil-Split"
            )
            _append_coil_split_records(
                records,
                row,
                spec,
                t,
                w,
                qty_ton,
                split_targets,
                match_rule=split_rule,
                cust_code=cust_code,
                maker_code=maker_code,
            )
            continue

        unmatched.append({
            "Mat_Spec": spec,
            "T": t,
            "W": w,
            "Quantity": round(qty_ton, 6),
            "CUST_CODE": cust_code,
            "MAKER_CODE": maker_code,
            "Reason": "No RD004 Spec/T/W match",
        })

    unmatched_df = pd.DataFrame(unmatched) if unmatched else empty_unmatched
    if not unmatched_df.empty:
        unmatched_df = (
            unmatched_df.groupby(
                ["Mat_Spec", "T", "W", "CUST_CODE", "MAKER_CODE", "Reason"],
                as_index=False,
            )["Quantity"]
            .sum()
            .round(6)
            .sort_values("Quantity", ascending=False)
        )

    if not records:
        summary = empty_summary
    else:
        detail = pd.DataFrame(records)
        # U-Stock-12: same Material Code sum weight
        summary = (
            detail.groupby(["Material_Code", "Common_Group"], as_index=False)["Quantity"]
            .sum()
            .round(6)
        )
        summary["Allocatable"] = True
        summary["Match_Rule"] = "U-Stock-12"

    if return_unmatched:
        return summary, unmatched_df
    return summary


def filter_open_mp008(mp008_df: pd.DataFrame) -> pd.DataFrame:
    """
    U4: keep in-transit PO lines for Inbound.
    Primary: Unshipped_Ratio > 20% (PO Wt − Received WT) / PO Wt.
    Fallback: Quantity > 0 when ratio columns absent.
    """
    if mp008_df.empty:
        return mp008_df
    out = mp008_df.copy()
    if "Unshipped_Ratio" in out.columns:
        out = out[_num_series(out["Unshipped_Ratio"]) > 0.20]
    elif "PO Balance WT" in out.columns:
        if "Close Flag" in out.columns:
            out = out[
                out["Close Flag"].astype(str).str.upper().isin(["FALSE", "0", "", "NAN"])
                | out["Close Flag"].isna()
            ]
        out = out[_num_series(out["PO Balance WT"]) > 0]
    if "Status" in out.columns:
        out = out[out["Status"].astype(str).str.upper() != "RECEIVED"]
    if "Quantity" in out.columns:
        out = out[_num_series(out["Quantity"]) > 0]
    return out


def _num_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _pick_material_for_stock(
    candidates: pd.DataFrame,
    cust_code: str,
    maker_code: str,
    stock_width: float | None = None,
) -> str:
    """
    One MS004 row → one Material Code (manual allocate behaviour).
    Priority: Main_Customer match → U-Stock-07 special-width Customer suffix
    → Common pool → Maker in code → first candidate.
    """
    if candidates.empty:
        return ""
    cust = _text(cust_code)
    if cust:
        for _, m in candidates.iterrows():
            if customer_matches(cust, m.get("Main_Customer", "")):
                return m["Material_Code"]
            # U-Stock-07: special-width mother coil suffix = Customer code
            parts = parse_material_code_parts(m.get("Material_Code", ""))
            suffix = _text(parts.get("Code_Suffix", ""))
            if suffix and customer_matches(cust, suffix):
                return m["Material_Code"]

    special_w = (
        stock_width is not None
        and stock_width > 0
        and round(stock_width, 0) != 1219
    )
    if special_w:
        # Prefer non-Common codes whose Parsed_W matches the special mother-coil width
        for _, m in candidates.iterrows():
            parts = parse_material_code_parts(m.get("Material_Code", ""))
            parsed_w = parts.get("Parsed_W")
            suffix = _norm_spec(parts.get("Code_Suffix", ""))
            if not parsed_w or round(float(parsed_w), 0) == 1219:
                continue
            if suffix.endswith("COMMON") or suffix == "COMMON":
                continue
            if round(float(parsed_w), 0) == round(stock_width, 0):
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
) -> list[tuple[pd.Series, float, int]]:
    """
    Mother-coil split allocations.
    Exclusive rules (BUSDE+Z fixed W×rolls) take priority over generic proportional split.
    Returns (rd004_row, product_width, rolls); weight = coil_wt × (product_W/raw_W) × rolls.
    """
    if raw_coil_width <= 0:
        return []

    exclusive = find_exclusive_coil_split_targets(mat_spec, thickness, raw_coil_width, rd004)
    if exclusive:
        return exclusive

    if pairing_blocks_coil_split(mat_spec, thickness, raw_coil_width, customer, cust_code):
        return []

    rules = rules or load_matching_rules()
    targets: list[tuple[pd.Series, float, int]] = []
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
        targets.append((m, product_w, 1))

    if len(targets) < 2:
        return []

    sum_product_w = sum(pw for _, pw, _ in targets)
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
        "PO Wt": _num(row.get("PO Wt", 0)),
        "Received WT": _num(row.get("Received WT", 0)),
        "Unshipped WT": _num(row.get("Unshipped WT", qty_ton * 1000)),
        "Unshipped_Ratio": _num(row.get("Unshipped_Ratio", 0)),
        "Inbound_Reason": _text(row.get("Inbound_Reason", "")),
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
        exclusive = resolve_customer_spec_pairing(spec, t, w, rd004, po_customer, "")
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
            split_rule = (
                "Coil-Split-Exclusive"
                if match_exclusive_coil_split(spec, t, w)
                else "Coil-Split"
            )
            for m, product_w, rolls in split_targets:
                ratio = (product_w / w) * rolls if w > 0 else 0.0
                rows.append(_mp008_alloc_row(
                    row,
                    m["Material_Code"],
                    qty_ton * ratio,
                    split_rule,
                    Raw_Coil_W=w,
                    Product_W=product_w,
                    Rolls=rolls,
                    Split_Ratio=round(ratio, 6),
                    Product_W_Sum=sum(pw * r for _, pw, r in split_targets),
                ))
            continue

        # Keep unmatched open PO visible in report (not used in Balance inbound)
        rows.append(_mp008_alloc_row(row, "", qty_ton, "Unmatched"))

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
