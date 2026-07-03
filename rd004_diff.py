# -*- coding: utf-8 -*-
"""Compare RD004/ folder inputs against baseline Material Master pairing rules."""

from __future__ import annotations

from collections import Counter
from pathlib import Path

import pandas as pd

from data_loaders import (
    BASE_RD004_COLUMNS,
    RD004_DIR,
    SA007_HISTORY_LABEL,
    TEMPLATE_REFERENCE_LABEL,
    _load_rd004_from_folder,
    _load_rd004_from_sample_balance,
    _text,
    _num,
    is_carrier_label,
    resolve_rd004_master_path,
    resolve_sample_path,
)
from stock_matching import (
    RULES_REFERENCE_SHEETS,
    load_embedded_rules_reference_sheets,
    load_rules_workbook_sheets,
    resolve_baseline_rules_path,
    resolve_rules_path,
)

MASTER_COMPARE_FIELDS = [c for c in BASE_RD004_COLUMNS if c != "Material_Code"]


def rd004_folder_active() -> bool:
    return resolve_rd004_master_path() is not None


def load_baseline_material_master(sample_path: Path | None = None) -> pd.DataFrame:
    """樣本參考基準（Sample Balance）；僅供 RD004 差異比對，非 Supply Plan 運行來源。"""
    sample_path = sample_path or resolve_sample_path()
    df = _load_rd004_from_sample_balance(sample_path)
    if df.empty:
        return df
    return df[~df["Main_Customer"].map(lambda v: is_carrier_label(v))].reset_index(drop=True)


def load_rd004_folder_material_master() -> pd.DataFrame:
    """Material Master as read from RD004/ folder (no supplement)."""
    df = _load_rd004_from_folder()
    if df.empty:
        return df
    return df[~df["Main_Customer"].map(lambda v: is_carrier_label(v))].reset_index(drop=True)


def _field_equal(field: str, baseline_val, current_val) -> bool:
    if field in ("Thickness", "Width", "MOQ"):
        return round(_num(baseline_val), 3) == round(_num(current_val), 3)
    return _text(baseline_val) == _text(current_val)


def _format_field_value(field: str, val) -> str:
    if field in ("Thickness", "Width", "MOQ"):
        n = _num(val)
        return f"{n:g}" if field != "MOQ" else f"{n:.3f}"
    return _text(val)


def compare_material_master(baseline: pd.DataFrame, current: pd.DataFrame) -> pd.DataFrame:
    """Row-level diff: added / removed / field changes per Material_Code."""
    rows: list[dict] = []
    base_idx = baseline.set_index("Material_Code", drop=False) if not baseline.empty else pd.DataFrame()
    curr_idx = current.set_index("Material_Code", drop=False) if not current.empty else pd.DataFrame()
    base_codes = set(baseline["Material_Code"]) if not baseline.empty else set()
    curr_codes = set(current["Material_Code"]) if not current.empty else set()

    for code in sorted(base_codes - curr_codes):
        b = base_idx.loc[code]
        if isinstance(b, pd.DataFrame):
            b = b.iloc[0]
        rows.append({
            "差異類型": "僅基準有(已移除)",
            "Material_Code": code,
            "欄位": "(整列)",
            "基準值_原本": _text(b.get("Spec", "")) or code,
            "RD004值": "",
            "基準_Common_Group": _text(b.get("Common_Group", "")),
            "RD004_Common_Group": "",
            "基準_FG_Code": _text(b.get("FG_Code", "")),
            "RD004_FG_Code": "",
        })

    for code in sorted(curr_codes - base_codes):
        c = curr_idx.loc[code]
        if isinstance(c, pd.DataFrame):
            c = c.iloc[0]
        rows.append({
            "差異類型": "僅RD004有(新增)",
            "Material_Code": code,
            "欄位": "(整列)",
            "基準值_原本": "",
            "RD004值": _text(c.get("Spec", "")) or code,
            "基準_Common_Group": "",
            "RD004_Common_Group": _text(c.get("Common_Group", "")),
            "基準_FG_Code": "",
            "RD004_FG_Code": _text(c.get("FG_Code", "")),
        })

    for code in sorted(base_codes & curr_codes):
        b = base_idx.loc[code]
        c = curr_idx.loc[code]
        if isinstance(b, pd.DataFrame):
            b = b.iloc[0]
        if isinstance(c, pd.DataFrame):
            c = c.iloc[0]
        for field in MASTER_COMPARE_FIELDS:
            if _field_equal(field, b.get(field), c.get(field)):
                continue
            rows.append({
                "差異類型": "欄位變更",
                "Material_Code": code,
                "欄位": field,
                "基準值_原本": _format_field_value(field, b.get(field)),
                "RD004值": _format_field_value(field, c.get(field)),
                "基準_Common_Group": _text(b.get("Common_Group", "")),
                "RD004_Common_Group": _text(c.get("Common_Group", "")),
                "基準_FG_Code": _text(b.get("FG_Code", "")),
                "RD004_FG_Code": _text(c.get("FG_Code", "")),
            })

    cols = [
        "差異類型", "Material_Code", "欄位",
        "基準值_原本", "RD004值",
        "基準_Common_Group", "RD004_Common_Group",
        "基準_FG_Code", "RD004_FG_Code",
    ]
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols]


def _sheet_row_tuples(df: pd.DataFrame) -> list[tuple[str, ...]]:
    if df.empty:
        return []
    out: list[tuple[str, ...]] = []
    for _, row in df.iterrows():
        out.append(tuple(_text(v) for v in row.tolist()))
    return out


def compare_rules_sheet(sheet_name: str, baseline: pd.DataFrame, current: pd.DataFrame) -> list[dict]:
    """Diff one rules reference sheet by row content."""
    rows: list[dict] = []
    base_cols = [_text(c) for c in baseline.columns] if not baseline.empty else []
    curr_cols = [_text(c) for c in current.columns] if not current.empty else []

    if base_cols != curr_cols and (base_cols or curr_cols):
        rows.append({
            "工作表": sheet_name,
            "差異類型": "欄位結構變更",
            "列識別": "(欄位)",
            "基準值_原本": " | ".join(base_cols) if base_cols else "(無)",
            "RD004值": " | ".join(curr_cols) if curr_cols else "(無)",
        })

    base_counter = Counter(_sheet_row_tuples(baseline))
    curr_counter = Counter(_sheet_row_tuples(current))

    for row_tuple, count in sorted(base_counter.items()):
        curr_count = curr_counter.get(row_tuple, 0)
        if curr_count >= count:
            continue
        label = " | ".join(row_tuple[:4]) if row_tuple else "(空白列)"
        for _ in range(count - curr_count):
            rows.append({
                "工作表": sheet_name,
                "差異類型": "僅基準有(已移除)",
                "列識別": label[:200],
                "基準值_原本": label[:500],
                "RD004值": "",
            })

    for row_tuple, count in sorted(curr_counter.items()):
        base_count = base_counter.get(row_tuple, 0)
        if base_count >= count:
            continue
        label = " | ".join(row_tuple[:4]) if row_tuple else "(空白列)"
        for _ in range(count - base_count):
            rows.append({
                "工作表": sheet_name,
                "差異類型": "僅RD004有(新增)",
                "列識別": label[:200],
                "基準值_原本": "",
                "RD004值": label[:500],
            })

    return rows


def load_baseline_rules_workbook_sheets() -> dict[str, pd.DataFrame]:
    """Original pairing rules outside RD004/; embedded defaults if no file."""
    path = resolve_baseline_rules_path()
    sheets: dict[str, pd.DataFrame] = {}
    if path and path.exists():
        try:
            xl = pd.ExcelFile(path)
            for name in RULES_REFERENCE_SHEETS:
                if name in xl.sheet_names:
                    sheets[name] = pd.read_excel(path, sheet_name=name)
        except Exception:
            pass
    if sheets:
        return sheets
    return load_embedded_rules_reference_sheets()


def compare_rules_workbooks(
    baseline_sheets: dict[str, pd.DataFrame],
    current_sheets: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    cols = ["工作表", "差異類型", "列識別", "基準值_原本", "RD004值"]
    all_names = sorted(set(baseline_sheets) | set(current_sheets))
    rows: list[dict] = []
    for name in all_names:
        base_df = baseline_sheets.get(name, pd.DataFrame())
        curr_df = current_sheets.get(name, pd.DataFrame())
        if name not in baseline_sheets and not curr_df.empty:
            rows.append({
                "工作表": name,
                "差異類型": "僅RD004有(新工作表)",
                "列識別": f"{len(curr_df)} 列",
                "基準值_原本": "(無此工作表)",
                "RD004值": f"{len(curr_df)} 列",
            })
            continue
        if name not in current_sheets and not base_df.empty:
            rows.append({
                "工作表": name,
                "差異類型": "僅基準有(工作表已移除)",
                "列識別": f"{len(base_df)} 列",
                "基準值_原本": f"{len(base_df)} 列",
                "RD004值": "(無此工作表)",
            })
            continue
        rows.extend(compare_rules_sheet(name, base_df, curr_df))
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[cols]


def list_rd004_folder_files() -> list[Path]:
    if not RD004_DIR.is_dir():
        return []
    files = [
        p for p in sorted(RD004_DIR.iterdir())
        if p.is_file() and not p.name.startswith("~$")
    ]
    return files


def build_rd004_diff_report() -> dict[str, pd.DataFrame]:
    """
    When RD004/ has files, compare against template sample baseline
    (Sample Balance export layout). History Balance/ is template-only.
    """
    if not rd004_folder_active():
        return {}

    baseline_master = load_baseline_material_master()
    current_master = load_rd004_folder_material_master()
    master_diff = compare_material_master(baseline_master, current_master)

    baseline_rules = load_baseline_rules_workbook_sheets()
    current_rules = load_rules_workbook_sheets()
    rules_diff = compare_rules_workbooks(baseline_rules, current_rules)

    rd004_master_path = resolve_rd004_master_path()
    rules_path = resolve_rules_path()
    baseline_rules_path = resolve_baseline_rules_path()

    summary_rows = [
        {
            "項目": "RD004 資料夾",
            "基準_原本": f"樣本參考 Sample Balance（{TEMPLATE_REFERENCE_LABEL}）",
            "RD004": str(RD004_DIR),
        },
        {
            "項目": "Material Master 檔",
            "基準_原本": resolve_sample_path().name,
            "RD004": rd004_master_path.name if rd004_master_path else "",
        },
        {
            "項目": "配對規則檔",
            "基準_原本": (
                baseline_rules_path.name
                if baseline_rules_path
                else "內建 U-Stock 預設規則 (export_matching_rules_excel)"
            ),
            "RD004": rules_path.name if rules_path.exists() else "",
        },
        {
            "項目": "主檔 Material 數",
            "基準_原本": len(baseline_master),
            "RD004": len(current_master),
        },
        {
            "項目": "主檔差異筆數",
            "基準_原本": "—",
            "RD004": len(master_diff),
        },
        {
            "項目": "配對規則差異筆數",
            "基準_原本": "—",
            "RD004": len(rules_diff),
        },
    ]
    for path in list_rd004_folder_files():
        role = "已納入比對"
        if path.suffix.lower() == ".md":
            role = "參考文件（不參與比對）"
        elif path.suffix.lower() not in (".xlsx", ".xls"):
            role = "非 Excel（僅列示）"
        summary_rows.append({
            "項目": f"RD004 檔案: {path.name}",
            "基準_原本": "見 RD004_差異_主檔 / RD004_差異_配對規則",
            "RD004": role,
        })

    type_counts = master_diff["差異類型"].value_counts() if not master_diff.empty else pd.Series(dtype=int)
    for label, count in type_counts.items():
        summary_rows.append({
            "項目": f"主檔 — {label}",
            "基準_原本": "—",
            "RD004": int(count),
        })

    rules_type_counts = rules_diff["差異類型"].value_counts() if not rules_diff.empty else pd.Series(dtype=int)
    for label, count in rules_type_counts.items():
        summary_rows.append({
            "項目": f"配對規則 — {label}",
            "基準_原本": "—",
            "RD004": int(count),
        })

    if master_diff.empty and rules_diff.empty:
        summary_rows.append({
            "項目": "結論",
            "基準_原本": "與原本 Material Master 配對規則一致",
            "RD004": "無差異",
        })
    else:
        summary_rows.append({
            "項目": "結論",
            "基準_原本": "詳見下方差異工作表",
            "RD004": f"主檔 {len(master_diff)} 筆、配對規則 {len(rules_diff)} 筆差異",
        })

    return {
        "RD004_差異摘要": pd.DataFrame(summary_rows),
        "RD004_差異_主檔": master_diff,
        "RD004_差異_配對規則": rules_diff,
    }
