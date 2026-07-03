# -*- coding: utf-8 -*-
"""
Material Balance Report Generator
Analyzes stock vs forecast demand and produces a balance sheet
similar to Sample Balance / Balance sheet (Update).
"""

from __future__ import annotations

import glob
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_PATH = BASE_DIR / "Sample Balance" / "All Customer review Jun '2026 review 20.06.2026.xlsx"
STOCK_PATH = BASE_DIR / "Stock" / "MS004-260619.xls"
FORECAST_DIR = BASE_DIR / "Forecast"
FORECAST_DIR_LEGACY = BASE_DIR / "Froecast"
OUTPUT_DIR = BASE_DIR / "Output"

MONTHS_SALES = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    "Jan", "Feb", "Mar",
]
MONTHS_BALANCE = ["Jun", "Jul", "Aug", "Sep", "Oct"]
STEEL_DENSITY = 7.85  # g/cm3
AL_DENSITY = 2.7


def num(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, str):
        value = value.strip().replace(",", "")
        if value in ("", "-", "nan"):
            return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def normalize_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_spec(value) -> str:
    text = normalize_text(value).upper()
    text = text.replace("P/O", "P/O").replace("?", "")
    return text


def parse_material_code(code: str) -> dict | None:
    code = normalize_text(code)
    if not code:
        return None

    # e.g. A5052 1.6X300.00X1250.00
    m = re.match(
        r"^([A-Z0-9][A-Z0-9\-\+\s/]+?)\s+([\d.]+)[xX_]([\d.]+)[xX_]([\d.]+)$",
        code,
        re.IGNORECASE,
    )
    if m:
        return {
            "spec": normalize_spec(m.group(1)),
            "t": num(m.group(2)),
            "w": num(m.group(3)),
            "suffix": normalize_text(m.group(4)),
            "code": code,
        }

    # e.g. SPFH590-P/O_2_1219_Common or BUSDE+Z-CSG+0 0/50_1.0x440_____
    if "_" in code:
        suffix = code.rsplit("_", 1)[-1]
        main = code.rsplit("_", 1)[0]
        parts = main.split("_")
        if len(parts) >= 3:
            t_raw = parts[-2].replace("t", "").strip()
            w_raw = parts[-1].replace("x", "").strip()
            spec = "_".join(parts[:-2])
            t_match = re.search(r"([\d.]+)", t_raw)
            w_match = re.search(r"([\d.]+)", w_raw)
            if t_match:
                return {
                    "spec": normalize_spec(spec),
                    "t": num(t_match.group(1)),
                    "w": num(w_match.group(1)) if w_match else 1219.0,
                    "suffix": suffix,
                    "code": code,
                }

    return {"spec": normalize_spec(code), "t": None, "w": None, "suffix": "", "code": code}


def calc_piece_weight_kg(t_mm, w_mm, l_mm, kind: str = "CR") -> float:
    t, w, l = num(t_mm), num(w_mm), num(l_mm)
    if t <= 0 or w <= 0 or l <= 0:
        return 0.0
    density = AL_DENSITY if str(kind).upper() == "AL" else STEEL_DENSITY
    return t * w * l * density / 1_000_000


def parse_dimensions_from_text(text: str) -> tuple[float, float, float]:
    text = normalize_text(text)
    m = re.search(
        r"([\d.]+)\s*[xX]\s*([\d.]+)\s*[xX]\s*([\d.]+)",
        text,
    )
    if m:
        return num(m.group(1)), num(m.group(2)), num(m.group(3))
    return 0.0, 0.0, 0.0


def load_balance_materials(sample_path: Path) -> pd.DataFrame:
    raw = pd.read_excel(sample_path, sheet_name="Balance sheet (Update)", header=None)
    rows = []
    for i in range(6, len(raw)):
        code = raw.iloc[i, 1]
        if pd.isna(code):
            continue
        sales = [num(raw.iloc[i, 5 + j]) for j in range(15)]
        rows.append(
            {
                "item": raw.iloc[i, 0],
                "material_code": normalize_text(code),
                "kind": normalize_text(raw.iloc[i, 2]),
                "t": raw.iloc[i, 3],
                "main_customer": normalize_text(raw.iloc[i, 4]),
                "sales": sales,
                "avg_sales": num(raw.iloc[i, 20]),
                "incoming_may": num(raw.iloc[i, 25]),
            }
        )
    return pd.DataFrame(rows)


def load_sa007_history_mapping(sample_path: Path) -> dict[tuple, str]:
    """Material spec+T+W → Material Code from Sample Balance SA007歷史銷售紀錄工作表（legacy fallback）。"""
    sa007_legacy = pd.read_excel(sample_path, sheet_name="Act order", header=None)
    mapping: dict[tuple, str] = {}
    for _, row in sa007_legacy.iloc[4:].iterrows():
        code = row[0]
        spec = row[1]
        t = row[8]
        w = row[9]
        if pd.isna(code) or pd.isna(spec):
            continue
        key = (normalize_spec(spec), num(t), num(w))
        mapping[key] = normalize_text(code)
    return mapping


def load_act_order_mapping(sample_path: Path) -> dict[tuple, str]:
    """Deprecated alias for load_sa007_history_mapping()."""
    return load_sa007_history_mapping(sample_path)


def spec_matches(stock_spec: str, target_spec: str) -> bool:
    a = normalize_spec(stock_spec)
    b = normalize_spec(target_spec)
    if not a or not b:
        return False
    if a == b:
        return True
    a_compact = re.sub(r"[^A-Z0-9]", "", a)
    b_compact = re.sub(r"[^A-Z0-9]", "", b)
    return a_compact == b_compact or a_compact.startswith(b_compact[:6]) or b_compact.startswith(a_compact[:6])


def load_stock_by_material(stock_path: Path, materials: pd.DataFrame) -> dict[str, float]:
    stock = pd.read_excel(stock_path)
    stock["REMAIN WT"] = pd.to_numeric(stock["REMAIN WT"], errors="coerce").fillna(0)
    stock["T"] = pd.to_numeric(stock["T"], errors="coerce")
    stock["W"] = pd.to_numeric(stock["W"], errors="coerce")

    result: dict[str, float] = {code: 0.0 for code in materials["material_code"]}

    for _, mat in materials.iterrows():
        parsed = parse_material_code(mat["material_code"])
        if not parsed:
            continue
        spec = parsed["spec"]
        t_val = parsed["t"]
        w_val = parsed["w"]

        mask = stock["MAT SPEC"].apply(lambda x: spec_matches(x, spec))
        if t_val is not None:
            mask &= stock["T"].round(3) == round(t_val, 3)
        if w_val is not None and w_val > 0:
            # Coil materials often use 1219 in code but various widths in stock.
            if w_val != 1219:
                mask &= stock["W"].round(0) == round(w_val, 0)

        matched_wt = stock.loc[mask, "REMAIN WT"].sum()
        if matched_wt > 0:
            result[mat["material_code"]] = matched_wt

    return result


def parse_cpk_forecast(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, header=None)
    header_row = None
    for r in range(min(20, len(df))):
        row = [normalize_text(v) for v in df.iloc[r].tolist()]
        if "ITEM" in row and "SPEC." in row:
            header_row = r
            break
    if header_row is None:
        return pd.DataFrame()

    month_cols: list[tuple[str, int]] = []
    for c in range(df.shape[1]):
        v = df.iloc[header_row, c]
        if isinstance(v, datetime):
            month_cols.append((v.strftime("%b"), c))

    records = []
    for r in range(header_row + 1, len(df)):
        item = df.iloc[r, 1]
        if pd.isna(item):
            continue

        # CPK layout: col3=T, col5=W, col7=L, col8=SPEC
        t = num(df.iloc[r, 3])
        w = num(df.iloc[r, 5])
        l = num(df.iloc[r, 7])
        spec = normalize_text(df.iloc[r, 8])
        if not spec or t <= 0:
            continue

        kind = "AL" if spec[:1] == "A" and len(spec) > 1 and spec[1].isdigit() else "CR"
        kg_per_pc = calc_piece_weight_kg(t, w, l, kind)
        monthly: dict[str, float] = {}
        for month_name, col in month_cols:
            pcs = num(df.iloc[r, col])
            monthly[month_name] = monthly.get(month_name, 0) + pcs * kg_per_pc
        if sum(monthly.values()) <= 0:
            continue

        w_code = int(w) if w else 1219
        records.append(
            {
                "source": "CPK",
                "spec": normalize_spec(spec),
                "t": t,
                "w": w if w else 1219.0,
                "material_code": f"{spec}_{t}_{w_code}_Common",
                **monthly,
            }
        )
    return pd.DataFrame(records)


def find_forecast_table(df: pd.DataFrame) -> tuple[int, dict[str, int], int, int] | None:
    month_map = {
        "07.2026": "Jul",
        "08.2026": "Aug",
        "09.2026": "Sep",
        "10.2026": "Oct",
        "11.2026": "Nov",
        "12.2026": "Dec",
        "06.2026": "Jun",
    }
    for r in range(min(25, len(df))):
        labels = [normalize_text(df.iloc[r, c]) for c in range(df.shape[1])]
        month_cols = {
            month_map[label]: c for c, label in enumerate(labels) if label in month_map
        }
        if len(month_cols) < 4:
            continue

        desc_col = next((c for c, label in enumerate(labels) if label == "Material Description"), None)
        mat_col = next((c for c, label in enumerate(labels) if label == "Material"), None)
        if desc_col is None:
            continue
        if mat_col is None:
            mat_col = desc_col - 1
        return r, month_cols, mat_col, desc_col
    return None


def parse_plant1410_forecast(path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(path)
    records = []

    for sheet in xl.sheet_names:
        if "total" in sheet.lower():
            continue
        df = pd.read_excel(path, sheet_name=sheet, header=None)
        table = find_forecast_table(df)
        if table is None:
            continue

        header_row, month_cols, mat_col, desc_col = table
        for r in range(header_row + 1, len(df)):
            desc = normalize_text(df.iloc[r, desc_col])
            if not desc or desc == "-":
                continue
            t, w, l = parse_dimensions_from_text(desc)
            spec_match = re.match(r"^([A-Z0-9][A-Z0-9\-/+ ]+)", desc.upper())
            spec = normalize_spec(spec_match.group(1)) if spec_match else ""
            if not spec or t <= 0:
                continue

            kg_per_she = calc_piece_weight_kg(t, w, l)
            monthly: dict[str, float] = {}
            for month, col in month_cols.items():
                qty = num(df.iloc[r, col])
                monthly[month] = monthly.get(month, 0) + qty * kg_per_she
            if sum(monthly.values()) <= 0:
                continue

            records.append(
                {
                    "source": "Plant1410",
                    "spec": spec,
                    "t": t,
                    "w": w if w else 1219.0,
                    "material_code": f"{spec}_{t}_{int(w) if w else 1219}_Common",
                    **monthly,
                }
            )
    return pd.DataFrame(records)


def load_forecast(forecast_dir: Path) -> pd.DataFrame:
    from pdf_forecast import load_pdf_forecasts

    frames = []
    for path in sorted(forecast_dir.glob("*.xlsx")):
        if "CARRIER" in path.name.upper():
            continue
        try:
            if "CPK" in path.name.upper():
                part = parse_cpk_forecast(path)
            else:
                part = parse_plant1410_forecast(path)
            if not part.empty:
                frames.append(part)
                print(f"  Parsed {path.name}: {len(part)} rows")
        except Exception as exc:
            print(f"Warning: failed to read {path.name}: {exc}")

    pdf_fc = load_pdf_forecasts(forecast_dir)
    if not pdf_fc.empty:
        frames.append(pdf_fc)

    if not frames:
        return pd.DataFrame()
    all_fc = pd.concat(frames, ignore_index=True)
    month_cols = [m for m in MONTHS_BALANCE if m in all_fc.columns]
    if not month_cols:
        return pd.DataFrame()
    grouped = (
        all_fc.groupby("material_code", as_index=False)[month_cols]
        .sum()
        .merge(
            all_fc.groupby("material_code", as_index=False)[["spec", "t", "w"]].first(),
            on="material_code",
        )
    )
    return grouped


def load_sample_forecast(sample_path: Path) -> pd.DataFrame:
    fc = pd.read_excel(sample_path, sheet_name="Forecast", header=None)
    month_kg_cols = {"Jun": 43, "Jul": 45, "Aug": 47, "Sep": 49, "Oct": 51}
    records = []
    for r in range(4, len(fc)):
        code = fc.iloc[r, 0]
        if pd.isna(code):
            continue
        row = {"material_code": normalize_text(code), "spec": normalize_spec(fc.iloc[r, 9])}
        for month, col in month_kg_cols.items():
            row[month] = num(fc.iloc[r, col])
        records.append(row)
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    month_cols = [m for m in MONTHS_BALANCE if m in df.columns]
    return df.groupby("material_code", as_index=False)[month_cols + ["spec"]].sum()


def match_forecast_to_material(material_code: str, forecast_df: pd.DataFrame) -> dict[str, float]:
    if forecast_df.empty:
        return {m: 0.0 for m in MONTHS_BALANCE}

    parsed = parse_material_code(material_code)
    if not parsed:
        return {m: 0.0 for m in MONTHS_BALANCE}

    # Exact code match
    exact = forecast_df[forecast_df["material_code"] == material_code]
    if not exact.empty:
        row = exact.iloc[0]
        return {m: num(row[m]) if m in row else 0.0 for m in MONTHS_BALANCE}

    # Spec + T match
    spec = parsed["spec"]
    t_val = parsed["t"]
    matched = forecast_df[
        forecast_df["spec"].apply(lambda s: spec_matches(s, spec))
    ]
    if t_val is not None:
        matched = matched[matched["t"].round(3) == round(t_val, 3)]

    result = {m: 0.0 for m in MONTHS_BALANCE}
    for _, row in matched.iterrows():
        for m in MONTHS_BALANCE:
            if m in row:
                result[m] += num(row[m])
    if sum(result.values()) == 0 and parsed:
        # Fallback: use average monthly sales from history
        return {}
    return result


def build_balance_data(
    materials: pd.DataFrame,
    stock_map: dict[str, float],
    forecast_df: pd.DataFrame,
) -> list[dict]:
    rows = []
    for idx, mat in materials.iterrows():
        code = mat["material_code"]
        sales = mat["sales"]
        avg = mat["avg_sales"] if mat["avg_sales"] > 0 else (
            sum(sales) / len(sales) if sales else 0.0
        )

        opening_stock = stock_map.get(code, 0.0)
        forecast = match_forecast_to_material(code, forecast_df)
        use_avg = not forecast or sum(forecast.values()) == 0

        monthly_forecast = {m: avg for m in MONTHS_BALANCE} if use_avg else forecast
        incoming = mat.get("incoming_may", 0.0)

        stock_end = opening_stock + incoming
        balance_row = {
            "item": mat["item"] if pd.notna(mat["item"]) else idx + 1,
            "material_code": code,
            "kind": mat["kind"],
            "t": mat["t"],
            "main_customer": mat["main_customer"],
            "sales": sales,
            "avg_sales": avg,
            "opening_stock": opening_stock,
            "incoming_may": incoming,
            "stock_after_may": stock_end,
        }

        for month in MONTHS_BALANCE:
            demand = monthly_forecast.get(month, avg)
            eta = 0.0
            stock_end = stock_end + eta - demand
            balance_row[f"{month}_eta"] = eta
            balance_row[f"{month}_forecast"] = demand
            balance_row[f"{month}_stock_end"] = stock_end

        balance_row["need_order"] = any(
            balance_row[f"{m}_stock_end"] < 0 for m in MONTHS_BALANCE
        )
        balance_row["max_shortage"] = min(
            balance_row[f"{m}_stock_end"] for m in MONTHS_BALANCE
        )
        rows.append(balance_row)
    return rows


def style_header(cell, bold=True, fill_color="D9E1F2"):
    cell.font = Font(bold=bold, size=10)
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_balance_excel(rows: list[dict], output_path: Path, report_date: datetime):
    wb = Workbook()
    ws = wb.active
    ws.title = "Balance sheet"

    title = f"Material balance review {report_date.strftime('%d %b %Y')}"
    ws.merge_cells("A1:AP1")
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=14)

    headers = [
        "Item", "Material Code", "Kind", "T", "Main customer",
        *MONTHS_SALES,
        "Average (kgs)",
        "Opening Stock (kg)",
        "Incoming May (kg)",
        "Stock after May (kg)",
    ]
    for month in MONTHS_BALANCE:
        headers.extend([f"{month} ETA", f"{month} Forecast", f"{month} Stock End"])

    headers.extend(["Need Order", "Max Shortage (kg)"])

    header_row = 4
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col, value=header)
        style_header(cell)

    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for r_idx, row in enumerate(rows, header_row + 1):
        values = [
            row["item"],
            row["material_code"],
            row["kind"],
            row["t"],
            row["main_customer"],
            *row["sales"],
            round(row["avg_sales"], 1),
            round(row["opening_stock"], 1),
            round(row["incoming_may"], 1),
            round(row["stock_after_may"], 1),
        ]
        for month in MONTHS_BALANCE:
            values.extend([
                round(row[f"{month}_eta"], 1),
                round(row[f"{month}_forecast"], 1),
                round(row[f"{month}_stock_end"], 1),
            ])
        values.extend([
            "YES" if row["need_order"] else "",
            round(row["max_shortage"], 1) if row["need_order"] else "",
        ])

        for c_idx, value in enumerate(values, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = border
            if isinstance(value, float):
                cell.number_format = "#,##0.0"
            if c_idx >= len(values) - 1 and value == "YES":
                cell.fill = PatternFill("solid", fgColor="FFC7CE")
                cell.font = Font(color="9C0006", bold=True)
            if c_idx == len(values) and isinstance(value, (int, float)) and value < 0:
                cell.fill = PatternFill("solid", fgColor="FFC7CE")

    # Summary row
    sum_row = r_idx + 2
    ws.cell(row=sum_row, column=1, value="Summary").font = Font(bold=True)
    for m_idx, month in enumerate(MONTHS_BALANCE):
        forecast_col = 5 + 15 + 4 + m_idx * 3 + 1
        stock_col = forecast_col + 1
        f_total = sum(row[f"{month}_forecast"] for row in rows)
        s_total = sum(row[f"{month}_stock_end"] for row in rows)
        ws.cell(row=sum_row, column=forecast_col, value=round(f_total, 1))
        ws.cell(row=sum_row, column=stock_col, value=round(s_total, 1))

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["E"].width = 28

    # Order recommendation sheet
    ws2 = wb.create_sheet("Order Required")
    order_headers = [
        "Item", "Material Code", "Kind", "T", "Main customer",
        "Opening Stock (kg)", "Jun Stock End", "Jul Stock End",
        "Aug Stock End", "Sep Stock End", "Oct Stock End",
        "Max Shortage (kg)", "Suggested Order (kg)",
    ]
    for col, header in enumerate(order_headers, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        style_header(cell, fill_color="FCE4D6")

    order_rows = [r for r in rows if r["need_order"]]
    order_rows.sort(key=lambda x: x["max_shortage"])
    for r_idx, row in enumerate(order_rows, 2):
        suggested = abs(row["max_shortage"]) if row["max_shortage"] < 0 else 0
        values = [
            row["item"], row["material_code"], row["kind"], row["t"],
            row["main_customer"], round(row["opening_stock"], 1),
            round(row["Jun_stock_end"], 1), round(row["Jul_stock_end"], 1),
            round(row["Aug_stock_end"], 1), round(row["Sep_stock_end"], 1),
            round(row["Oct_stock_end"], 1),
            round(row["max_shortage"], 1), round(suggested, 1),
        ]
        for c_idx, value in enumerate(values, 1):
            ws2.cell(row=r_idx, column=c_idx, value=value)

    OUTPUT_DIR.mkdir(exist_ok=True)
    wb.save(output_path)
    return len(order_rows)


def main():
    report_date = datetime.now()
    print("Loading materials from sample balance sheet...")
    materials = load_balance_materials(SAMPLE_PATH)
    print(f"  Materials: {len(materials)}")

    print("Loading stock...")
    stock_map = load_stock_by_material(STOCK_PATH, materials)
    stocked = sum(1 for v in stock_map.values() if v > 0)
    print(f"  Materials with stock: {stocked}")

    print("Loading forecast...")
    from data_loaders import resolve_forecast_dir

    forecast_df = load_forecast(resolve_forecast_dir())
    if forecast_df.empty:
        print("  No xlsx forecast parsed, using sample forecast...")
        forecast_df = load_sample_forecast(SAMPLE_PATH)
    print(f"  Forecast entries: {len(forecast_df)}")

    print("Building balance...")
    rows = build_balance_data(materials, stock_map, forecast_df)

    output_name = f"Material Balance Review {report_date.strftime('%d.%m.%Y_%H%M')}.xlsx"
    output_path = OUTPUT_DIR / output_name
    order_count = write_balance_excel(rows, output_path, report_date)

    total_stock = sum(stock_map.values())
    need_order = sum(1 for r in rows if r["need_order"])
    print(f"\nReport saved: {output_path}")
    print(f"  Total stock (matched): {total_stock:,.0f} kg")
    print(f"  Materials needing order: {need_order}")
    print(f"  Order list items: {order_count}")


if __name__ == "__main__":
    main()
