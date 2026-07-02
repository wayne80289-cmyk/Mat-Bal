# -*- coding: utf-8 -*-
"""
Penta Thick Spot & Future Stock Balance System
現貨備貨與期貨備貨雙軌統計 → 現貨緊急調貨計畫 + 期貨備貨計畫
"""

from __future__ import annotations

import datetime
from pathlib import Path

import pandas as pd
from dateutil.relativedelta import relativedelta

from data_loaders import (
    TARGET_MONTHS,
    get_data_source_summary,
    load_client_forecast,
    load_mp008,
    load_ms004,
    load_order_history,
    load_rd004_master,
    load_so003,
)
from stock_matching import (
    allocate_ms004_stock,
    allocate_mp008_inbound,
    enrich_material_master,
    inbound_by_material_month,
    load_rules_workbook_sheets,
    resolve_rules_path,
)

OUTPUT_DIR = Path(__file__).resolve().parent / "Output"


class MaterialOrderBalanceSystem:
    def __init__(self, buffer_stock_policy_days=30):
        self.buffer_days = buffer_stock_policy_days
        self.LT_PRODUCTION = 2.0
        self.LT_SHIPPING = 0.5
        self.LT_PROCESSING = 1.0
        self.LT_SAFETY_BUFFER = 1.0
        self.TOTAL_LEAD_TIME_MONTHS = 4.5

    def integrate_sales_forecast(self, client_forecast_df, order_history_df, target_months):
        integrated_forecast = {}
        fg_codes = set()
        if not client_forecast_df.empty and "Material_Code" in client_forecast_df.columns:
            fg_codes.update(client_forecast_df["Material_Code"].unique())
        if not order_history_df.empty:
            fg_codes.update(order_history_df["FG_Code"].unique())

        for fg_code in fg_codes:
            integrated_forecast[fg_code] = {}
            for month in target_months:
                client_f = 0.0
                if (
                    not client_forecast_df.empty
                    and "Material_Code" in client_forecast_df.columns
                    and month in client_forecast_df.columns
                ):
                    client_f = client_forecast_df[
                        client_forecast_df["Material_Code"] == fg_code
                    ][month].sum()

                calculated_f = 0.0
                if client_f == 0 and not order_history_df.empty:
                    history = order_history_df[order_history_df["FG_Code"] == fg_code]
                    if not history.empty:
                        calculated_f = history[["M-1", "M-2", "M-3"]].mean(axis=1).values[0] / 1000.0

                integrated_forecast[fg_code][month] = client_f + calculated_f
        return integrated_forecast

    def process_so003_orders(self, so003_df):
        so_summary = {}
        df = so003_df.copy()
        df["Due_Month"] = pd.to_datetime(df["Due Date"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m")

        for _, row in df.iterrows():
            fg_code = row.get("FG Code") or row.get("Basic Code")
            due_month = row.get("Due_Month")
            delivery_ton = row.get("Delivery Kg", 0) / 1000.0
            bal_ton = row.get("Bal Kg", 0) / 1000.0

            if pd.isna(due_month) or not fg_code:
                continue

            so_summary.setdefault(fg_code, {})
            so_summary[fg_code].setdefault(due_month, {"Delivery_Ton": 0, "Bal_Ton": 0})
            so_summary[fg_code][due_month]["Delivery_Ton"] += delivery_ton
            so_summary[fg_code][due_month]["Bal_Ton"] += bal_ton
        return so_summary

    def clean_and_allocate_stock(self, ms004_raw_df, rd004_master_df):
        """U-Stock rules: PTT + spec/T/W match → allocatable Initial_Stock pool."""
        return allocate_ms004_stock(ms004_raw_df, rd004_master_df)

    def filter_on_way_po(self, mp008_raw_df, rd004_master_df):
        """U4: Close Flag=False; Mat Spec match with customer-first allocation."""
        return allocate_mp008_inbound(mp008_raw_df, rd004_master_df)

    def run_material_balance(
        self,
        cleaned_stock_df,
        active_po_df,
        integrated_forecast,
        so_summary,
        rd004_master_df,
        target_months,
        buffer_limit_ton=10.0,
    ):
        balance_reports = []
        all_materials = rd004_master_df["Material_Code"].unique()

        fg_to_mat: dict[str, str] = {}
        for _, m in rd004_master_df.iterrows():
            fg_to_mat[m["FG_Code"]] = m["Material_Code"]
            for fg in str(m.get("FG_Codes_All", "")).split(","):
                if fg.strip():
                    fg_to_mat[fg.strip()] = m["Material_Code"]

        inbound_pool = inbound_by_material_month(active_po_df)
        inbound_customer_pool: dict[tuple[str, str], float] = {}
        inbound_common_pool: dict[tuple[str, str], float] = {}
        if not active_po_df.empty and "Alloc_Priority" in active_po_df.columns:
            for _, prow in active_po_df.iterrows():
                key = (str(prow["Material_Code"]).strip(), str(prow.get("ETA_Month", "")).strip())
                if not key[0] or not key[1]:
                    continue
                qty = float(prow.get("Quantity", 0) or 0)
                if prow.get("Alloc_Priority") == "Customer-Match":
                    inbound_customer_pool[key] = inbound_customer_pool.get(key, 0.0) + qty
                else:
                    inbound_common_pool[key] = inbound_common_pool.get(key, 0.0) + qty

        for mat_code in all_materials:
            mat_info = rd004_master_df[rd004_master_df["Material_Code"] == mat_code].iloc[0]
            fg_code_mapped = mat_info["FG_Code"]

            current_stock_qty = cleaned_stock_df[
                cleaned_stock_df["Material_Code"] == mat_code
            ]["Quantity"].sum()
            available_balance = current_stock_qty

            mat_timeline = {
                "Material_Code": mat_code,
                "FG_Code": fg_code_mapped,
                "Main_Customer": mat_info.get("Main_Customer", ""),
                "Initial_Stock_Ton": round(current_stock_qty, 3),
                "Initial_Stock_Allocatable": round(current_stock_qty, 3),
            }

            for month_str in target_months:
                forecast_demand = integrated_forecast.get(mat_code, {}).get(month_str, 0)
                if forecast_demand == 0:
                    forecast_demand = integrated_forecast.get(fg_code_mapped, {}).get(month_str, 0)

                so_data = {"Delivery_Ton": 0, "Bal_Ton": 0}
                for fg, months in so_summary.items():
                    if fg_to_mat.get(fg, fg) != mat_code and fg != fg_code_mapped:
                        continue
                    if month_str in months:
                        so_data["Delivery_Ton"] += months[month_str]["Delivery_Ton"]
                        so_data["Bal_Ton"] += months[month_str]["Bal_Ton"]

                so_delivery = so_data["Delivery_Ton"]
                so_balance = so_data["Bal_Ton"]
                effective_demand = max(forecast_demand, so_delivery + so_balance) - so_delivery

                inbound = inbound_pool.get((mat_code, month_str), 0.0)
                if inbound == 0.0 and not active_po_df.empty:
                    inbound = active_po_df[
                        (active_po_df["Material_Code"] == mat_code)
                        & (active_po_df["ETA_Month"] == month_str)
                    ]["Quantity"].sum()
                inbound_customer = inbound_customer_pool.get((mat_code, month_str), 0.0)
                inbound_common = inbound_common_pool.get((mat_code, month_str), 0.0)

                available_balance = available_balance + inbound - effective_demand

                mat_timeline[f"{month_str}_Forecast_Ton"] = round(forecast_demand, 3)
                mat_timeline[f"{month_str}_SO_Delivered"] = round(so_delivery, 3)
                mat_timeline[f"{month_str}_SO_Balance(欠交)"] = round(so_balance, 3)
                mat_timeline[f"{month_str}_Inbound_PO"] = round(inbound, 3)
                mat_timeline[f"{month_str}_Inbound_PO_同客戶"] = round(inbound_customer, 3)
                mat_timeline[f"{month_str}_Inbound_PO_共通池"] = round(inbound_common, 3)
                mat_timeline[f"{month_str}_Effective_Demand"] = round(effective_demand, 3)
                mat_timeline[f"{month_str}_Final_Balance"] = round(available_balance, 3)

                if available_balance < buffer_limit_ton:
                    shortage_qty = abs(available_balance - buffer_limit_ton)
                    po_deadline_date = datetime.datetime.strptime(
                        f"{month_str}-15", "%Y-%m-%d"
                    ) - relativedelta(months=int(self.TOTAL_LEAD_TIME_MONTHS))
                    mat_timeline[f"{month_str}_Alert"] = "SHORTAGE"
                    mat_timeline[f"{month_str}_Shortage_Ton"] = round(shortage_qty, 3)
                    mat_timeline[f"{month_str}_PO_Deadline"] = po_deadline_date.strftime("%Y-%m-%d")
                    available_balance = buffer_limit_ton
                else:
                    mat_timeline[f"{month_str}_Alert"] = "SAFE"
                    mat_timeline[f"{month_str}_Shortage_Ton"] = 0.0
                    mat_timeline[f"{month_str}_PO_Deadline"] = "-"

            balance_reports.append(mat_timeline)

        return pd.DataFrame(balance_reports)


def build_spot_urgent_plan(
    balance_df: pd.DataFrame,
    so003_df: pd.DataFrame,
    rd004: pd.DataFrame,
    target_months: list[str],
    near_months: int = 2,
) -> pd.DataFrame:
    """
    現貨緊急調貨計畫：SO003 欠交 + 近期需求 > MS004 現貨 → 需緊急調撥。
    """
    fg_to_mat: dict[str, str] = {}
    for _, m in rd004.iterrows():
        fg_to_mat[m["FG_Code"]] = m["Material_Code"]
        for fg in str(m.get("FG_Codes_All", "")).split(","):
            if fg.strip():
                fg_to_mat[fg.strip()] = m["Material_Code"]

    focus_months = target_months[:near_months]
    plans = []

    for _, row in balance_df.iterrows():
        mat = row["Material_Code"]
        stock = row.get("Initial_Stock_Ton", 0)
        total_backorder = sum(row.get(f"{m}_SO_Balance(欠交)", 0) for m in focus_months)
        total_demand = sum(row.get(f"{m}_Effective_Demand", 0) for m in focus_months)
        first_month = focus_months[0] if focus_months else target_months[0]
        first_balance = row.get(f"{first_month}_Final_Balance", stock)

        if total_backorder <= 0 and first_balance >= 10:
            continue

        shortfall = max(total_backorder + total_demand - stock, 0)
        if total_backorder > 0 and first_balance < 10:
            urgency = "CRITICAL"
            action = "立即調撥現貨 / 同 Common Group 挪用"
        elif total_backorder > 0:
            urgency = "HIGH"
            action = "優先調撥滿足 SO003 欠交"
        elif first_balance < 10:
            urgency = "MEDIUM"
            action = "補充現貨至緩衝水位"
        else:
            continue

        mat_so = so003_df.copy()
        mat_so["Mat_Code"] = mat_so["FG Code"].map(lambda x: fg_to_mat.get(x, ""))
        mat_so = mat_so[(mat_so["Mat_Code"] == mat) & (mat_so["Bal Kg"] > 0)]
        customers = ", ".join(sorted(mat_so["Customer"].unique()[:5]))

        plans.append({
            "Material_Code": mat,
            "FG_Code": row.get("FG_Code", ""),
            "Main_Customer": row.get("Main_Customer", ""),
            "SO003_Customers_欠交": customers,
            "現貨庫存_Ton": round(stock, 3),
            "SO003_欠交_Ton": round(total_backorder, 3),
            f"近{len(focus_months)}月_有效需求_Ton": round(total_demand, 3),
            "缺口_Ton": round(shortfall, 3),
            "緊急程度": urgency,
            "調貨建議": action,
        })

    df = pd.DataFrame(plans)
    if df.empty:
        return df
    order = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2}
    df["_sort"] = df["緊急程度"].map(order)
    return df.sort_values(["_sort", "缺口_Ton"], ascending=[True, False]).drop(columns="_sort")


def build_futures_stocking_plan(balance_df: pd.DataFrame, rd004: pd.DataFrame, target_months: list[str]) -> pd.DataFrame:
    """期貨備貨計畫：SHORTAGE 月份需下 PO，含逆推 Deadline 與建議採購量。"""
    moq_map = rd004.set_index("Material_Code")["MOQ"].to_dict() if "MOQ" in rd004.columns else {}
    plans = []

    for _, row in balance_df.iterrows():
        mat = row["Material_Code"]
        for month in target_months:
            if row.get(f"{month}_Alert") != "SHORTAGE":
                continue
            shortage = row.get(f"{month}_Shortage_Ton", 0)
            moq = moq_map.get(mat, 0) or 0
            order_qty = max(shortage, moq) if moq > 0 else shortage
            plans.append({
                "Material_Code": mat,
                "FG_Code": row.get("FG_Code", ""),
                "Main_Customer": row.get("Main_Customer", ""),
                "需求月份": month,
                "Forecast_Ton": row.get(f"{month}_Forecast_Ton", 0),
                "SO003_欠交_Ton": row.get(f"{month}_SO_Balance(欠交)", 0),
                "在途_PO_Ton": row.get(f"{month}_Inbound_PO", 0),
                "缺口_Ton": round(shortage, 3),
                "MOQ_Ton": moq,
                "建議下單_Ton": round(order_qty, 3),
                "PO_Deadline": row.get(f"{month}_PO_Deadline", ""),
                "備貨類型": "期貨採購",
            })

    df = pd.DataFrame(plans)
    if df.empty:
        return df
    return df.sort_values(["PO_Deadline", "缺口_Ton"], ascending=[True, False])


def run_full_pipeline(target_months: list[str] | None = None) -> pd.DataFrame:
    target_months = target_months or TARGET_MONTHS
    sources = get_data_source_summary()
    print("=== Penta Thick 現貨/期貨整合計畫系統 ===")
    print("資料來源:")
    for k, v in sources.items():
        print(f"  {k}: {v}")

    rd004 = load_rd004_master()
    ms004 = load_ms004()
    so003 = load_so003()
    mp008_raw = load_mp008()
    order_history = load_order_history()
    client_forecast = load_client_forecast(target_months)

    engine = MaterialOrderBalanceSystem()
    print(f"\n  RD004 materials: {len(rd004)}")
    print(f"  MS004 PTT stock rows: {len(ms004)}")
    print(f"  SO003 order rows: {len(so003)}")
    forecast_map = engine.integrate_sales_forecast(client_forecast, order_history, target_months)
    so_summary = engine.process_so003_orders(so003)
    cleaned_stock = engine.clean_and_allocate_stock(ms004, rd004)
    active_po = engine.filter_on_way_po(mp008_raw, rd004)
    mp008 = active_po

    print(f"  Open PO rows (allocatable): {len(active_po)}")
    print(f"  Allocatable stock materials: {len(cleaned_stock)}")
    print(f"  Forecast materials: {len(client_forecast)}")

    result = engine.run_material_balance(
        cleaned_stock, active_po, forecast_map, so_summary, rd004, target_months
    )

    spot_plan = build_spot_urgent_plan(result, so003, rd004, target_months)
    future_plan = build_futures_stocking_plan(result, rd004, target_months)

    material_master = enrich_material_master(
        rd004,
        allocatable_stock=cleaned_stock,
        allocated_po=active_po,
    )
    rules_sheets = load_rules_workbook_sheets()

    OUTPUT_DIR.mkdir(exist_ok=True)
    ts = datetime.datetime.now().strftime("%d.%m.%Y_%H%M")
    out_path = OUTPUT_DIR / f"Supply Plan {ts}.xlsx"
    shortage = result[result.filter(like="_Alert").apply(lambda r: (r == "SHORTAGE").any(), axis=1)]

    source_rows = pd.DataFrame([{"項目": k, "檔案": v} for k, v in sources.items()])

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        spot_plan.to_excel(writer, sheet_name="現貨緊急調貨計畫", index=False)
        future_plan.to_excel(writer, sheet_name="期貨備貨計畫", index=False)
        result.to_excel(writer, sheet_name="Balance", index=False)
        shortage.to_excel(writer, sheet_name="Order Required", index=False)
        so003.to_excel(writer, sheet_name="SO003_Source", index=False)
        mp008.to_excel(writer, sheet_name="MP008_OnWay", index=False)
        cleaned_stock.to_excel(writer, sheet_name="MS004_Allocatable_Stock", index=False)
        client_forecast.to_excel(writer, sheet_name="Forecast_Source", index=False)
        source_rows.to_excel(writer, sheet_name="Data Sources", index=False)
        material_master.to_excel(writer, sheet_name="Material Master", index=False)
        for sheet_name, sheet_df in rules_sheets.items():
            safe = f"Rules_{sheet_name}"[:31]
            sheet_df.to_excel(writer, sheet_name=safe, index=False)

    print(f"\nReport saved: {out_path}")
    print(f"  Balance materials: {len(result)}")
    print(f"  現貨緊急調貨: {len(spot_plan)} 項")
    print(f"  期貨備貨計畫: {len(future_plan)} 項")
    print(f"  Material Master (rules: {resolve_rules_path().name}): {len(material_master)} 項")
    return result


if __name__ == "__main__":
    df = run_full_pipeline()
    cols = [
        "Material_Code", "Initial_Stock_Ton",
        "2026-06_SO_Balance(欠交)", "2026-06_Final_Balance", "2026-06_Alert",
    ]
    show = [c for c in cols if c in df.columns]
    print("\n" + "=" * 90)
    print("Sample Balance output:")
    print(df[show].head(10).to_string(index=False))
    print("=" * 90)
