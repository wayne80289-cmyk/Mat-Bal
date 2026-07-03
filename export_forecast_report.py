# -*- coding: utf-8 -*-
"""Export integrated Forecast/ vendor report (partial customers only, U1)."""

from __future__ import annotations

import datetime
from pathlib import Path

import pandas as pd

from data_loaders import FORECAST_SCOPE_LABEL, OUTPUT_DIR, build_forecast_integrated_report

BASE_DIR = Path(__file__).resolve().parent


def main() -> Path:
    print(f"=== Forecast 各家整合報告 ===")
    print(f"範圍: {FORECAST_SCOPE_LABEL}")
    sheets = build_forecast_integrated_report()
    info = sheets["Forecast_說明"]
    detail = sheets["Forecast_各家明細"]
    summary = sheets["Forecast_整合彙總"]
    print(f"各家明細: {len(detail)} 列")
    print(f"整合彙總 Material: {len(summary)} 項")
    if not detail.empty and "Source_File" in detail.columns:
        print(f"檔案數: {detail['Source_File'].nunique()}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    ts = datetime.datetime.now().strftime("%d.%m.%Y_%H%M")
    out_path = OUTPUT_DIR / f"Forecast整合報告 {ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    print(f"\nSaved: {out_path}")
    return out_path


if __name__ == "__main__":
    main()
