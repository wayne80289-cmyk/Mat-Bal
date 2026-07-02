# -*- coding: utf-8 -*-
"""Export Stock Material Code matching rules to Excel for review."""

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT_DIR = Path(__file__).resolve().parent / "History Balance"
OUT_PATH = OUT_DIR / "Stock-Material-Code-Matching-Rules.xlsx"


def write_sheets(writer: pd.ExcelWriter) -> None:
  # Sheet 1: 說明
    pd.DataFrame([
        {"項目": "文件名稱", "內容": "Stock Material Code 配對庫存規則清單"},
        {"項目": "分析來源", "內容": "History Balance/ 2026-01～06 共 7 份工作簿 Stock 工作表"},
        {"項目": "對照工作表", "內容": "Balance sheet (Update)、Act order、Original material code（01月）"},
        {"項目": "產出日期", "內容": datetime.now().strftime("%Y-%m-%d %H:%M")},
        {"項目": "用途", "內容": "供檢查 MS004 庫存配對 Material Code 之業務規則"},
    ]).to_excel(writer, sheet_name="說明", index=False)

    # Sheet 2: Stock 頁結構
    pd.DataFrame([
        {"項目": "工作表名稱", "規則": "Stock {日期}，例：Stock 19.05.2026、Stock 22.06.26 0930"},
        {"項目": "標題列", "規則": "第 4 列（Excel row 4）"},
        {"項目": "Material Code 欄", "規則": "A 欄，欄名 Material Code 或 Material code"},
        {"項目": "資料起始列", "規則": "第 5 列（標題列下一列）"},
        {"項目": "A1 註記", "規則": "常放目前篩選/操作中的 Material Code 範例"},
        {"項目": "資料來源", "規則": "MS004 匯出（與 Stock/MS004-*.xls 相同結構）"},
        {"項目": "庫存重量加總", "規則": "P REMAIN WT（kg）"},
    ]).to_excel(writer, sheet_name="Stock頁結構", index=False)

    # Sheet 3: MS004 欄位
    pd.DataFrame([
        {"MS004欄位": "OWNER", "用途": "庫存歸屬，配對前必查（僅 PTT）"},
        {"MS004欄位": "MAT SPEC", "用途": "實際鋼種規格（可與 Material Code 前綴不同）"},
        {"MS004欄位": "T", "用途": "厚度"},
        {"MS004欄位": "W", "用途": "寬度"},
        {"MS004欄位": "L", "用途": "長度"},
        {"MS004欄位": "MAKER CODE", "用途": "製造商，影響 Code 尾碼"},
        {"MS004欄位": "CUST CODE", "用途": "客戶，判斷是否為專案專料"},
        {"MS004欄位": "P REMAIN WT", "用途": "剩餘重量（kg），加總後 ÷1000 為噸"},
    ]).to_excel(writer, sheet_name="MS004欄位對照", index=False)

    # Sheet 4: 命名格式
    pd.DataFrame([
        {"格式類型": "格式 A（現行，2026-02起）", "結構": "{Common_Group}_{厚度}_{寬度}_{尾碼}", "範例": "SPCC-SD_0.6_1219_Common"},
        {"格式類型": "格式 A", "結構": "{Common_Group}_{厚度}_{寬度}_{尾碼}", "範例": "SECC-16/16_1.2_1219_CHINA STEEL"},
        {"格式類型": "格式 A", "結構": "{Common_Group}_{厚度}_{寬度}_{尾碼}", "範例": "SAPH440-PO_2.9_1219_COMMON"},
        {"格式類型": "格式 A", "結構": "{Common_Group}_{厚度}_{寬度}_{尾碼}", "範例": "SPHC-P/O_2.6_1086_CASH"},
        {"格式類型": "格式 A", "結構": "{Common_Group}_{厚度}_{寬度}_{尾碼}", "範例": "JSC270C_0.6_1226_Special control"},
        {"格式類型": "格式 B（2026-01 Original material code）", "結構": "{MAT SPEC}_{T}_{W}_0_{Maker或PTT}", "範例": "SGCD1-F08_0.7_1260_0_China Steel TW"},
        {"格式類型": "格式 B", "結構": "{MAT SPEC}_{T}_{W}_0_{Maker或PTT}", "範例": "CR_2.2_1219_0_PTT"},
    ]).to_excel(writer, sheet_name="Code命名格式", index=False)

    # Sheet 5: 配對規則主表
    rules = [
        ("U-Stock-01", "OWNER 篩選", "必要", "僅 OWNER=PTT 納入配對與 Balance；非 PTT 不填 Code、不計平衡"),
        ("U-Stock-02", "主檔 RD004", "必要", "Material Code 來自 RD004（Act order + Balance sheet）；一 Code 對一 Common Group"),
        ("U-Stock-03", "配對主鍵", "必要", "1.Common Group 2.厚度T 3.寬度W；MS004 無原生 Code，需 allocate 貼上"),
        ("U-Stock-04", "MAT SPEC 跨規格", "對照", "實際 MAT SPEC 可與 Code 前綴不同，同 RD004 群組即可（見 MAT SPEC對照表）"),
        ("U-Stock-05", "尾碼 _Common", "尾碼", "跨客戶共通備貨池；同 Code 多列加總進 Balance"),
        ("U-Stock-06", "尾碼 Maker", "尾碼", "依 MAKER CODE 區分，例：_CHINA STEEL、_SSI、_BAOSHAN"),
        ("U-Stock-07", "尾碼 _CASH", "尾碼", "非標準寬度或現金採購，例：SPHC-P/O_2.6_1086_CASH"),
        ("U-Stock-08", "厚度容差", "容差", "0.55→0.5、3.02→3、0.75→0.7；建議 round(T,1) 或依 RD004"),
        ("U-Stock-09", "寬度容差", "容差", "Coil 共通料 Code 常固定 1219；實際 W 1270/1225/1200/1195/1260 可同碼"),
        ("U-Stock-10", "客戶專料", "排除", "專案專購 CUST CODE 特定客戶 → Material Code 留空（約85~90% PTT列）"),
        ("U-Stock-11", "Allocate 方式", "作業", "手動貼上/向下複製，非公式；約600~800列有 Code（10~15%）"),
        ("U-Stock-12", "同 Code 合併", "加總", "相同 Material Code 之 P REMAIN WT 加總 → Balance 期初現貨（噸）"),
        ("U-Stock-13", "Balance 對應", "對應", "群組型 Code 與 Balance 重疊約88~100%；另有尺寸型命名"),
        ("U-Stock-14", "不納入配對", "排除", "OWNER≠PTT、Code空白、Code=0、REMAIN WT=0 等（見排除清單）"),
        ("U-Stock-15", "CARRIER 例外", "例外", "無 Material Code 欄；改用 Master list G00料號+Spec+T+W+L"),
    ]
    pd.DataFrame(rules, columns=["規則編號", "規則名稱", "類別", "規則內容"]).to_excel(
        writer, sheet_name="配對規則清單", index=False
    )

    # Sheet 6: MAT SPEC 對照
    pd.DataFrame([
        {"MS004_MAT_SPEC": "SPCC", "配對_Code前綴": "SPCC-SD", "備註": "同 Common Group"},
        {"MS004_MAT_SPEC": "SPCC-M", "配對_Code前綴": "SPCC-SD", "備註": ""},
        {"MS004_MAT_SPEC": "SGCC-Z08", "配對_Code前綴": "SGCC-Z22", "備註": ""},
        {"MS004_MAT_SPEC": "SGCD2-Z08", "配對_Code前綴": "SGCC-Z22", "備註": ""},
        {"MS004_MAT_SPEC": "SGCD1-Z18", "配對_Code前綴": "SGCC-Z18", "備註": ""},
        {"MS004_MAT_SPEC": "SAPH440-P/O", "配對_Code前綴": "SAPH440-PO", "備註": ""},
        {"MS004_MAT_SPEC": "SAPH400-P/O", "配對_Code前綴": "SAPH440-P/O", "備註": ""},
        {"MS004_MAT_SPEC": "JSH590R-P/O", "配對_Code前綴": "SPHC-PO", "備註": ""},
        {"MS004_MAT_SPEC": "FC440", "配對_Code前綴": "JSC440W", "備註": ""},
        {"MS004_MAT_SPEC": "SPFC440", "配對_Code前綴": "JSC440W", "備註": ""},
        {"MS004_MAT_SPEC": "SECC-AF,E16/E16", "配對_Code前綴": "SECC-16/16", "備註": ""},
        {"MS004_MAT_SPEC": "SPCD", "配對_Code前綴": "SPCC-SD 或 SPCEN-SD", "備註": "依群組"},
        {"MS004_MAT_SPEC": "SPCEN-SD", "配對_Code前綴": "SPCC-SD 或 SPCEN-SD", "備註": "依群組"},
    ]).to_excel(writer, sheet_name="MAT SPEC對照", index=False)

    # Sheet 7: 排除情況
    pd.DataFrame([
        {"情況": "OWNER ≠ PTT", "處理": "排除，不填 Material Code"},
        {"情況": "Material Code 空白", "處理": "不計入 Balance（專料或待處理）"},
        {"情況": "Material Code = 0", "處理": "視為未配對（01月常見）"},
        {"情況": "P REMAIN WT = 0", "處理": "可保留列但重量不計"},
        {"情況": "CUST CODE 為專案客戶（如 MING TAI）", "處理": "通常不配碼"},
    ]).to_excel(writer, sheet_name="排除情況", index=False)

    # Sheet 8: 各月統計
    pd.DataFrame([
        {"月份": "01/2026", "檔案": "All small volume…22.01.2026", "Stock工作表": "Stock 17.12.25", "總列數": 5183, "已填Code": 601, "唯一Code數": 115},
        {"月份": "02/2026", "檔案": "All small volume…14.02.2026", "Stock工作表": "Stock 13.02.2026", "總列數": 4887, "已填Code": 621, "唯一Code數": 100},
        {"月份": "03/2026", "檔案": "All small volume…16.03.2026", "Stock工作表": "Stock 20.03.26 0930", "總列數": 5070, "已填Code": 651, "唯一Code數": 100},
        {"月份": "04/2026", "檔案": "All small volume…18.04.2026", "Stock工作表": "Stock 18.04.2026", "總列數": 5290, "已填Code": 808, "唯一Code數": 98},
        {"月份": "05/2026", "檔案": "All Customer review May…", "Stock工作表": "Stock 19.05.2026", "總列數": 5196, "已填Code": 612, "唯一Code數": 91},
        {"月份": "05/2026", "檔案": "CARRIER material review…", "Stock工作表": "Stock 18.05.02", "總列數": 1356, "已填Code": "無此欄", "唯一Code數": "—"},
        {"月份": "06/2026", "檔案": "All Customer review Jun…", "Stock工作表": "Stock 22.06.26 0930", "總列數": 5026, "已填Code": 581, "唯一Code數": 88},
    ]).to_excel(writer, sheet_name="各月統計", index=False)

    # Sheet 9: 程式對照
    pd.DataFrame([
        {"規則": "OWNER=PTT", "Excel實際": "✓", "程式現況": "✓", "建議": "保持"},
        {"規則": "Spec 群組模糊比對", "Excel實際": "RD004 多規格", "程式現況": "contains(spec[:6])", "建議": "需 RD004 Common Group 表"},
        {"規則": "厚度", "Excel實際": "容差比對", "程式現況": "round(3) 精確", "建議": "加入容差"},
        {"規則": "寬度", "Excel實際": "常忽略（1219群組）", "程式現況": "非1219才篩W", "建議": "符合 U-Stock-09"},
        {"規則": "客戶專料", "Excel實際": "不配碼", "程式現況": "未區分", "建議": "可加 CUST CODE 規則"},
        {"規則": "Maker 尾碼", "Excel實際": "影響 Code", "程式現況": "未使用", "建議": "進階版可加"},
    ]).to_excel(writer, sheet_name="程式實作對照", index=False)

    # Sheet 10: 配對流程
    pd.DataFrame([
        {"步驟": 1, "動作": "MS004 匯出 → Stock 頁"},
        {"步驟": 2, "動作": "檢查 OWNER = PTT？否 → 跳過"},
        {"步驟": 3, "動作": "查 RD004 Common Group（Spec + T [+ W]）"},
        {"步驟": 4, "動作": "是否專案專料（CUST CODE）？是 → Material Code 留空"},
        {"步驟": 5, "動作": "決定尾碼（Common / Maker / CASH）"},
        {"步驟": 6, "動作": "填入 A 欄 Material Code（手動 allocate）"},
        {"步驟": 7, "動作": "同 Code 之 P REMAIN WT 加總 → Balance 表期初現貨"},
    ]).to_excel(writer, sheet_name="配對流程", index=False)


def style_workbook(path: Path) -> None:
    wb = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    wrap = Alignment(wrap_text=True, vertical="top")

    for ws in wb.worksheets:
        if ws.max_row < 1:
            continue
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.alignment = wrap
        for col in range(1, ws.max_column + 1):
            letter = get_column_letter(col)
            max_len = 12
            for row in range(1, min(ws.max_row + 1, 200)):
                val = ws.cell(row=row, column=col).value
                if val is not None:
                    max_len = max(max_len, min(len(str(val)) + 2, 60))
            ws.column_dimensions[letter].width = max_len
        ws.freeze_panes = "A2"

    wb.save(path)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
        write_sheets(writer)
    style_workbook(OUT_PATH)
    print(f"Saved: {OUT_PATH}")


if __name__ == "__main__":
    main()
