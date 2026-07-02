# -*- coding: utf-8 -*-
"""PDF forecast parsers for material balance."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError:
    RapidOCR = None

# Shared helpers imported from material_balance at runtime to avoid circular imports
# These will be passed in or duplicated minimally

MONTHS_BALANCE = ["Jun", "Jul", "Aug", "Sep", "Oct"]

MONTH_ALIASES = {
    "JUNE": "Jun", "JUN": "Jun", "06/2026": "Jun", "06.2026": "Jun",
    "JULY": "Jul", "JUL": "Jul", "07/2026": "Jul", "07.2026": "Jul",
    "AUGUST": "Aug", "AUG": "Aug", "08/2026": "Aug", "08.2026": "Aug",
    "SEPTEMBER": "Sep", "SEP": "Sep", "09/2026": "Sep", "09.2026": "Sep",
    "OCTOBER": "Oct", "OCT": "Oct", "10/2026": "Oct", "10.2026": "Oct",
    "NOVEMBER": "Nov", "NOV": "Nov", "11/2026": "Nov", "11.2026": "Nov",
}


def _num(value) -> float:
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


def _text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _norm_spec(value) -> str:
    return _text(value).upper().replace("?", "")


def _parse_dims(text: str) -> tuple[float, float, float]:
    text = _text(text)
    m = re.search(r"([\d.]+)\s*[xX]\s*([\d.]+)\s*[xX]\s*([\d.]+)", text)
    if m:
        return _num(m.group(1)), _num(m.group(2)), _num(m.group(3))
    return 0.0, 0.0, 0.0


def _calc_kg(t_mm, w_mm, l_mm, kind: str = "CR") -> float:
    t, w, l = _num(t_mm), _num(w_mm), _num(l_mm)
    if t <= 0 or w <= 0 or l <= 0:
        return 0.0
    density = 2.7 if str(kind).upper() == "AL" else 7.85
    return t * w * l * density / 1_000_000


def _make_record(source: str, spec: str, t: float, w: float, monthly: dict[str, float]) -> dict:
    w_code = int(w) if w else 1219
    t_code = str(t).rstrip("0").rstrip(".") if t == int(t) else str(t)
    return {
        "source": source,
        "spec": _norm_spec(spec),
        "t": t,
        "w": w if w else 1219.0,
        "material_code": f"{_norm_spec(spec)}_{t_code}_{w_code}_Common",
        **{m: monthly.get(m, 0.0) for m in MONTHS_BALANCE},
    }


def _find_month_columns(header_row: list) -> dict[str, int]:
    month_cols: dict[str, int] = {}
    for c, cell in enumerate(header_row):
        label = _text(cell).upper().replace(" (PCS.)", "").replace("(PCS.)", "")
        label = label.split("\n")[0].strip()
        if label in MONTH_ALIASES:
            month_cols[MONTH_ALIASES[label]] = c
        for key, month in MONTH_ALIASES.items():
            if key in label:
                month_cols[month] = c
    return month_cols


def parse_pdf_tsk(path: Path) -> pd.DataFrame:
    records = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table or len(table) < 3:
                    continue
                header = table[0]
                if not any(_text(h) == "SPEC" for h in header):
                    continue
                month_cols = _find_month_columns(header)
                if len(month_cols) < 2:
                    continue
                # T, W, L in sub-header row 1
                for row in table[2:]:
                    if not row or not row[0]:
                        continue
                    spec = _text(row[2] if len(row) > 2 else "")
                    if not spec:
                        continue
                    t = _num(row[3] if len(row) > 3 else 0)
                    w = _num(row[4] if len(row) > 4 else 0)
                    l = _num(row[5] if len(row) > 5 else 1219)
                    kg_pc = _calc_kg(t, w, l)
                    monthly = {
                        m: _num(row[c]) * kg_pc if c < len(row) else 0.0
                        for m, c in month_cols.items()
                    }
                    if sum(monthly.values()) <= 0:
                        continue
                    records.append(_make_record("TSK", spec, t, w, monthly))
    return pd.DataFrame(records)


def parse_pdf_cps(path: Path) -> pd.DataFrame:
    records = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table:
                    continue
                header_row_idx = None
                month_cols: dict[str, int] = {}
                for ri, row in enumerate(table):
                    labels = [_text(c) for c in row]
                    for c, label in enumerate(labels):
                        for key, month in MONTH_ALIASES.items():
                            if key in label.upper() and "SHEET" in label.upper():
                                month_cols[month] = c
                    if "SPEC" in labels and len(month_cols) >= 4:
                        header_row_idx = ri
                        break
                    month_cols = {}

                if header_row_idx is None:
                    continue

                for row in table[header_row_idx + 2:]:
                    if not row or not _text(row[1] if len(row) > 1 else ""):
                        continue
                    spec_raw = _text(row[1])
                    size_raw = _text(row[2] if len(row) > 2 else "")
                    spec = spec_raw.split("/")[0].strip()
                    t, w, l = _parse_dims(size_raw)
                    if t <= 0:
                        continue
                    kg_she = _calc_kg(t, w, l)
                    monthly = {
                        m: _num(row[c]) * kg_she if c < len(row) else 0.0
                        for m, c in month_cols.items()
                    }
                    if sum(monthly.values()) <= 0:
                        continue
                    records.append(_make_record("CPS", spec, t, w, monthly))
    return pd.DataFrame(records)


def parse_pdf_1430(path: Path) -> pd.DataFrame:
    records = []
    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    # Line pattern: item ... material_no SPEC dims ... quantities ... UNIT
    line_re = re.compile(
        r"^\s*(\d+)\s+R\d+\s+.+?\s+(\d{8})\s+"
        r"([A-Z0-9][A-Z0-9\-/ ]+?)\s+"
        r"([\d.]+)\s*[xX]\s*([\d.]+)\s*[xX]\s*(\w+)"
        r".*?"
        r"((?:\d[\d,]*\s+){3,7}\d[\d,]*)\s+(SHE|KG)\b",
        re.IGNORECASE | re.MULTILINE,
    )
    for m in line_re.finditer(text):
        spec = _norm_spec(m.group(3).split()[0])
        t, w = _num(m.group(4)), _num(m.group(5))
        length_token = m.group(6).upper()
        qtys = [_num(x) for x in m.group(7).split()]
        unit = m.group(8).upper()
        l = 1219 if length_token == "COIL" else _num(length_token)

        month_order = ["Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        monthly: dict[str, float] = {}
        for i, month in enumerate(month_order):
            if i >= len(qtys):
                break
            qty = qtys[i]
            if unit == "KG":
                monthly[month] = qty
            else:
                monthly[month] = qty * _calc_kg(t, w, l if l > 0 else 1219)

        if sum(monthly.get(m, 0) for m in MONTHS_BALANCE) <= 0:
            continue
        records.append(_make_record("Plant1430", spec, t, w, monthly))

    # Fallback: looser regex per line from raw text
    if not records:
        for line in text.splitlines():
            if not re.search(r"\d{8}", line):
                continue
            mat_m = re.search(r"(\d{8})\s+([A-Z][A-Z0-9\-]+)", line)
            dim_m = re.search(r"([\d.]+)\s*[xX]\s*([\d.]+)\s*[xX]\s*(\w+)", line)
            if not mat_m or not dim_m:
                continue
            spec = _norm_spec(mat_m.group(2))
            t, w = _num(dim_m.group(1)), _num(dim_m.group(2))
            l_token = dim_m.group(3).upper()
            l = 1219 if l_token == "COIL" else _num(l_token)
            nums = re.findall(r"(?<!\d)(\d{2,6})(?!\d)", line)
            if len(nums) < 4:
                continue
            qtys = [float(n) for n in nums[-7:]]
            is_kg = "KG" in line.upper()
            monthly = {}
            for i, month in enumerate(["Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]):
                if i >= len(qtys):
                    break
                monthly[month] = qtys[i] if is_kg else qtys[i] * _calc_kg(t, w, l)
            if sum(monthly.get(m, 0) for m in MONTHS_BALANCE) <= 0:
                continue
            records.append(_make_record("Plant1430", spec, t, w, monthly))

    return pd.DataFrame(records)


def _ocr_probending_image(image_path: Path) -> str:
    if RapidOCR is None:
        return ""
    ocr = RapidOCR()
    result, _ = ocr(str(image_path))
    if not result:
        return ""
    # Sort by Y then X
    items = sorted(result, key=lambda x: (x[0][0][1], x[0][0][0]))
    return " ".join(item[1] for item in items)


def _render_pdf_page(pdf_path: Path, out_image: Path, dpi: int = 200) -> bool:
    if fitz is None:
        return False
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi)
    pix.save(str(out_image))
    return out_image.exists()


def parse_pdf_probending(path: Path, cache_dir: Path | None = None) -> pd.DataFrame:
    """Parse scanned Probending forecast PDF via OCR."""
    records = []
    cache_dir = cache_dir or path.parent
    image_path = cache_dir / "_probending_ocr.png"

    if not _render_pdf_page(path, image_path):
        return pd.DataFrame()

    if RapidOCR is None:
        print("  Warning: rapidocr-onnxruntime not installed, skipping Probending OCR")
        return pd.DataFrame()

    ocr = RapidOCR()
    result, _ = ocr(str(image_path))
    if not result:
        return pd.DataFrame()

    from collections import defaultdict

    rows: dict[int, list[tuple[float, str]]] = defaultdict(list)
    for box, text, _conf in result:
        y = round((box[0][1] + box[2][1]) / 2 / 8) * 8
        x = (box[0][0] + box[1][0]) / 2
        rows[y].append((x, _text(text)))

    raw_lines: list[str] = []
    for y in sorted(rows):
        cells = sorted(rows[y], key=lambda t: t[0])
        line = " ".join(c[1] for c in cells)
        if re.search(r"770\d{4,5}", line) or re.search(r"JSH|JSC", line, re.I):
            raw_lines.append(line)

    # Normalize OCR artifacts
    def normalize_line(s: str) -> str:
        s = re.sub(r"(JSH\d+[A-Z-]*)(\d)", r"\1 \2", s, flags=re.I)
        s = re.sub(r"(JSC\d+[A-Z-]*)(\d)", r"\1 \2", s, flags=re.I)
        s = re.sub(r"(\d)([xX])", r"\1 \2", s)
        s = re.sub(r"([xX])(\d)", r"\1 \2", s)
        s = re.sub(r"(770\d{5})(JSH|JSC)", r"\1 \2", s, flags=re.I)
        return re.sub(r"\s+", " ", s).strip()

    lines = [normalize_line(ln) for ln in raw_lines]

    def extract_code(s: str) -> str:
        m = re.search(r"(770\d{5})", s)
        return m.group(1) if m else ""

    def extract_spec(s: str) -> tuple[str, float, float, float]:
        m = re.search(
            r"((?:JSH|JSC)[A-Z0-9\-]*)\s+([\d.]+)\s*[xX]\s*([\d.]+)\s*[xX]\s*([\d.]+)",
            s,
            re.I,
        )
        if not m:
            return "", 0.0, 0.0, 0.0
        return _norm_spec(m.group(1)), _num(m.group(2)), _num(m.group(3)), _num(m.group(4))

    def extract_pcs(s: str) -> list[int]:
        cleaned = re.sub(r"770\d{5}", " ", s)
        cleaned = re.sub(
            r"(?:JSH|JSC)[A-Z0-9\-]*\s*[\d.]+\s*[xX]\s*[\d.]+\s*[xX]\s*[\d.]+",
            " ",
            cleaned,
            flags=re.I,
        )
        cleaned = re.sub(r"^\s*\d{1,2}\s+", "", cleaned)  # row number
        vals = [int(x) for x in re.findall(r"\b(\d{1,4})\b", cleaned)]
        return vals[-6:] if len(vals) >= 6 else vals

    # Pair adjacent fragmented lines (max 2 lines)
    merged: list[str] = []
    i = 0
    while i < len(lines):
        chunk = lines[i]
        code = extract_code(chunk)
        spec, t, w, l = extract_spec(chunk)
        pcs = extract_pcs(chunk)

        if (not code or not spec or len(pcs) < 3) and i + 1 < len(lines):
            chunk2 = lines[i + 1]
            code = code or extract_code(chunk2)
            s2, t2, w2, l2 = extract_spec(chunk2)
            if not spec:
                spec, t, w, l = s2, t2, w2, l2
            pcs2 = extract_pcs(chunk2)
            pcs = pcs if len(pcs) >= len(pcs2) else pcs2
            i += 1

        if code and spec and t > 0 and len(pcs) >= 3:
            merged.append((code, spec, t, w, l, pcs))
        i += 1

    for _code, spec, t, w, l, pcs_list in merged:
        month_order = ["Jun", "Jul", "Aug", "Sep", "Oct", "Nov"]
        kg_pc = _calc_kg(t, w, l)
        monthly = {
            month_order[i]: pcs_list[i] * kg_pc
            for i in range(min(len(pcs_list), len(month_order)))
        }
        if sum(monthly.values()) <= 0:
            continue
        records.append(_make_record("Probending", spec, t, w, monthly))

    return pd.DataFrame(records)


def parse_pdf_forecast(path: Path) -> pd.DataFrame:
    name = path.name.upper()
    if "TSK" in name:
        return parse_pdf_tsk(path)
    if "CPS" in name:
        return parse_pdf_cps(path)
    if "1430" in name:
        return parse_pdf_1430(path)
    if "PROBENDING" in name or "PROBEND" in name:
        return parse_pdf_probending(path)
    # Try by content
    with pdfplumber.open(path) as pdf:
        sample = (pdf.pages[0].extract_text() or "").upper()
    if "THAI SERVICE KONLAKAN" in sample or "TSK" in sample:
        return parse_pdf_tsk(path)
    if "CPS" in sample and "HONDA" in sample:
        return parse_pdf_cps(path)
    if "1430" in sample or "PLANT. 1430" in sample:
        return parse_pdf_1430(path)
    if "PROBENDING" in sample or "STRIP SHEET" in sample:
        return parse_pdf_probending(path)
    return pd.DataFrame()


def load_pdf_forecasts(forecast_dir: Path) -> pd.DataFrame:
    frames = []
    for path in sorted(forecast_dir.glob("*.pdf")):
        try:
            part = parse_pdf_forecast(path)
            if not part.empty:
                frames.append(part)
                print(f"  Parsed PDF {path.name}: {len(part)} rows")
            else:
                print(f"  Warning: no data from PDF {path.name}")
        except Exception as exc:
            print(f"  Warning: failed PDF {path.name}: {exc}")
    if not frames:
        return pd.DataFrame()
    all_fc = pd.concat(frames, ignore_index=True)
    month_cols = [m for m in MONTHS_BALANCE if m in all_fc.columns]
    if not month_cols:
        return pd.DataFrame()
    return (
        all_fc.groupby("material_code", as_index=False)[month_cols]
        .sum()
        .merge(
            all_fc.groupby("material_code", as_index=False)[["spec", "t", "w"]].first(),
            on="material_code",
        )
    )
