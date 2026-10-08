from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook


SOURCE = Path(r"D:\market-research\宠速星\对货表.xlsx")
TARGET_CODE = "csx104"


def clean(value):
    if value is None:
        return None
    if isinstance(value, str):
        value = value.strip()
        return value or None
    return value


def as_number(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    if isinstance(value, str):
        match = re.fullmatch(r"\s*[-+]?\d+(?:\.\d+)?\s*", value)
        if match:
            return float(value)
    return None


def find_header(ws):
    for row_index, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), 1):
        values = [str(clean(v) or "") for v in row]
        if "货号" in values and "参考进货价" in values:
            return row_index, {value: i for i, value in enumerate(values) if value}
    return None, {}


def main():
    wb = load_workbook(SOURCE, read_only=True, data_only=True, keep_links=False)
    report = {
        "source": str(SOURCE),
        "size_mb": round(SOURCE.stat().st_size / 1024 / 1024, 2),
        "sheets": [],
    }

    carton_patterns = [
        re.compile(r"(?:整箱|一箱|箱装|\d+\s*[罐袋盒支瓶包片条个枚粒杯桶]|\*\s*\d+|×\s*\d+|x\s*\d+)", re.I),
        re.compile(r"\d+\s*(?:罐|袋|盒|支|瓶|包|片|条|个|枚|粒|杯|桶)\s*/\s*(?:箱|盒|包|组)", re.I),
    ]

    for ws in wb.worksheets:
        header_row, headers = find_header(ws)
        sheet_info = {
            "name": ws.title,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "header_row": header_row,
            "headers": headers,
        }
        if not header_row:
            report["sheets"].append(sheet_info)
            continue

        code_idx = headers.get("货号")
        cost_idx = headers.get("参考进货价")
        name_idx = headers.get("商品名称")
        spec_idx = headers.get("商品规格")
        shelf_idx = headers.get("货架码")
        qty_idx = headers.get("参考拿货数量")
        total_idx = headers.get("合计")

        stats = Counter()
        codes = defaultdict(list)
        target_rows = []
        packaging_candidates = []
        rows_with_cost = []

        for excel_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
            values = [clean(v) for v in row]
            code = values[code_idx] if code_idx is not None and code_idx < len(values) else None
            if code is None:
                continue
            code_text = str(code).strip().lower()
            if not code_text:
                continue

            stats["data_rows"] += 1
            codes[code_text].append(excel_row)
            cost = values[cost_idx] if cost_idx is not None and cost_idx < len(values) else None
            cost_number = as_number(cost)
            if cost_number is None:
                stats["missing_or_non_numeric_cost"] += 1
            else:
                stats["numeric_cost"] += 1

            name = values[name_idx] if name_idx is not None and name_idx < len(values) else None
            spec = values[spec_idx] if spec_idx is not None and spec_idx < len(values) else None
            shelf = values[shelf_idx] if shelf_idx is not None and shelf_idx < len(values) else None
            qty = values[qty_idx] if qty_idx is not None and qty_idx < len(values) else None
            total = values[total_idx] if total_idx is not None and total_idx < len(values) else None
            text = f"{name or ''} {spec or ''}"
            looks_packaged = any(pattern.search(text) for pattern in carton_patterns)
            if looks_packaged:
                stats["packaging_text_candidate"] += 1

            record = {
                "row": excel_row,
                "货号": code_text,
                "货架码": shelf,
                "商品名称": name,
                "商品规格": spec,
                "参考拿货数量": qty,
                "参考进货价": cost,
                "合计": total,
            }
            if code_text == TARGET_CODE:
                target_rows.append(record)
            if looks_packaged and len(packaging_candidates) < 50:
                packaging_candidates.append(record)
            if cost_number is not None and len(rows_with_cost) < 5:
                rows_with_cost.append(record)

        duplicate_codes = {code: rows for code, rows in codes.items() if len(rows) > 1}
        sheet_info.update(
            {
                "stats": dict(stats),
                "unique_codes": len(codes),
                "duplicate_code_count": len(duplicate_codes),
                "duplicate_examples": dict(list(duplicate_codes.items())[:20]),
                "target_csx104": target_rows,
                "packaging_candidates_sample": packaging_candidates,
                "first_cost_rows": rows_with_cost,
            }
        )
        report["sheets"].append(sheet_info)

    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
