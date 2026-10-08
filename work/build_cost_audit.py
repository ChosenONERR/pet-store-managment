from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from openpyxl import load_workbook


PACK_UNITS = "罐|袋|支|瓶|盒|包|片|个|枚|条|杯|桶"
MULTIPLIER_RE = re.compile(
    rf"(?P<size>\d+(?:\.\d+)?)\s*(?P<unit>kg|g|ml|l)\s*[*x×]\s*(?P<count>\d+)\s*(?P<piece>{PACK_UNITS})?",
    re.IGNORECASE,
)
COUNT_PACK_RE = re.compile(rf"(?P<count>\d+)\s*(?P<piece>{PACK_UNITS})\s*(?:整箱|整盒|整包|装)")
PACKAGE_SIGNAL_RE = re.compile(rf"(?:整箱|一箱|箱装|\d+\s*(?:{PACK_UNITS})|[*x×]\s*\d+)", re.IGNORECASE)


def text(value: object) -> str:
    return "" if value is None else str(value).strip()


def number(value: object) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def normalize_code(value: object) -> str:
    return text(value).lower()


def load_cost_rows(path: Path) -> dict[str, dict[str, object]]:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    sheet = workbook.active
    rows = sheet.iter_rows(values_only=True)
    header = None
    for excel_row, row in enumerate(rows, 1):
        values = [text(value) for value in row]
        if "货号" in values and "参考进货价" in values:
            header = {name: index for index, name in enumerate(values) if name}
            header_row = excel_row
            break
    if not header:
        raise RuntimeError("对货表没有找到货号和参考进货价表头")

    output: dict[str, dict[str, object]] = {}
    for excel_row, row in enumerate(sheet.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
        values = list(row)
        code = normalize_code(values[header["货号"]])
        if not code:
            continue
        output[code] = {
            "source_row": excel_row,
            "csx_code": code,
            "product_name": text(values[header.get("商品名称", -1)]) if "商品名称" in header else "",
            "purchase_unit": text(values[header.get("商品规格", -1)]) if "商品规格" in header else "",
            "reference_purchase_quantity": number(values[header.get("参考拿货数量", -1)]) if "参考拿货数量" in header else None,
            "reference_pack_cost": number(values[header["参考进货价"]]),
        }
    return output


def load_inventory_rows(path: Path) -> list[dict[str, object]]:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    sheet = workbook.active
    # Some Meituan exports declare an incorrect worksheet dimension (A1:A...),
    # although the XML contains eight columns. Force the known export width.
    header_values = [
        text(value)
        for value in next(sheet.iter_rows(min_row=1, max_row=1, max_col=8, values_only=True))
    ]
    header = {name: index for index, name in enumerate(header_values) if name}
    required = ["商品名称", "店内码/货号", "库存", "月售", "货架码/位置码"]
    missing = [name for name in required if name not in header]
    if missing:
        raise RuntimeError(f"库存文件缺少字段：{','.join(missing)}")

    output = []
    for row in sheet.iter_rows(min_row=3, max_col=8, values_only=True):
        values = list(row)
        code = normalize_code(values[header["店内码/货号"]])
        if not code:
            continue
        output.append(
            {
                "kw_code": code,
                "title": text(values[header["商品名称"]]),
                "inventory": number(values[header["库存"]]),
                "monthly_sales": number(values[header["月售"]]),
                "status": text(values[header.get("售卖状态", -1)]) if "售卖状态" in header else "",
                "csx_code": normalize_code(values[header["货架码/位置码"]]),
            }
        )
    return output


def same_measure_in_title(title: str, size: str, unit: str) -> bool:
    normalized = re.sub(r"\s+", "", title).lower()
    target = re.sub(r"\.0$", "", size.lower()) + unit.lower()
    return target in normalized


def title_mentions_pack(title: str, count: int, piece: str) -> bool:
    normalized = re.sub(r"\s+", "", title).lower()
    patterns = [
        rf"[*x×]{count}(?:{re.escape(piece)})?",
        rf"{count}{re.escape(piece)}(?:装|整箱|整盒|整包)?",
    ]
    return any(re.search(pattern, normalized, re.IGNORECASE) for pattern in patterns)


def infer_conversion(source: dict[str, object], listing_title: str) -> tuple[int, str, str]:
    local_name = text(source.get("product_name"))
    purchase_unit = text(source.get("purchase_unit"))
    purchase_singular = re.search(rf"/\s*(?P<piece>{PACK_UNITS})", purchase_unit)
    if purchase_singular and re.search(
        rf"/\s*{re.escape(purchase_singular.group('piece'))}", listing_title
    ):
        return 1, "默认整件", "采购单位与美团销售单位一致"
    signals = []
    for match in MULTIPLIER_RE.finditer(local_name):
        count = int(match.group("count"))
        if count <= 1:
            continue
        size = match.group("size")
        unit = match.group("unit")
        piece = match.group("piece") or ""
        if same_measure_in_title(listing_title, size, unit):
            signals.append((count, size, unit, piece))

    high_confidence = []
    for count, size, unit, piece in signals:
        title_has_pack = title_mentions_pack(listing_title, count, piece) if piece else bool(
            re.search(rf"[*x×]{count}(?:\D|$)", re.sub(r"\s+", "", listing_title), re.IGNORECASE)
        )
        purchase_unit_match = next(
            (
                candidate
                for candidate in MULTIPLIER_RE.finditer(purchase_unit)
                if int(candidate.group("count")) == count
            ),
            None,
        )
        purchase_is_outer_pack = (
            purchase_unit in {"箱", "件"}
            or "整箱" in local_name
            or "一箱" in local_name
            or purchase_unit_match is not None
        )
        title_is_single = (piece and re.search(rf"/\s*{re.escape(piece)}", listing_title)) or not title_has_pack
        if purchase_is_outer_pack and title_is_single and not title_has_pack:
            high_confidence.append((count, size, unit, piece))

    if high_confidence:
        counts = {item[0] for item in high_confidence}
        if len(counts) == 1:
            count, size, unit, piece = high_confidence[0]
            return count, "已自动换算", f"采购为{purchase_unit or '外包装'}；本地含{size}{unit}×{count}{piece}，美团按单个销售"

    explicit_piece_signals = [item for item in signals if item[3]]
    if explicit_piece_signals and all(
        title_mentions_pack(listing_title, count, piece)
        for count, _size, _unit, piece in explicit_piece_signals
    ):
        return 1, "默认整件", "美团销售标题与本地包装数量一致，按整件销售"

    for count, size, unit, _piece in signals:
        normalized_title = re.sub(r"\s+", "", listing_title).lower()
        if re.search(rf"{count}(?:{PACK_UNITS})", normalized_title):
            return 1, "默认整件", "美团标题包含与采购包装一致的件数，按整件销售"
        try:
            total_size = float(size) * count
        except ValueError:
            continue
        total_texts = {f"{total_size:g}{unit.lower()}"}
        if unit.lower() == "g" and total_size >= 1000:
            total_texts.add(f"{total_size / 1000:g}kg")
        if any(total_text in normalized_title for total_text in total_texts):
            return 1, "默认整件", "美团标题总规格与采购包装总规格一致，按整件销售"

    package_candidate = bool(PACKAGE_SIGNAL_RE.search(local_name)) and any(
        int(match.group("count")) > 1 for match in MULTIPLIER_RE.finditer(local_name)
    )
    if package_candidate:
        return 1, "待人工核对", "本地描述含多件包装，但无法确认美团销售单位是否拆零"
    return 1, "默认整件", "未发现需要拆分的高置信度证据"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cost-file", required=True)
    parser.add_argument("--inventory-file", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    cost_path = Path(args.cost_file)
    inventory_path = Path(args.inventory_file)
    cost_rows = load_cost_rows(cost_path)
    inventory_rows = load_inventory_rows(inventory_path)

    audit_rows = []
    counts = {"已自动换算": 0, "待人工核对": 0, "默认整件": 0, "缺少成本": 0, "缺少CSX映射": 0}
    for inventory in inventory_rows:
        csx_code = text(inventory.get("csx_code"))
        source = cost_rows.get(csx_code)
        if not csx_code:
            conversion, audit_status, reason = 1, "缺少CSX映射", "美团货号没有货架码/CSX映射"
            pack_cost = None
            purchase_unit = ""
            source_name = ""
        elif not source:
            conversion, audit_status, reason = 1, "缺少成本", "CSX货架码在本地对货表中没有成本"
            pack_cost = None
            purchase_unit = ""
            source_name = ""
        else:
            conversion, audit_status, reason = infer_conversion(source, text(inventory.get("title")))
            pack_cost = source.get("reference_pack_cost")
            purchase_unit = text(source.get("purchase_unit"))
            source_name = text(source.get("product_name"))
        counts[audit_status] = counts.get(audit_status, 0) + 1
        unit_cost = None if pack_cost is None else round(float(pack_cost) / conversion, 6)
        audit_rows.append(
            {
                **inventory,
                "reference_pack_cost": pack_cost,
                "purchase_unit": purchase_unit,
                "units_per_pack": conversion,
                "baseline_unit_cost": unit_cost,
                "audit_status": audit_status,
                "conversion_reason": reason,
                "source_product_name": source_name,
            }
        )

    payload = {
        "cost_file": str(cost_path),
        "inventory_file": str(inventory_path),
        "source_cost_skus": len(cost_rows),
        "inventory_rows": len(inventory_rows),
        "counts": counts,
        "rows": audit_rows,
    }
    Path(args.output).write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: value for key, value in payload.items() if key != "rows"}, ensure_ascii=False))


if __name__ == "__main__":
    main()
