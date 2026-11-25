#!/usr/bin/env python3
"""Build a store hierarchy migration workbook from a source workbook and template."""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from copy import copy
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill


REQUIRED_SHEETS = ["新旧参照(支店)", "新旧参照(エリア)", "エリア", "支店", "支社"]
YELLOW = PatternFill("solid", fgColor="FFF2CC")


def norm(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_header(value: Any) -> str:
    return norm(value).replace(" ", "").replace("　", "")


def is_store_code(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return isinstance(value, str) and value.strip().isdigit()


def canonical_code(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str) and value.strip().isdigit():
        return int(value.strip())
    return value


def find_source_layout(workbook):
    aliases = {
        "old_company_cd": {"支社CD(旧)"},
        "old_company_name": {"支社名(旧)"},
        "old_branch_cd": {"支店CD(旧)"},
        "old_branch_name": {"支店名(旧)"},
        "old_area_cd": {"エリアCD(旧)"},
        "old_area_name": {"エリア名(旧)"},
        "new_company_cd": {"支社CD(新)"},
        "new_company_name": {"支社名(新)"},
        "new_branch_cd": {"支店CD(新)"},
        "new_branch_name": {"支店名(新)"},
        "new_area_cd": {"エリアCD(新)"},
        "new_area_name": {"エリア名(新)"},
        "store_cd": {"店舗CD", "店舗コード"},
        "store_name": {"店舗名", "店舗名称", "店名"},
    }
    for ws in workbook.worksheets:
        for row_number in range(1, min(ws.max_row, 20) + 1):
            headers = {}
            for col in range(1, ws.max_column + 1):
                header = normalize_header(ws.cell(row_number, col).value)
                for key, names in aliases.items():
                    if header in names and key not in headers:
                        headers[key] = col
            if all(key in headers for key in aliases):
                return ws, row_number, headers
    raise ValueError("源文件中找不到完整的新旧店舗階層字段。")


def read_records(source: Path):
    workbook = load_workbook(source, data_only=True)
    ws, header_row, cols = find_source_layout(workbook)
    end_row = ws.max_row
    for row in range(header_row + 1, ws.max_row + 1):
        if any("閉店店舗の編成が必要な場合" in norm(ws.cell(row, col).value) for col in range(1, ws.max_column + 1)):
            end_row = row - 1
            break

    def value(row, key):
        raw = ws.cell(row, cols[key]).value
        return None if raw in (None, "") else raw

    records = []
    for row in range(header_row + 1, end_row + 1):
        store_cd = value(row, "store_cd")
        if not is_store_code(store_cd):
            continue
        record = {
            "source_row": row,
            "store_cd": canonical_code(store_cd),
            "store_name": value(row, "store_name"),
            "old": [
                (value(row, "old_company_cd"), value(row, "old_company_name")),
                (value(row, "old_branch_cd"), value(row, "old_branch_name")),
                (value(row, "old_area_cd"), value(row, "old_area_name")),
            ],
            "new": [
                (value(row, "new_company_cd"), value(row, "new_company_name")),
                (value(row, "new_branch_cd"), value(row, "new_branch_name")),
                (value(row, "new_area_cd"), value(row, "new_area_name")),
            ],
        }
        records.append(record)

    records.sort(key=lambda item: (isinstance(item["store_cd"], str), item["store_cd"]))
    store_codes = [item["store_cd"] for item in records]
    duplicates = [code for code, count in Counter(store_codes).items() if count > 1]
    if duplicates:
        raise ValueError(f"店舗CD重复：{duplicates[:20]}")
    if not records:
        raise ValueError("未读取到有效店舗。")
    return records, ws.title


def cd_path(record, side, level):
    return tuple(norm(record[side][index][0]) for index in range(level + 1))


def name_path(record, side, level):
    return tuple(norm(record[side][index][1]) for index in range(level + 1))


def node_path(record, side, level):
    return tuple((norm(record[side][index][0]), norm(record[side][index][1])) for index in range(level + 1))


def unique_nodes(records, side, level):
    nodes = {}
    for record in records:
        key = node_path(record, side, level)
        if not any(code or name for code, name in key):
            continue
        nodes.setdefault(key, {"pairs": record[side][: level + 1], "stores": []})["stores"].append(record)
    return nodes


def sort_part(value):
    if value in (None, ""):
        return 2, ""
    try:
        return 0, float(value)
    except (TypeError, ValueError):
        return 1, str(value)


def analyze_level(records, level):
    old_nodes = unique_nodes(records, "old", level)
    new_nodes = unique_nodes(records, "new", level)
    old_by_cd = defaultdict(list)
    old_by_name = defaultdict(list)
    for node in old_nodes.values():
        representative = node["stores"][0]
        old_by_cd[cd_path(representative, "old", level)].append(node)
        old_by_name[name_path(representative, "old", level)].append(node)

    results = []
    for node in new_nodes.values():
        representative = node["stores"][0]
        current_cd_path = cd_path(representative, "new", level)
        current_name_path = name_path(representative, "new", level)
        cd_candidates = old_by_cd.get(current_cd_path, [])
        name_candidates = old_by_name.get(current_name_path, [])
        target_cd = representative["new"][level][0]
        name_match = bool(name_candidates)
        cd_match = (
            target_cd not in (None, "") and bool(cd_candidates)
        ) or (
            target_cd in (None, "") and name_match and bool(cd_candidates)
        )

        if cd_match and name_match:
            pattern = "変更なし"
        elif cd_match:
            pattern = "名称変更(CD移動要確認)"
        elif name_match:
            pattern = "名称一致(CD変更あり)"
        else:
            pattern = "新規"

        counterpart = cd_candidates[0] if cd_match else name_candidates[0] if name_match else None
        reference = None
        reference_count = 0
        ratio = None
        if pattern == "新規":
            votes = Counter()
            old_pairs = {}
            for store in node["stores"]:
                key = node_path(store, "old", level)
                votes[key] += 1
                old_pairs[key] = store["old"][: level + 1]
            if votes:
                winner, reference_count = votes.most_common(1)[0]
                reference = {"pairs": old_pairs[winner]}
                ratio = reference_count / len(node["stores"])

        results.append(
            {
                "new": node,
                "counterpart": counterpart,
                "reference": reference,
                "cd": "〇" if cd_match else "×",
                "name": "〇" if name_match else "×",
                "pattern": pattern,
                "reference_count": reference_count,
                "total": len(node["stores"]),
                "ratio": ratio,
                "needs_review": ratio is not None and ratio < 0.70,
            }
        )

    results.sort(
        key=lambda item: tuple(
            sort_part(code) + (norm(name),) for code, name in item["new"]["pairs"]
        )
    )
    return results


def copy_row_style(ws, source_row, target_row, max_column):
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for column in range(1, max_column + 1):
        source = ws.cell(source_row, column)
        target = ws.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)


def clear_values(ws, start_row, max_column):
    for row in ws.iter_rows(
        min_row=start_row, max_row=ws.max_row, min_col=1, max_col=max_column
    ):
        for cell in row:
            cell.value = None
            cell.comment = None
            cell.hyperlink = None


def ensure_rows(ws, start_row, end_row, style_row, max_column):
    for row in range(start_row, end_row + 1):
        copy_row_style(ws, style_row, row, max_column)


def write_reference_sheets(workbook, records):
    ws = workbook["新旧参照(支店)"]
    clear_values(ws, 3, 13)
    ensure_rows(ws, 3, len(records) + 2, 3, 13)
    ws["A1"], ws["H1"] = "現階層", "新階層"
    for row, record in enumerate(records, 3):
        old, new = record["old"], record["new"]
        values = [
            record["store_cd"], record["store_name"], old[1][0], old[1][1],
            old[0][0], old[0][1], None, record["store_cd"], record["store_name"],
            new[1][0], new[1][1], new[0][0], new[0][1],
        ]
        for column, value in enumerate(values, 1):
            ws.cell(row, column, value)
        if old[:2] != new[:2]:
            for column in range(8, 14):
                ws.cell(row, column).fill = copy(YELLOW)
    ws.auto_filter.ref = f"A2:M{len(records) + 2}"
    ws.freeze_panes = "A3"

    ws = workbook["新旧参照(エリア)"]
    clear_values(ws, 3, 17)
    ensure_rows(ws, 3, len(records) + 2, 3, 17)
    ws["A1"], ws["J1"] = "現階層", "新階層"
    for row, record in enumerate(records, 3):
        old, new = record["old"], record["new"]
        values = [
            record["store_cd"], record["store_name"], old[2][0], old[2][1],
            old[1][0], old[1][1], old[0][0], old[0][1], None,
            record["store_cd"], record["store_name"], new[2][0], new[2][1],
            new[1][0], new[1][1], new[0][0], new[0][1],
        ]
        for column, value in enumerate(values, 1):
            ws.cell(row, column, value)
        if old != new:
            for column in range(10, 18):
                ws.cell(row, column).fill = copy(YELLOW)
    ws.auto_filter.ref = f"A2:Q{len(records) + 2}"
    ws.freeze_panes = "A3"


def write_summary_sheet(workbook, records, name, level):
    ws = workbook[name]
    max_column = {"支社": 9, "支店": 13, "エリア": 17}[name]
    results = analyze_level(records, level)
    clear_values(ws, 4, max_column)
    ensure_rows(ws, 4, len(results) + 3, 4, max_column)
    ws["A1"] = "新階層"
    ws.cell(1, {"支社": 6, "支店": 8, "エリア": 10}[name], "現階層")

    for row, result in enumerate(results, 4):
        new_pairs = result["new"]["pairs"]
        counterpart = result["counterpart"]
        old_pairs = counterpart["pairs"] if counterpart else [(None, None)] * (level + 1)
        reference = result["reference"]
        reference_pairs = reference["pairs"] if reference else [(None, None)] * (level + 1)

        if name == "支社":
            values = [
                *new_pairs[0], result["cd"], result["name"], result["pattern"],
                *(old_pairs[0] if counterpart else ("#N/A", "#N/A")),
                *(reference_pairs[0] if reference else (None, None)),
            ]
            pattern_column = 5
        elif name == "支店":
            values = [
                *new_pairs[0], *new_pairs[1], result["cd"], result["name"],
                result["pattern"],
                *(old_pairs[0] if counterpart else ("#N/A", "#N/A")),
                *(old_pairs[1] if counterpart else ("#N/A", "#N/A")),
                *(reference_pairs[1] if reference else (None, None)),
            ]
            pattern_column = 7
        else:
            values = [
                *new_pairs[0], *new_pairs[1], *new_pairs[2],
                result["cd"], result["name"], result["pattern"],
                *(old_pairs[0] if counterpart else ("#N/A", "#N/A")),
                *(old_pairs[1] if counterpart else ("#N/A", "#N/A")),
                *(old_pairs[2] if counterpart else ("#N/A", "#N/A")),
                *(reference_pairs[2] if reference else (None, None)),
            ]
            pattern_column = 9

        for column, value in enumerate(values, 1):
            ws.cell(row, column, value)
        if result["pattern"] == "新規":
            for column in range(1, max_column + 1):
                ws.cell(row, column).fill = copy(YELLOW)
            ratio = result["ratio"] or 0
            note = (
                f"参照根拠：店舗CD一致 "
                f"{result['reference_count']}/{result['total']} "
                f"（主継承率 {ratio:.1%}）"
            )
            if result["needs_review"]:
                note += "／要確認"
            ws.cell(row, pattern_column).comment = Comment(note, "Cursor")

    end_row = len(results) + 3
    end_column = {"支社": "I", "支店": "M", "エリア": "Q"}[name]
    ws.auto_filter.ref = f"A3:{end_column}{end_row}"
    ws.freeze_panes = "A4"
    return results


def validate_output(path: Path, expected_stores: int):
    workbook = load_workbook(path, data_only=False)
    formulas = []
    ref_errors = []
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas.append((ws.title, cell.coordinate))
                    if "#REF!" in cell.value:
                        ref_errors.append((ws.title, cell.coordinate))
    for sheet in ("新旧参照(支店)", "新旧参照(エリア)"):
        ws = workbook[sheet]
        count = sum(ws.cell(row, 1).value is not None for row in range(3, ws.max_row + 1))
        if count != expected_stores:
            raise ValueError(f"{sheet} 店舗数不一致：{count} != {expected_stores}")
    if formulas or ref_errors:
        raise ValueError(f"输出中残留公式：{formulas[:10]}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    if args.output.resolve() in {args.source.resolve(), args.template.resolve()}:
        raise ValueError("输出文件不得覆盖源文件或模板。")

    records, source_sheet = read_records(args.source)
    workbook = load_workbook(args.template)
    missing = [sheet for sheet in REQUIRED_SHEETS if sheet not in workbook.sheetnames]
    if missing:
        raise ValueError(f"模板缺少Sheet：{missing}")

    write_reference_sheets(workbook, records)
    results = {
        "支社": write_summary_sheet(workbook, records, "支社", 0),
        "支店": write_summary_sheet(workbook, records, "支店", 1),
        "エリア": write_summary_sheet(workbook, records, "エリア", 2),
    }
    for ws in workbook.worksheets:
        ws["A1"].comment = Comment(
            f"作成元：{args.source.name}\nテンプレート：{args.template.name}"
            f"\n源Sheet：{source_sheet}\n閉店店舗参考区は対象外。",
            "Cursor",
        )
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = None

    args.output.parent.mkdir(parents=True, exist_ok=True)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(args.output)
    validate_output(args.output, len(records))

    hierarchy_changes = sum(record["old"] != record["new"] for record in records)
    branch_changes = sum(record["old"][:2] != record["new"][:2] for record in records)
    area_changes = hierarchy_changes
    print(f"output={args.output}")
    print(f"stores={len(records)}")
    print(f"hierarchy_changed_stores={hierarchy_changes}")
    print(f"branch_changed_stores={branch_changes}")
    print(f"area_changed_stores={area_changes}")
    for name in ("支社", "支店", "エリア"):
        new_nodes = [item for item in results[name] if item["pattern"] == "新規"]
        print(f"{name}_new_nodes={len(new_nodes)}")
        for item in new_nodes:
            new_pair = item["new"]["pairs"][-1]
            reference = item["reference"]["pairs"][-1] if item["reference"] else (None, None)
            print(f"  {new_pair[0]} {new_pair[1]} <- {reference[0]} {reference[1]}")


if __name__ == "__main__":
    main()
