# -*- coding: utf-8 -*-
"""SIREJAN_RCMDORDER_LAST（7項目ファイル）と shT_OrderQty_test の比較三資料。"""
import importlib.util
from openpyxl import Workbook
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "SIREJAN_対_shT_OrderQty_test"

spec = importlib.util.spec_from_file_location(
    "seisen_fmt", HERE / "create_seisen_template_format_20261019.py")
fmt = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fmt)

fill = fmt.fill
put = fmt.put
banner = fmt.banner
section = fmt.section
apply_sheet = fmt.apply_sheet
paint_row = fmt.paint_row
TEAL, NAVY, NAVY2 = fmt.TEAL, fmt.NAVY, fmt.NAVY2
WHITE, JUDGE, PRIO, SLATE = fmt.WHITE, fmt.JUDGE, fmt.PRIO, fmt.SLATE
C_WHITE, C_TITLE, C_LEGEND = fmt.C_WHITE, fmt.C_TITLE, fmt.C_LEGEND
C_CELL, C_JUDGE, C_RULE, C_NO = fmt.C_CELL, fmt.C_JUDGE, fmt.C_RULE, fmt.C_NO
CENTER, LEFT = fmt.CENTER, fmt.LEFT

# 左が受信先 shT_OrderQty_test。右が SIREJAN_RCMDORDER_LAST の7項目。
# LAST_SUB も同じ7項目。着地フォルダだけ違う。
ROWS = [
    ("idx", "001", "○", "JAN / ProdJAN", "varchar / 13", "仕入発注JAN",
     "仕入発注JAN / SireJAN", "固定長 / 20", "△", "統合後。切捨てなし"),
    ("idx", "002", "○", "店CD / BranchCD", "int / 4", "店舗CD",
     "店舗CD / BranchCD", "固定長 / 4", "△", "文字→int。桁は統合"),
    ("idx", "003", "○", "ベンダーCD / SupplierCD", "int / 4", "ベンダーCD",
     "ベンダーCD", "固定長 / 8", "△", "文字→int。桁は統合"),
    ("idx", "004", "○", "納品日 / DeliveryDate", "datetime / 8", "納品日",
     "納品日 / ShippingDate", "固定長 / 8  yyyymmdd", "△", "yyyymmdd → datetime。年月日のみ"),
    ("idx", "005", "○", "便 / Bin", "int / 4", "TRIALの便",
     "—", "—", "―", "表側でTRIALの便"),
    ("body", "006", "－", "推奨数 / RecQuantity", "money / 8  NULL可", "推奨発注数量",
     "推奨発注数量 / SireJAN_rcmd_end", "可変長", "○", "そのまま"),
    ("body", "007", "－", "発注数 / OrderQty", "money / 8  NULL可", "推奨発注数量",
     "推奨発注数量 / SireJAN_rcmd_end", "可変長", "○", "006と同じ"),
    ("body", "008", "－", "客注数 / CustOrderQty", "money / 8  NULL可", "NULL",
     "—", "—", "―", "空値"),
    ("body", "009", "○", "発注種類CD / OrderKindCD", "tinyint / 1", "1",
     "—", "—", "―", "固定1。通常発注"),
    ("body", "010", "○", "登録FLG / EntryFLG", "tinyint / 1", "0",
     "—", "—", "―", "固定0。1=店舗登録済、-1以外=店舗未登録"),
    ("body", "011", "－", "担当CD / UserCD", "int / 4  NULL可", "未定",
     "—", "—", "✕", "空にしない。運用サポートへ新規システムCDを申請"),
    ("body", "012", "○", "登録日 / CreateTime", "datetime / 8", "getdate()",
     "—", "—", "―", "現在日時。時分秒あり"),
    ("body", "013", "－", "更新日 / UpdateTime", "datetime / 8  NULL可", "—",
     "—", "—", "―", "列なし"),
]

# (No, カテゴリ, 優先度, ファイル, 項目, 送信元, 受信先, 方針, 状態)
GAPS = [
    (1, "商品コード", "—", "SIREJAN_RCMDORDER_LAST", "仕入発注JAN",
     "固定長20", "varchar13",
     "統合後。切捨てなし", "確認済"),
    (2, "店CD", "—", "SIREJAN_RCMDORDER_LAST", "店舗CD",
     "固定長4", "int4",
     "文字→int。桁は統合", "確認済"),
    (3, "仕入先", "—", "SIREJAN_RCMDORDER_LAST", "ベンダーCD",
     "固定長8", "int4",
     "文字→int。桁は統合", "確認済"),
    (4, "納品日", "—", "SIREJAN_RCMDORDER_LAST", "納品日",
     "yyyymmdd", "datetime",
     "年月日のみ", "確認済"),
    (5, "数量", "—", "SIREJAN_RCMDORDER_LAST", "推奨発注数量",
     "推奨発注数量", "推奨数・発注数",
     "両方に推奨発注数量", "確認済"),
    (6, "便", "—", "SIREJAN_RCMDORDER_LAST", "便",
     "列なし", "TRIALの便",
     "表側でTRIALの便を入れる", "確認済"),
    (7, "発注種類CD", "—", "SIREJAN_RCMDORDER_LAST", "発注種類CD",
     "—", "1",
     "固定1。通常発注", "確認済"),
    (8, "登録日", "—", "SIREJAN_RCMDORDER_LAST", "登録日",
     "—", "getdate()",
     "現在日時。時分秒あり", "確認済"),
    (9, "登録FLG", "—", "SIREJAN_RCMDORDER_LAST", "登録FLG",
     "—", "0",
     "固定0", "確認済"),
    (10, "担当CD", "★高", "SIREJAN_RCMDORDER_LAST", "担当CD",
     "—", "空にしない",
     "運用サポートへ新規システムCDを申請", "未確認"),
    (11, "同一キー", "★高", "SIREJAN_RCMDORDER_LAST", "キー",
     "店+JAN", "JAN・店・ベンダー・納品日・便",
     "ファイルのキーと表の唯一キーが違う", "未確認"),
    (12, "客注数", "—", "SIREJAN_RCMDORDER_LAST", "客注数",
     "列なし", "NULL可",
     "空値", "確認済"),
    (13, "ファイル更新日", "★高", "SIREJAN_RCMDORDER_LAST", "更新日",
     "yyyymmdd", "—",
     "登録日には入れない。入れ先未定", "未確認"),
]

RULES = [
    (1, "店舗CD", "固定長(4)", "店CD", "int/4", "統合後を int", "—", "非数値はエラー"),
    (2, "仕入発注JAN", "固定長(20)", "JAN", "varchar/13", "統合後をそのまま", "—", "切捨てなし"),
    (3, "ベンダーCD", "固定長(8)", "ベンダーCD", "int/4", "統合後を int", "—", "非数値はエラー"),
    (4, "納品日", "yyyymmdd", "納品日", "datetime", "yyyymmdd → datetime", "20261020", "年月日のみ"),
    (5, "推奨発注数量", "可変長", "推奨数・発注数", "money", "両方にそのまま", "—", "—"),
    (6, "Ordertype", "固定長(1)", "—", "—", "対応しない", "0 / 1", "発注種類CDは固定1。通常発注"),
    (7, "更新日", "yyyymmdd", "—", "—", "登録日には入れない", "—", "入れ先未定"),
]

FIXED = [
    ("表側", "便", "TRIALの便", "ファイルに列なし。表側で具体の便を入れる", "LAST"),
    ("固定", "発注種類CD", "1", "通常発注", "LAST"),
    ("固定", "登録FLG", "0", "1=店舗登録済、-1以外=店舗未登録", "LAST"),
    ("未定", "担当CD", "空にしない", "運用サポートへ新規システムCDを申請", "LAST"),
    ("NULL", "客注数", "NULL", "空値", "LAST"),
    ("既定", "登録日", "getdate()", "現在日時。時分秒あり", "LAST"),
    ("対象外", "Ordertype", "入れない", "0非日配 / 1日配・生鮮。発注種類CDではない", "LAST"),
    ("未定", "ファイルの更新日", "—", "yyyymmdd。登録日には入れない。入れ先未定", "LAST"),
]


def build_mapping(wb):
    ws = wb.active
    ws.title = "比較"
    apply_sheet(ws, {
        "A": 12, "B": 12, "C": 28, "D": 24, "E": 22, "F": 36, "G": 24, "H": 16, "I": 42,
    }, "A4")
    banner(ws, 1, 9, "項目マッピング表 — SIREJAN_RCMDORDER_LAST → shT_OrderQty_test",
           C_TITLE, fill(TEAL), 26)
    banner(ws, 2, 9,
           "○直連携　△変換　✕未決　―なし"
           "　　ファイル LV3/SIREJAN_RCMDORDER_LAST/YYYYMMDD/"
           "　　表 192.168.1.129 / dborder / shT_OrderQty_test"
           "　　LAST_SUB も同じ7項目",
           C_LEGEND, fill(fmt.LEGEND), 32)
    headers = ["表 No.", "必須", "表の項目（shT_OrderQty_test）", "表の型/桁", "設定内容",
               "SIREJAN 項目", "SIREJAN 型/桁", "判定", "比較メモ"]
    for i, h in enumerate(headers, 1):
        put(ws, 3, i, h, C_WHITE, fill(TEAL), CENTER)
    ws.row_dimensions[3].height = 28
    row = 4
    section(ws, row, 9, "■ 表の唯一キー（JAN / 店CD / ベンダーCD / 納品日 / 便）")
    row = 5
    for kind, no, req, t_name, t_type, setting, s_name, s_type, judge, rule in ROWS:
        if kind == "body" and no == "006":
            section(ws, row, 9, "■ 数量・区分・時刻")
            row += 1
        vals = [no, req, t_name, t_type, setting, s_name, s_type, judge, rule]
        fonts = [C_CELL, C_CELL, C_CELL, C_CELL, C_CELL, C_CELL, C_CELL, C_JUDGE, C_RULE]
        fills = [fill(WHITE)] * 7 + [fill(JUDGE), fill(JUDGE)]
        paint_row(ws, row, vals, fonts, fills, 48 if len(rule) > 28 else 36)
        row += 1
    ws.print_title_rows = "1:3"


def build_gap(wb):
    ws = wb.active
    ws.title = "GAP分析一覧"
    apply_sheet(ws, {
        "A": 8, "B": 16, "C": 10, "D": 28, "E": 28, "F": 28, "G": 28, "H": 42, "I": 12,
    }, "A4")
    banner(ws, 1, 9, "GAP分析書 — SIREJAN_RCMDORDER_LAST と shT_OrderQty_test",
           C_TITLE, fill(NAVY), 26)
    banner(ws, 2, 9, "LV3/SIREJAN_RCMDORDER_LAST/YYYYMMDD/  →  192.168.1.129 / dborder / shT_OrderQty_test",
           fmt.Font(name=fmt.FONT, size=9, color=WHITE), fill(NAVY2), 20)
    headers = ["No.", "GAPカテゴリ", "優先度", "ファイル", "受信先（表）",
               "送信元（SIREJAN）", "項目", "対応方針", "ステータス"]
    for i, h in enumerate(headers, 1):
        put(ws, 3, i, h, C_WHITE, fill(TEAL), CENTER)
    ws.row_dimensions[3].height = 28
    for i, item in enumerate(GAPS):
        row = 4 + i
        no, cat, prio, files, target, sinops, trial, policy, status = item
        vals = [no, cat, prio, files, trial, sinops, target, policy, status]
        for c, val in enumerate(vals, 1):
            font = C_NO if c == 1 else (C_JUDGE if c in (3, 9) else C_CELL)
            bg = TEAL if c == 1 else (PRIO if c == 3 else (JUDGE if c == 9 else WHITE))
            put(ws, row, c, val, font, fill(bg), CENTER if c in (1, 3, 9) else LEFT)
        ws.row_dimensions[row].height = 48 if len(policy) > 24 else 32
    ws.auto_filter.ref = f"A3:I{3 + len(GAPS)}"
    ws.print_title_rows = "1:3"


def build_rules(wb):
    ws = wb.active
    ws.title = "型変換ルール"
    apply_sheet(ws, {
        "A": 6, "B": 22, "C": 16, "D": 22, "E": 16, "F": 36, "G": 16, "H": 28,
    }, "A3")
    banner(ws, 1, 8, "変換ルール定義書 — SIREJAN_RCMDORDER_LAST → shT_OrderQty_test",
           C_TITLE, fill(NAVY), 26)
    headers = ["No.", "表の項目（受信）", "表の型", "SIREJAN項目（送信）", "SIREJAN型",
               "変換（ファイル → 表）", "具体例", "注意事項"]
    for i, h in enumerate(headers, 1):
        put(ws, 2, i, h, C_WHITE, fill(TEAL), CENTER)
    ws.row_dimensions[2].height = 28
    for i, item in enumerate(RULES):
        row = 3 + i
        no, s_item, s_type, t_item, t_type, logic, example, note = item
        vals = [no, t_item, t_type, s_item, s_type, logic, example, note]
        for c, val in enumerate(vals, 1):
            bg = SLATE if i % 2 else WHITE
            font = C_NO if c == 1 else C_CELL
            put(ws, row, c, val, font, fill(TEAL if c == 1 else bg), CENTER if c == 1 else LEFT)
        ws.row_dimensions[row].height = 36
    ws.auto_filter.ref = f"A2:H{2 + len(RULES)}"
    ws.print_title_rows = "1:2"

    ws2 = wb.create_sheet("ファイルに無い列")
    apply_sheet(ws2, {"A": 14, "B": 28, "C": 16, "D": 62, "E": 14}, "A3", landscape=True)
    banner(ws2, 1, 5, "ファイルに無い列・別項目", C_TITLE, fill(NAVY), 26)
    headers2 = ["設定種別", "表の項目", "設定値", "根拠", "ファイル"]
    for i, h in enumerate(headers2, 1):
        put(ws2, 2, i, h, C_WHITE, fill(NAVY2), CENTER)
    ws2.row_dimensions[2].height = 22
    for i, item in enumerate(FIXED):
        row = 3 + i
        kind, name, val, reason, file_ = item
        bg = SLATE if kind == "NULL" else (WHITE if kind == "対象外" else JUDGE)
        vals = [kind, name, val, reason, file_]
        for c, v in enumerate(vals, 1):
            put(ws2, row, c, v, C_JUDGE if c in (1, 3) else C_CELL,
                fill(bg if c in (1, 3) else WHITE),
                CENTER if c in (1, 3, 5) else LEFT)
        ws2.row_dimensions[row].height = 36
    ws2.auto_filter.ref = f"A2:E{2 + len(FIXED)}"
    ws2.print_title_rows = "1:2"


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    names = [
        ("①項目マッピング表(SIREJAN_RCMDORDER_LAST).xlsx", build_mapping),
        ("②GAP分析書(SIREJAN_RCMDORDER_LAST).xlsx", build_gap),
        ("③変換ルール定義書(SIREJAN_RCMDORDER_LAST).xlsx", build_rules),
    ]
    for name, builder in names:
        wb = Workbook()
        builder(wb)
        path = OUT / name
        wb.save(path)
        print(path)


if __name__ == "__main__":
    main()
