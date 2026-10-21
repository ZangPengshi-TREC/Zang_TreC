# -*- coding: utf-8 -*-
"""生鮮発注勧告の三資料を、Downloads の発注勧告テンプレと同じ体裁で出す。"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pathlib import Path

OUT = Path(__file__).resolve().parent / "生鮮_発注勧告"

TEAL = "0D7377"
NAVY = "1A2E4A"
NAVY2 = "243A5E"
SECTION = "D6E4F0"
LEGEND = "FFFBF0"
JUDGE = "FEF3C7"
WHITE = "FFFFFF"
PRIO = "D1FAE5"
SLATE = "F1F5F9"
GREEN = "D1FAE5"

THIN = Border(
    left=Side(style="thin", color="D0D5DD"),
    right=Side(style="thin", color="D0D5DD"),
    top=Side(style="thin", color="D0D5DD"),
    bottom=Side(style="thin", color="D0D5DD"),
)
FONT = "Meiryo"
C_WHITE = Font(name=FONT, size=9, bold=True, color=WHITE)
C_TITLE = Font(name=FONT, size=12, bold=True, color=WHITE)
C_LEGEND = Font(name=FONT, size=9, color="595959")
C_SECTION = Font(name="宋体", size=11, bold=True, color="2F5496")
C_CELL = Font(name=FONT, size=9)
C_JUDGE = Font(name=FONT, size=9, bold=True, color="92400E")
C_RULE = Font(name=FONT, size=8, bold=True, color="92400E")
C_NO = Font(name=FONT, size=10, bold=True, color=WHITE)

CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def apply_sheet(ws, widths, freeze, landscape=True):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = freeze
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.4
    ws.sheet_view.zoomScale = 110
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def put(ws, row, col, value, font, fill_=None, align=None, border=True):
    cell = ws.cell(row, col, value)
    cell.font = font
    cell.alignment = align or LEFT
    if fill_:
        cell.fill = fill_
    if border:
        cell.border = THIN
    return cell


def banner(ws, row, cols, text, font, fill_, height=22):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = put(ws, row, 1, text, font, fill_, CENTER)
    for c in range(2, cols + 1):
        put(ws, row, c, None, font, fill_, CENTER)
    ws.row_dimensions[row].height = height
    return cell


def section(ws, row, cols, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    put(ws, row, 1, text, C_SECTION, fill(SECTION), LEFT)
    for c in range(2, cols + 1):
        put(ws, row, c, None, C_SECTION, fill(SECTION), LEFT)
    ws.row_dimensions[row].height = 20


def paint_row(ws, row, values, fonts, fills, height=42):
    for i, val in enumerate(values, 1):
        put(ws, row, i, val, fonts[i - 1], fills[i - 1], LEFT if i not in (1, 2, 8) else CENTER)
    ws.row_dimensions[row].height = height


# 方向は Sinops → TRIAL。行は受信先（TRIAL表）の順。右が送信元。
# 判定: ○直連携  △変換  ✕未決  ―固定・NULL・送らない
# (区分, No, 必須, TRIAL項目, TRIAL型, 設定, sinops項目, sinops型, 判定, ルール)
ROWS = [
    ("idx", "001", "○", "JAN / ProdJAN", "varchar / 13", "JAN",
     "商品コード", "文字 / 13桁", "△", "統合後。切捨て・ゼロ埋めなし"),
    ("idx", "002", "○", "店CD / BranchCD", "int / 4", "店",
     "店舗コード", "文字 / 6桁", "△", "統合後を int。桁は統合"),
    ("idx", "003", "○", "ベンダーCD / SupplierCD", "int / 4", "仕入先",
     "仕入先コード", "文字 / 9桁", "△", "統合後を int。桁は統合"),
    ("idx", "004", "○", "納品日 / DeliveryDate", "datetime / 8", "納品日",
     "納品日", "yyyymmdd", "△", "yyyymmdd → datetime。時刻未定"),
    ("idx", "006", "○", "便 / Bin", "int / 4", "便なし=9999",
     "便区分", "文字 / 2桁", "✕", "int化。便なし=9999。固定1は使わない"),
    ("body", "009", "－", "推奨数 / RecQuantity", "money / 8  NULL可", "推奨数",
     "発注バラ数", "整数 / 8桁", "✕", "OrderQtyとの分担未定。×1000しない"),
    ("body", "010", "○", "発注数 / OrderQty", "money / 8", "必須",
     "—", "—", "✕", "対応未定"),
    ("body", "011", "－", "発注数 / CustOrderQty", "money / 8  NULL可", "NULL",
     "—", "—", "―", "NULL"),
    ("body", "012", "○", "発注種類CD / OrderKindCD", "tinyint / 1", "1",
     "—", "—", "―", "固定1（普通発注）"),
    ("body", "013", "○", "登録FLG / EntryFLG", "tinyint / 1", "0",
     "—", "—", "―", "固定0。1=店舗登録済、-1以外=店舗未登録"),
    ("body", "014", "－", "担当CD / UserCD", "int / 4  NULL可", "NULL",
     "—", "—", "―", "NULL"),
    ("body", "015", "○", "登録日 / CreateTime", "datetime / 8", "getdate",
     "発注日", "yyyymmdd", "✕", "発注日を書くか未定"),
    ("body", "016", "－", "更新日 / UpdateTime", "datetime / 8  NULL可", "getdate",
     "発注日", "yyyymmdd", "✕", "同上"),
    ("drop", "—", "—", "—", "—", "対象外",
     "発注日", "yyyymmdd", "―", "列なし。015/016参照"),
    ("drop", "—", "—", "—", "—", "対象外",
     "大分類コード", "文字 / 2桁", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "中分類コード", "文字 / 4桁", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "小分類コード", "文字 / 8桁", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "細目コード", "文字", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "発注単位数", "整数 / 8桁", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "標準仕入単価", "整数", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "標準売価", "整数", "―", "列なし"),
    ("drop", "—", "—", "—", "—", "対象外",
     "発注単位", "整数 / 8桁", "―", "列なし"),
]


SHEETS = [
    ("kankoku.txt (当日分)",
     "項目マッピング表 — 生鮮発注勧告 当日（kankoku.txt）",
     "通常採用"),
    ("subkankoku.txt (翌日以降分)",
     "項目マッピング表 — 生鮮発注勧告 翌日以降（subkankoku.txt）",
     "障害時のみ"),
]


def build_mapping(wb, sheets=None):
    sheets = SHEETS if sheets is None else sheets
    headers = ["TRIAL No.", "必須", "TRIAL 項目名（受信）", "TRIAL 型/桁", "設定内容",
               "sinops 項目名（送信）", "sinops 型/桁", "判定", "変換ルール（Sinops → TRIAL）"]
    for name, title, adopt in sheets:
        ws = wb.create_sheet(name)
        apply_sheet(ws, {
            "A": 12, "B": 12, "C": 18, "D": 24, "E": 28, "F": 32, "G": 22, "H": 16, "I": 62,
        }, "A4")
        banner(ws, 1, 9, title, C_TITLE, fill(TEAL), 26)
        banner(ws, 2, 9,
               "○直連携　△変換　✕未決　―固定・NULL・対象外"
               "　　192.168.1.129 / dborder / shT_OrderQty_test　　" + adopt,
               C_LEGEND, fill(LEGEND), 32)
        for i, h in enumerate(headers, 1):
            put(ws, 3, i, h, C_WHITE, fill(TEAL), CENTER)
        ws.row_dimensions[3].height = 28
        row = 4
        section(ws, row, 9, "■ 索引（JAN / 店CD / ベンダーCD / 納品日 / 便）")
        row = 5
        for kind, no, req, t_name, t_type, setting, s_name, s_type, judge, rule in ROWS:
            if kind == "body" and no == "009":
                section(ws, row, 9, "■ その他")
                row += 1
            if kind == "drop" and s_name == "発注日":
                section(ws, row, 9, "■ 対象外")
                row += 1
            vals = [no, req, t_name, t_type, setting, s_name, s_type, judge, rule]
            fonts = [C_CELL, C_CELL, C_CELL, C_CELL, C_CELL, C_CELL, C_CELL, C_JUDGE, C_RULE]
            fills = [fill(WHITE)] * 7 + [fill(JUDGE), fill(JUDGE)]
            paint_row(ws, row, vals, fonts, fills, 48 if len(rule) > 40 else 36)
            row += 1
        ws.print_title_rows = "1:3"
        ws.oddHeader.left.text = name
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]


# (No, カテゴリ, 優先度, ファイル, 項目, 送信元sinops, 受信先TRIAL, 方針, 状態)
GAPS = [
    (1, "数量", "★最高", "kankoku / subkankoku", "発注バラ数",
     "発注バラ数", "推奨数 / OrderQty",
     "分担未定。×1000しない", "未確認"),
    (2, "便", "★高", "kankoku / subkankoku", "便",
     "便区分", "int。便なし=9999",
     "9999にする値未定", "未確認"),
    (3, "発注日", "★高", "kankoku / subkankoku", "発注日",
     "発注日", "列なし。既定 getdate",
     "書く列未定", "未確認"),
    (4, "納品日", "★中", "kankoku / subkankoku", "納品日",
     "yyyymmdd", "datetime",
     "時刻未定", "未確認"),
    (5, "同一キー", "★高", "kankoku / subkankoku", "索引5列",
     "—", "PK記載なし",
     "追加か更新か未定", "未確認"),
    (6, "表名", "★中", "kankoku / subkankoku", "shT_OrderQty_test",
     "—", "192.168.1.129 / dborder",
     "test のままか未定", "未確認"),
    (7, "採用", "—", "kankoku / subkankoku", "—",
     "当日 / 翌日以降", "同一表",
     "通常は当日。障害時は前日 subkankoku の当日分。時刻未定", "確認済"),
    (8, "桁", "—", "kankoku / subkankoku", "店・JAN・仕入先",
     "6桁 / 13桁 / 9桁", "int4 / varchar13 / int4",
     "マスタ統合", "確認済"),
    (9, "対象外", "—", "kankoku / subkankoku", "分類・単価・単位",
     "分類・単価・単位", "列なし",
     "送らない", "確認済"),
    (10, "発注種類CD", "—", "kankoku / subkankoku", "発注種類CD",
     "—", "tinyint",
     "固定 1", "確認済"),
    (11, "登録FLG", "—", "kankoku / subkankoku", "登録FLG",
     "—", "tinyint",
     "固定 0", "確認済"),
    (12, "担当CD", "—", "kankoku / subkankoku", "担当CD",
     "—", "NULL可",
     "NULL", "確認済"),
]


def build_gap(wb, title=None, file_label=None):
    ws = wb.active
    ws.title = "GAP分析一覧"
    apply_sheet(ws, {
        "A": 8, "B": 22, "C": 10, "D": 28, "E": 28, "F": 36, "G": 36, "H": 55, "I": 12,
    }, "A4")
    headers = ["No.", "GAPカテゴリ", "優先度", "影響ファイル", "受信先（TRIAL）",
               "送信元（sinops）", "項目", "対応方針", "ステータス"]
    banner(ws, 1, 9, title or "GAP分析書 — 生鮮発注勧告（sinops kankoku / subkankoku → shT_OrderQty_test）",
           C_TITLE, fill(NAVY), 26)
    banner(ws, 2, 9, "192.168.1.129 / dborder / shT_OrderQty_test",
           Font(name=FONT, size=9, color=WHITE), fill(NAVY2), 20)
    for i, h in enumerate(headers, 1):
        put(ws, 3, i, h, C_WHITE, fill(TEAL), CENTER)
    ws.row_dimensions[3].height = 28
    for i, item in enumerate(GAPS):
        row = 4 + i
        no, cat, prio, files, target, sinops, trial, policy, status = item
        if file_label:
            files = file_label
        vals = [no, cat, prio, files, trial, sinops, target, policy, status]
        for c, val in enumerate(vals, 1):
            font = C_NO if c == 1 else (C_JUDGE if c in (3, 9) else C_CELL)
            bg = TEAL if c == 1 else (PRIO if c == 3 else (JUDGE if c == 9 else WHITE))
            put(ws, row, c, val, font, fill(bg), CENTER if c in (1, 3, 9) else LEFT)
        ws.row_dimensions[row].height = 56 if len(policy) > 40 else 36
    ws.auto_filter.ref = f"A3:I{3 + len(GAPS)}"
    ws.print_title_rows = "1:3"


RULES = [
    (1, "店舗コード", "文字(6)", "店CD", "int/4", "統合後を int", "—", "非数値はエラー"),
    (2, "商品コード", "文字(13)", "JAN", "varchar/13", "統合後をそのまま", "—", "切捨て・ゼロ埋めなし"),
    (3, "仕入先コード", "文字(9)", "ベンダーCD", "int/4", "統合後を int", "—", "非数値はエラー"),
    (4, "納品日", "yyyymmdd", "納品日", "datetime", "yyyymmdd → datetime", "20261020", "時刻未定"),
    (5, "便区分", "文字(2)", "便", "int/4", "int化。便なし=9999", "—", "対象値未定"),
    (6, "発注日", "yyyymmdd", "登録日・更新日", "datetime", "未定。空なら getdate", "—", "列未定"),
    (7, "発注バラ数", "整数(8)", "推奨数 / OrderQty", "money", "分担未定。×1000しない", "—", "—"),
    (8, "大分類〜細目", "文字", "—", "—", "対象外", "—", "列なし"),
    (9, "発注単位数・単価・売価・発注単位", "整数", "—", "—", "対象外", "—", "列なし"),
]

FIXED = [
    ("固定", "発注種類CD", "1", "普通発注", "両ファイル"),
    ("固定", "登録FLG", "0", "—", "両ファイル"),
    ("NULL", "担当CD", "NULL", "—", "両ファイル"),
    ("NULL", "CustOrderQty", "NULL", "—", "両ファイル"),
    ("既定", "登録日・更新日", "getdate", "発注日の書込は未定", "両ファイル"),
    ("対象外", "分類・単価・単位", "送らない", "列なし", "両ファイル"),
]


def build_rules(wb, title=None):
    ws = wb.active
    ws.title = "型変換ルール"
    apply_sheet(ws, {
        "A": 6, "B": 28, "C": 18, "D": 22, "E": 18, "F": 46, "G": 32, "H": 28,
    }, "A3")
    banner(ws, 1, 8, title or "変換ルール定義書 — 型変換ルール（Sinops → shT_OrderQty_test）",
           C_TITLE, fill(NAVY), 26)
    headers = ["No.", "TRIAL項目（受信）", "TRIAL型", "sinops項目（送信）", "sinops型",
               "変換ロジック（Sinops → TRIAL）", "具体例", "注意事項"]
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
            bg_hex = TEAL if c == 1 else bg
            put(ws, row, c, val, font, fill(bg_hex), CENTER if c == 1 else LEFT)
        ws.row_dimensions[row].height = 36
    ws.auto_filter.ref = f"A2:H{2 + len(RULES)}"
    ws.print_title_rows = "1:2"

    ws2 = wb.create_sheet("NULL・固定値設定ルール")
    apply_sheet(ws2, {"A": 14, "B": 36, "C": 18, "D": 72, "E": 18}, "A3", landscape=True)
    banner(ws2, 1, 5, "NULL・固定値設定ルール定義", C_TITLE, fill(NAVY), 26)
    headers2 = ["設定種別", "受信項目（TRIAL）", "設定値", "送信元・根拠", "ファイル"]
    for i, h in enumerate(headers2, 1):
        put(ws2, 2, i, h, C_WHITE, fill(NAVY2), CENTER)
    ws2.row_dimensions[2].height = 22
    for i, item in enumerate(FIXED):
        row = 3 + i
        kind, name, val, reason, file_ = item
        bg = JUDGE if kind == "固定" else (SLATE if kind == "NULL" else WHITE)
        vals = [kind, name, val, reason, file_]
        for c, v in enumerate(vals, 1):
            put(ws2, row, c, v, C_JUDGE if c in (1, 3) else C_CELL, fill(bg if c in (1, 3) else WHITE),
                CENTER if c in (1, 3, 5) else LEFT)
        ws2.row_dimensions[row].height = 36
    ws2.auto_filter.ref = f"A2:E{2 + len(FIXED)}"
    ws2.print_title_rows = "1:2"


def write_set(folder, names, mapping_sheets, gap_title, gap_file, rule_title):
    folder.mkdir(parents=True, exist_ok=True)
    specs = [
        (names[0], lambda wb: build_mapping(wb, mapping_sheets)),
        (names[1], lambda wb: build_gap(wb, gap_title, gap_file)),
        (names[2], lambda wb: build_rules(wb, rule_title)),
    ]
    for name, builder in specs:
        wb = Workbook()
        builder(wb)
        path = folder / name
        wb.save(path)
        print(path)


def main():
    write_set(
        OUT,
        [
            "発注勧告_①_項目マッピング表(生鮮kankoku subkankoku).xlsx",
            "発注勧告_②_GAP分析書(生鮮kankoku subkankoku).xlsx",
            "発注勧告_③_変換ルール定義書(生鮮kankoku subkankoku).xlsx",
        ],
        SHEETS,
        None,
        None,
        None,
    )
    write_set(
        OUT.parent / "生鮮_発注勧告翌日以降",
        [
            "発注勧告_①_項目マッピング表(生鮮subkankoku).xlsx",
            "発注勧告_②_GAP分析書(生鮮subkankoku).xlsx",
            "発注勧告_③_変換ルール定義書(生鮮subkankoku).xlsx",
        ],
        [SHEETS[1]],
        "GAP分析書 — 生鮮発注勧告 翌日以降（subkankoku.txt → shT_OrderQty_test）",
        "subkankoku.txt",
        "変換ルール定義書 — 型変換ルール（生鮮発注勧告 翌日以降 → shT_OrderQty_test）",
    )


if __name__ == "__main__":
    main()
