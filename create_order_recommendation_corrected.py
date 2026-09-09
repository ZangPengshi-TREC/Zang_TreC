from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


OUT = Path(__file__).parent / "発注勧告_訂正版"
OUT.mkdir(exist_ok=True)

BLUE = "1F4E78"
LIGHT_BLUE = "D9EAF7"
LIGHT_YELLOW = "FFF2CC"
LIGHT_RED = "FCE4D6"
LIGHT_GREEN = "E2F0D9"
GRAY = "E7E6E6"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="808080")

TARGET = [
    (1, "店舗CD", "BranchCD", "KEY", "固定長", "4"),
    (2, "仕入発注JAN", "SireJAN", "KEY", "固定長", "20"),
    (3, "推奨発注数量", "SireJAN_rcmd_end", "VALUE", "可変長", "0"),
    (4, "ベンダーCD", "（名称なし）", "VALUE", "固定長", "8"),
    (5, "納品日", "ShippingDate", "VALUE", "固定長", "8"),
    (6, "Ordertype", "（名称なし）", "VALUE", "固定長", "1"),
    (7, "更新日", "（名称なし）", "VALUE", "固定長", "8"),
]

STORE_MAPPING = [
    [1, "店舗CD", "BranchCD", "KEY", "固定長4", "No.1 店舗コード", "文字・実態6桁",
     "△", "6桁→4桁のコード変換が必要。単純な左ゼロ除去は未確定", "G02"],
    [2, "仕入発注JAN", "SireJAN", "KEY", "固定長20", "No.3 商品コード", "文字・実態13桁",
     "○", "文字列のまま設定し、先頭ゼロを保持する", "G11"],
    [3, "推奨発注数量", "SireJAN_rcmd_end", "VALUE", "可変長（桁数欄0）",
     "No.11 発注バラ数", "整数8桁",
     "○", "発注単位調整後のバラ総数を直接設定。No.11＝No.12 発注単位数×No.15 発注単位。×1000しない", "G04"],
    [4, "ベンダーCD", "（名称なし）", "VALUE", "固定長8", "No.6 仕入先コード", "文字・実態9桁",
     "×", "9桁→8桁は単純切捨て不可。変換表または正式な8桁コードが必要", "G03"],
    [5, "納品日", "ShippingDate", "VALUE", "固定長8", "No.4 納品日", "yyyymmdd・8桁",
     "○", "yyyymmddをそのまま設定", ""],
    [6, "Ordertype", "（名称なし）", "VALUE", "固定長1", "対応項目なし", "—",
     "△", "固定値または業務区分から導出。値定義を確認", "G05"],
    [7, "更新日", "（名称なし）", "VALUE", "固定長8", "No.2 発注日（候補）", "yyyymmdd・8桁",
     "△", "発注日を更新日とする候補。DB更新日との意味差を確認", "G07"],
]

WAREHOUSE_MAPPING = [
    [1, "店舗CD", "BranchCD", "KEY", "固定長4", "ファイル名の倉庫コード（候補）", "実態6桁",
     "×", "項目意味が店舗CDであり、倉庫コードを設定可能か未確認", "G01 G02"],
    [2, "仕入発注JAN", "SireJAN", "KEY", "固定長20", "No.7 品番", "文字・実態13桁",
     "△", "JANを文字列のまま設定する候補。仕入発注JANとの同一性を確認", "G11"],
    [3, "推奨発注数量", "SireJAN_rcmd_end", "VALUE", "可変長（桁数欄0）",
     "No.11 バラ換算数量", "整数8桁",
     "○", "発注単位調整後のバラ総数を直接設定。No.11＝No.9 注文数量×No.12 発注ロット。×1000しない", "G04"],
    [4, "ベンダーCD", "（名称なし）", "VALUE", "固定長8", "No.3 仕入先コード", "文字・実態9桁",
     "×", "9桁→8桁は単純切捨て不可。変換表または正式な8桁コードが必要", "G03"],
    [5, "納品日", "ShippingDate", "VALUE", "固定長8", "No.14 納品日", "yyyymmdd・8桁",
     "○", "yyyymmddをそのまま設定", ""],
    [6, "Ordertype", "（名称なし）", "VALUE", "固定長1", "対応項目なし", "—",
     "△", "固定値または業務区分から導出。値定義を確認", "G05"],
    [7, "更新日", "（名称なし）", "VALUE", "固定長8", "No.1 発行日（候補）", "yyyymmdd・8桁",
     "△", "発行日を更新日とする候補。DB更新日との意味差を確認", "G07"],
]

GAPS = [
    ["G01", "適用範囲", "★最高", "倉庫系",
     "SIREJAN_RCMDORDER_LASTは店舗CDをKEYとする。倉庫系ka/subkaに店舗CDはない",
     "同表が店舗系だけの参照先か、倉庫系も対象かを確認。対象外なら倉庫系の別レイアウトを受領", "未確認"],
    ["G02", "店舗・倉庫コード桁差", "★最高", "共通",
     "sinops実態6桁に対しBranchCDは固定長4桁",
     "正式な変換規則・コード対照表を確認。単純な切捨ては禁止", "未確認"],
    ["G03", "ベンダーコード桁差", "★最高", "共通",
     "sinops実態9桁に対し対象表は固定長8桁",
     "8桁ベンダーCDの取得元または9→8桁対照表を確認", "未確認"],
    ["G04", "推奨数量の定義", "—", "共通",
     "SIREJAN_RCMDORDER_LASTの推奨発注数量は、発注単位調整後で発注単位の整数倍となるバラ総数",
     "店舗系No.11 発注バラ数／倉庫系No.11 バラ換算数量を直接設定。従来の×1000ルールは廃止", "確認済"],
    ["G05", "Ordertype値定義", "★高", "共通",
     "対象表は固定長1桁だがコード一覧がなく、sinopsに直接対応項目もない",
     "通常・特殊等の値定義とファイル別設定値を確認", "未確認"],
    ["G06", "KEY粒度・最終値化", "★最高", "共通",
     "対象表KEYはBranchCD＋SireJANのみ。sinops側は発注日・納品日等を含み、同一KEYが複数になり得る",
     "どのレコードを最終推奨数として残すか、並び順・比較項目・同値時優先順位を確定", "未確認"],
    ["G07", "更新日の意味", "★高", "共通",
     "対象表の更新日が業務日付かDB更新日か不明",
     "店舗系発注日／倉庫系発行日を設定してよいか確認", "未確認"],
    ["G08", "当日・翌日以降の反映順", "★最高", "共通",
     "kankokuとsubkankoku、kaとsubkaで同一KEYが重複すると後勝ちで結果が変わる",
     "ファイル処理順と優先規則を確定し、順序依存を排除", "未確認"],
    ["G09", "ゼロ数量・削除", "★高", "共通",
     "推奨数0を更新するか、対象外とするか、既存値を削除するか不明",
     "0件・0数量・取消時の更新仕様を確認", "未確認"],
    ["G10", "固定長編集", "★中", "共通",
     "固定長項目の左/右詰め、空白・ゼロ埋め、文字コードが画像だけでは不明",
     "物理ファイル仕様またはサンプルを受領して編集規則を確定", "未確認"],
    ["G11", "仕入発注JANの定義", "★高", "共通",
     "SireJANが通常JANと常に同一か、仕入用JANへの変換が必要か不明",
     "JAN変換マスタの要否と先頭ゼロ・20桁格納規則を確認", "未確認"],
]

RULES = [
    ["R01", "抽出", "参照先", "CoreSaver数理発注データ", "SIREJAN_RCMDORDER_LAST",
     "以降は画像の7項目を基準とし、CoreSaver固有項目は比較対象外", "確定"],
    ["R02", "コード", "BranchCD", "store_idへ数値変換", "6桁コードから固定長4桁へ変換",
     "正式規則が決まるまで実装しない", "G02"],
    ["R03", "JAN", "SireJAN", "item_id(int32)へ数値変換", "文字列で設定し先頭ゼロを保持",
     "int32超過GAPは解消。ただし仕入発注JANの定義確認が必要", "G11"],
    ["R04", "数量", "SireJAN_rcmd_end", "バラ数×1000をsetsへ設定",
     "店舗系No.11 発注バラ数／倉庫系No.11 バラ換算数量を直接設定",
     "発注単位調整後のバラ総数。発注単位の整数倍であることを検証し、×1000は行わない", "確定（G04）"],
    ["R05", "コード", "ベンダーCD", "vendor_idへ数値変換", "固定長8桁へ設定",
     "9→8桁の正式変換元が必要。切捨て禁止", "G03"],
    ["R06", "日付", "ShippingDate", "date型へ変換", "yyyymmdd固定長8桁で設定",
     "店舗No.4／倉庫No.14", "確定"],
    ["R07", "区分", "Ordertype", "odr_type等へ設定", "1桁コードを固定または導出",
     "コード定義受領後に確定", "G05"],
    ["R08", "日付", "更新日", "reg_dateへ変換", "店舗No.2／倉庫No.1を設定する候補",
     "項目意味確認後に確定", "G07"],
    ["R09", "集約", "最終推奨数", "複合KEYごとに保持", "BranchCD＋SireJANごとに1件化",
     "最終判定ロジック・同値時優先順位を確定する", "G06 G08"],
    ["R10", "入力", "ゼロ数量", "CoreSaver取込仕様に従う", "0更新・削除・無視のいずれか",
     "業務仕様確認後に確定", "G09"],
]


def style_sheet(ws, widths):
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = ws.dimensions
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(color=WHITE, bold=True, size=14)
    for cell in ws[4]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.font = Font(bold=True)
    for row in range(5, ws.max_row + 1):
        verdict = str(ws.cell(row, 8).value or "") if ws.max_column >= 8 else ""
        fill = LIGHT_GREEN if verdict == "○" else LIGHT_YELLOW if verdict == "△" else LIGHT_RED if verdict == "×" else WHITE
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor=fill)


def add_mapping_sheet(wb, title, rows, source_files):
    ws = wb.create_sheet(title)
    headers = ["No.", "対象項目", "Field name", "Type", "対象型/桁", "sinops対応元", "sinops型/桁", "判定", "変換ルール・留意事項", "関連GAP"]
    ws.append([f"発注勧告 項目マッピング — {title}（訂正版）"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.append([f"参照先：SIREJAN_RCMDORDER_LAST／対象：{source_files}"])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.append(["訂正：CoreSaver数理発注データとの比較を廃止。推奨発注数量は発注単位調整後のバラ総数を直接設定する。コード桁・最終値化は未確定GAPとして残す。"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_sheet(ws, [7, 18, 22, 10, 18, 27, 18, 8, 53, 13])


def write_target_sheet(wb):
    ws = wb.create_sheet("参照レイアウト", 0)
    headers = ["No.", "Label", "Field name", "Type", "Expression Type", "Number of digits"]
    ws.append(["参照レイアウト転記 — SIREJAN_RCMDORDER_LAST"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.append(["Note：仕入JAN_最終推奨数の反映"])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.append(["出典：ユーザー提供画像（2026-09-04）。画像にない物理編集規則・コード定義は推測せずGAPに残す。"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))
    ws.append(headers)
    for row in TARGET:
        ws.append(row)
    ws.freeze_panes = "A5"
    for i, w in enumerate([7, 20, 25, 12, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(color=WHITE, bold=True, size=14)
    for cell in ws[4]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.font = Font(bold=True)


def build_mapping():
    wb = Workbook()
    wb.remove(wb.active)
    write_target_sheet(wb)
    add_mapping_sheet(wb, "店舗系_当日", STORE_MAPPING, "kankoku.txt")
    add_mapping_sheet(wb, "店舗系_翌日以降", STORE_MAPPING, "subkankoku.txt")
    add_mapping_sheet(wb, "倉庫系_当日", WAREHOUSE_MAPPING, "【倉庫CD】ka.txt")
    add_mapping_sheet(wb, "倉庫系_翌日以降", WAREHOUSE_MAPPING, "【倉庫CD】subka.txt")
    wb.save(OUT / "発注勧告_①_項目マッピング表.xlsx")


def build_gap():
    wb = Workbook()
    ws = wb.active
    ws.title = "GAP分析一覧"
    headers = ["No.", "GAPカテゴリ", "優先度", "影響範囲", "現状・差異", "確認・対応方針", "ステータス"]
    ws.append(["GAP分析書 — 発注勧告IF（SIREJAN_RCMDORDER_LAST基準）"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.append(["旧CoreSaver比較のGAPは失効。新レイアウトで必要な論点だけを再整理。"])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.append(["特にG01/G02/G03/G04/G06/G08は実装開始前に解決が必要。"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))
    ws.append(headers)
    for row in GAPS:
        ws.append(row)
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = ws.dimensions
    for i, w in enumerate([8, 24, 12, 15, 58, 58, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(color=WHITE, bold=True, size=14)
    for cell in ws[4]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.font = Font(bold=True)
    for row in range(5, ws.max_row + 1):
        priority = str(ws.cell(row, 3).value)
        fill = LIGHT_RED if priority == "★最高" else LIGHT_YELLOW if priority == "★高" else GRAY
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor=fill)
    wb.save(OUT / "発注勧告_②_GAP分析書.xlsx")


def build_rules():
    wb = Workbook()
    ws = wb.active
    ws.title = "変更ルール"
    headers = ["Rule ID", "分類", "対象", "旧ルール", "新ルール", "備考", "状態/関連GAP"]
    ws.append(["変更ルール定義書 — 発注勧告IF（参照先変更）"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.append(["CoreSaver比較からSIREJAN_RCMDORDER_LAST比較へ変更"])
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    ws.append(["未確定ルールはGAP解決まで実装条件として使用しない。"])
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))
    ws.append(headers)
    for row in RULES:
        ws.append(row)
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = ws.dimensions
    for i, w in enumerate([11, 13, 23, 38, 46, 48, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(color=WHITE, bold=True, size=14)
    for cell in ws[4]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        cell.font = Font(bold=True)
    for row in range(5, ws.max_row + 1):
        state = str(ws.cell(row, 7).value or "")
        fill = LIGHT_GREEN if state.startswith("確定") else LIGHT_YELLOW
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor=fill)
    wb.save(OUT / "発注勧告_③_変更ルール定義書.xlsx")


def verify():
    expected = {
        "発注勧告_①_項目マッピング表.xlsx": 5,
        "発注勧告_②_GAP分析書.xlsx": 1,
        "発注勧告_③_変更ルール定義書.xlsx": 1,
    }
    for name, sheet_count in expected.items():
        path = OUT / name
        wb = load_workbook(path, data_only=True, read_only=True)
        assert len(wb.sheetnames) == sheet_count, (name, wb.sheetnames)
        assert all(wb[s].max_row >= 10 for s in wb.sheetnames), (name, wb.sheetnames)


if __name__ == "__main__":
    build_mapping()
    build_gap()
    build_rules()
    verify()
    print(f"created: {OUT}")
