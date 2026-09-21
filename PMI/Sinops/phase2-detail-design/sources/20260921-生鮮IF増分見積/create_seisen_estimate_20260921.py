#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate standalone 生鮮 IF 増分見積 — 標準案 / リスク込み案 を別ファイルで出力。

既存 403 workbook は変更しない。
Rev.2026-09-21e: 分位用語を使わず「標準案」「リスク込み案」で分冊。各ファイルに PJ管理按分を含む。
"""

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent
UNIT = 27500
UNIT_PJ = 40000
BLUE = "1F4E78"
LIGHT = "D9EAF7"
GREEN = "E2F0D9"
ORANGE = "FCE4D6"
THIN = Side(style="thin", color="808080")
PJ_RATIO = 50.0 / 499.5  # 既存: PJ管理50 / 開発スコープ499.5

P1 = [
    ("生鮮発注勧告(当日分)", "Sinops-S24", 1.5, 2.0, 1.0, 1.5, 6.0,
     "非生鮮24と項目・ファイルが違うため新規。連携先=生鮮発注統合前提"),
    ("生鮮発注勧告(翌日以降分)", "Sinops-S25", 1.0, 1.5, 1.0, 1.0, 4.5,
     "当日とレイアウト類似。未来スロット差分"),
    ("生鮮入荷実績", "Sinops-S11", 2.0, 2.5, 1.5, 2.0, 8.0,
     "双源：PC/CK=生鮮基幹、取引先=MD基幹"),
    ("生鮮入荷予定", "Sinops-S12", 1.5, 2.0, 1.0, 1.5, 6.0,
     "生鮮基幹新規ソース。L1別作"),
]
P1_TOTAL = sum(r[6] for r in P1)

L1 = [
    ("生鮮発注勧告(当日分)", "Sinops-S24", "TBD(≠R6)", 6.0, 6.0, 5.0, 17.0,
     "基準24の14に+3：着地が生鮮発注統合"),
    ("生鮮発注勧告(翌日以降分)", "Sinops-S25", "TBD(≠R6)", 5.0, 5.0, 4.0, 14.0,
     "基準25の14。見積行は2本"),
    ("生鮮入荷実績", "Sinops-S11", "TBD(~6+合流)", 6.0, 6.0, 5.0, 17.0,
     "基準11の10に+7：双源合流"),
    ("生鮮入荷予定", "Sinops-S12", "TBD(~8)", 5.0, 5.0, 4.0, 14.0,
     "基準12の10に+4：生鮮基幹新規"),
]
L1_TOTAL = sum(r[6] for r in L1)

L3 = [
    ("生鮮発注勧告(当日分)", "Sinops-S24", "TBD", 1.5, 1.5, 1.5, 4.5,
     "基準24の3に+1.5"),
    ("生鮮発注勧告(翌日以降分)", "Sinops-S25", "TBD", 1.0, 1.0, 1.0, 3.0,
     "基準25の3を維持"),
    ("生鮮入荷実績", "Sinops-S11", "TBD", 1.0, 1.5, 1.5, 4.0,
     "基準11の2.5に+1.5：合流後取込"),
    ("生鮮入荷予定", "Sinops-S12", "TBD", 0.5, 1.0, 1.0, 2.5,
     "基準12の2.5を維持"),
]
L3_TOTAL = sum(r[6] for r in L3)
TRAN_DETAIL = P1_TOTAL + L1_TOTAL + L3_TOTAL  # 100.5（IF別明細合計＝標準案A）
TRAN_RISK = 120.0

MASTER_MUST = [
    ("生鮮マスタ棚卸・一覧化", 6.0,
     "生鮮専用マスタの種類・所在・GCS有無・Sinops参照関係を一覧化"),
    ("変換ルール単位の流用判定（1本ずつ）", 10.0,
     "既存01–05/13–15等の③変換と突合。流用／差分／新規の三択を付ける"),
    ("既存Sinops横断影響調査", 10.0,
     "マスタ前提が変わる場合の01–25設計・マッピング・単体観点への波及洗い出し"),
    ("生鮮側サーバー→GCS 上げ接続確認", 5.0,
     "会議: GCSに既にあれば再開発不要、無い見込み。見積に上げ作業を含める"),
    ("判定結果のゲート資料・再見積更新", 4.0,
     "判定表→増分IF採番→本見積のマスタIF枠を確定値に更新"),
]
MASTER_MUST_STD = sum(r[1] for r in MASTER_MUST)  # 35
MASTER_MUST_RISK = 40.0
MASTER_IF_STD = 51.0
MASTER_IF_RISK = 95.0
IMPACT_403_STD = 15.0
IMPACT_403_RISK = 45.0

SCENARIOS = {
    "standard": {
        "label": "標準案",
        "outfile": "【見積】生鮮IF増分_標準案_調査Mappingから詳細設計開発単体_20260921.xlsx",
        "sibling": "リスク込み案",
        "tran": TRAN_DETAIL,
        "b1": MASTER_MUST_STD,
        "b2": MASTER_IF_STD,
        "b3": IMPACT_403_STD,
        "a_note": "IF別明細の合計",
        "b1_note": "必須5作業の合計",
        "b2_note": "仮置き：新規／大幅差分 3本×17人日",
        "b3_note": "軽微パッチ・注記・共通変換の追い込み",
        "positioning": "マスタ流用が概ね成立し、波及が限定的な場合の標準工数",
    },
    "risk": {
        "label": "リスク込み案",
        "outfile": "【見積】生鮮IF増分_リスク込み案_調査Mappingから詳細設計開発単体_20260921.xlsx",
        "sibling": "標準案",
        "tran": TRAN_RISK,
        "b1": MASTER_MUST_RISK,
        "b2": MASTER_IF_RISK,
        "b3": IMPACT_403_RISK,
        "a_note": "IF別明細に上振れを加味",
        "b1_note": "必須作業＋判定遅延バッファ",
        "b2_note": "仮置き：5本×17＋予備",
        "b3_note": "複数IF再設計・再単体",
        "positioning": "生鮮専用マスタの新規本数・既存波及が大きい場合の工数",
    },
}

OLD_FILES = [
    "【試算見積】生鮮IF増分_調査Mappingから詳細設計開発単体_20260921.xlsx",
    "【見積】生鮮IF増分_P80採用_調査Mappingから詳細設計開発単体_20260921.xlsx",
    "【見積】生鮮IF増分_P50_調査Mappingから詳細設計開発単体_20260921.xlsx",
    "【見積】生鮮IF増分_P80_調査Mappingから詳細設計開発単体_20260921.xlsx",
]

IF_SUM = [
    ("生鮮発注勧告(当日分)", "Sinops-S24", 6.0, 17.0, 0.0, 4.5, 27.5),
    ("生鮮発注勧告(翌日以降分)", "Sinops-S25", 4.5, 14.0, 0.0, 3.0, 21.5),
    ("生鮮入荷実績", "Sinops-S11", 8.0, 17.0, 0.0, 4.0, 29.0),
    ("生鮮入荷予定", "Sinops-S12", 6.0, 14.0, 0.0, 2.5, 22.5),
]


def style_range(ws, start_row=1):
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header_fill(ws, row, cols):
    for c in range(1, cols + 1):
        ws.cell(row, c).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row, c).font = Font(bold=True)


def title(ws, text, cols=8):
    ws["A1"] = text
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=13)
    ws["A1"].fill = PatternFill("solid", fgColor=BLUE)


def section_label(ws, cell, text):
    ws[cell] = text
    ws[cell].font = Font(bold=True, size=11, color=BLUE)


def fill_row(ws, row, cols, color):
    for c in range(1, cols + 1):
        ws.cell(row, c).fill = PatternFill("solid", fgColor=color)


def build_summary(ws, sc, label, sibling, direct, pj, yen_direct, yen_pj, yen_total, total_md):
    """サマリー：提示値 → 内訳1表 → 単価 → 短い前提。IF明細は他シートへ。"""
    cols = 5
    title(ws, f"西友MD基幹統合　生鮮IF増分 見積（{label}）", cols)

    # --- メタ（ラベル列＋値列、読みやすい2行）---
    meta = [
        (2, "見積区分", f"{label}　※対になる【{sibling}】は別ファイル"),
        (3, "作成日", "2026/09/21"),
        (4, "対象範囲", "生鮮増分のみ（調査・Mapping〜詳細設計・開発・単体＋PJ按分）"),
        (5, "既存403", "契約額は据え置き（本表は増分のみ）"),
    ]
    for r, k, v in meta:
        ws.cell(r, 1, k).font = Font(bold=True)
        ws.cell(r, 2, v)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=cols)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=LIGHT)

    # --- 提示値（最初に見る数字）---
    section_label(ws, "A7", "1. 提示値（このファイルの結論）")
    for i, h in enumerate(["項目", "人日", "金額（円・税抜）", "単価", ""], 1):
        ws.cell(8, i, h)
    header_fill(ws, 8, 4)
    headline = [
        ("直接工数（A〜B3）", direct, yen_direct, f"{UNIT:,} 円/人日"),
        ("PJ管理（按分）", pj, yen_pj, f"{UNIT_PJ:,} 円/人日"),
        (f"合計　【{label}】", total_md, yen_total, "—"),
    ]
    for i, (name, md, yen, unit) in enumerate(headline, 9):
        ws.cell(i, 1, name)
        ws.cell(i, 2, md)
        ws.cell(i, 3, yen)
        ws.cell(i, 4, unit)
        ws.cell(i, 2).number_format = "0.0"
        ws.cell(i, 3).number_format = "#,##0"
        if name.startswith("合計"):
            fill_row(ws, i, 4, GREEN)
            for c in range(1, 5):
                ws.cell(i, c).font = Font(bold=True, size=12)

    ws["A12"] = sc["positioning"]
    ws.merge_cells("A12:E12")
    ws["A12"].font = Font(italic=True, color="666666")

    # --- 内訳（1表だけ）---
    section_label(ws, "A14", "2. 内訳")
    for i, h in enumerate(["#", "ブロック", "人日", "金額（円）", "内容"], 1):
        ws.cell(15, i, h)
    header_fill(ws, 15, 5)
    blocks = [
        ("A", "先行トラン（新規IF 4本）", sc["tran"], int(sc["tran"] * UNIT),
         "勧告当日／翌日、入荷実績、入荷予定"),
        ("B1", "マスタ棚卸・判定（必須）", sc["b1"], int(sc["b1"] * UNIT),
         "一覧・流用判定・既存影響・GCS上げ確認"),
        ("B2", "マスタIF本体（仮置き）", sc["b2"], int(sc["b2"] * UNIT),
         sc["b2_note"]),
        ("B3", "既存403への波及", sc["b3"], int(sc["b3"] * UNIT),
         sc["b3_note"]),
        ("", "直接小計", direct, yen_direct, "A＋B1＋B2＋B3"),
        ("C", "PJ管理", pj, yen_pj, "進捗・品質・横断課題・MT（既存比率で按分）"),
        ("", f"合計（{label}）", total_md, yen_total, "＝ 提示値"),
    ]
    for i, (num, name, md, yen, note) in enumerate(blocks, 16):
        ws.cell(i, 1, num)
        ws.cell(i, 2, name)
        ws.cell(i, 3, md)
        ws.cell(i, 4, yen)
        ws.cell(i, 5, note)
        ws.cell(i, 3).number_format = "0.0"
        ws.cell(i, 4).number_format = "#,##0"
        if name.startswith("合計") or name.startswith("直接"):
            fill_row(ws, i, 5, GREEN)
            for c in range(1, 6):
                ws.cell(i, c).font = Font(bold=True)
        elif num in ("B1", "B2", "B3", "C"):
            fill_row(ws, i, 5, ORANGE)

    # --- 単価（小さく）---
    section_label(ws, "A24", "3. 単価")
    for i, h in enumerate(["対象", "単価（円/人日）", "備考"], 1):
        ws.cell(25, i, h)
    header_fill(ws, 25, 3)
    ws["A26"] = "調査〜単体（直接工数）"
    ws["B26"] = UNIT
    ws["C26"] = "既存見積（2026-09-03）と同じ"
    ws["A27"] = "PJ管理"
    ws["B27"] = UNIT_PJ
    ws["C27"] = f"直接工数 × (50÷499.5) ≒ {pj}人日"
    ws["B26"].number_format = "#,##0"
    ws["B27"].number_format = "#,##0"

    # --- 前提（短文・箇条書き）---
    section_label(ws, "A29", "4. 前提・注意（短文）")
    notes = [
        f"・このファイルは【{label}】。詳細な対案は【{sibling}】ファイルを参照。",
        "・生鮮専用マスタの本数は未確定。B1判定後にB2を確定値へ置換する。",
        "・既存403・共通基盤43・既存PJ管理50の契約行は変更しない。",
        "・結合以降・生鮮基幹本体・共通基盤の二重構築は含まない。",
        "・AのIF別明細は「A_先行トラン明細」シート。Bの内訳は「B_マスタ横断リスク」シート。",
    ]
    for i, t in enumerate(notes, 30):
        ws.cell(i, 1, t)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=cols)

    # 列幅
    widths = [6, 28, 12, 16, 48]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 22
    for r in range(9, 12):
        ws.row_dimensions[r].height = 20

    # 罫線は表だけに（全文囲まない）
    for r in list(range(8, 12)) + list(range(15, 23)) + list(range(25, 28)):
        for c in range(1, 6 if r >= 15 and r <= 22 else (4 if r <= 11 else 4)):
            if r >= 25:
                maxc = 3
            elif r >= 15:
                maxc = 5
            else:
                maxc = 4
            if c <= maxc:
                ws.cell(r, c).border = Border(
                    left=THIN, right=THIN, top=THIN, bottom=THIN
                )
                ws.cell(r, c).alignment = Alignment(vertical="center", wrap_text=True)


def build_one(scenario_key: str) -> Path:
    sc = SCENARIOS[scenario_key]
    label = sc["label"]
    sibling = sc["sibling"]
    direct = sc["tran"] + sc["b1"] + sc["b2"] + sc["b3"]
    pj = round(direct * PJ_RATIO)
    yen_direct = int(direct * UNIT)
    yen_pj = int(pj * UNIT_PJ)
    yen_total = yen_direct + yen_pj
    total_md = direct + pj

    wb = Workbook()
    ws = wb.active
    ws.title = "サマリー"
    build_summary(ws, sc, label, sibling, direct, pj, yen_direct, yen_pj, yen_total, total_md)

    # A明細
    ws2 = wb.create_sheet("A_先行トラン明細")
    title(ws2, f"【A】先行トラン4本 — Phase1＋Phase2（IF別明細／本ファイルA={sc['tran']}）", 9)
    ws2["A2"] = f"単位：人日。本見積ファイルのAブロック値は {sc['tran']}（{label}）。"
    r = 4
    ws2.cell(r, 1, "■ Phase1 調査用表・Mapping")
    r = 5
    for i, h in enumerate(["IF名称", "仮IF No.", "調査用表", "①", "②", "③", "小計", "備考"], 1):
        ws2.cell(r, i, h)
    header_fill(ws2, r, 8)
    r = 6
    for row in P1:
        for c, v in enumerate([row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]], 1):
            ws2.cell(r, c, v)
        r += 1
    ws2.cell(r, 1, "小計 Phase1")
    ws2.cell(r, 7, P1_TOTAL)

    r += 2
    ws2.cell(r, 1, "■ Layer①")
    r += 1
    for i, h in enumerate(["IF名称", "仮IF No.", "項目数", "詳細設計", "開発", "単体", "小計", "備考"], 1):
        ws2.cell(r, i, h)
    header_fill(ws2, r, 8)
    r += 1
    for row in L1:
        for c, v in enumerate(row, 1):
            ws2.cell(r, c, v)
        r += 1
    ws2.cell(r, 1, "小計 Layer①")
    ws2.cell(r, 7, L1_TOTAL)

    r += 2
    ws2.cell(r, 1, "■ Layer② = 0（全件）")
    r += 2
    ws2.cell(r, 1, "■ Layer③")
    r += 1
    for i, h in enumerate(["IF名称", "仮IF No.", "項目数", "詳細設計", "開発", "単体", "小計", "備考"], 1):
        ws2.cell(r, i, h)
    header_fill(ws2, r, 8)
    r += 1
    for row in L3:
        for c, v in enumerate(row, 1):
            ws2.cell(r, c, v)
        r += 1
    ws2.cell(r, 1, "小計 Layer③")
    ws2.cell(r, 7, L3_TOTAL)
    r += 2
    ws2.cell(r, 1, f"Aブロック（{label}）")
    ws2.cell(r, 7, sc["tran"])
    ws2.cell(r, 8, f"IF別明細合計={TRAN_DETAIL}　／　本ファイル={sc['tran']}")
    for c in range(1, 9):
        ws2.cell(r, c).fill = PatternFill("solid", fgColor=GREEN)
        ws2.cell(r, c).font = Font(bold=True)
    for col, w in enumerate([30, 14, 12, 12, 10, 10, 10, 44], 1):
        ws2.column_dimensions[get_column_letter(col)].width = w
    style_range(ws2, 4)

    # B
    ws3 = wb.create_sheet("B_マスタ横断リスク")
    title(ws3, f"【B】生鮮専用マスタ横断リスク（{label}）", 6)
    ws3["A2"] = "生鮮専用マスタは先行トランに閉じない。非生鮮中心のSinops全体に効く。"
    ws3.merge_cells("A2:F2")
    ws3["A2"].fill = PatternFill("solid", fgColor=ORANGE)

    ws3["A4"] = f"■ B1 必須作業／本ファイル={sc['b1']}"
    for i, h in enumerate(["作業", "人日（明細）", "成果物", "備考"], 1):
        ws3.cell(5, i, h)
    header_fill(ws3, 5, 4)
    deliverables = [
        "マスタ一覧・所在マップ",
        "IF別 流用/差分/新規 判定表",
        "既存01–25影響リスト・改修要否",
        "GCS接続・上げ手順の確認結果",
        "増分IF採番案・見積更新版",
    ]
    for i, ((name, days, note), d) in enumerate(zip(MASTER_MUST, deliverables), 6):
        ws3.cell(i, 1, name)
        ws3.cell(i, 2, days)
        ws3.cell(i, 3, d)
        ws3.cell(i, 4, note)
    ws3.cell(11, 1, f"B1 本ファイル（{label}）")
    ws3.cell(11, 2, sc["b1"])
    ws3.cell(11, 1).font = Font(bold=True)

    ws3["A13"] = "■ B2 マスタIF本体（仮置き）"
    for i, h in enumerate(["想定", "人日", "算定", "置換条件"], 1):
        ws3.cell(14, i, h)
    header_fill(ws3, 14, 4)
    ws3["A15"] = "新規または大幅差分のマスタIF"
    ws3["B15"] = sc["b2"]
    ws3["C15"] = sc["b2_note"]
    ws3["D15"] = "B1判定後に行を分解して置換"
    ws3["A16"] = "候補例（未確定）"
    ws3["C16"] = "TM商品、店舗商品差分、カテゴリ/仕入先、停止フラグ、発注スケジュール、新商品在庫15 等"
    ws3.merge_cells("C16:D16")

    ws3["A18"] = "■ B3 既存403への波及"
    for i, h in enumerate(["内容", "人日", "説明"], 1):
        ws3.cell(19, i, h)
    header_fill(ws3, 19, 3)
    ws3["A20"] = "既存Sinops-01〜25の手戻り"
    ws3["B20"] = sc["b3"]
    ws3["C20"] = sc["b3_note"] + "。403契約額は据え置き"

    ws3["A22"] = "■ 影響が及び得る既存IF（例）"
    for i, h in enumerate(["既存IF", "なぜ効くか", "リスク"], 1):
        ws3.cell(23, i, h)
    header_fill(ws3, 23, 3)
    impacts = [
        ("01/02 商品・店別", "TM商品・停止・発注期間が生鮮専用の可能性", "変換流用不可→別IF or 分岐"),
        ("03/04 カテゴリ・仕入先", "生鮮階層・仕入先体系が別", "コード体系の再マッピング"),
        ("05 ケースバラ", "生鮮は計量・パック差", "パターン追加"),
        ("13 発注スケジュール", "生鮮締め・便・曜日が別源", "L1ソース分岐"),
        ("14/15 棚割・新商品在庫", "生鮮は店舗調達を使わない可能性", "15個別対応・14ダミー可否"),
        ("06–10 実績・在庫", "営業在庫の生鮮適用未検証", "計算前提の再確認"),
        ("11/12/24/25", "先行トラン自体が新ソース", "本見積Aで別計上済"),
    ]
    for i, row in enumerate(impacts, 24):
        for c, v in enumerate(row, 1):
            ws3.cell(i, c, v)

    b_sum = sc["b1"] + sc["b2"] + sc["b3"]
    ws3["A32"] = f"B合計（{label}）"
    ws3["B32"] = b_sum
    ws3["C32"] = f"B1 {sc['b1']} + B2 {sc['b2']} + B3 {sc['b3']}"
    for c in range(1, 4):
        ws3.cell(32, c).fill = PatternFill("solid", fgColor=ORANGE)
        ws3.cell(32, c).font = Font(bold=True)
    for col, w in enumerate([36, 14, 48, 40], 1):
        ws3.column_dimensions[get_column_letter(col)].width = w
    style_range(ws3, 4)

    # C PJ
    ws_pj = wb.create_sheet("C_PJ管理")
    title(ws_pj, f"【C】PJ管理（増分按分・{label}）", 5)
    ws_pj["A2"] = (
        "既存PJ管理50は開発スコープ499.5横断分。"
        "本増分は同じ比率50/499.5を本ファイル直接工数に適用。"
    )
    ws_pj.merge_cells("A2:E2")
    for i, h in enumerate(["項目", "算定", "人日", "単価", "金額(円)"], 1):
        ws_pj.cell(4, i, h)
    header_fill(ws_pj, 4, 5)
    ws_pj["A5"] = f"PJ管理（{label}）"
    ws_pj["B5"] = f"{direct} × (50/499.5) ≈ {pj}"
    ws_pj["C5"] = pj
    ws_pj["D5"] = UNIT_PJ
    ws_pj["E5"] = yen_pj
    ws_pj["A7"] = "含む作業"
    ws_pj["B7"] = "進捗・品質管理、横断課題対応、会議体（MT）対応など"
    ws_pj.merge_cells("B7:E7")
    ws_pj["A8"] = "既存契約"
    ws_pj["B8"] = "既存PJ管理50の契約行は変更しない。本表は生鮮増分の追加PJ。"
    ws_pj.merge_cells("B8:E8")
    for col, w in enumerate([22, 40, 10, 12, 14], 1):
        ws_pj.column_dimensions[get_column_letter(col)].width = w
    style_range(ws_pj, 4)

    # 対既存
    ws4 = wb.create_sheet("対既存差分")
    title(ws4, "既存IF流用可否と新規判定", 6)
    ws4["A2"] = "トランは新規IF必須。マスタは判定待ち（B1）。"
    for i, h in enumerate(["対象", "既存対照", "レイアウト", "ソース", "着地", "判定"], 1):
        ws4.cell(4, i, h)
    header_fill(ws4, 4, 6)
    diffs = [
        ("生鮮勧告 当日/翌日", "24/25", "項目・ファイルが違う", "Sinops復路", "生鮮発注統合", "新規IF必須"),
        ("生鮮入荷実績", "11", "近い可能性（要精查）", "生鮮基幹+MD基幹", "Sinops", "新規IF必須（双源）"),
        ("生鮮入荷予定", "12", "近い可能性（要精查）", "生鮮基幹", "Sinops", "新規IF必須"),
        ("生鮮専用マスタ群", "01–05/13–15等", "不明（大量）", "生鮮側サーバー", "Sinops", "棚卸後に流用/新規判定"),
    ]
    for i, row in enumerate(diffs, 5):
        for c, v in enumerate(row, 1):
            ws4.cell(i, c, v)
        if i == 8:
            for c in range(1, 7):
                ws4.cell(i, c).fill = PatternFill("solid", fgColor=ORANGE)
    for col, w in enumerate([22, 16, 28, 28, 18, 28], 1):
        ws4.column_dimensions[get_column_letter(col)].width = w
    style_range(ws4, 4)

    # 前提
    ws5 = wb.create_sheet("前提条件・注記")
    title(ws5, f"前提条件・注記（生鮮IF増分 {label} Rev.e）", 5)
    notes = [
        ("■ 見積範囲", ""),
        ("1", f"本ファイルは【{label}】のみ。対になる【{sibling}】は別Excel。"),
        ("2", "増分見積。既存403・共通基盤43・BO・既存PJ50の契約行は変更しない。"),
        ("3", "工程：Phase1調査・Mapping＋Phase2詳細設計・開発・単体＋増分PJ按分。結合以降は含まない。"),
        ("4", f"単価：直接 {UNIT:,}円／PJ {UNIT_PJ:,}円（税抜）。"),
        ("5", sc["positioning"] + "。"),
        ("", ""),
        ("■ 生鮮専用マスタ（最重要リスク）", ""),
        ("1", "生鮮は自己専用マスタを多数持つ。非生鮮変換で足りるか未確定。"),
        ("2", "影響はSinops全体。B1をAと並行必須。"),
        ("3", "マスタ源は生鮮側サーバー→GCS→IFクラウド。GCS無ければ上げ作業込み。"),
        ("4", "B2は仮置き。B1判定後に確定値へ置換。"),
        ("", ""),
        ("■ 本ファイル数値", ""),
        ("1", f"直接={direct}人日（{yen_direct:,}円）＋PJ={pj}人日（{yen_pj:,}円）＝合計{yen_total:,}円。"),
        ("2", f"内訳 A={sc['tran']}＋B1={sc['b1']}＋B2={sc['b2']}＋B3={sc['b3']}。"),
        ("3", "PJ按分＝直接×(50/499.5)。"),
        ("", ""),
        ("■ 除外", ""),
        ("1", "生鮮基幹／生鮮発注統合本体の改修（先方）。"),
        ("2", "@rms／市場EOS等 自動補充対象外。"),
        ("3", "共通基盤の二重構築。"),
        ("4", "結合・総合・移行・運用。"),
    ]
    r = 3
    for a, b in notes:
        ws5.cell(r, 1, a)
        ws5.cell(r, 2, b)
        if a.startswith("■"):
            ws5.cell(r, 1).font = Font(bold=True)
            ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
            if "マスタ" in a:
                ws5.cell(r, 1).fill = PatternFill("solid", fgColor=ORANGE)
        else:
            ws5.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        r += 1
    ws5.column_dimensions["A"].width = 8
    ws5.column_dimensions["B"].width = 96
    style_range(ws5, 3)

    # 算定根拠
    ws6 = wb.create_sheet("算定根拠")
    title(ws6, f"算定根拠（{label}）", 4)
    ws6["A2"] = f"本ファイル提示値は{label}。対になる{sibling}は別ファイル。"
    for i, h in enumerate(["項目", "基準/仮定", "人日", "根拠"], 1):
        ws6.cell(4, i, h)
    header_fill(ws6, 4, 4)
    rows = [
        ("A ブロック", sc["a_note"], sc["tran"], "勧告2＋入荷実績＋入荷予定"),
        ("B1 必須", sc["b1_note"], sc["b1"], "棚卸5作業ベース"),
        ("B2 マスタIF", sc["b2_note"], sc["b2"], "01/02級を中央値"),
        ("B3 403波及", sc["b3_note"], sc["b3"], "全Sinops影響"),
        ("直接小計", "", direct, ""),
        ("C PJ管理", "50/499.5 按分", pj, f"単価{UNIT_PJ:,}円"),
        (f"合計（{label}）", "", total_md, f"{yen_total:,}円"),
    ]
    for i, row in enumerate(rows, 5):
        for c, v in enumerate(row, 1):
            ws6.cell(i, c, v)
        if row[0].startswith("合計") or row[0].startswith("直接"):
            for c in range(1, 5):
                ws6.cell(i, c).fill = PatternFill("solid", fgColor=GREEN)
                ws6.cell(i, c).font = Font(bold=True)
    ws6["A13"] = "出典: 20260918生鮮IF会議、センター認識合わせ、既存見積Sinops明細・サマリー"
    ws6.merge_cells("A13:D13")
    for col, w in enumerate([18, 36, 12, 36], 1):
        ws6.column_dimensions[get_column_letter(col)].width = w
    style_range(ws6, 4)

    path = OUT / sc["outfile"]
    wb.save(path)
    print(f"saved {path.name}: direct={direct} pj={pj} total={total_md} yen={yen_total}")
    return path


def build():
    paths = [build_one("standard"), build_one("risk")]
    for name in OLD_FILES:
        old = OUT / name
        if old.exists():
            old.unlink()
            print("removed", name)
    return paths


if __name__ == "__main__":
    build()
