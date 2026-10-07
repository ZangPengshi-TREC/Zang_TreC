#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate standalone 生鮮 IF 増分見積 — 標準案 / リスク込み案 を別ファイルで出力。

既存 403 workbook は変更しない。
Rev.2026-09-21e: 分位用語を使わず「標準案」「リスク込み案」で分冊。各ファイルに PJ管理按分を含む。
Rev.2026-09-23: 書式整理（表頭・罫線・数値揃え・列幅・シート体裁）。
Rev.2026-09-30k: マスタ流用は先方協議の仮置き。当社未同席・実データ未確認を前提条件に明示。
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent
UNIT_P1 = 40000   # Phase1（調査・Mapping）
UNIT_P2 = 27500   # Phase2 直接（詳細設計・開発・単体）※PJ管理以外
UNIT_PJ = 40000   # PJ管理
UNIT = UNIT_P2    # 互換用エイリアス（Phase2直接）
PJ_RATIO = 50.0 / 499.5  # 既存: PJ管理50 / 開発スコープ499.5

LAB_A = "A（新規／差分IF）"
LAB_B1 = "B1（流用確認・残件）"
LAB_B2 = "B2（マスタ本体IF）"
LAB_B3 = "B3（非生鮮Phase2波及）"
LAB_C = "C（PJ管理按分）"
SHEET_A = "A_新規差分IF"
SHEET_B = "B_横断"
SHEET_C = "C_PJ管理"

# 正式IF番号は未定。対照用の仮番号。説明はIF名称で読む。
S11 = "S11（仮採番）"  # 生鮮入荷実績
S12 = "S12（仮採番）"  # 生鮮入荷予定
S13 = "S13（仮採番）"  # 生鮮発注スケジュール
S24 = "S24（仮採番）"  # 生鮮発注勧告(当日分)
S25 = "S25（仮採番）"  # 生鮮発注勧告(翌日以降分)
S26 = "S26（仮採番）"  # 生鮮勧告→発注統合変換

# --- palette ---
NAVY = "1F4E78"
NAVY_SOFT = "2E75B6"
LIGHT = "D6E3F0"
GREEN = "C6EFCE"
GREEN_TXT = "006100"
ORANGE = "FCE4D6"
ORANGE_TXT = "C65911"
ZEBRA = "F5F8FB"
GRAY_TXT = "595959"
THIN = Side(style="thin", color="B0B0B0")
MED = Side(style="medium", color=NAVY)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FONT = "Yu Gothic"
FONT_BODY = Font(name=FONT, size=10, color="2F2F2F")
FONT_BOLD = Font(name=FONT, size=10, bold=True, color="2F2F2F")
FONT_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
FONT_TITLE = Font(name=FONT, size=14, bold=True, color="FFFFFF")
FONT_SECTION = Font(name=FONT, size=11, bold=True, color=NAVY)
FONT_NOTE = Font(name=FONT, size=9, color=GRAY_TXT)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)

# Mapping 難易度: 高5 / 中4 / 低3。小計を調査用表・①②③へ配分。
# 高5＝1.0+2.0+1.0+1.0、中4＝1.0+1.5+0.5+1.0、0.5本＝高5の半分。
P1_STD = [
    ("生鮮発注勧告(当日分)", S24, 1.0, 2.0, 1.0, 1.0, 5.0,
     "難易度高5。Sinops勧告ファイル。新規IF"),
    ("生鮮発注勧告(翌日以降分)", S25, 0.5, 1.0, 0.5, 0.5, 2.5,
     "当日と類似。2ファイルを1.5本換算のため0.5本（高5の半分）"),
    ("生鮮勧告→発注統合変換", S26, 1.0, 2.0, 1.0, 1.0, 5.0,
     "難易度高5。現行の生鮮→発注統合IFは流用不可。非生鮮発注勧告を参照して新規"),
    ("生鮮入荷実績", S11, 1.0, 1.5, 0.5, 1.0, 4.0,
     "難易度中4。ベンダーは流用。伝票区分差分＋PC/CKは振替"),
    ("生鮮入荷予定", S12, 1.0, 2.0, 1.0, 1.0, 5.0,
     "難易度高5。ベンダーは流用。PC/CKは生鮮基幹。nyuka_yotei。SQLレイアウトはLayer①のDBGET前提"),
    ("生鮮発注スケジュール", S13, 1.0, 2.0, 1.0, 1.0, 5.0,
     "難易度高5。ベンダーは流用。PC/CKは生鮮基幹。sii.txt。SQLレイアウトはLayer①のDBGET前提"),
]
L1_STD = [
    ("生鮮発注勧告(当日分)", S24, "未確定（非生鮮発注勧告は使わない）", 6.0, 6.0, 5.0, 17.0,
     "非生鮮発注勧告 Layer①14人日に+3：生鮮勧告レイアウト差分"),
    ("生鮮発注勧告(翌日以降分)", S25, "未確定（非生鮮発注勧告は使わない）", 3.0, 3.0, 2.5, 8.5,
     "当日レイアウト流用。1.5本換算の0.5本（発注勧告当日のLayer①の半分）"),
    ("生鮮勧告→発注統合変換", S26, "未確定（非生鮮発注IF参照）", 5.0, 5.0, 4.0, 14.0,
     "非生鮮発注勧告 Layer①相当。当日・翌日を1ジョブで発注統合向けに変換。現行生鮮IFは使わない"),
    ("生鮮入荷実績", S11, "未確定（差分）", 4.0, 5.0, 3.0, 12.0,
     "非生鮮入荷実績 Layer①10人日に+2：伝票区分差分・PC/CK振替。2系統合流は置かない"),
    ("生鮮入荷予定", S12, "未確定（SQL取得）", 5.0, 6.5, 3.5, 15.0,
     "ベンダー流用。PC/CKは別SQLサーバ。DBGET2系統＋作成時刻の待ち合わせ"),
    ("生鮮発注スケジュール", S13, "未確定（SQL取得）", 5.0, 6.5, 3.5, 15.0,
     "ベンダー流用。PC/CKは別SQLサーバ。DBGET2系統＋作成時刻の待ち合わせ"),
]
L3_STD = [
    ("生鮮発注勧告(当日分)", S24, "未確定", 1.5, 1.5, 1.5, 4.5,
     "非生鮮発注勧告 Layer③3人日に+1.5"),
    ("生鮮発注勧告(翌日以降分)", S25, "未確定", 0.5, 1.0, 1.0, 2.5,
     "1.5本換算の0.5本（発注勧告当日のLayer③の半分相当）"),
    ("生鮮勧告→発注統合変換", S26, "未確定", 1.0, 1.0, 1.0, 3.0,
     "発注統合側の取込ジョブ。本体改修は含まない"),
    ("生鮮入荷実績", S11, "未確定", 0.5, 1.0, 1.0, 2.5,
     "非生鮮入荷実績 Layer③2.5人日。合流加算なし"),
    ("生鮮入荷予定", S12, "未確定", 0.5, 1.0, 1.0, 2.5,
     "非生鮮入荷予定 Layer③2.5人日を維持"),
    ("生鮮発注スケジュール", S13, "未確定", 0.5, 1.0, 1.0, 2.5,
     "非生鮮発注スケジュールのLayer③相当"),
]
P1_RISK = [
    ("生鮮発注勧告(当日分)", S24, 1.5, 2.0, 1.0, 1.5, 6.0,
     "非生鮮発注勧告と項目・ファイルが違うため新規。連携先=生鮮発注統合"),
    ("生鮮発注勧告(翌日以降分)", S25, 0.5, 1.0, 0.5, 1.0, 3.0,
     "当日と類似。1.5本換算の0.5本"),
    ("生鮮勧告→発注統合変換", S26, 1.5, 2.0, 1.0, 1.5, 6.0,
     "発注統合レイアウトが非生鮮発注IFから大きく外れる場合"),
    ("生鮮入荷実績", S11, 2.0, 2.5, 1.5, 2.0, 8.0,
     "判定覆り：2系統ソースに戻る場合"),
    ("生鮮入荷予定", S12, 1.5, 2.0, 1.0, 1.5, 6.0,
     "判定覆り：生鮮基幹を全量新規ソースとする場合。SQL2系統も前提"),
    ("生鮮発注スケジュール", S13, 1.5, 2.0, 1.0, 1.5, 6.0,
     "PC/CKのSQL接続・作成時刻が想定より重い場合"),
]
L1_RISK = [
    ("生鮮発注勧告(当日分)", S24, "未確定（非生鮮発注勧告は使わない）", 6.0, 6.0, 5.0, 17.0,
     "非生鮮発注勧告 Layer①14人日に+3：生鮮勧告レイアウト差分"),
    ("生鮮発注勧告(翌日以降分)", S25, "未確定（非生鮮発注勧告は使わない）", 3.0, 3.0, 2.5, 8.5,
     "1.5本換算の0.5本"),
    ("生鮮勧告→発注統合変換", S26, "未確定（発注IF差が大きい場合）", 6.0, 6.0, 5.0, 17.0,
     "項目差分が大きく、非生鮮発注勧告相当では足りない場合"),
    ("生鮮入荷実績", S11, "未確定（2系統合流）", 6.0, 6.0, 5.0, 17.0,
     "2系統合流を戻す"),
    ("生鮮入荷予定", S12, "未確定（SQL取得）", 6.0, 8.0, 4.0, 18.0,
     "DBGET2系統＋作成時刻の揺れが大きい場合の再設計"),
    ("生鮮発注スケジュール", S13, "未確定（SQL取得）", 6.0, 8.0, 4.0, 18.0,
     "DBGET2系統＋作成時刻の揺れが大きい場合の再設計"),
]
L3_RISK = [
    ("生鮮発注勧告(当日分)", S24, "未確定", 1.5, 1.5, 1.5, 4.5, "非生鮮発注勧告 Layer③3人日に+1.5"),
    ("生鮮発注勧告(翌日以降分)", S25, "未確定", 0.5, 1.0, 1.0, 2.5, "1.5本換算の0.5本"),
    ("生鮮勧告→発注統合変換", S26, "未確定", 1.5, 1.5, 1.5, 4.5, "取込が想定より重い場合"),
    ("生鮮入荷実績", S11, "未確定", 1.0, 1.5, 1.5, 4.0, "合流後取込"),
    ("生鮮入荷予定", S12, "未確定", 0.5, 1.0, 1.0, 2.5, "非生鮮入荷予定 Layer③相当"),
    ("生鮮発注スケジュール", S13, "未確定", 0.5, 1.0, 1.0, 2.5, "非生鮮発注スケジュール相当"),
]

B1_STD = [
    ("店舗系12本の流用突合（先方協議・当社未確認）", 8.0, "既存③との突合メモ",
     "商品・店別・カテゴリ・仕入先・ケースバラ・POS系・廃棄・在庫修正・棚割14・新商品在庫15。先方協議では流用可。当社は未同席のため実データで確認する"),
    ("倉庫系対象外の整理（DC在庫なし）", 2.0, "対象外リスト",
     "倉庫マスタ・休日・発注曜日・発注実績・受払・倉庫勧告は生鮮対象外。"
     "倉庫商品マスタ（mst.txt）にTRIAL青果水煮相当は無し（確認済）"),
    ("GCSアップロード要否の最終確認", 3.0, "GCS確認結果",
     "会議見込：既存GCSに無い。確認と必要時のアップロード"),
]
B1_RISK = [
    ("流用判定の再確認", 10.0, "再判定表", "12本の一部が差分／新規に戻る場合"),
    ("既存Sinops横断影響調査", 8.0, "影響リスト", "01–25の再設計・再単体"),
    ("GCSアップロード・接続", 5.0, "接続確認", "想定より重い場合"),
]


def _sum6(rows):
    return sum(r[6] for r in rows)


def _sum_b1(rows):
    return sum(r[1] for r in rows)

A_STD = _sum6(P1_STD) + _sum6(L1_STD) + _sum6(L3_STD)
A_RISK = _sum6(P1_RISK) + _sum6(L1_RISK) + _sum6(L3_RISK)

SCENARIOS = {
    "standard": {
        "label": "標準案",
        "outfile": "【見積】生鮮IF増分_標準案_調査Mappingから詳細設計開発単体_20260921.xlsx",
        "tran": A_STD,
        "p1": _sum6(P1_STD),
        "p1_rows": P1_STD,
        "l1_rows": L1_STD,
        "l3_rows": L3_STD,
        "b1_rows": B1_STD,
        "b1": _sum_b1(B1_STD),
        "b2": 0.0,
        "b3": 5.0,
        "a_note": "勧告は2ファイルだが1.5本換算＋発注統合変換1本（新規）＋入荷2本＋スケジュール1本（差分）",
        "b1_note": "流用12本突合＋倉庫対象外＋GCS",
        "b2_note": "流用前提。未確認。本体IFは0",
        "b3_note": "伝票区分・振替の軽微注記",
        "positioning": "店舗系12本は先方協議では流用（当社未同席）。勧告→発注統合変換は新規。入荷予定・発注スケジュールは別SQLからDBGET",
        "tab": NAVY,
    },
    "risk": {
        "label": "リスク込み案",
        "outfile": "【見積】生鮮IF増分_リスク込み案_調査Mappingから詳細設計開発単体_20260921.xlsx",
        "tran": A_RISK,
        "p1": _sum6(P1_RISK),
        "p1_rows": P1_RISK,
        "l1_rows": L1_RISK,
        "l3_rows": L3_RISK,
        "b1_rows": B1_RISK,
        "b1": _sum_b1(B1_RISK),
        "b2": 34.0,
        "b3": 20.0,
        "a_note": "判定覆りを加味したIF別明細",
        "b1_note": "再判定＋横断影響＋GCS",
        "b2_note": "流用覆り：新規マスタ2本×17",
        "b3_note": "複数IF再設計・再単体",
        "positioning": "流用判定が一部覆る、または入荷予定・発注スケジュールのSQL／作成時刻が想定より重い場合。"
                       "倉庫商品マスタに水煮相当は無し（確認済）",
        "tab": ORANGE_TXT,
    },
}

OLD_FILES = [
    "【試算見積】生鮮IF増分_調査Mappingから詳細設計開発単体_20260921.xlsx",
    "【見積】生鮮IF増分_P80採用_調査Mappingから詳細設計開発単体_20260921.xlsx",
    "【見積】生鮮IF増分_P50_調査Mappingから詳細設計開発単体_20260921.xlsx",
    "【見積】生鮮IF増分_P80_調査Mappingから詳細設計開発単体_20260921.xlsx",
]


def fill(cell, color):
    cell.fill = PatternFill("solid", fgColor=color)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def paint_title(ws, text, cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    cell = ws.cell(1, 1, text)
    cell.font = FONT_TITLE
    fill(cell, NAVY)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    for c in range(2, cols + 1):
        fill(ws.cell(1, c), NAVY)


def paint_section(ws, row, text, cols, accent=LIGHT):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.font = FONT_SECTION
    fill(cell, accent)
    cell.alignment = LEFT
    ws.row_dimensions[row].height = 22
    for c in range(2, cols + 1):
        fill(ws.cell(row, c), accent)


def paint_note(ws, row, text, cols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.font = FONT_NOTE
    cell.alignment = LEFT_TOP
    ws.row_dimensions[row].height = 18


def paint_header(ws, row, headers):
    ws.row_dimensions[row].height = 32
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row, c, h)
        cell.font = FONT_HEAD
        fill(cell, NAVY)
        cell.alignment = CENTER
        cell.border = BORDER


def style_data_row(ws, row, cols, *, num_cols=(), money_cols=(), zebra=False, total=False, warn=False):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.border = BORDER
        if total:
            fill(cell, GREEN)
            cell.font = Font(name=FONT, size=10, bold=True, color=GREEN_TXT)
        elif warn:
            fill(cell, ORANGE)
            cell.font = FONT_BOLD
        elif zebra:
            fill(cell, ZEBRA)
            cell.font = FONT_BODY
        else:
            cell.font = FONT_BODY

        if c in num_cols:
            cell.alignment = CENTER
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.0"
        elif c in money_cols:
            cell.alignment = CENTER
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
        else:
            cell.alignment = LEFT if c == 1 or c == cols else CENTER
    ws.row_dimensions[row].height = 20 if not total else 22


def apply_sheet_chrome(ws, tab_color=NAVY):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.sheet_properties.tabColor = tab_color
    ws.freeze_panes = "A2"


def money_breakdown(sc):
    """Phase1=4万／Phase2直接=2.75万／PJ=4万 で金額を分解する。"""
    a_p1 = sc["p1"]
    a_p2 = sc["tran"] - a_p1
    if a_p2 < 0:
        raise ValueError(f"A Phase2 negative: tran={sc['tran']} p1={a_p1}")
    phase1 = a_p1 + sc["b1"]
    phase2 = a_p2 + sc["b2"] + sc["b3"]
    direct = phase1 + phase2
    pj = round(direct * PJ_RATIO)
    yen_a = int(a_p1 * UNIT_P1 + a_p2 * UNIT_P2)
    yen_b1 = int(sc["b1"] * UNIT_P1)
    yen_b2 = int(sc["b2"] * UNIT_P2)
    yen_b3 = int(sc["b3"] * UNIT_P2)
    yen_p1 = int(phase1 * UNIT_P1)
    yen_p2 = int(phase2 * UNIT_P2)
    yen_direct = yen_p1 + yen_p2
    yen_pj = int(pj * UNIT_PJ)
    return {
        "a_p1": a_p1,
        "a_p2": a_p2,
        "phase1": phase1,
        "phase2": phase2,
        "direct": direct,
        "pj": pj,
        "total_md": direct + pj,
        "yen_a": yen_a,
        "yen_b1": yen_b1,
        "yen_b2": yen_b2,
        "yen_b3": yen_b3,
        "yen_p1": yen_p1,
        "yen_p2": yen_p2,
        "yen_direct": yen_direct,
        "yen_pj": yen_pj,
        "yen_total": yen_direct + yen_pj,
    }


def build_summary(ws, sc, label, m):
    cols = 5
    paint_title(ws, "西友MD基幹統合　生鮮IF増分 見積", cols)
    apply_sheet_chrome(ws, sc["tab"])

    meta = [
        ("対象範囲", "調査・Mapping〜詳細設計・開発・単体（PJ管理按分を含む）"),
    ]
    for i, (k, v) in enumerate(meta, 2):
        kcell = ws.cell(i, 1, k)
        kcell.font = FONT_BOLD
        fill(kcell, LIGHT)
        kcell.alignment = CENTER
        kcell.border = BORDER
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=cols)
        vcell = ws.cell(i, 2, v)
        vcell.font = FONT_BODY
        vcell.alignment = LEFT
        vcell.border = BORDER
        for c in range(3, cols + 1):
            ws.cell(i, c).border = BORDER
            fill(ws.cell(i, c), "FFFFFF")
        ws.row_dimensions[i].height = 20

    paint_section(ws, 4, "1. 提示値", cols)
    paint_header(ws, 5, ["項目", "人日", "金額（円・税抜）", "単価", "備考"])
    headline = [
        ("Phase1（調査・Mapping）", m["phase1"], m["yen_p1"], f"{UNIT_P1:,}", "", False),
        ("Phase2（詳細設計・開発・単体）", m["phase2"], m["yen_p2"], f"{UNIT_P2:,}", "PJ管理を除く", False),
        ("直接小計", m["direct"], m["yen_direct"], "—", "A〜B3", False),
        ("PJ管理", m["pj"], m["yen_pj"], f"{UNIT_PJ:,}", "直接×50/499.5", False),
        ("合計", m["total_md"], m["yen_total"], "—", "", True),
    ]
    for i, (name, md, yen, unit, note, is_total) in enumerate(headline, 6):
        ws.cell(i, 1, name)
        ws.cell(i, 2, md)
        ws.cell(i, 3, yen)
        ws.cell(i, 4, unit)
        ws.cell(i, 5, note)
        style_data_row(ws, i, 5, num_cols={2}, money_cols={3}, total=is_total)
        if name == "直接小計":
            for c in range(1, 6):
                fill(ws.cell(i, c), LIGHT)
                ws.cell(i, c).font = FONT_BOLD
        if is_total:
            for c in range(1, 6):
                ws.cell(i, c).font = Font(name=FONT, size=12, bold=True, color=GREEN_TXT)

    paint_note(ws, 11, f"{sc['positioning']}　単価は税抜・人日（Phase1 {UNIT_P1:,}／Phase2 {UNIT_P2:,}／PJ {UNIT_PJ:,}）。", cols)

    paint_section(ws, 13, "2. 内訳", cols)
    paint_header(ws, 14, ["区分", "内容", "人日", "金額（円）", "備考"])
    blocks = [
        (LAB_A, "勧告1.5本＋変換＋入荷2本＋スケジュール", sc["tran"], m["yen_a"],
         f"P1 {m['a_p1']}／P2 {m['a_p2']}", False, False),
        (LAB_B1, "突合、倉庫対象外、GCS", sc["b1"], m["yen_b1"],
         "Phase1単価", False, False),
        (LAB_B2, "店舗系マスタは流用前提", sc["b2"], m["yen_b2"],
         sc["b2_note"], False, False),
        (LAB_B3, "伝票区分・振替の注記", sc["b3"], m["yen_b3"],
         sc["b3_note"], False, True),
        ("直接小計", "", m["direct"], m["yen_direct"], "", True, False),
        (LAB_C, "直接×50/499.5", m["pj"], m["yen_pj"],
         "進捗・品質・会議", False, False),
        ("合計", "", m["total_md"], m["yen_total"], "", True, False),
    ]
    for i, (num, name, md, yen, note, is_total, is_warn) in enumerate(blocks, 15):
        ws.cell(i, 1, num)
        ws.cell(i, 2, name)
        ws.cell(i, 3, md)
        ws.cell(i, 4, yen)
        ws.cell(i, 5, note)
        style_data_row(
            ws, i, 5, num_cols={3}, money_cols={4},
            total=is_total, warn=is_warn and not is_total,
            zebra=(i % 2 == 0 and not is_total and not is_warn),
        )
        ws.cell(i, 1).alignment = LEFT
        ws.cell(i, 2).alignment = LEFT
        ws.cell(i, 5).alignment = LEFT

    paint_section(ws, 23, "3. 単価", cols)
    paint_header(ws, 24, ["対象", "単価（円/人日）", "備考", "", ""])
    for c in range(4, 6):
        fill(ws.cell(24, c), NAVY)
        ws.cell(24, c).border = BORDER
    unit_rows = [
        ("Phase1", UNIT_P1, "調査・Mapping、B1"),
        ("Phase2", UNIT_P2, "詳細設計・開発・単体（AのLayer①③、B2、B3）"),
        ("PJ管理", UNIT_PJ, f"直接 {m['direct']} × 50/499.5 ≒ {m['pj']}人日"),
    ]
    for i, (a, b, c) in enumerate(unit_rows, 25):
        ws.cell(i, 1, a)
        ws.cell(i, 2, b)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=5)
        ws.cell(i, 3, c)
        style_data_row(ws, i, 3, money_cols={2})
        for col in range(4, 6):
            ws.cell(i, col).border = BORDER

    paint_section(ws, 29, "4. 前提", cols)
    notes = [
        "店舗系12本の流用は先方協議の結果。当社は未同席で、実データは未確認。提示は仮置きであり、確定請負ではない。",
        "勧告は2ファイルで1.5本。発注統合向け変換は新規（現行の生鮮→統合IFは使わない）。入荷・スケジュールは差分。倉庫は対象外。",
        "Aに新規・差分、B1に突合と残件、B2は0、B3に波及。"
        "倉庫はDC在庫なしで対象外。倉庫商品マスタに水煮相当は無し（確認済）。",
        "覆る場合（流用不可、SQL／作成時刻が想定外など）は再見積。非生鮮Phase2・共通基盤43・既存PJ50は動かさない。",
    ]
    for i, t in enumerate(notes, 30):
        paint_note(ws, i, "・" + t, cols)
        ws.row_dimensions[i].height = 30
        cell = ws.cell(i, 1)
        cell.font = Font(name=FONT, size=9, color="2F2F2F")
        bg = ORANGE if i == 30 else ZEBRA
        fill(cell, bg)
        for c in range(2, cols + 1):
            fill(ws.cell(i, c), bg)

    set_widths(ws, [44, 28, 12, 16, 42])
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[14].height = 22
    for r in range(15, 22):
        ws.row_dimensions[r].height = 22


def write_if_table(ws, start_row, section_title, headers, rows, total_label, total_value, cols=8):
    """Write section banner + header + data + subtotal. Returns next free row."""
    paint_section(ws, start_row, section_title, cols)
    hr = start_row + 1
    paint_header(ws, hr, headers)
    r = hr + 1
    num_cols = {i for i, h in enumerate(headers, 1) if h in (
        "調査用表", "①項目マッピング表", "②GAP分析書", "③変換ルール定義書",
        "詳細設計", "開発", "単体テスト", "小計", "人日", "人日（明細）",
    )}
    for idx, row in enumerate(rows):
        for c, v in enumerate(row, 1):
            ws.cell(r, c, v)
        style_data_row(
            ws, r, cols, num_cols=num_cols,
            zebra=(idx % 2 == 1),
        )
        # 備考列 left
        ws.cell(r, cols).alignment = LEFT
        ws.cell(r, 1).alignment = LEFT
        if cols >= 3 and headers[2] == "項目数":
            ws.cell(r, 3).alignment = LEFT
        r += 1
    # subtotal
    ws.cell(r, 1, total_label)
    for c in range(2, cols):
        ws.cell(r, c, "")
    # put total in 小計 column (usually 7) or col 2 if short
    subtotal_col = 7 if cols >= 7 else 2
    ws.cell(r, subtotal_col, total_value)
    style_data_row(ws, r, cols, num_cols={subtotal_col}, total=True)
    return r + 2


def build_sheet_a(ws, sc, label):
    cols = 8
    paint_title(ws, "A 新規・差分IF明細", cols)
    apply_sheet_chrome(ws, sc["tab"])
    paint_note(
        ws, 2,
        "人日。Phase1は調査用表／①マッピング／②GAP／③変換ルール。S11〜S26は仮採番。"
        " Layer①＝TRIAL↔GCS、Layer②＝GCS→DWH（本件0）、Layer③＝DWH・GCS↔業務。",
        cols,
    )
    ws.row_dimensions[2].height = 32

    r = 4
    r = write_if_table(
        ws, r,
        "■ Phase1（調査・Mapping：調査用表＋①項目マッピング表／②GAP分析書／③変換ルール定義書）",
        ["IF名称", "採番", "調査用表", "①項目マッピング表", "②GAP分析書", "③変換ルール定義書", "小計", "備考"],
        [[x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7]] for x in sc["p1_rows"]],
        "小計 Phase1", sc["p1"], cols,
    )
    r = write_if_table(
        ws, r,
        "■ Layer① TRIAL↔GCS（外部IFプログラム開発）",
        ["IF名称", "採番", "項目数", "詳細設計", "開発", "単体テスト", "小計", "備考"],
        [list(x) for x in sc["l1_rows"]],
        "小計 Layer①", _sum6(sc["l1_rows"]), cols,
    )
    paint_section(ws, r, "■ Layer② GCS→DWH（DataSpider取込ジョブ）＝ 0（全件）", cols, LIGHT)
    r += 2
    r = write_if_table(
        ws, r,
        "■ Layer③ DWH・GCS↔業務システム（DataSpider連携ジョブ）",
        ["IF名称", "採番", "項目数", "詳細設計", "開発", "単体テスト", "小計", "備考"],
        [list(x) for x in sc["l3_rows"]],
        "小計 Layer③", _sum6(sc["l3_rows"]), cols,
    )

    ws.cell(r, 1, "A 計")
    for c in range(2, 7):
        ws.cell(r, c, "")
    ws.cell(r, 7, sc["tran"])
    ws.cell(r, 8, "")
    style_data_row(ws, r, cols, num_cols={7}, total=True)
    ws.cell(r, 8).alignment = LEFT

    set_widths(ws, [26, 16, 28, 14, 12, 14, 8, 42])


def build_sheet_b(ws, sc, label):
    cols = 4
    paint_title(ws, "B 横断（流用確認・マスタIF・波及）", cols)
    apply_sheet_chrome(ws, sc["tab"])
    paint_section(
        ws, 2,
        "店舗系12本は先方協議では流用。当社未同席。本体IFは置かず、B1で突合する。",
        cols, ORANGE,
    )
    ws.cell(2, 1).font = Font(name=FONT, size=10, bold=True, color=ORANGE_TXT)

    paint_section(ws, 4, f"■ B1 流用確認・残件　{sc['b1']}人日", cols)
    paint_header(ws, 5, ["作業", "人日（明細）", "成果物", "備考"])
    b1_rows = sc["b1_rows"]
    for i, (name, days, deliverable, note) in enumerate(b1_rows, 6):
        ws.cell(i, 1, name)
        ws.cell(i, 2, days)
        ws.cell(i, 3, deliverable)
        ws.cell(i, 4, note)
        style_data_row(ws, i, cols, num_cols={2}, zebra=(i % 2 == 0))
        ws.cell(i, 1).alignment = LEFT
        ws.cell(i, 3).alignment = LEFT
        ws.cell(i, 4).alignment = LEFT_TOP
        ws.row_dimensions[i].height = 28
    total_r = 6 + len(b1_rows)
    ws.cell(total_r, 1, "B1 計")
    ws.cell(total_r, 2, sc["b1"])
    ws.cell(total_r, 3, "")
    ws.cell(total_r, 4, "")
    style_data_row(ws, total_r, cols, num_cols={2}, total=True)

    b2_r = total_r + 2
    paint_section(ws, b2_r, "■ B2 マスタ本体IF", cols)
    paint_header(ws, b2_r + 1, ["想定", "人日", "算定", "置換条件"])
    ws.cell(b2_r + 2, 1, "新規マスタIF")
    ws.cell(b2_r + 2, 2, sc["b2"])
    ws.cell(b2_r + 2, 3, sc["b2_note"])
    ws.cell(b2_r + 2, 4, "標準案は0。覆ったら再見積")
    style_data_row(ws, b2_r + 2, cols, num_cols={2}, warn=sc["b2"] > 0)
    ws.cell(b2_r + 2, 1).alignment = LEFT
    ws.cell(b2_r + 2, 3).alignment = LEFT
    ws.cell(b2_r + 2, 4).alignment = LEFT

    b3_r = b2_r + 5
    paint_section(ws, b3_r, "■ B3 非生鮮Phase2波及", cols)
    paint_header(ws, b3_r + 1, ["内容", "人日", "説明", ""])
    fill(ws.cell(b3_r + 1, 4), NAVY)
    ws.cell(b3_r + 1, 4).border = BORDER
    ws.cell(b3_r + 2, 1, "既存Sinopsへの手戻り")
    ws.cell(b3_r + 2, 2, sc["b3"])
    ws.merge_cells(start_row=b3_r + 2, start_column=3, end_row=b3_r + 2, end_column=4)
    ws.cell(b3_r + 2, 3, sc["b3_note"] + "。非生鮮Phase2の契約額は据え置き")
    style_data_row(ws, b3_r + 2, 3, num_cols={2}, warn=True)
    ws.cell(b3_r + 2, 4).border = BORDER
    ws.cell(b3_r + 2, 1).alignment = LEFT
    ws.cell(b3_r + 2, 3).alignment = LEFT

    imp_r = b3_r + 4
    paint_section(ws, imp_r, "■ 判定後の扱い", cols)
    paint_header(ws, imp_r + 1, ["既存IF", "判定", "見積上の扱い", ""])
    fill(ws.cell(imp_r + 1, 4), NAVY)
    ws.cell(imp_r + 1, 4).border = BORDER
    impacts = [
        ("商品・店別・カテゴリ・仕入先・ケースバラ", "先方協議では非生鮮と同じ", "流用（B1で突合）"),
        ("POS・時間帯・来客・廃棄・在庫修正", "先方協議では非生鮮と同じ", "流用（B1で突合）"),
        ("棚割・新商品在庫", "先方協議では非生鮮と同じ", "流用（B1で突合）"),
        ("入荷実績", "ベンダー流用／伝票区分差分／PC/CK振替", "A 差分"),
        ("入荷予定", "ベンダー流用／PC・CKは別SQL・DBGET", "A 差分"),
        ("発注スケジュール", "ベンダー流用／PC・CKは別SQL・DBGET", "A 差分"),
        ("発注勧告（当日／翌日以降）", "2ファイル／1.5本換算", "A 新規"),
        ("勧告→発注統合変換", "現行生鮮→統合IFは使わない", "A 新規"),
        ("倉庫マスタ・休日・実績・勧告", "DC在庫なし", "対象外"),
    ]
    for i, row in enumerate(impacts, imp_r + 2):
        ws.cell(i, 1, row[0])
        ws.cell(i, 2, row[1])
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=4)
        ws.cell(i, 3, row[2])
        style_data_row(ws, i, 3, zebra=(i % 2 == 0))
        ws.cell(i, 4).border = BORDER
        ws.cell(i, 1).alignment = LEFT
        ws.cell(i, 2).alignment = LEFT
        ws.cell(i, 3).alignment = LEFT

    sum_r = imp_r + 2 + len(impacts) + 1
    b_sum = sc["b1"] + sc["b2"] + sc["b3"]
    ws.cell(sum_r, 1, "B 計")
    ws.cell(sum_r, 2, b_sum)
    ws.merge_cells(start_row=sum_r, start_column=3, end_row=sum_r, end_column=4)
    ws.cell(sum_r, 3, f"B1 {sc['b1']} ＋ B2 {sc['b2']} ＋ B3 {sc['b3']}")
    style_data_row(ws, sum_r, 3, num_cols={2}, total=True)
    ws.cell(sum_r, 4).border = BORDER
    ws.cell(sum_r, 3).alignment = LEFT

    set_widths(ws, [42, 14, 36, 42])


def build_sheet_c(ws, sc, label, direct, pj, yen_pj):
    cols = 5
    paint_title(ws, "C PJ管理按分", cols)
    apply_sheet_chrome(ws, sc["tab"])
    paint_note(
        ws, 2,
        "既存PJ管理50は開発スコープ499.5横断分。本増分は同じ比率50/499.5を本ファイル直接工数に適用。",
        cols,
    )
    paint_header(ws, 4, ["項目", "算定", "人日", "単価", "金額(円)"])
    ws.cell(5, 1, "PJ管理")
    ws.cell(5, 2, f"{direct} × (50/499.5) ≈ {pj}")
    ws.cell(5, 3, pj)
    ws.cell(5, 4, UNIT_PJ)
    ws.cell(5, 5, yen_pj)
    style_data_row(ws, 5, cols, num_cols={3}, money_cols={4, 5}, total=True)
    ws.cell(5, 2).alignment = LEFT

    ws.cell(7, 1, "含む作業")
    ws.merge_cells("B7:E7")
    ws.cell(7, 2, "進捗・品質管理、横断課題対応、会議体（MT）対応など")
    style_data_row(ws, 7, 2)
    for c in range(3, 6):
        ws.cell(7, c).border = BORDER
    ws.cell(7, 1).alignment = CENTER
    fill(ws.cell(7, 1), LIGHT)
    ws.cell(7, 2).alignment = LEFT

    ws.cell(8, 1, "既存契約")
    ws.merge_cells("B8:E8")
    ws.cell(8, 2, "既存PJ管理50の契約行は変更しない。本表は生鮮増分の追加PJ管理。")
    style_data_row(ws, 8, 2)
    for c in range(3, 6):
        ws.cell(8, c).border = BORDER
    ws.cell(8, 1).alignment = CENTER
    fill(ws.cell(8, 1), LIGHT)
    ws.cell(8, 2).alignment = LEFT

    set_widths(ws, [18, 42, 10, 12, 14])


def build_sheet_diff(ws, sc):
    cols = 6
    paint_title(ws, "流用可否", cols)
    apply_sheet_chrome(ws, sc["tab"])
    paint_note(ws, 2, "出典：2026-09-29 マスタデータ判定。倉庫系はDC在庫なしで対象外。", cols)
    paint_header(ws, 4, ["対象", "既存対照", "判定メモ", "ソース／連携先", "本見積", "判定"])
    diffs = [
        ("商品・店別・カテゴリ・仕入先・ケースバラ", "非生鮮同名マスタ", "先方協議では非生鮮と同じ", "既存", "B1", "流用（未確認）"),
        ("POS・時間帯・来客・廃棄・在庫修正", "非生鮮同名実績", "先方協議では非生鮮と同じ", "既存", "B1", "流用（未確認）"),
        ("棚割明細・新商品在庫", "非生鮮棚割・新商品在庫", "先方協議では非生鮮と同じ", "既存", "B1", "流用（未確認）"),
        ("入荷実績", "非生鮮入荷実績", "伝票区分差分／PC/CKは振替", "ベンダー＝既存", "A", "差分"),
        ("入荷予定", "非生鮮入荷予定", "ベンダー同じ／PC・CKは別SQL", "2系統DBGET", "A", "差分"),
        ("発注スケジュール", "非生鮮発注スケジュール", "ベンダー同じ／PC・CKは別SQL", "2系統DBGET", "A", "差分"),
        ("発注勧告 当日/翌日", "非生鮮発注勧告", "類似のため1.5本換算", "Sinops復路", "A", "新規"),
        ("勧告→発注統合変換", "非生鮮発注勧告", "現行生鮮→統合IFは使わない", "非生鮮発注勧告を参照", "A", "新規"),
        ("倉庫マスタ・休日・実績・勧告", "倉庫", "DC在庫なし", "—", "—", "対象外"),
    ]
    for i, row in enumerate(diffs, 5):
        for c, v in enumerate(row, 1):
            ws.cell(i, c, v)
        style_data_row(ws, i, cols, warn=("対象外" in row[-1]), zebra=(i % 2 == 0))
        for c in range(1, cols + 1):
            ws.cell(i, c).alignment = LEFT if c in (1, 3, 4, 6) else CENTER
    set_widths(ws, [22, 22, 26, 22, 28, 24])


def build_sheet_notes(ws, sc, label, m):
    paint_title(ws, "見積前提条件", 5)
    apply_sheet_chrome(ws, sc["tab"])
    paint_section(
        ws, 2,
        "店舗系12本の流用は先方協議の結果。当社は未同席。実データ未確認。確定請負ではない。",
        5, ORANGE,
    )
    ws.cell(2, 1).font = Font(name=FONT, size=11, bold=True, color=ORANGE_TXT)

    notes = [
        ("■ 範囲", True, False),
        ("調査・Mapping〜詳細設計・開発・単体。PJ管理按分を含む。結合以降、生鮮基幹／発注統合本体、@rms／EOS、倉庫系は含まない。", False, False),
        (f"単価（税抜・人日）：Phase1 {UNIT_P1:,}円、Phase2 {UNIT_P2:,}円、PJ管理 {UNIT_PJ:,}円。非生鮮Phase2・共通基盤43・既存PJ50は動かさない。", False, False),
        ("", False, False),
        ("■ 流用・新規・差分", True, False),
        ("店舗系12本は先方協議では非生鮮と同じ。標準案はB2＝0とし、B1で突合する。覆ったら再見積。", False, False),
        ("発注勧告は当日・翌日の2ファイル、工数は1.5本。発注統合向け変換は新規（現行の生鮮→統合IFは使わない）。", False, False),
        ("入荷実績はベンダー流用＋伝票区分差分、PC/CKは振替。入荷予定・発注スケジュールはベンダー流用、PC/CKは別SQLからDBGET。", False, False),
        ("倉庫はDC在庫なしで対象外。倉庫商品マスタ（mst.txt）に水煮相当は無し（確認済）。", False, False),
        ("", False, False),
        ("■ 数値の置き方", True, False),
        (f"A {sc['tran']}人日、B1 {sc['b1']}、B2 {sc['b2']}、B3 {sc['b3']}、直接 {m['direct']}、PJ {m['pj']}、合計 {m['total_md']}人日。", False, False),
        ("Mappingは高5／中4／低3。Phase1小計を調査用表・①②③に分けている。S11〜S26は仮採番。", False, False),
        ("流用が覆る、統合レイアウト差が大きい、SQL／作成時刻が想定外、のときは再見積。", False, False),
    ]
    r = 3
    for text, is_sec, is_warn in notes:
        if not text:
            r += 1
            continue
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell = ws.cell(r, 1, text)
        if is_sec:
            cell.font = FONT_SECTION
            bg = ORANGE if is_warn else LIGHT
            fill(cell, bg)
            for c in range(2, 6):
                fill(ws.cell(r, c), bg)
            ws.row_dimensions[r].height = 22
        else:
            cell.font = FONT_BODY
            cell.alignment = LEFT_TOP
            ws.row_dimensions[r].height = 40 if len(text) > 140 else (22 if len(text) > 60 else 18)
        r += 1
    set_widths(ws, [110, 12, 12, 12, 12])
    ws.column_dimensions["A"].width = 118
    ws.sheet_view.zoomScale = 90


def build_sheet_basis(ws, sc, label, m):
    cols = 4
    paint_title(ws, "算定根拠", cols)
    apply_sheet_chrome(ws, sc["tab"])
    paint_note(
        ws, 2,
        f"単価：Phase1 {UNIT_P1:,}／Phase2 {UNIT_P2:,}／PJ {UNIT_PJ:,}（円・人日・税抜）。",
        cols,
    )
    paint_header(ws, 4, ["項目", "内容", "人日", "金額"])
    rows = [
        ("A Phase1", "調査・Mapping", m["a_p1"], f"{int(m['a_p1']*UNIT_P1):,}", False),
        ("A Phase2", "Layer①③", m["a_p2"], f"{int(m['a_p2']*UNIT_P2):,}", False),
        ("A 計", sc["a_note"], sc["tran"], f"{m['yen_a']:,}", False),
        ("B1", sc["b1_note"], sc["b1"], f"{m['yen_b1']:,}", False),
        ("B2", sc["b2_note"], sc["b2"], f"{m['yen_b2']:,}", False),
        ("B3", sc["b3_note"], sc["b3"], f"{m['yen_b3']:,}", False),
        ("直接小計", f"Phase1 {m['phase1']}＋Phase2 {m['phase2']}", m["direct"], f"{m['yen_direct']:,}", True),
        ("C PJ", "50/499.5", m["pj"], f"{m['yen_pj']:,}", False),
        ("合計", "", m["total_md"], f"{m['yen_total']:,}", True),
    ]
    for i, (a, b, c, d, is_total) in enumerate(rows, 5):
        ws.cell(i, 1, a)
        ws.cell(i, 2, b)
        ws.cell(i, 3, c)
        ws.cell(i, 4, d)
        style_data_row(
            ws, i, cols, num_cols={3}, total=is_total,
            warn=(a.startswith("B") and not is_total),
            zebra=(i % 2 == 0 and not is_total),
        )
        ws.cell(i, 1).alignment = LEFT
        ws.cell(i, 2).alignment = LEFT
        ws.cell(i, 4).alignment = LEFT
    paint_note(ws, 15, "出典: 20260929マスタデータの判定、20260918生鮮IF会議、既存見積Sinops明細・サマリー", cols)
    set_widths(ws, [42, 40, 10, 36])


def build_one(scenario_key: str) -> Path:
    sc = SCENARIOS[scenario_key]
    label = sc["label"]
    m = money_breakdown(sc)

    wb = Workbook()
    ws = wb.active
    ws.title = "サマリー"
    build_summary(ws, sc, label, m)

    ws5 = wb.create_sheet("見積前提条件・注記")
    build_sheet_notes(ws5, sc, label, m)

    ws2 = wb.create_sheet(SHEET_A)
    build_sheet_a(ws2, sc, label)

    ws3 = wb.create_sheet(SHEET_B)
    build_sheet_b(ws3, sc, label)

    ws_pj = wb.create_sheet(SHEET_C)
    build_sheet_c(ws_pj, sc, label, m["direct"], m["pj"], m["yen_pj"])

    ws4 = wb.create_sheet("対既存差分")
    build_sheet_diff(ws4, sc)

    ws6 = wb.create_sheet("算定根拠")
    build_sheet_basis(ws6, sc, label, m)

    path = OUT / sc["outfile"]
    wb.save(path)
    print(
        f"saved {path.name}: direct={m['direct']} "
        f"(P1={m['phase1']}@{UNIT_P1}/P2={m['phase2']}@{UNIT_P2}) "
        f"pj={m['pj']} total={m['total_md']} yen={m['yen_total']}"
    )
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
