# -*- coding: utf-8 -*-
"""生鮮IF増分 標準案 WBS（2026-10-12着手、単体は入荷3本が12/18、勧告3本が12/11）。

横断管理台帳 D-2 の列に合わせる。台帳 (3) の手動調整後の版に合わせて生成する。
統合変換（S26）はTRIAL自動発注の内部処理。Mappingとレビューはなく、詳細設計から着手する。
Phase1のMappingは10/21まで。レビューの期限は10/22。全レビュー合計は1.0人日以内で、計画工数には含めない。
開発・単体の担当は未定。
"""
from __future__ import annotations

import math
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
OUT_XLSX = HERE / "【WBS】生鮮IF増分_標準案_20261012.xlsx"
OUT_MD = HERE / "標準案_日程_20261012.md"
OUT_HTML = HERE / "標準案_日程_20261012.html"

START = date(2026, 10, 12)
PROJECT_END = date(2026, 12, 18)
TIMELINE_START = date(2026, 10, 1)
TIMELINE_END = date(2026, 12, 31)

DEV_OWNER = "未定（調整中）"
CROSS_OWNER = "未定（横断）"

# code, name, kubun, difficulty, owner,
# p1, l1_dd, l1_dev, l1_test, l3_dd, l3_dev, l3_test, gate, note, test_end
IFS = [
    ("S24", "発注勧告(当日分)", "A", "高", "臧鹏仕",
     5.0, 6.0, 6.0, 5.0, 1.5, 1.5, 1.5, "－",
     "Sinops勧告ファイル。新規。2ファイルを1.5本換算の当日分", date(2026, 12, 11)),
    ("S25", "発注勧告(翌日以降分)", "A", "高", "臧鹏仕",
     2.5, 3.0, 3.0, 2.5, 0.5, 1.0, 1.0, "－",
     "当日と類似。1.5本換算の0.5本", date(2026, 12, 11)),
    ("S12", "入荷予定", "C", "高", "尹启林",
     5.0, 5.0, 6.5, 3.5, 0.5, 1.0, 1.0, "DBGET・作成時刻",
     "ベンダー流用。PC/CKは別SQL。DBGET 2系統＋作成時刻の待ち合わせ", date(2026, 12, 18)),
    ("S11", "入荷実績", "C", "中", "王静",
     4.0, 4.0, 5.0, 3.0, 0.5, 1.0, 1.0, "伝票区分",
     "ベンダー流用。伝票区分差分。PC/CKは振替", date(2026, 12, 18)),
    ("S13", "発注スケジュール", "C", "高", "杨艳",
     5.0, 5.0, 6.5, 3.5, 0.5, 1.0, 1.0, "DBGET・作成時刻",
     "ベンダー流用。PC/CKは別SQL。DBGET 2系統＋作成時刻の待ち合わせ", date(2026, 12, 18)),
    # 統合変換: TRIAL自動発注の内部処理。Mapping（p1）は 0。詳細設計から着手する。
    ("S26", "勧告→発注統合変換", "A", "高", "臧鹏仕",
     0.0, 5.0, 5.0, 4.0, 1.0, 1.0, 1.0, "非生鮮発注勧告参照",
     "TRIAL自動発注の内部処理。Mappingはせず詳細設計から着手。現行の生鮮→統合IFは使わない。非生鮮発注勧告を参照して新規", date(2026, 12, 11)),
]
ZANG_ORDER = ["S24", "S25", "S26"]

NAVY = "1A365D"
BLUE = "2B6CB0"
GREEN = "276749"
ORANGE = "C05621"
PURPLE = "6B46C1"
BROWN = "744210"
GRAY = "718096"
GOLD = "B7791F"
OCT_LAST = date(2026, 10, 30)
# Phase1レビュー。各IFのMapping完了後、期限は10/22。全レビュー合計は1.0人日以内。
# 計画工数の列には入れない（別掲）。
REVIEW_DEADLINE = date(2026, 10, 22)
REVIEW_CAP = 1.0
PLAN_EFFORT = 152.5  # 見積本体157.5 − 統合変換のMapping 5.0
ZEBRA = "F7FAFC"
LINE = "E2E8F0"
WEEKEND = "EDF2F7"
WHITE = "FFFFFF"
SECTION = "E2E8F0"

THIN = Border(
    left=Side(style="thin", color="CBD5E0"),
    right=Side(style="thin", color="CBD5E0"),
    top=Side(style="thin", color="CBD5E0"),
    bottom=Side(style="thin", color="CBD5E0"),
)
FONT = "Yu Gothic"
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def is_work(d: date) -> bool:
    return d.weekday() < 5


def next_work(d: date) -> date:
    while not is_work(d):
        d += timedelta(days=1)
    return d


def span(start: date, days: float):
    """ceil(days) 営業日。戻り値は (開始, 終了, 次の営業日)。"""
    if days <= 0:
        return None, None, next_work(start)
    n = max(1, math.ceil(days - 1e-9))
    s = next_work(start)
    e = s
    left = n - 1
    while left:
        e += timedelta(days=1)
        if is_work(e):
            left -= 1
    return s, e, next_work(e + timedelta(days=1))


def workdays(a: date, b: date):
    d = a
    out = []
    while d <= b:
        if is_work(d):
            out.append(d)
        d += timedelta(days=1)
    return out


def daterange(a: date, b: date):
    d = a
    while d <= b:
        yield d
        d += timedelta(days=1)


def fmt(d: date | None) -> str:
    if d is None:
        return "－"
    return f"{d.month}/{d.day}"


def split_window(days, parts):
    """parts: [(key, effort), ...] を連続営業日へ割り当て。末尾が終了日に着く。"""
    n = len(days)
    efforts = [p[1] for p in parts]
    counts = [max(1, math.ceil(e - 1e-9)) if e > 0 else 0 for e in efforts]
    extra = n - sum(counts)
    total = sum(efforts)
    if extra > 0 and total > 0:
        fracs = [extra * e / total for e in efforts]
        floors = [math.floor(f) for f in fracs]
        rem = extra - sum(floors)
        order = sorted(range(len(parts)), key=lambda i: (fracs[i] - floors[i], efforts[i]), reverse=True)
        for i in range(rem):
            floors[order[i % len(order)]] += 1
        counts = [counts[i] + floors[i] for i in range(len(parts))]
    elif extra < 0:
        while sum(counts) > n:
            i = max(range(len(counts)), key=lambda k: counts[k])
            floor = max(1, math.ceil(efforts[i] - 1e-9)) if efforts[i] > 0 else 0
            if counts[i] > floor:
                counts[i] -= 1
            else:
                break
    gap = n - sum(counts)
    if gap > 0:
        counts[-1] += gap
    if sum(counts) != n:
        raise RuntimeError(f"window split {sum(counts)} != {n} {counts} {efforts}")
    idx = 0
    out = {}
    for (key, _e), c in zip(parts, counts):
        seg = days[idx:idx + c]
        idx += c
        out[key] = (seg[0], seg[-1])
    return out


def dd_gate(code, row):
    """詳細設計の前提ゲート。統合変換はMapping・レビューがないので、レビュー後を付けない。"""
    if code == "S26":
        return row[12]
    return "レビュー後" if row[12] == "－" else "レビュー後／" + row[12]


def build_plan():
    by_code = {row[0]: row for row in IFS}
    design = {}

    # 臧は当日・翌日のMappingを10月内に終わらせてから詳細設計。
    # 統合変換はMappingなし（p1=0）。翌日のMapping完了の次の営業日から詳細設計。
    # レビューの期限は10/22。計画工数には入れない。
    cursor = START
    for code in ZANG_ORDER:
        row = by_code[code]
        if row[5] <= 0:
            design[code] = {"p1": None, "rev": None}
            continue
        s1, e1, cursor = span(cursor, row[5])
        design[code] = {"p1": (s1, e1), "rev": (e1, REVIEW_DEADLINE)}
    dd_cursor = cursor
    for code in ZANG_ORDER:
        row = by_code[code]
        s2, e2, dd_cursor = span(dd_cursor, row[6])
        s3, e3, dd_cursor = span(dd_cursor, row[9])
        design[code]["l1dd"] = (s2, e2)
        design[code]["l3dd"] = (s3, e3)

    for row in IFS:
        code = row[0]
        if code in ZANG_ORDER:
            continue
        cursor = START
        s1, e1, cursor = span(cursor, row[5])
        s2, e2, cursor = span(cursor, row[6])
        s3, e3, _ = span(cursor, row[9])
        design[code] = {"p1": (s1, e1), "rev": (e1, REVIEW_DEADLINE), "l1dd": (s2, e2), "l3dd": (s3, e3)}

    for code, slot in design.items():
        if slot["p1"] and slot["p1"][1] > OCT_LAST:
            raise RuntimeError(f"{code} mapping ends {slot['p1'][1]}, after October")

    later = {}
    for row in IFS:
        code = row[0]
        _s, dd_end = design[code]["l3dd"]
        dev_start = next_work(dd_end + timedelta(days=1))
        test_end = row[14]
        days = workdays(dev_start, test_end)
        parts = [
            ("l1dev", row[7]),
            ("l3dev", row[10]),
            ("l1test", row[8]),
            ("l3test", row[11]),
        ]
        later[code] = split_window(days, parts)
        later[code]["_window"] = (dev_start, test_end)

    b1_s, b1_e, nxt = span(START, 8.0)
    wh_s, wh_e, nxt = span(nxt, 2.0)
    gcs_s, gcs_e, _nxt = span(nxt, 3.0)
    b3_s, b3_e, _ = span(date(2026, 11, 16), 5.0)

    tasks = []

    def add(wbs, phase, major, minor, if_id, kubun, owner, days, diff, start, end,
            state, gate, note, color, kind="task"):
        tasks.append({
            "wbs": wbs, "phase": phase, "major": major, "minor": minor,
            "if": if_id, "kubun": kubun, "owner": owner, "days": days,
            "diff": diff, "start": start, "end": end, "state": state,
            "gate": gate, "note": note, "color": color, "kind": kind,
        })

    add("1", "調査・Mapping", "", "", "", "", "", None, "", None, None,
        "", "", "Phase1は10/21まで。統合変換はTRIAL自動発注の内部処理のためMappingなし。担当は詳細設計まで確定", NAVY, "section")
    mapped = [row for row in IFS if design[row[0]]["p1"]]
    for i, row in enumerate(mapped, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["p1"]
        add(f"1.{i}", "調査・Mapping", "Phase1", name, code, kubun, owner, row[5], diff,
            s, e, "未着手", row[12],
            "調査用表／①項目マッピング／②GAP／③変換ルール。" + row[13], BLUE)

    add("R", "レビュー", "", "", "", "", "", None, "", None, None,
        "", "", "各IFのMapping完了後にレビュー。期限10/22。全レビュー合計は1.0人日以内（計画工数には含めない）", NAVY, "section")
    for i, row in enumerate(mapped, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["rev"]
        add(f"R.{i}", "レビュー", "Phase1レビュー", name, code, kubun, owner, None, "－",
            s, e, "未着手", "詳細設計の前提",
            "Mapping完了後にレビュー。期限10/22。全レビュー合計1.0人日以内。工数は別掲", GOLD)

    add("2", "詳細設計", "", "", "", "", "", None, "", None, None,
        "", "", "統合変換を除く各IFはMapping後に着手。統合変換はMappingなしで詳細設計から。Layer①＝TRIAL↔GCS、Layer③＝DWH・GCS↔業務。Layer②は0", NAVY, "section")
    add("2.1", "詳細設計", "Layer①", "", "", "", "", None, "", None, None,
        "", "", "外部IFプログラム", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["l1dd"]
        gate = dd_gate(code, row)
        add(f"2.1.{i}", "詳細設計", "Layer①", name, code, kubun, owner, row[6], diff,
            s, e, "未着手", gate, row[13], BLUE)
    add("2.2", "詳細設計", "Layer②", "GCS→DWH", "全IF", "対象外", "－", 0, "－",
        None, None, "対象外", "－", "本件は0。DataSpider取込ジョブは作らない", GRAY)
    add("2.3", "詳細設計", "Layer③", "", "", "", "", None, "", None, None,
        "", "", "業務側連携ジョブの設計", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["l3dd"]
        gate = dd_gate(code, row)
        add(f"2.3.{i}", "詳細設計", "Layer③", name, code, kubun, owner, row[9], diff,
            s, e, "未着手", gate, "Layer③詳細設計。" + row[13], BLUE)

    add("3", "開発", "", "", "", "", "", None, "", None, None,
        "", "", "担当は未定（調整中）。日程は計画工数より長く置く", NAVY, "section")
    add("3.1", "開発", "Layer①", "", "", "", "", None, "", None, None,
        "", "", "外部IFプログラム開発", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l1dev"]
        add(f"3.1.{i}", "開発", "Layer①", name, code, kubun, DEV_OWNER, row[7], diff,
            s, e, "未着手", row[12],
            "日程は工数より長く置く。工数は標準案のまま。担当は調整中", GREEN)
    add("3.2", "開発", "Layer②", "GCS→DWH", "全IF", "対象外", "－", 0, "－",
        None, None, "対象外", "－", "本件は0", GRAY)
    add("3.3", "開発", "Layer③", "", "", "", "", None, "", None, None,
        "", "", "", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l3dev"]
        add(f"3.3.{i}", "開発", "Layer③", name, code, kubun, DEV_OWNER, row[10], diff,
            s, e, "未着手", row[12],
            "日程は工数より長く置く。工数は標準案のまま。担当は調整中", GREEN)

    add("4", "単体テスト", "", "", "", "", "", None, "", None, None,
        "", "", "結合以降は含まない。勧告3本は12/11、入荷3本は12/18（12/20は日曜）", NAVY, "section")
    add("4.1", "単体テスト", "Layer①", "", "", "", "", None, "", None, None,
        "", "", "", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l1test"]
        add(f"4.1.{i}", "単体テスト", "Layer①", name, code, kubun, DEV_OWNER, row[8], diff,
            s, e, "未着手", row[12],
            "日程は工数より長く置く。工数は標準案のまま。担当は調整中", ORANGE)
    add("4.2", "単体テスト", "Layer②", "GCS→DWH", "全IF", "対象外", "－", 0, "－",
        None, None, "対象外", "－", "本件は0", GRAY)
    add("4.3", "単体テスト", "Layer③", "", "", "", "", None, "", None, None,
        "", "", "", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l3test"]
        add(f"4.3.{i}", "単体テスト", "Layer③", name, code, kubun, DEV_OWNER, row[11], diff,
            s, e, "未着手", row[12],
            "日程は工数より長く置く。勧告3本は12/11、入荷3本は12/18まで。担当は調整中。結合以降は含まない", ORANGE)

    add("5", "横断", "", "", "", "", "", None, "", None, None,
        "", "", "B1／B3／PJ。B2は0", NAVY, "section")
    add("5.1", "横断", "B1", "店舗系12本の流用突合", "流用12本", "横断", CROSS_OWNER, 8.0, "－",
        b1_s, b1_e, "未着手", "先方協議・当社未確認",
        "商品・店別・カテゴリ・仕入先・ケースバラ・POS系・廃棄・在庫修正・棚割・新商品在庫。実データで確認", PURPLE)
    add("5.2", "横断", "B1", "倉庫系対象外の整理", "倉庫系", "横断", CROSS_OWNER, 2.0, "－",
        wh_s, wh_e, "未着手", "DC在庫なし",
        "倉庫マスタ・休日・発注曜日・発注実績・受払・倉庫勧告は対象外。水煮相当は無し（確認済）", PURPLE)
    add("5.3", "横断", "B1", "GCSアップロード要否の確認", "共通", "横断", CROSS_OWNER, 3.0, "－",
        gcs_s, gcs_e, "未着手", "－",
        "既存GCSに無い見込。確認と、必要な場合のアップロード", PURPLE)
    add("5.4", "横断", "B2", "マスタ本体IF", "流用12本", "対象外", "－", 0, "－",
        None, None, "対象外", "流用前提",
        "標準案は0。流用が覆ったら再見積", GRAY)
    add("5.5", "横断", "B3", "非生鮮Phase2への波及", "非生鮮Phase2", "横断", CROSS_OWNER, 5.0, "－",
        b3_s, b3_e, "未着手", "－",
        "伝票区分・振替の軽微注記。非生鮮Phase2の契約額は据え置き。全量の手戻りではない", BROWN)
    add("5.6", "横断", "PJ管理", "進捗・品質・会議体", "全IF", "横断", CROSS_OWNER, 14.0, "－",
        START, PROJECT_END, "未着手", "－",
        "直接143.5 × 50/499.5。既存PJ管理50の契約行は動かさない", GRAY)

    effort = sum(t["days"] or 0 for t in tasks)
    if abs(effort - PLAN_EFFORT) > 0.01:
        raise RuntimeError(f"effort {effort} != {PLAN_EFFORT}")
    if any(t["days"] for t in tasks if t["phase"] == "レビュー"):
        raise RuntimeError("review effort must stay out of 計画工数")
    return tasks, design, later


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(size=10, bold=False, color="1A202C"):
    return Font(name=FONT, size=size, bold=bold, color=color)


def write_xlsx(tasks):
    wb = Workbook()
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True

    ws = wb.active
    ws.title = "D-2_WBS"

    dates = list(daterange(TIMELINE_START, TIMELINE_END))
    first_col = 19  # S
    last_col = first_col + len(dates) - 1

    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=18)
    ws["B1"] = "生鮮IF増分　WBS（作業分解・日程）　標準案"
    ws["B1"].font = font(16, True, WHITE)
    ws["B1"].fill = fill(NAVY)
    ws["B1"].alignment = LEFT
    for col in range(2, 19):
        ws.cell(1, col).fill = fill(NAVY)
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=18)
    ws["B2"] = (
        "2026/10/12着手。単体は勧告3本が12/11、入荷3本が12/18（12/20は日曜）。"
        "Phase1（Mapping）は10/21まで。統合変換はTRIAL自動発注の内部処理でMappingなし、詳細設計から。"
        "レビューは各IFのMapping後、期限10/22。合計1.0人日以内で計画工数には含めない。"
        "計画工数152.5人日（見積157.5－統合変換のMapping 5.0）。開発・単体は未定（調整中）。"
        "土日以外は稼働。結合以降は含まない。　○＝開始　→＝作業中　▶＝終了"
    )
    ws["B2"].font = font(9, False, "1A202C")
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 48

    headers = [
        (2, "WBS No"), (3, "段階"), (4, "大分類"), (5, "小分類（作業）"),
        (6, "対象IF"), (7, "区分"), (8, "担当者"), (9, "計画工数"),
        (10, "難易度"), (11, "開始予定"), (12, "終了予定"),
        (13, "開始実績"), (14, "終了実績"), (15, "状態"),
        (16, "進捗率"), (17, "前提ゲート"), (18, "備考"),
    ]
    for col, text in headers:
        cell = ws.cell(6, col, text)
        cell.font = font(9, True, WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = CENTER
        cell.border = THIN
    ws.row_dimensions[6].height = 30
    ws.auto_filter.ref = f"B6:R{6 + len(tasks)}"
    ws.freeze_panes = "S7"
    ws.auto_filter.ref = f"B6:R{6 + len(tasks)}"

    # month headers on row 4, day number row 5, weekday row 6
    month_fill = {"10": "2C5282", "11": "2B6CB0", "12": "1A365D"}
    month_start = None
    prev_key = None
    for i, d in enumerate(dates):
        col = first_col + i
        key = f"{d.month}"
        letter = get_column_letter(col)
        head = ws.cell(4, col, f"{d.month}月" if d.day == 1 or i == 0 else None)
        head.font = font(9, True, WHITE)
        head.fill = fill(month_fill[key])
        head.alignment = CENTER
        day = ws.cell(5, col, d)
        day.number_format = "d"
        day.font = font(8, False, "4A5568")
        day.alignment = CENTER
        wd = "月火水木金土日"[d.weekday()]
        wcell = ws.cell(6, col, wd)
        wcell.font = font(8, True, "718096" if not is_work(d) else WHITE)
        wcell.alignment = CENTER
        if not is_work(d):
            day.fill = fill(WEEKEND)
            wcell.fill = fill(WEEKEND)
        else:
            day.fill = fill("EBF8FF")
            wcell.fill = fill(month_fill[key])
            wcell.font = font(8, True, WHITE)
        ws.column_dimensions[letter].width = 3.3
        if prev_key is None:
            month_start = col
            prev_key = key
        elif key != prev_key:
            if col - 1 > month_start:
                ws.merge_cells(start_row=4, start_column=month_start, end_row=4, end_column=col - 1)
            month_start = col
            prev_key = key
    if month_start and last_col > month_start:
        ws.merge_cells(start_row=4, start_column=month_start, end_row=4, end_column=last_col)
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 16

    ws.merge_cells(start_row=4, start_column=2, end_row=5, end_column=18)
    ws["B4"] = "タイムライン　2026/10/1–12/31　　青＝調査・詳細設計　金＝レビュー　緑＝開発　橙＝単体　紫＝B1　茶＝B3　灰＝PJ・対象外"
    ws["B4"].font = font(9, True, NAVY)
    ws["B4"].alignment = LEFT
    ws["B4"].fill = fill("EBF8FF")
    for col in range(2, 19):
        ws.cell(4, col).fill = fill("EBF8FF")
        ws.cell(5, col).fill = fill("EBF8FF")

    phase_font_color = {
        BLUE: WHITE, GREEN: WHITE, ORANGE: WHITE, PURPLE: WHITE,
        BROWN: WHITE, GRAY: WHITE, NAVY: WHITE,
    }

    status_rows = []
    for i, t in enumerate(tasks):
        r = 7 + i
        values = [
            t["wbs"], t["phase"], t["major"], t["minor"], t["if"], t["kubun"],
            t["owner"], t["days"] if t["days"] is not None else None, t["diff"] or None,
            t["start"], t["end"], None, None, t["state"] or None,
            0 if t["kind"] == "task" and t["state"] != "対象外" else None,
            t["gate"] or None, t["note"] or None,
        ]
        is_section = t["kind"] == "section"
        for c, v in enumerate(values, 2):
            cell = ws.cell(r, c, v)
            cell.border = THIN
            cell.alignment = CENTER if c in (2, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17) else LEFT
            if c == 18:
                cell.alignment = LEFT
            if is_section:
                cell.font = font(10, True, WHITE)
                cell.fill = fill(NAVY)
            elif t["state"] == "対象外":
                cell.font = font(9, False, "718096")
                cell.fill = fill("F7FAFC")
            else:
                cell.font = font(9)
                if i % 2 == 1 and c <= 18:
                    cell.fill = fill(ZEBRA)
        if t["days"] is not None and not is_section:
            ws.cell(r, 9).number_format = "0.0"
        if t["start"]:
            ws.cell(r, 11).number_format = "yyyy/mm/dd"
            ws.cell(r, 12).number_format = "yyyy/mm/dd"
        if t["kind"] == "task" and t["state"] != "対象外":
            ws.cell(r, 16).number_format = "0%"
            status_rows.append(r)
        ws.row_dimensions[r].height = 22 if is_section else 32

        if t["start"] and t["end"] and not is_section:
            for j, d in enumerate(dates):
                col = first_col + j
                cell = ws.cell(r, col)
                letter = get_column_letter(col)
                cell.value = (
                    f'=IF(OR($K{r}="",$L{r}=""),"",'
                    f'IF(WEEKDAY({letter}$5,2)>5,"",'
                    f'IF(AND({letter}$5>=$K{r},{letter}$5<=$L{r}),'
                    f'IF(NETWORKDAYS($K{r},{letter}$5)=1,"○",'
                    f'IF(NETWORKDAYS({letter}$5,$L{r})=1,"▶","→")),"")))'
                )
                cell.font = font(8, True, WHITE)
                cell.alignment = CENTER
                cell.border = THIN
                if not is_work(d):
                    cell.fill = fill(WEEKEND)
            start_letter = get_column_letter(first_col)
            end_letter = get_column_letter(last_col)
            rule = FormulaRule(
                formula=[f'AND({start_letter}{r}<>"",WEEKDAY({start_letter}$5,2)<=5)'],
                fill=fill(t["color"]),
                font=font(8, True, WHITE),
            )
            ws.conditional_formatting.add(f"{start_letter}{r}:{end_letter}{r}", rule)
        else:
            for j, d in enumerate(dates):
                col = first_col + j
                cell = ws.cell(r, col, "")
                cell.border = THIN
                if is_section:
                    cell.fill = fill(NAVY)
                elif not is_work(d):
                    cell.fill = fill(WEEKEND)

    if status_rows:
        dv = DataValidation(type="list", formula1='"未着手,進行中,レビュー中,完了,対象外,中止"', allow_blank=True)
        dv.error = "リストから選択"
        dv.errorTitle = "状態"
        ws.add_data_validation(dv)
        dv.add(f"O{status_rows[0]}:O{status_rows[-1]}")

    widths = {
        "A": 3, "B": 10, "C": 16, "D": 14, "E": 28, "F": 16, "G": 10,
        "H": 16, "I": 10, "J": 8, "K": 12, "L": 12, "M": 12, "N": 12,
        "O": 12, "P": 8, "Q": 20, "R": 42,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "生鮮IF増分 WBS"
    ws.oddFooter.right.text = "計画工数152.5　単体まで　12/11・12/18"
    ws.print_title_rows = "1:6"
    ws.page_setup.scale = 40
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = NAVY

    total_row = 7 + len(tasks)
    ws.cell(total_row, 5, "計画工数 合計（レビュー1.0以内は別掲）")
    ws.cell(total_row, 9, f"=SUM(I7:I{total_row - 1})")
    ws.cell(total_row, 9).number_format = "0.0"
    ws.cell(total_row, 9).font = font(11, True, NAVY)
    ws.cell(total_row, 5).font = font(10, True, NAVY)

    write_summary(wb, tasks)
    write_help(wb)
    wb.active = wb["サマリー"]
    wb.save(OUT_XLSX)


def write_summary(wb, tasks):
    ws = wb.create_sheet("サマリー", 0)
    ws.sheet_properties.tabColor = ORANGE
    ws.sheet_view.showGridLines = False
    ws["B2"] = "生鮮IF増分　標準案　日程（2026-10-12着手）"
    ws["B2"].font = font(16, True, WHITE)
    ws.merge_cells("B2:G2")
    for col in range(2, 8):
        ws.cell(2, col).fill = fill(NAVY)
    ws.row_dimensions[2].height = 28

    lines = [
        "着手は 2026/10/12（月）。単体完了は勧告3本が 12/11（金）、入荷3本が 12/18（金）。12/20は日曜。結合以降は含まない。",
        "見積本体は直接143.5＋PJ14＝157.5人日。統合変換はMappingなし（5.0）なので、計画工数は152.5人日。レビューは合計1.0人日以内で別掲。",
        "Phase1のMappingは10/21まで。レビューは各IFのMapping後、期限10/22。詳細設計の開始日はレビューでは後ろへずらさない。",
        "臧鹏仕は当日→翌日のMappingのあと詳細設計。統合変換はTRIAL自動発注の内部処理で、Mappingなしで詳細設計から。入荷3本はMapping完了の翌営業日から詳細設計。",
        "開発・単体の担当は未定（調整中）。IFごとに並行で置き、バーは計画工数より長くして 12月下旬に着地させる。",
        "土日以外は稼働。10/12・11/3・11/23 は日本の祝日だが、着手日が 10/12 なので稼働日として置く。",
        "流用12本は B1 の突合のみ。倉庫は対象外。Layer②は 0。流用が覆る、または作成時刻が待ち合わせに収まらない場合は再見積。",
    ]
    for i, text in enumerate(lines):
        ws.cell(4 + i, 2, text)
        ws.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=7)
        ws.cell(4 + i, 2).font = font(10)
        ws.cell(4 + i, 2).alignment = LEFT
        ws.row_dimensions[4 + i].height = 22

    headers = ["対象", "担当（調査・詳細設計）", "調査開始", "調査完了", "詳細設計完了", "開発開始", "単体完了", "計画工数"]
    hr = 11
    for c, h in enumerate(headers, 2):
        cell = ws.cell(hr, c, h)
        cell.font = font(9, True, WHITE)
        cell.fill = fill(NAVY)
        cell.alignment = CENTER
        cell.border = THIN

    # one row per IF: earliest start of mapping, latest of design, earliest dev, latest test, sum of days
    if_order = []
    seen = []
    for t in tasks:
        if t["if"] in ("S24", "S25", "S12", "S11", "S13", "S26") and t["if"] not in seen and t["kind"] == "task":
            seen.append(t["if"])
    for code in seen:
        mine = [t for t in tasks if t["if"] == code and t["kind"] == "task" and t["phase"] != "横断"]
        design_owner = next(t["owner"] for t in mine if t["phase"] in ("調査・Mapping", "詳細設計"))
        name = next(t["minor"] for t in mine if t["phase"] in ("調査・Mapping", "詳細設計"))
        p1 = next((t for t in mine if t["phase"] == "調査・Mapping"), None)
        dd = [t for t in mine if t["phase"] == "詳細設計"]
        dev = [t for t in mine if t["phase"] == "開発"]
        test = [t for t in mine if t["phase"] == "単体テスト"]
        days = sum((t["days"] or 0) for t in mine if t["phase"] != "レビュー")
        if_order.append((name, design_owner, p1["start"] if p1 else "Mappingなし", p1["end"] if p1 else "－",
                         max(t["end"] for t in dd),
                         min(t["start"] for t in dev), max(t["end"] for t in test), days))

    for i, row in enumerate(if_order):
        r = hr + 1 + i
        for c, v in enumerate(row, 2):
            cell = ws.cell(r, c, v)
            cell.font = font(9)
            cell.border = THIN
            cell.alignment = CENTER if c > 2 else LEFT
            if isinstance(v, date):
                cell.number_format = "yyyy/mm/dd"
            if isinstance(v, float):
                cell.number_format = "0.0"
            if i % 2:
                cell.fill = fill(ZEBRA)
        ws.row_dimensions[r].height = 20

    cross = [t for t in tasks if t["phase"] == "横断" and t["kind"] == "task"]
    r0 = hr + 2 + len(if_order)
    ws.cell(r0, 2, "横断")
    ws.cell(r0, 2).font = font(10, True, WHITE)
    ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=9)
    for c in range(2, 10):
        ws.cell(r0, c).fill = fill(NAVY)
    for i, t in enumerate(cross):
        r = r0 + 1 + i
        vals = [t["minor"], t["owner"], t["start"], t["end"], t["days"], t["note"]]
        for c, v in enumerate(vals, 2):
            cell = ws.cell(r, c, v)
            cell.font = font(9)
            cell.border = THIN
            cell.alignment = LEFT if c in (2, 7) else CENTER
            if isinstance(v, date):
                cell.number_format = "yyyy/mm/dd"
            if isinstance(v, float):
                cell.number_format = "0.0"
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 28

    for col, w in zip("BCDEFGHI", [28, 22, 14, 14, 16, 14, 14, 40]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 8
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "生鮮IF増分　日程サマリー"
    ws.oddFooter.right.text = "単体まで　勧告3本 12/11・入荷3本 12/18"
    ws.sheet_view.zoomScale = 110
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.5


def write_help(wb):
    ws = wb.create_sheet("使い方")
    ws.sheet_properties.tabColor = GRAY
    ws["B2"] = "使い方（横断管理台帳 D-2 に合わせた置き方）"
    ws["B2"].font = font(14, True, WHITE)
    ws.merge_cells("B2:D2")
    for col in range(2, 5):
        ws.cell(2, col).fill = fill(NAVY)
    notes = [
        "列はテンプレの D-2_WBS と同じ。WBS No／段階／大分類／小分類／対象IF／区分／担当者／計画工数／難易度／開始予定／終了予定／開始実績／終了実績／状態／進捗率／前提ゲート／備考。",
        "タイムラインは K列・L列の日付から自動計算。開始予定・終了予定を変えると、○ → ▶ の位置が動く。",
        "○＝開始日　→＝作業中　▶＝終了日。土日は飛ばす。",
        "背景色　青＝調査・詳細設計　金＝レビュー　緑＝開発　橙＝単体テスト　紫＝B1　茶＝B3　灰＝PJ・対象外。",
        "Phase1のMappingは10/21まで。統合変換はTRIAL自動発注の内部処理のためMappingとレビューはなく、詳細設計から。レビューは各IFのMapping後、期限10/22。合計1.0人日以内で、計画工数の列には入れない。",
        "調査・詳細設計の担当者名は確定。開発・単体は「未定（調整中）」。名前が決まったら H列だけ差し替える。",
        "見積本体は157.5人日。統合変換のMapping 5.0を行わないので、計画工数は152.5。開発・単体のバーが工数より長いのは、勧告3本12/11・入荷3本12/18に着地させるため。",
        "非稼働日を足す場合は、開始・終了を手で1日ずらす。本ファイルは土日だけを非稼働にしている。",
        "採番 S11／S12／S13／S24／S25／S26 は見積上の仮採番。台帳の正式IF IDではない。",
    ]
    for i, text in enumerate(notes):
        ws.cell(4 + i, 2, text)
        ws.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=4)
        ws.cell(4 + i, 2).alignment = LEFT
        ws.cell(4 + i, 2).font = font(10)
        ws.row_dimensions[4 + i].height = 28
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["D"].width = 36
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.oddHeader.left.text = "生鮮IF増分 WBS　使い方"


def weeks():
    out = []
    d = START
    while d <= PROJECT_END:
        out.append(d)
        d += timedelta(days=7)
    return out


def week_label(d: date) -> str:
    return f"{d.month}/{d.day}"


def overlaps(task, week_start):
    if not task["start"] or not task["end"]:
        return False
    week_end = week_start + timedelta(days=4)
    return task["start"] <= week_end and task["end"] >= week_start


def phase_class(task):
    return {
        "調査・Mapping": "p1",
        "レビュー": "rev",
        "詳細設計": "l1",
        "開発": "dev",
        "単体テスト": "l3",
        "横断": "b1" if task["major"] == "B1" else ("b3" if task["major"] == "B3" else "pj"),
    }.get(task["phase"], "pj")


def write_docs(tasks):
    wk = weeks()
    people = ["臧鹏仕", "尹启林", "王静", "杨艳"]
    ifs = ["S24", "S25", "S12", "S11", "S13", "S26"]
    names = {}
    for t in tasks:
        if t["if"] in ifs and t["kind"] == "task" and t["if"] not in names:
            names[t["if"]] = t["minor"]

    def person_cells(owner):
        rows = []
        mine = [t for t in tasks if t["owner"] == owner and t["kind"] == "task"]
        for w in wk:
            hit = [t for t in mine if overlaps(t, w)]
            if not hit:
                rows.append("")
            else:
                # shortest label
                lab = "／".join(dict.fromkeys(
                    ("Mapping" if t["phase"] == "調査・Mapping" else "詳細設計") + " " + t["minor"].split("(")[0][:6]
                    for t in hit
                ))
                rows.append((phase_class(hit[-1]), lab))
        return rows

    def if_cells(code):
        mine = [t for t in tasks if t["if"] == code and t["phase"] in ("開発", "単体テスト")]
        out = []
        for w in wk:
            hit = [t for t in mine if overlaps(t, w)]
            if not hit:
                out.append("")
            else:
                lab = "／".join(dict.fromkeys(
                    ("開発" if t["phase"] == "開発" else "単体") + t["major"].replace("Layer", "")
                    for t in hit
                ))
                out.append((phase_class(hit[-1]), lab))
        return out

    def cross_cells(major):
        mine = [t for t in tasks if t["major"] == major and t["kind"] == "task"]
        out = []
        for w in wk:
            hit = [t for t in mine if overlaps(t, w)]
            if not hit:
                out.append("")
            else:
                out.append((phase_class(hit[0]), hit[0]["minor"][:10]))
        return out

    # key dates for narrative
    def task(code, phase, major=None):
        for t in tasks:
            if t["if"] == code and t["phase"] == phase and (major is None or t["major"] == major):
                return t
        return None

    zang_dd_end = max(
        task(c, "詳細設計", "Layer③")["end"] for c in ZANG_ORDER
    )
    b1 = next(t for t in tasks if t["wbs"] == "5.1")
    wh = next(t for t in tasks if t["wbs"] == "5.2")
    gcs = next(t for t in tasks if t["wbs"] == "5.3")
    b3 = next(t for t in tasks if t["wbs"] == "5.5")

    # peak overlap of dev/test
    peak = 0
    peak_week = wk[0]
    for w in wk:
        n = 0
        for code in ifs:
            if any(overlaps(t, w) for t in tasks if t["if"] == code and t["phase"] in ("開発", "単体テスト")):
                n += 1
        if n > peak:
            peak = n
            peak_week = w

    OUT_MD.write_text(build_bilingual_md(tasks, zang_dd_end, peak, peak_week, b1, wh, gcs, b3), encoding="utf-8")
    OUT_HTML.write_text(build_html(tasks, wk, people, ifs, names, zang_dd_end, peak, peak_week, b1, wh, gcs, b3), encoding="utf-8")


def if_summary_rows(tasks):
    """IFごとの1行集計。統合変換はMappingとレビューがない。"""
    out = []
    for code in ["S24", "S25", "S12", "S11", "S13", "S26"]:
        mine = [t for t in tasks if t["if"] == code and t["kind"] == "task"]
        p1 = next((t for t in mine if t["phase"] == "調査・Mapping"), None)
        rev = next((t for t in mine if t["phase"] == "レビュー"), None)
        first = next(t for t in mine if t["phase"] in ("調査・Mapping", "詳細設計"))
        dd = max(t["end"] for t in mine if t["phase"] == "詳細設計")
        dev = min(t["start"] for t in mine if t["phase"] == "開発")
        ut = max(t["end"] for t in mine if t["phase"] == "単体テスト")
        days = sum((t["days"] or 0) for t in mine if t["phase"] != "レビュー")
        out.append({
            "name": first["minor"], "owner": first["owner"],
            "mapping": f"{fmt(p1['start'])}–{fmt(p1['end'])}" if p1 else "なし（詳細設計から）",
            "review": fmt(rev["end"]) if rev else "－",
            "dd": fmt(dd), "dev": fmt(dev), "ut": fmt(ut), "days": days,
        })
    return out


def build_bilingual_md(tasks, zang_dd_end, peak, peak_week, b1, wh, gcs, b3):
    rows = [
        f"| {r['name']} | {r['owner']} | {r['mapping']} | {r['review']} | {r['dd']} | {r['dev']} | {r['ut']} | {r['days']:.1f} |"
        for r in if_summary_rows(tasks)
    ]
    table = "\n".join(rows)
    return f"""# 生鮮IF増分・標準案　日程（2026-10-12着手）

- 日付: 2026-10-12
- 工数: 見積本体 **157.5人日**から、統合変換のMapping **5.0人日** を除いた **計画工数152.5人日**（レビューは合計1.0人日以内。別掲で、計画工数に含めない）
- 着手: **2026-10-12（月）**
- 単体完了: 勧告3本 **2026-12-11（金）**、入荷3本 **2026-12-18（金）**（12/20は日曜）
- 体制: 調査・詳細設計は4名確定。開発・単体は未定（調整中）
- WBS: `【WBS】生鮮IF増分_標準案_20261012.xlsx`
- HTML: `標準案_日程_20261012.html`
- 範囲: 調査・Mapping〜詳細設計・開発・単体。結合以降は含まない
- 前の拍子: `標準案_4人体制_概算日程_20260929.md`（10/5着手・12/8前後）

---

## 結論

### 中文

10月12日（周一）开工。Mapping 收到 **10月21日**。每本 IF 在 Mapping 完成后做 review，期限 **10月22日**。全部 review 合计不超过 **1.0人日**，不占额外日历，也不计入计划工数。

勧告 3 本（当日、翌日以降、统合转换）的单体收到 **12月11日**。入荷 3 本的单体收到 **12月18日**（12月20日是周日）。

統合转换是 TRIAL 自动订货的内部处理，**不做 Mapping，也没有单独的 review**，从详细设计开始。臧鹏仕 10/21 做完翌日的 Mapping，10/22 直接进入当日的详细设计，统合转换的详细设计在 **{fmt(zang_dd_end)}** 结束。

见积本体 157.5 人日里，统合转换的 Mapping 5.0 人日不再做，所以计划工数是 **152.5人日**。

日历按周末休息。10/12、11/3、11/23 是日本祝日，这次开工日就是 10/12，所以这三天按出勤放。

### 日文

2026年10月12日（月）に着手する。Mappingは **10月21日** まで。各IFのレビューはMapping後で、期限は **10月22日**。全レビューの合計は **1.0人日以内** で、カレンダーは足さず、計画工数にも含めない。

勧告3本（当日、翌日以降、統合変換）の単体は **12月11日** まで。入荷3本の単体は **12月18日** まで（12月20日は日曜）。

統合変換はTRIAL自動発注の内部処理なので、**Mappingは行わず、単独のレビューもない**。詳細設計から着手する。臧鹏仕は10/21に翌日のMappingを終え、10/22から当日の詳細設計へ入る。統合変換の詳細設計は **{fmt(zang_dd_end)}** に終わる。

見積本体157.5人日のうち、統合変換のMapping 5.0人日は行わない。計画工数は **152.5人日**。

カレンダーは土日だけ休む。10/12、11/3、11/23 は日本の祝日だが、着手日が 10/12 なので稼働日として置く。

---

## 担当と日付

| IF | 調査・詳細設計 | Mapping | レビュー期限 | 詳細設計完了 | 開発開始 | 単体完了 | 人日 |
|---|---|---|---|---|---|---|---:|
{table}

- 発注勧告当日／翌日以降、勧告→発注統合変換: 臧鹏仕（順次）
- 入荷予定: 尹启林
- 入荷実績: 王静
- 発注スケジュール: 杨艳
- 開発・単体: 未定（調整中）
- レビュー: 合計1.0人日以内。各IFのMapping後、期限10/22。計画工数に含めない

---

## 週次

| 週 | 日付 | 中心 |
|---|---|---|
| W1 | 10/12–10/16 | 4名でMapping開始。入荷実績は10/15にMapping完了、10/16から詳細設計。B1突合開始 |
| W2 | 10/19–10/23 | 翌日のMapping（10/21まで）。Phase1レビューの期限10/22。入荷予定・スケジュールはLayer①詳細設計。臧の当日詳細設計は10/22開始。B1突合完了、倉庫対象外 |
| W3 | 10/26–10/30 | 入荷予定・スケジュールの詳細設計完了、開発開始。GCS確認。当日のLayer①詳細設計完了（10/29） |
| W4 | 11/2–11/6 | 当日の詳細設計完了、開発開始（11/3）。翌日の詳細設計（11/3–11/6） |
| W5 | 11/9–11/13 | 統合変換のLayer①詳細設計（Mappingなし）。翌日の開発開始 |
| W6 | 11/16–11/20 | 統合変換のLayer③詳細設計完了（{fmt(zang_dd_end)}）、開発開始（11/17）。B3 |
| W7–W9 | 11/23–12/11 | 開発の残りと単体。勧告3本の単体は **12/11** |
| W10 | 12/14–12/18 | 入荷3本の単体。全体の完了は **12/18** |

同じ週に開発または単体へ入る IF は最大 **{peak}本**（{fmt(peak_week)} の週）。この並行を開発要員が持てないときは、勧告3本を直列に組み直す。1名に勧告3本の開発・単体をすべて載せると、12/11 には収まらない。

---

## 横断

| 作業 | 人日 | 日程 | 担当 |
|---|---:|---|---|
| B1 店舗系12本の流用突合 | 8.0 | {fmt(b1['start'])}–{fmt(b1['end'])} | 未定（横断） |
| B1 倉庫系対象外の整理 | 2.0 | {fmt(wh['start'])}–{fmt(wh['end'])} | 未定（横断） |
| B1 GCS確認 | 3.0 | {fmt(gcs['start'])}–{fmt(gcs['end'])} | 未定（横断） |
| B2 マスタ本体IF | 0 | 対象外 | － |
| B3 非生鮮Phase2への波及 | 5.0 | {fmt(b3['start'])}–{fmt(b3['end'])} | 未定（横断） |
| PJ管理 | 14.0 | 10/12–12/18 | 未定（横断） |

---

## ずれるとき

### 中文

- Review 合计若超过 1 人日并占住日历：当日的详细设计（10/22 开始）会往后滑，统合转换的详细设计也跟着往后，勧告 3 本在 12/11 前的余量变少。
- 11月以后开发人员到不齐：勧告 3 本的单体会从 12/11 往后滑，入荷 3 本会从 12/18 往后滑。
- 若 11/3 或 11/23 休息：臧鹏仕的详细设计各往后 1 天。统合转换的开发到 12/11 仍有余量。
- 统合转换若发现需要 Mapping（不是纯内部处理）：要补 Mapping，臧鹏仕的详细设计整体往后，这部分是 5.0 人日。
- 12本流用被推翻，或入荷予定・发注日程的作成时刻等不住：先再見積再重排。风险込み案 251人日放不进这张日程。
- 非生鲜 Phase2、共通基盤 43、既有 PJ50 的契约行不动。

### 日文

- レビュー合計が1人日を超え、カレンダーを取る場合、当日の詳細設計（10/22開始）が後ろへずれ、統合変換の詳細設計も続けてずれる。勧告3本の12/11までの余白が減る。
- 11月以降に開発要員が揃わない場合、勧告3本の単体は 12/11 から、入荷3本は 12/18 から後ろへずれる。
- 11/3 または 11/23 を休む場合、臧鹏仕の詳細設計が1日ずつ後ろへずれる。統合変換の開発・単体は12/11まで余白がある。
- 統合変換が内部処理だけでは済まず、Mappingが必要になった場合は、5.0人日を足し、臧鹏仕の詳細設計が全体に後ろへずれる。
- 流用12本が覆る、または入荷予定・発注スケジュールの作成時刻が待ち合わせに収まらない場合は、再見積してから組み直す。リスク込み案の251人日はこの日程に載らない。
- 非生鮮Phase2、共通基盤43、既存PJ50の契約行は動かさない。
"""


def bar(cell):
    if not cell:
        return '<td class="cell"></td>'
    css, lab = cell
    return f'<td class="cell"><span class="bar {css}">{lab}</span></td>'


def build_html(tasks, wk, people, ifs, names, zang_dd_end, peak, peak_week, b1, wh, gcs, b3):
    def cells_for(pred):
        out = []
        mine = [t for t in tasks if pred(t)]
        for w in wk:
            hit = [t for t in mine if overlaps(t, w)]
            if not hit:
                out.append("")
                continue
            bits = []
            for t in hit:
                if t["phase"] == "調査・Mapping":
                    bits.append("Mapping")
                elif t["phase"] == "レビュー":
                    bits.append("Review")
                elif t["phase"] == "詳細設計":
                    bits.append("設計" + t["major"].replace("Layer", ""))
                elif t["phase"] == "開発":
                    bits.append("開発" + t["major"].replace("Layer", ""))
                elif t["phase"] == "単体テスト":
                    bits.append("単体" + t["major"].replace("Layer", ""))
                else:
                    bits.append(t["minor"][:8])
            # unique preserve order
            seen = []
            for b in bits:
                if b not in seen:
                    seen.append(b)
            out.append((phase_class(hit[-1]), " ".join(seen)))
        return out

    def row(label, pred):
        tds = "".join(bar(c) for c in cells_for(pred))
        return f'<tr><td class="lane">{label}</td>{tds}</tr>'

    week_heads = "".join(f"<th>{week_label(w)}</th>" for w in wk)
    body = []
    design_phases = ("調査・Mapping", "レビュー", "詳細設計")
    body.append(row("臧鹏仕<br>調査・レビュー・設計", lambda t: t["owner"] == "臧鹏仕" and t["phase"] in design_phases))
    body.append(row("尹启林<br>入荷予定", lambda t: t["owner"] == "尹启林" and t["phase"] in design_phases))
    body.append(row("王静<br>入荷実績", lambda t: t["owner"] == "王静" and t["phase"] in design_phases))
    body.append(row("杨艳<br>スケジュール", lambda t: t["owner"] == "杨艳" and t["phase"] in design_phases))
    for code in ifs:
        body.append(row(f"開発・単体<br>{names[code]}<br>未定", lambda t, code=code: t["if"] == code and t["phase"] in ("開発", "単体テスト")))
    body.append(row("B1 横断<br>未定", lambda t: t["major"] == "B1"))
    body.append(row("B3 波及<br>未定", lambda t: t["major"] == "B3"))
    body.append(row("PJ管理<br>未定", lambda t: t["major"] == "PJ管理"))

    if_rows = []
    for r in if_summary_rows(tasks):
        if_rows.append(
            f"<tr><td>{r['name']}</td><td>{r['owner']}</td><td>{r['mapping']}</td>"
            f"<td>{r['review']}</td><td>{r['dd']}</td><td>{r['dev']}</td><td>{r['ut']}</td>"
            f"<td class='num'>{r['days']:.1f}</td></tr>"
        )

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>生鮮IF増分・標準案　日程 2026-10-12</title>
<style>
  :root {{ --navy:#1a365d; --line:#e2e8f0; --bg:#f7fafc; }}
  * {{ box-sizing: border-box; }}
  body {{ margin:0; font-family:"Hiragino Sans","Yu Gothic","Noto Sans JP","Microsoft YaHei",sans-serif; background:var(--bg); color:#1a202c; line-height:1.7; }}
  header {{ background:linear-gradient(135deg,var(--navy),#2c5282); color:#fff; padding:1.3rem 1.1rem 1.1rem; }}
  header h1 {{ margin:0 0 .4rem; font-size:1.22rem; }}
  header p {{ margin:.2rem 0; font-size:.9rem; opacity:.93; }}
  .nums {{ display:grid; grid-template-columns:repeat(4,1fr); gap:.6rem; max-width:1100px; margin:1rem auto 0; padding:0 1rem; }}
  .num {{ background:#fff; border:1px solid var(--line); border-radius:8px; padding:.7rem .8rem; }}
  .num b {{ display:block; font-size:1.05rem; color:var(--navy); }}
  .num span {{ font-size:.75rem; color:#4a5568; }}
  main {{ max-width:1100px; margin:0 auto; padding:1rem 1rem 2.6rem; }}
  section {{ background:#fff; border:1px solid var(--line); border-radius:8px; margin:.9rem 0; overflow:hidden; }}
  h2 {{ margin:0; padding:.55rem .9rem; background:var(--navy); color:#fff; font-size:.98rem; }}
  .cn, .jp {{ padding:.7rem .95rem; }}
  .cn {{ background:#f0fff4; border-bottom:1px solid var(--line); }}
  .jp {{ background:#ebf8ff; }}
  .lang {{ font-size:.72rem; font-weight:700; color:#4a5568; margin-bottom:.2rem; }}
  p {{ margin:.28rem 0; }}
  rt {{ font-size:.55em; color:#c05621; }}
  strong {{ color:var(--navy); }}
  table {{ width:100%; border-collapse:collapse; font-size:.82rem; }}
  th, td {{ border:1px solid var(--line); padding:.35rem .45rem; vertical-align:top; }}
  th {{ background:#edf2f7; text-align:left; }}
  .num {{ text-align:right; }}
  .gantt {{ width:100%; border-collapse:separate; border-spacing:0; font-size:.68rem; }}
  .gantt th {{ background:#2c5282; color:#fff; text-align:center; padding:.35rem .15rem; border:none; }}
  .gantt td {{ border-bottom:1px solid var(--line); padding:.22rem .12rem; }}
  .lane {{ width:8.2rem; font-weight:700; color:var(--navy); background:#f7fafc; font-size:.72rem; }}
  .bar {{ display:block; min-height:1.7rem; border-radius:4px; color:#fff; font-size:.62rem; line-height:1.25; text-align:center; padding:.15rem .15rem; }}
  .p1 {{ background:#2b6cb0; }}
  .rev {{ background:#b7791f; }}
  .l1 {{ background:#1a365d; }}
  .dev {{ background:#276749; }}
  .l3 {{ background:#c05621; }}
  .b1 {{ background:#6b46c1; }}
  .b3 {{ background:#744210; }}
  .pj {{ background:#718096; }}
  .legend {{ display:flex; flex-wrap:wrap; gap:.45rem; padding:.65rem .9rem; font-size:.78rem; }}
  .legend i {{ display:inline-block; width:.75rem; height:.75rem; border-radius:2px; margin-right:.25rem; vertical-align:-1px; }}
  .note {{ font-size:.82rem; color:#4a5568; padding:.2rem .95rem 1rem; }}
  @media (max-width:800px) {{ .nums {{ grid-template-columns:1fr 1fr; }} .lane {{ width:5.2rem; }} }}
</style>
</head>
<body>
<header>
  <h1>生鮮IF増分・標準案　日程</h1>
  <p>着手 2026-10-12。単体完了は勧告3本が 12/11、入荷3本が 12/18。計画工数152.5人日。結合以降は含まない。</p>
  <p>Phase1のMappingは10/21まで。統合変換はMappingなし。レビューは合計1.0人日以内（別掲）。開発・単体は未定。</p>
</header>
<div class="nums">
  <div class="num"><span>着手</span><b>10/12（月）</b></div>
  <div class="num"><span>単体完了</span><b>勧告 12/11・入荷 12/18</b></div>
  <div class="num"><span>計画工数</span><b>152.5（見積157.5－Mapping5.0）</b></div>
  <div class="num"><span>詳細設計の出口</span><b>統合変換 {fmt(zang_dd_end)}</b></div>
</div>
<main>
<section>
  <h2>結論</h2>
  <div class="cn"><div class="lang">中文</div>
    <p>10月12日开工。Mapping 收到 <strong>10月21日</strong>。Review 期限 10/22，合计不超过 <strong>1.0人日</strong>，不计入计划工数。勧告三本单体收到 <strong>12月11日</strong>，入荷三本收到 <strong>12月18日</strong>。</p>
    <p>统合转换是 TRIAL 自动订货的内部处理，<strong>不做 Mapping，也没有 review</strong>，从详细设计开始。见积 157.5 人日去掉这部分 5.0，计划工数是 <strong>152.5人日</strong>。</p>
    <p>臧鹏仕 10/21 做完翌日的 Mapping，10/22 进入当日的详细设计，统合转换在 <strong>{fmt(zang_dd_end)}</strong> 结束详细设计。入荷三本在 Mapping 次日进入详细设计。</p>
    <p>开发和单体的人未定。详细设计的下一工作日开工，单体验收到12月下旬。</p>
  </div>
  <div class="jp"><div class="lang">日文</div>
    <p>10月12日に着手する。Mappingは <strong>10月21日</strong> まで。レビューの期限は10/22、合計 <strong>1.0人日以内</strong> で計画工数に含めない。勧告3本の単体は <strong>12月11日</strong>、入荷3本は <strong>12月18日</strong> まで。</p>
    <p>統合変換はTRIAL自動発注の内部処理なので、<strong>Mappingもレビューもなく</strong>、詳細設計から着手する。見積157.5人日からMapping 5.0を除いて、計画工数は <strong>152.5人日</strong>。</p>
    <p>臧鹏仕は10/21に翌日のMappingを終え、10/22から当日の詳細設計に入る。統合変換の詳細設計は <strong>{fmt(zang_dd_end)}</strong> に終わる。入荷の3本はMappingの翌営業日から詳細設計へ入る。</p>
    <p>開発と単体の担当は未定。詳細設計の翌営業日から着手し、単体を12月下旬に置く。</p>
  </div>
</section>
<section>
  <h2>週次</h2>
  <div class="legend">
    <span><i class="p1"></i>調査・Mapping</span>
    <span><i class="rev"></i>レビュー</span>
    <span><i class="l1"></i>詳細設計</span>
    <span><i class="dev"></i>開発</span>
    <span><i class="l3"></i>単体</span>
    <span><i class="b1"></i>B1</span>
    <span><i class="b3"></i>B3</span>
    <span><i class="pj"></i>PJ</span>
  </div>
  <div style="overflow-x:auto; padding:0 .4rem 1rem;">
  <table class="gantt">
    <thead><tr><th class="lane"> </th>{week_heads}</tr></thead>
    <tbody>
    {''.join(body)}
    </tbody>
  </table>
  </div>
  <p class="note">同じ週に開発または単体へ入る IF は最大 {peak}本（{fmt(peak_week)} の週）。1名に勧告3本の開発・単体をすべて載せると、12/11 には収まらない。</p>
</section>
<section>
  <h2>IF別</h2>
  <table>
    <thead><tr><th>IF</th><th>調査・詳細設計</th><th>Mapping</th><th>レビュー期限</th><th>詳細設計完了</th><th>開発開始</th><th>単体完了</th><th>人日</th></tr></thead>
    <tbody>{''.join(if_rows)}</tbody>
  </table>
  <p class="note">B1突合 {fmt(b1['start'])}–{fmt(b1['end'])}。倉庫対象外 {fmt(wh['start'])}–{fmt(wh['end'])}。GCS {fmt(gcs['start'])}–{fmt(gcs['end'])}。B3 {fmt(b3['start'])}–{fmt(b3['end'])}。PJは 10/12–12/18。レビュー合計1.0人日以内（別掲）。担当は未定（横断）。</p>
</section>
<section>
  <h2>ずれるとき</h2>
  <div class="cn"><div class="lang">中文</div>
    <p>Review 合计若超过 1 人日并占住日历，当日的详细设计和统合转换的详细设计会往后滑。11月以后开发人员到不齐，勧告三本的单体会从 12/11 往后滑，入荷三本会从 12/18 往后滑。</p>
    <p>统合转换若发现需要 Mapping，要补 5.0 人日，臧鹏仕的详细设计整体往后。</p>
    <p>12本流用被推翻，或入荷予定・发注日程的作成时刻等不住，先再見積再重排。风险込み案 251人日放不进这张日程。</p>
  </div>
  <div class="jp"><div class="lang">日文</div>
    <p>11月以降に開発要員が揃わない場合、勧告3本の単体は 12/11 から、入荷3本は 12/18 から後ろへずれる。11/3 または 11/23 を休む場合、詳細設計が1日ずつ後ろへずれる。統合変換にMappingが必要になった場合は、5.0人日を足して詳細設計が後ろへずれる。</p>
    <p><ruby>流用<rt>りゅうよう</rt></ruby>12本が覆る、または作成時刻が待ち合わせに収まらない場合は、再見積してから組み直す。リスク込み案の251人日はこの日程に載らない。</p>
  </div>
</section>
</main>
</body>
</html>
"""


def main():
    tasks, _design, _later = build_plan()
    write_xlsx(tasks)
    write_docs(tasks)
    print("xlsx", OUT_XLSX)
    print("md", OUT_MD)
    for t in tasks:
        if t["kind"] != "task":
            continue
        if t["wbs"] in ("1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "2.1.1", "2.3.6", "3.1.6", "4.3.6", "5.1", "5.5", "5.6"):
            print(f"{t['wbs']:6} {t['minor'][:18]:18} {t['owner']:10} {t['days']} {fmt(t['start'])} {fmt(t['end'])}")


if __name__ == "__main__":
    main()
