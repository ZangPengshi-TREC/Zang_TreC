# -*- coding: utf-8 -*-
"""生鮮IF増分 標準案 WBS コンパクト版（2026-10-12着手、単体完了 11/30）。

横断管理台帳 D-2 の列に合わせる。12/18版（create_seisen_wbs_20261012.py）を圧縮した版。
  調査開始 10/12 → 調査完了 10/21 → 概要完了 10/27 → 詳細完了 11/16 → 開発完了 11/20 → 単体完了 11/30
圧縮の考え方:
  - 開発・単体は Layer① と Layer③ を並行する。
  - 各Layerは自分の詳細設計が終わった翌営業日から開発に入る（全IF一斉の11/17開始をやめる）。
  - 臧鹏仕以外の3名の詳細設計は工数どおりに詰める（早く終わる）。臧鹏仕の3本は従来どおり順に詰める。
統合変換（S26）はTRIAL自動発注の内部処理。Mappingとレビューはなく、概要設計から着手する。
概要設計の工数はLayer①詳細設計から切り出す（計画工数152.5は変えない）。
Phase1レビューの期限は10/22。全レビュー合計は1.0人日以内で、計画工数には含めない。
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
OUT_XLSX = HERE / "【WBS】生鮮IF増分_標準案_コンパクト_20261013.xlsx"
OUT_MD = HERE / "標準案_日程_コンパクト_20261013.md"
OUT_HTML = HERE / "標準案_日程_コンパクト_20261013.html"

START = date(2026, 10, 12)
PROJECT_END = date(2026, 11, 30)
TIMELINE_START = date(2026, 10, 1)
TIMELINE_END = date(2026, 12, 31)

# 全IF共通の節目（生鮮は調査が多くリスクが高い前提）
INVESTIGATE_END = date(2026, 10, 21)   # 調査完了
OVERVIEW_START = date(2026, 10, 22)
OVERVIEW_END = date(2026, 10, 27)      # 概要完了
DETAIL_START = date(2026, 10, 28)
DETAIL_END = date(2026, 11, 16)        # 詳細完了
DEV_END = date(2026, 11, 20)           # 開発完了（各Layerは自分の詳細設計の翌営業日から開発に入る）
UT_START = date(2026, 11, 23)
UT_END = date(2026, 11, 30)            # 単体完了
# 概要設計の工数。Layer①詳細設計の工数から切り出す。総工数は変えない。
OVERVIEW_EFFORT = {"S24": 2.0, "S25": 1.0, "S26": 1.0, "S12": 1.0, "S11": 1.0, "S13": 1.0}

DEV_OWNER = "未定（調整中）"
CROSS_OWNER = "未定（横断）"

# code, name, kubun, difficulty, owner,
# p1, l1_dd, l1_dev, l1_test, l3_dd, l3_dev, l3_test, gate, note, test_end
IFS = [
    ("S24", "発注勧告(当日分)", "A", "高", "臧鹏仕",
     5.0, 6.0, 6.0, 5.0, 1.5, 1.5, 1.5, "－",
     "Sinops勧告ファイル。新規。2ファイルを1.5本換算の当日分", UT_END),
    ("S25", "発注勧告(翌日以降分)", "A", "高", "臧鹏仕",
     2.5, 3.0, 3.0, 2.5, 0.5, 1.0, 1.0, "－",
     "当日と類似。1.5本換算の0.5本", UT_END),
    ("S12", "入荷予定", "C", "高", "尹启林",
     5.0, 5.0, 6.5, 3.5, 0.5, 1.0, 1.0, "DBGET・作成時刻",
     "ベンダー流用。PC/CKは別SQL。DBGET 2系統＋作成時刻の待ち合わせ", UT_END),
    ("S11", "入荷実績", "C", "中", "王静",
     4.0, 4.0, 5.0, 3.0, 0.5, 1.0, 1.0, "伝票区分",
     "ベンダー流用。伝票区分差分。PC/CKは振替", UT_END),
    ("S13", "発注スケジュール", "C", "高", "杨艳",
     5.0, 5.0, 6.5, 3.5, 0.5, 1.0, 1.0, "DBGET・作成時刻",
     "ベンダー流用。PC/CKは別SQL。DBGET 2系統＋作成時刻の待ち合わせ", UT_END),
    # 統合変換: TRIAL自動発注の内部処理。Mapping（p1）は 0。詳細設計から着手する。
    ("S26", "勧告→発注統合変換", "A", "高", "臧鹏仕",
     0.0, 5.0, 5.0, 4.0, 1.0, 1.0, 1.0, "非生鮮発注勧告参照",
     "TRIAL自動発注の内部処理。Mappingはせず詳細設計から着手。現行の生鮮→統合IFは使わない。非生鮮発注勧告を参照して新規", UT_END),
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
TEAL = "2C7A7B"
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
    """前提ゲート。統合変換はMapping・レビューがないので、レビュー後を付けない。"""
    if code == "S26":
        return row[12]
    return "レビュー後" if row[12] == "－" else "レビュー後／" + row[12]


def build_plan():
    by_code = {row[0]: row for row in IFS}
    design = {}

    # 調査（Mapping）。臧は当日→翌日を順に。他3人は10/12から並行。
    # 入荷3本は調査が重い前提で、調査完了日（10/21）まで日程を置く。計画工数は変えない。
    cursor = START
    for code in ZANG_ORDER:
        row = by_code[code]
        if row[5] <= 0:
            design[code] = {"p1": None, "rev": None}
            continue
        s1, e1, cursor = span(cursor, row[5])
        design[code] = {"p1": (s1, e1), "rev": (e1, REVIEW_DEADLINE)}
    for row in IFS:
        code = row[0]
        if code in ZANG_ORDER:
            continue
        design[code] = {"p1": (START, INVESTIGATE_END), "rev": (INVESTIGATE_END, REVIEW_DEADLINE)}
    for code, slot in design.items():
        if slot["p1"] and slot["p1"][1] > INVESTIGATE_END:
            raise RuntimeError(f"{code} mapping ends {slot['p1'][1]}, after {INVESTIGATE_END}")

    # 概要設計 10/22–10/27。臧の3本は順に（統合変換もここから着手）。他3人は枠いっぱいに置く。
    ov_cursor = OVERVIEW_START
    for code in ZANG_ORDER:
        s, e, ov_cursor = span(ov_cursor, OVERVIEW_EFFORT[code])
        design[code]["ov"] = (s, e)
        if e > OVERVIEW_END:
            raise RuntimeError(f"{code} overview ends {e}, after {OVERVIEW_END}")
    for row in IFS:
        if row[0] not in ZANG_ORDER:
            design[row[0]]["ov"] = (OVERVIEW_START, OVERVIEW_END)

    # 詳細設計 10/28–11/16。臧は順に詰める（余白なし）。他3人は枠いっぱいに置く。
    dd_cursor = DETAIL_START
    for code in ZANG_ORDER:
        row = by_code[code]
        s2, e2, dd_cursor = span(dd_cursor, row[6] - OVERVIEW_EFFORT[code])
        s3, e3, dd_cursor = span(dd_cursor, row[9])
        design[code]["l1dd"] = (s2, e2)
        design[code]["l3dd"] = (s3, e3)
    zang_end = design["S26"]["l3dd"][1]
    if zang_end > DETAIL_END:
        raise RuntimeError(f"臧の詳細設計が {zang_end} に終わる。{DETAIL_END} を超える")
    # 臧鹏仕以外は工数どおりに詰める。Layer①→Layer③の順。早く終わる。
    for row in IFS:
        code = row[0]
        if code in ZANG_ORDER:
            continue
        s2, e2, nx = span(DETAIL_START, row[6] - OVERVIEW_EFFORT[code])
        s3, e3, _ = span(nx, row[9])
        design[code]["l1dd"] = (s2, e2)
        design[code]["l3dd"] = (s3, e3)

    # 開発・単体。Layer①とLayer③を並行。各Layerは自分の詳細設計の翌営業日から開発に入る。
    # 開発は 開発完了 11/20 まで、単体は 11/23–11/30。区間いっぱいに置く（バーは工数より長い）。
    ut_days = workdays(UT_START, UT_END)
    later = {}
    for row in IFS:
        code = row[0]
        slot = {}
        for lay, dd_key, dev_eff, test_eff in (
            ("l1", "l1dd", row[7], row[8]),
            ("l3", "l3dd", row[10], row[11]),
        ):
            ds = next_work(design[code][dd_key][1] + timedelta(days=1))
            dev_days = workdays(ds, DEV_END)
            if len(dev_days) < math.ceil(dev_eff - 1e-9):
                raise RuntimeError(f"{code} {lay} 開発が {len(dev_days)} 営業日に収まらない")
            slot[lay + "dev"] = (dev_days[0], dev_days[-1])
            ut_seg = ut_days
            if len(ut_seg) < math.ceil(test_eff - 1e-9):
                raise RuntimeError(f"{code} {lay} 単体が収まらない")
            slot[lay + "test"] = (ut_seg[0], ut_seg[-1])
        later[code] = slot

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
        "", "", "調査は10/12〜10/21。統合変換はTRIAL自動発注の内部処理のためMappingなし。入荷3本は調査が重い前提で10/21まで置く。担当は詳細設計まで確定", NAVY, "section")
    mapped = [row for row in IFS if design[row[0]]["p1"]]
    for i, row in enumerate(mapped, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["p1"]
        add(f"1.{i}", "調査・Mapping", "Phase1", name, code, kubun, owner, row[5], diff,
            s, e, "未着手", row[12],
            "調査用表／①項目マッピング／②GAP／③変換ルール。" + row[13], BLUE)

    add("R", "レビュー", "", "", "", "", "", None, "", None, None,
        "", "", "各IFのMapping後にレビュー。期限10/22。全レビュー合計は1.0人日以内（計画工数には含めない）", NAVY, "section")
    for i, row in enumerate(mapped, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["rev"]
        add(f"R.{i}", "レビュー", "Phase1レビュー", name, code, kubun, owner, None, "－",
            s, e, "未着手", "概要設計の前提",
            "Mapping後にレビュー。期限10/22。全レビュー合計1.0人日以内。工数は別掲", GOLD)

    add("2", "概要設計", "", "", "", "", "", None, "", None, None,
        "", "", "概要完了 10/27。IFレイアウトの概要（送受信・頻度・時刻・件数・レイヤー別スコープ）。工数はLayer①詳細設計から切り出し、総工数は変えない。統合変換はここから着手", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["ov"]
        add(f"2.{i}", "概要設計", "概要", name, code, kubun, owner, OVERVIEW_EFFORT[code], diff,
            s, e, "未着手", dd_gate(code, row).replace("レビュー後", "Phase1レビュー後"),
            "概要設計。" + row[13], TEAL)

    add("3", "詳細設計", "", "", "", "", "", None, "", None, None,
        "", "", "詳細完了 11/16（臧鹏仕の統合変換）。Layer①＝TRIAL↔GCS、Layer③＝DWH・GCS↔業務。Layer②は0。臧鹏仕の3本は順に詰めて余白なし。他3名は工数どおりに詰め、11/3までに終わる", NAVY, "section")
    add("3.1", "詳細設計", "Layer①", "", "", "", "", None, "", None, None,
        "", "", "外部IFプログラム", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["l1dd"]
        add(f"3.1.{i}", "詳細設計", "Layer①", name, code, kubun, owner, row[6] - OVERVIEW_EFFORT[code], diff,
            s, e, "未着手", "概要設計後", row[13] if code in ZANG_ORDER else "日程は工数より長く置く。" + row[13], BLUE)
    add("3.2", "詳細設計", "Layer②", "GCS→DWH", "全IF", "対象外", "－", 0, "－",
        None, None, "対象外", "－", "本件は0。DataSpider取込ジョブは作らない", GRAY)
    add("3.3", "詳細設計", "Layer③", "", "", "", "", None, "", None, None,
        "", "", "業務側連携ジョブの設計", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, owner = row[:5]
        s, e = design[code]["l3dd"]
        add(f"3.3.{i}", "詳細設計", "Layer③", name, code, kubun, owner, row[9], diff,
            s, e, "未着手", "概要設計後", "Layer③詳細設計。" + row[13], BLUE)

    add("4", "開発", "", "", "", "", "", None, "", None, None,
        "", "", "開発完了 11/20。Layer①とLayer③は並行。各Layerは自分の詳細設計の翌営業日から開始。担当は未定（調整中）。日程は計画工数より長く置く", NAVY, "section")
    add("4.1", "開発", "Layer①", "", "", "", "", None, "", None, None,
        "", "", "外部IFプログラム開発", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l1dev"]
        add(f"4.1.{i}", "開発", "Layer①", name, code, kubun, DEV_OWNER, row[7], diff,
            s, e, "未着手", row[12],
            "Layer①・③は並行。自分の詳細設計の翌営業日から開始。日程は工数より長い。担当は調整中", GREEN)
    add("4.2", "開発", "Layer②", "GCS→DWH", "全IF", "対象外", "－", 0, "－",
        None, None, "対象外", "－", "本件は0", GRAY)
    add("4.3", "開発", "Layer③", "", "", "", "", None, "", None, None,
        "", "", "", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l3dev"]
        add(f"4.3.{i}", "開発", "Layer③", name, code, kubun, DEV_OWNER, row[10], diff,
            s, e, "未着手", row[12],
            "Layer①・③は並行。自分の詳細設計の翌営業日から開始。日程は工数より長い。担当は調整中", GREEN)

    add("5", "単体テスト", "", "", "", "", "", None, "", None, None,
        "", "", "単体完了 11/30。11/23〜11/30。Layer①とLayer③は並行。結合以降は含まない", NAVY, "section")
    add("5.1", "単体テスト", "Layer①", "", "", "", "", None, "", None, None,
        "", "", "", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l1test"]
        add(f"5.1.{i}", "単体テスト", "Layer①", name, code, kubun, DEV_OWNER, row[8], diff,
            s, e, "未着手", row[12],
            "Layer①・③は並行。自分の詳細設計の翌営業日から開始。日程は工数より長い。担当は調整中", ORANGE)
    add("5.2", "単体テスト", "Layer②", "GCS→DWH", "全IF", "対象外", "－", 0, "－",
        None, None, "対象外", "－", "本件は0", GRAY)
    add("5.3", "単体テスト", "Layer③", "", "", "", "", None, "", None, None,
        "", "", "", NAVY, "section")
    for i, row in enumerate(IFS, 1):
        code, name, kubun, diff, _owner = row[:5]
        s, e = later[code]["l3test"]
        add(f"5.3.{i}", "単体テスト", "Layer③", name, code, kubun, DEV_OWNER, row[11], diff,
            s, e, "未着手", row[12],
            "日程は工数より長く置く。単体は11/30まで。担当は調整中。結合以降は含まない", ORANGE)

    add("6", "横断", "", "", "", "", "", None, "", None, None,
        "", "", "B1／B3／PJ。B2は0", NAVY, "section")
    add("6.1", "横断", "B1", "店舗系12本の流用突合", "流用12本", "横断", CROSS_OWNER, 8.0, "－",
        b1_s, b1_e, "未着手", "先方協議・当社未確認",
        "商品・店別・カテゴリ・仕入先・ケースバラ・POS系・廃棄・在庫修正・棚割・新商品在庫。実データで確認", PURPLE)
    add("6.2", "横断", "B1", "倉庫系対象外の整理", "倉庫系", "横断", CROSS_OWNER, 2.0, "－",
        wh_s, wh_e, "未着手", "DC在庫なし",
        "倉庫マスタ・休日・発注曜日・発注実績・受払・倉庫勧告は対象外。水煮相当は無し（確認済）", PURPLE)
    add("6.3", "横断", "B1", "GCSアップロード要否の確認", "共通", "横断", CROSS_OWNER, 3.0, "－",
        gcs_s, gcs_e, "未着手", "－",
        "既存GCSに無い見込。確認と、必要な場合のアップロード", PURPLE)
    add("6.4", "横断", "B2", "マスタ本体IF", "流用12本", "対象外", "－", 0, "－",
        None, None, "対象外", "流用前提",
        "標準案は0。流用が覆ったら再見積", GRAY)
    add("6.5", "横断", "B3", "非生鮮Phase2への波及", "非生鮮Phase2", "横断", CROSS_OWNER, 5.0, "－",
        b3_s, b3_e, "未着手", "－",
        "伝票区分・振替の軽微注記。非生鮮Phase2の契約額は据え置き。全量の手戻りではない", BROWN)
    add("6.6", "横断", "PJ管理", "進捗・品質・会議体", "全IF", "横断", CROSS_OWNER, 14.0, "－",
        START, PROJECT_END, "未着手", "－",
        "直接143.5 × 50/499.5。既存PJ管理50の契約行は動かさない", GRAY)

    for t in tasks:
        if t["kind"] != "task":
            continue
        # 統合変換は外部IFのLayerではない。台帳では大分類を「-」にする。内部ではlayerに残す。
        if t["if"] == "S26" and t["phase"] in ("詳細設計", "開発", "単体テスト"):
            t["layer"] = t["major"]
            t["major"] = "-"
        # 単体テストのLayer①は台帳に合わせて新規／差分
        if t["phase"] == "単体テスト" and t.get("layer", t["major"]) == "Layer①":
            t["kubun"] = {"A": "新規", "C": "差分"}.get(t["kubun"], t["kubun"])

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
        "2026/10/12着手。コンパクト版。調査 10/12〜10/21 → 概要完了 10/27 → 詳細完了 11/16 → 開発完了 11/20 → 単体完了 11/30。"
        "開発・単体はLayer①とLayer③を並行し、各Layerは自分の詳細設計の翌営業日から開発に入る。統合変換はTRIAL自動発注の内部処理でMappingなし、概要設計から。"
        "レビューは各IFのMapping後、期限10/22。合計1.0人日以内で計画工数には含めない。"
        "計画工数152.5人日（見積157.5－統合変換のMapping 5.0。概要設計はLayer①詳細設計から切り出し）。開発・単体は未定（調整中）。"
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
    ws["B4"] = "タイムライン　2026/10/1–12/31　　青＝調査・詳細設計　金＝レビュー　緑青＝概要設計　緑＝開発　橙＝単体　紫＝B1　茶＝B3　灰＝PJ・対象外"
    ws["B4"].font = font(9, True, NAVY)
    ws["B4"].alignment = LEFT
    ws["B4"].fill = fill("EBF8FF")
    for col in range(2, 19):
        ws.cell(4, col).fill = fill("EBF8FF")
        ws.cell(5, col).fill = fill("EBF8FF")

    phase_font_color = {
        BLUE: WHITE, GREEN: WHITE, ORANGE: WHITE, PURPLE: WHITE, TEAL: WHITE, GOLD: WHITE,
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
    ws.oddFooter.right.text = "計画工数152.5　単体まで　11/30"
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
        "コンパクト版の節目: 調査 10/12〜10/21 → 概要完了 10/27 → 詳細完了 11/16 → 開発完了 11/20 → 単体完了 11/30。結合以降は含まない。",
        "入荷3本は調査を10/21まで置く。入荷3本の詳細設計は11/3までに終わる（工数どおりに詰める）。統合変換はTRIAL自動発注の内部処理で、Mappingなしで概要設計から。",
        "見積本体は直接143.5＋PJ14＝157.5人日。統合変換のMapping 5.0を行わないので計画工数は152.5人日。概要設計はLayer①詳細設計の工数から切り出し、総工数は変えない。レビューは合計1.0人日以内で別掲、期限10/22。",
        "臧鹏仕の詳細設計（10/28〜11/16）は14営業日で14人日。余白なし。概要設計は当日10/22〜23、翌日10/26、統合変換10/27。",
        "開発・単体の担当は未定（調整中）。Layer①とLayer③は並行。各Layerは自分の詳細設計の翌営業日から開発（入荷3本は11/2〜11/4、勧告は11/3〜、統合変換は11/16〜）。単体は11/23〜11/30。",
        "土日以外は稼働。10/12・11/3・11/23 は日本の祝日だが、着手日が 10/12 なので稼働日として置く。",
        "流用12本は B1 の突合のみ。倉庫は対象外。Layer②は 0。流用が覆る、または作成時刻が待ち合わせに収まらない場合は再見積。",
    ]
    for i, text in enumerate(lines):
        ws.cell(4 + i, 2, text)
        ws.merge_cells(start_row=4 + i, start_column=2, end_row=4 + i, end_column=7)
        ws.cell(4 + i, 2).font = font(10)
        ws.cell(4 + i, 2).alignment = LEFT
        ws.row_dimensions[4 + i].height = 22

    headers = ["対象", "担当（調査・設計）", "調査", "概要完了", "詳細完了", "開発完了", "単体完了", "計画工数"]
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
        design_owner = next(t["owner"] for t in mine if t["phase"] in ("調査・Mapping", "概要設計"))
        name = next(t["minor"] for t in mine if t["phase"] in ("調査・Mapping", "概要設計"))
        p1 = next((t for t in mine if t["phase"] == "調査・Mapping"), None)
        ov = [t for t in mine if t["phase"] == "概要設計"]
        dd = [t for t in mine if t["phase"] == "詳細設計"]
        dev = [t for t in mine if t["phase"] == "開発"]
        test = [t for t in mine if t["phase"] == "単体テスト"]
        days = sum((t["days"] or 0) for t in mine if t["phase"] != "レビュー")
        mapping = f"{fmt(p1['start'])}–{fmt(p1['end'])}" if p1 else "なし"
        if_order.append((name, design_owner, mapping, max(t["end"] for t in ov),
                         max(t["end"] for t in dd), max(t["end"] for t in dev),
                         max(t["end"] for t in test), days))

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
    ws.oddFooter.right.text = "単体完了　2026/11/30"
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
        "背景色　青＝調査・詳細設計　金＝レビュー　緑青＝概要設計　緑＝開発　橙＝単体テスト　紫＝B1　茶＝B3　灰＝PJ・対象外。",
        "節目はコンパクト版。調査 10/12〜10/21、概要完了 10/27、詳細完了 11/16、開発完了 11/20、単体完了 11/30。統合変換はTRIAL自動発注の内部処理のためMappingとレビューはなく、概要設計から。レビューは各IFのMapping後、期限10/22。合計1.0人日以内で、計画工数の列には入れない。",
        "調査・詳細設計の担当者名は確定。開発・単体は「未定（調整中）」。名前が決まったら H列だけ差し替える。",
        "見積本体は157.5人日。統合変換のMapping 5.0を行わないので、計画工数は152.5。概要設計の工数はLayer①詳細設計から切り出し（総工数は変えない）。開発・単体のバーが工数より長いのは、開発完了11/20・単体完了11/30の区間いっぱいに置くため。",
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
        "概要設計": "ov",
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
                    ("開発" if t["phase"] == "開発" else "単体") + t.get("layer", t["major"]).replace("Layer", "")
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
            if t["if"] == code and t["phase"] == phase and (major is None or t.get("layer", t["major"]) == major):
                return t
        return None

    zang_dd_end = max(
        task(c, "詳細設計", "Layer③")["end"] for c in ZANG_ORDER
    )
    b1 = next(t for t in tasks if t["wbs"] == "6.1")
    wh = next(t for t in tasks if t["wbs"] == "6.2")
    gcs = next(t for t in tasks if t["wbs"] == "6.3")
    b3 = next(t for t in tasks if t["wbs"] == "6.5")

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
        first = next(t for t in mine if t["phase"] in ("調査・Mapping", "概要設計"))
        ov = max(t["end"] for t in mine if t["phase"] == "概要設計")
        dd = max(t["end"] for t in mine if t["phase"] == "詳細設計")
        dev = max(t["end"] for t in mine if t["phase"] == "開発")
        ut = max(t["end"] for t in mine if t["phase"] == "単体テスト")
        days = sum((t["days"] or 0) for t in mine if t["phase"] != "レビュー")
        out.append({
            "name": first["minor"], "owner": first["owner"],
            "mapping": f"{fmt(p1['start'])}–{fmt(p1['end'])}" if p1 else "なし",
            "ov": fmt(ov), "dd": fmt(dd), "dev": fmt(dev), "ut": fmt(ut), "days": days,
        })
    return out


def load_numbers(tasks):
    """開発・単体の合計人日と、日ごとの必要人数（工数÷バーの営業日数）のピーク。"""
    from collections import defaultdict
    dev = sum((t["days"] or 0) for t in tasks if t["phase"] == "開発")
    ut = sum((t["days"] or 0) for t in tasks if t["phase"] == "単体テスト")
    load = {"開発": defaultdict(float), "単体テスト": defaultdict(float)}
    for t in tasks:
        if t["kind"] == "task" and t["phase"] in load and t["start"]:
            wd = workdays(t["start"], t["end"])
            for d in wd:
                load[t["phase"]][d] += (t["days"] or 0) / len(wd)
    dev_peak = max(load["開発"].values())
    ut_peak = max(load["単体テスト"].values())
    # 参考: 全IF一斉に 11/17 から始める場合の必要人数（11/17〜11/30）
    alt = (dev + ut) / len(workdays(date(2026, 11, 17), UT_END))
    return dev, ut, dev_peak, ut_peak, alt


def build_bilingual_md(tasks, zang_dd_end, peak, peak_week, b1, wh, gcs, b3):
    rows = [
        f"| {r['name']} | {r['owner']} | {r['mapping']} | {r['ov']} | {r['dd']} | {r['dev']} | {r['ut']} | {r['days']:.1f} |"
        for r in if_summary_rows(tasks)
    ]
    table = "\n".join(rows)
    dev, ut, dev_peak, ut_peak, alt = load_numbers(tasks)
    return f"""# 生鮮IF増分・標準案　日程 コンパクト版（2026-10-12着手、単体完了 11/30）

- 日付: 2026-10-13
- 工数: 見積本体 **157.5人日**から、統合変換のMapping **5.0人日** を除いた **計画工数152.5人日**（レビューは合計1.0人日以内。別掲で、計画工数に含めない）
- 着手: **2026-10-12（月）**
- 節目: 調査 **10/12〜10/21** → 概要完了 **10/27** → 詳細完了 **11/16** → 開発完了 **11/20** → 単体完了 **11/30**
- 体制: 調査・設計は4名確定。開発・単体は未定（調整中）
- WBS: `【WBS】生鮮IF増分_標準案_コンパクト_20261013.xlsx`
- HTML: `標準案_日程_コンパクト_20261013.html`
- 12/18版: `標準案_日程_20261012.md`（開発11/17〜12/3、単体12/4〜12/18。人員が揃わない場合の安全側）
- 範囲: 調査・Mapping〜概要・詳細設計・開発・単体。結合以降は含まない

---

## 結論

### 中文

目标是 **11/30 结束单体测试**，比 12/18 版提前 14 个工作日。压缩靠两点：

1. 开发・单体里 **Layer① 和 Layer③ 并行**（Layer② 本次为 0）。
2. 每个 Layer **在自己的详细设计完成的次日就开始开发**，不再等全部 IF 在 11/17 一起开工。臧鹏仕以外三人的详细设计按工数排满（不再拉长到 11/16），入荷三本在 **11/3 前**设计完成，开发 11/2～11/4 就能开始。

节点：调查 10/21 → 概要 10/27 → 详细 11/16 → 开发 **11/20** → 单体 **11/30**。开发 11/20 结束后，单体 11/23～11/30 共 6 个工作日。

**前提（需要确认）**：开发人员 11/2 起就位，并且和 11/16 前的详细设计并行。开发约 **{dev_peak:.1f} 人**（峰值）、单体约 **{ut_peak:.1f} 人**（峰值）。如果开发只能 11/17 全体一起开工，11/17～11/30 只有 10 个工作日，需要约 **{alt:.1f} 人**，现实中放不下，这时请用 12/18 版。

没有余量的地方：臧鹏仕的详细设计（10/28～11/16，14 个工作日对 14 人日）；统合转换 Layer① 开发 11/16～11/20（5 个工作日对 5 人日）；单体 11/23～11/30 里 勧告当日 Layer① 5 人日只有 6 个工作日。

统合转换是 TRIAL 自动订货的内部处理，**不做 Mapping，也没有 review**，从概要设计开始。概要设计工数从 Layer① 详细设计里切出，总工数不变，计划工数 **152.5人日**。Review 期限 10/22，合计不超过 1.0 人日，不计入计划工数。

### 日文

目標は **11/30 に単体テストを終える** こと。12/18版より14営業日早い。圧縮の方法は2つ。

1. 開発・単体で **Layer①とLayer③を並行** する（Layer②は本件は0）。
2. 各Layerは **自分の詳細設計が終わった翌営業日に開発へ入る**。全IFが11/17に一斉に始めるのをやめる。臧鹏仕以外の3名の詳細設計は工数どおりに詰め（11/16まで引き伸ばさない）、入荷3本は **11/3まで** に設計が終わり、11/2〜11/4に開発へ入れる。

節目: 調査 10/21 → 概要 10/27 → 詳細 11/16 → 開発 **11/20** → 単体 **11/30**。開発完了11/20のあと、単体は11/23〜11/30の6営業日。

**前提（要確認）**: 開発要員が11/2から入り、11/16までの詳細設計と並行で動けること。必要人数は開発が **ピーク約{dev_peak:.1f}人**、単体が **ピーク約{ut_peak:.1f}人**。開発が全体で11/17にしか始められない場合は、11/17〜11/30の10営業日で **約{alt:.1f}人** が必要になり、現実的でない。その場合は12/18版を使う。

余白がないところ: 臧鹏仕の詳細設計（10/28〜11/16、14営業日に14人日）。統合変換のLayer①開発 11/16〜11/20（5営業日に5人日）。単体 11/23〜11/30 のうち勧告（当日分）Layer①は5人日に対して6営業日。

統合変換はTRIAL自動発注の内部処理なので、**Mappingは行わず、単独のレビューもない**。概要設計から着手する。概要設計の工数はLayer①詳細設計から切り出し、総工数は変えず、計画工数は **152.5人日**。レビュー期限10/22、合計1.0人日以内で計画工数に含めない。

---

## 担当と日付

| IF | 調査・設計 | 調査 | 概要完了 | 詳細完了 | 開発完了 | 単体完了 | 人日 |
|---|---|---|---|---|---|---|---:|
{table}

- 発注勧告当日／翌日以降、勧告→発注統合変換: 臧鹏仕（順次）
- 入荷予定: 尹启林
- 入荷実績: 王静
- 発注スケジュール: 杨艳
- 開発・単体: 未定（調整中）
- 概要設計: 臧鹏仕は当日 10/22–10/23、翌日 10/26、統合変換 10/27（順次）。他3名は10/22–10/27
- 詳細設計: 臧鹏仕は当日 10/28–11/4、翌日 11/5–11/9、統合変換 11/10–11/16（順次）。入荷予定・発注スケジュールは10/28–11/3、入荷実績は10/28–11/2
- 開発の開始（Layer①）: 入荷実績 11/2、入荷予定・発注スケジュール・勧告当日 11/3、勧告翌日 11/9、統合変換 11/16。Layer③は各Layer③の詳細設計の翌営業日
- 開発は全IF 11/20まで、単体は全IF 11/23〜11/30（Layer①・③は並行）
- レビュー: 合計1.0人日以内。各IFのMapping後、期限10/22。計画工数に含めない

---

## 週次

| 週 | 日付 | 中心 |
|---|---|---|
| W1 | 10/12–10/16 | 調査。臧は当日。入荷3本の調査（10/21まで）。B1突合開始 |
| W2 | 10/19–10/23 | 臧は翌日の調査（10/21まで）。**調査完了 10/21**。Phase1レビュー期限10/22。概要設計開始（10/22）。B1突合完了、倉庫対象外 |
| W3 | 10/26–10/30 | 臧は翌日・統合変換の概要設計。**概要完了 10/27**。詳細設計開始（10/28）。入荷実績は詳細設計をLayer①→③と10/30まで。GCS確認 |
| W4 | 11/2–11/6 | 入荷3本の**詳細設計が終わる**（11/2〜11/3）。入荷3本と勧告（当日）の開発開始（11/2〜11/5）。臧は当日のLayer③（11/3–11/4）、翌日のLayer①（11/5–11/6） |
| W5 | 11/9–11/13 | 臧は翌日のLayer③（11/9）、統合変換のLayer①（11/10–11/13）。勧告（翌日）の開発開始（11/9〜11/10） |
| W6 | 11/16–11/20 | **詳細完了 11/16**（統合変換のLayer③）。統合変換の開発開始（11/16〜）。B3。**開発完了 11/20** |
| W7 | 11/23–11/27 | 単体（Layer①・③並行） |
| W8 | 11/30 | 単体。**単体完了 11/30** |

同じ週に開発または単体へ入る IF は最大 **{peak}本**（{fmt(peak_week)} の週）。開発は計 **{dev:.1f}人日**（ピーク約{dev_peak:.1f}人）、単体は計 **{ut:.1f}人日**（ピーク約{ut_peak:.1f}人）。

---

## 横断

| 作業 | 人日 | 日程 | 担当 |
|---|---:|---|---|
| B1 店舗系12本の流用突合 | 8.0 | {fmt(b1['start'])}–{fmt(b1['end'])} | 未定（横断） |
| B1 倉庫系対象外の整理 | 2.0 | {fmt(wh['start'])}–{fmt(wh['end'])} | 未定（横断） |
| B1 GCS確認 | 3.0 | {fmt(gcs['start'])}–{fmt(gcs['end'])} | 未定（横断） |
| B2 マスタ本体IF | 0 | 対象外 | － |
| B3 非生鮮Phase2への波及 | 5.0 | {fmt(b3['start'])}–{fmt(b3['end'])} | 未定（横断） |
| PJ管理 | 14.0 | 10/12–11/30 | 未定（横断） |

---

## ずれるとき

### 中文

- 11/2 起开发人员到不齐：单体 11/30 保不住。每晚 1 天，单体就要往后推；开发峰值约 {dev_peak:.1f} 人，单体峰值约 {ut_peak:.1f} 人。
- 调查 10/21 做不完：概要、详细依次后滑。臧鹏仕的详细设计没有余量，晚 1 天就超过 11/16，统合转换的开发也跟着晚（11/16～11/20 没有余量）。11/3 或 11/23 休息也一样。
- Review 合计超过 1 人日并占住日历：10/22 开始的概要设计后滑，臧鹏仕的链条会顶到 11/16 之后。
- 统合转换若发现需要 Mapping：要补 5.0 人日，臧鹏仕的概要・详细设计整体后移，11/30 保不住。
- 12本流用被推翻，或入荷予定・发注日程的作成时刻等不住：先再見積再重排。风险込み案 251人日放不进这张日程。
- 非生鲜 Phase2、共通基盤 43、既有 PJ50 的契约行不动。

### 日文

- 11/2 から開発要員が揃わない場合、単体 11/30 は守れない。1日遅れるごとに単体が後ろへずれる。開発はピーク約{dev_peak:.1f}人、単体はピーク約{ut_peak:.1f}人。
- 調査が10/21に終わらない場合、概要・詳細が順にずれる。臧鹏仕の詳細設計は余白がなく、1日遅れると11/16を超え、統合変換の開発（11/16〜11/20は余白なし）も遅れる。11/3・11/23を休む場合も同じ。
- レビュー合計が1人日を超え、カレンダーを取る場合、10/22開始の概要設計が後ろへずれ、臧鹏仕の連鎖が11/16を超える。
- 統合変換にMappingが必要になった場合は、5.0人日を足し、臧鹏仕の概要・詳細設計が全体に後ろへずれ、11/30は守れない。
- 流用12本が覆る、または入荷予定・発注スケジュールの作成時刻が待ち合わせに収まらない場合は、再見積してから組み直す。リスク込み案の251人日はこの日程に載らない。
- 非生鮮Phase2、共通基盤43、既存PJ50の契約行は動かさない。
"""


def bar(cell):
    if not cell:
        return '<td class="cell"></td>'
    css, lab = cell
    return f'<td class="cell"><span class="bar {css}">{lab}</span></td>'


def build_html(tasks, wk, people, ifs, names, zang_dd_end, peak, peak_week, b1, wh, gcs, b3):
    dev, ut, dev_peak, ut_peak, alt = load_numbers(tasks)
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
                elif t["phase"] == "概要設計":
                    bits.append("概要")
                elif t["phase"] == "詳細設計":
                    bits.append("設計" + t.get("layer", t["major"]).replace("Layer", ""))
                elif t["phase"] == "開発":
                    bits.append("開発" + t.get("layer", t["major"]).replace("Layer", ""))
                elif t["phase"] == "単体テスト":
                    bits.append("単体" + t.get("layer", t["major"]).replace("Layer", ""))
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
    design_phases = ("調査・Mapping", "レビュー", "概要設計", "詳細設計")
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
            f"<td>{r['ov']}</td><td>{r['dd']}</td><td>{r['dev']}</td><td>{r['ut']}</td>"
            f"<td class='num'>{r['days']:.1f}</td></tr>"
        )

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>生鮮IF増分・標準案　日程 コンパクト版 2026-10-13</title>
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
  .ov {{ background:#2c7a7b; }}
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
  <h1>生鮮IF増分・標準案　日程（コンパクト版）</h1>
  <p>コンパクト版。着手 2026-10-12。調査 10/21 → 概要 10/27 → 詳細 11/16 → 開発 11/20 → 単体 11/30。計画工数152.5人日。結合以降は含まない。</p>
  <p>Layer①・③を並行し、各Layerは自分の詳細設計の翌営業日から開発。統合変換はMappingなし。レビューは合計1.0人日以内（別掲）。開発・単体は未定。</p>
</header>
<div class="nums">
  <div class="num"><span>着手</span><b>10/12（月）</b></div>
  <div class="num"><span>調査→概要→詳細→開発→単体</span><b>10/21→10/27→11/16→11/20→11/30</b></div>
  <div class="num"><span>計画工数</span><b>152.5（見積157.5－Mapping5.0）</b></div>
  <div class="num"><span>詳細設計の出口（臧・余白なし）</span><b>統合変換 {fmt(zang_dd_end)}</b></div>
</div>
<main>
<section>
  <h2>結論</h2>
  <div class="cn"><div class="lang">中文</div>
    <p>目标 <strong>11/30 结束单体</strong>，比 12/18 版提前 14 个工作日。靠两点压缩：开发・单体里 <strong>Layer① 与 Layer③ 并行</strong>；每个 Layer <strong>在自己的详细设计完成的次日就开发</strong>，不等 11/17 一起开工。</p>
    <p>节点：调查 <strong>10/21</strong> → 概要 <strong>10/27</strong> → 详细 <strong>11/16</strong> → 开发 <strong>11/20</strong> → 单体 <strong>11/30</strong>（11/23～11/30 共 6 个工作日）。统合转换不做 Mapping、没有 review，从概要设计开始，计划工数 <strong>152.5人日</strong>。</p>
    <p><strong>前提</strong>：开发人员 11/2 起就位并与详细设计并行。开发峰值约 <strong>{dev_peak:.1f}</strong> 人，单体峰值约 <strong>{ut_peak:.1f}</strong> 人；若 11/17 才一起开工需约 <strong>{alt:.1f}</strong> 人，放不下，请用 12/18 版。无余量处：臧鹏仕的详细设计（{fmt(zang_dd_end)}）。</p>
  </div>
  <div class="jp"><div class="lang">日文</div>
    <p>目標は <strong>11/30 に単体を終える</strong> こと（12/18版より14営業日早い）。圧縮は2点。開発・単体で <strong>Layer①とLayer③を並行</strong>。各Layerは <strong>自分の詳細設計が終わった翌営業日に開発へ入る</strong>（11/17一斉開始をやめる）。</p>
    <p>節目: 調査 <strong>10/21</strong> → 概要 <strong>10/27</strong> → 詳細 <strong>11/16</strong> → 開発 <strong>11/20</strong> → 単体 <strong>11/30</strong>（11/23〜11/30の6営業日）。統合変換はMappingもレビューもなく概要設計から。計画工数 <strong>152.5人日</strong>。</p>
    <p><strong>前提</strong>: 開発要員が11/2から入り、詳細設計と並行で動けること。必要人数は開発ピーク約 <strong>{dev_peak:.1f}</strong> 人、単体ピーク約 <strong>{ut_peak:.1f}</strong> 人。11/17一斉開始だと約 <strong>{alt:.1f}</strong> 人が必要で現実的でないため、その場合は12/18版を使う。余白がないのは臧鹏仕の詳細設計（{fmt(zang_dd_end)}）。</p>
  </div>
</section>
<section>
  <h2>週次</h2>
  <div class="legend">
    <span><i class="p1"></i>調査・Mapping</span>
    <span><i class="rev"></i>レビュー</span>
    <span><i class="ov"></i>概要設計</span>
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
  <p class="note">同じ週に開発または単体へ入る IF は最大 {peak}本（{fmt(peak_week)} の週）。開発は11/20まで、単体は11/23〜11/30。Layer①・③は並行。</p>
</section>
<section>
  <h2>IF別</h2>
  <table>
    <thead><tr><th>IF</th><th>調査・設計</th><th>調査</th><th>概要完了</th><th>詳細完了</th><th>開発完了</th><th>単体完了</th><th>人日</th></tr></thead>
    <tbody>{''.join(if_rows)}</tbody>
  </table>
  <p class="note">B1突合 {fmt(b1['start'])}–{fmt(b1['end'])}。倉庫対象外 {fmt(wh['start'])}–{fmt(wh['end'])}。GCS {fmt(gcs['start'])}–{fmt(gcs['end'])}。B3 {fmt(b3['start'])}–{fmt(b3['end'])}。PJは 10/12–11/30。レビュー合計1.0人日以内（別掲）。担当は未定（横断）。</p>
</section>
<section>
  <h2>ずれるとき</h2>
  <div class="cn"><div class="lang">中文</div>
    <p>调查 10/21 做不完（生鲜风险高）：概要、详细依次后滑。臧鹏仕的详细设计没有余量，晚 1 天就超过 11/16；11/3 或 11/23 休息也一样。Review 合计超过 1 人日并占住日历，同样会顶到 11/16 之后。</p>
    <p>11/2 起开发人员到不齐：单体 11/30 保不住。调查 10/21 做不完则概要、详细依次后滑；臧鹏仕的详细设计没有余量，晚 1 天就超过 11/16，统合转换开发（11/16～11/20）也没有余量。11/3 或 11/23 休息、review 合计超过 1 人日占住日历也一样。</p>
    <p>统合转换若发现需要 Mapping，要补 5.0 人日，11/30 保不住。12本流用被推翻，或入荷予定・发注日程的作成时刻等不住，先再見積再重排。风险込み案 251人日放不进这张日程。</p>
  </div>
  <div class="jp"><div class="lang">日文</div>
    <p>11/2 から開発要員が揃わない場合、単体11/30は守れない。調査が10/21に終わらない場合は概要・詳細が順にずれる。臧鹏仕の詳細設計は余白がなく、1日遅れると11/16を超え、統合変換の開発（11/16〜11/20）も余白がない（11/3・11/23を休む場合、レビュー合計が1人日を超えてカレンダーを取る場合も同じ）。統合変換にMappingが必要になった場合は、5.0人日を足すと11/30は守れない。</p>
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
        if t["wbs"] in ("1.1", "1.2", "1.3", "1.4", "1.5", "2.1", "2.6", "3.1.1", "3.3.6", "4.1.6", "5.3.6", "6.1", "6.5", "6.6"):
            print(f"{t['wbs']:6} {t['minor'][:18]:18} {t['owner']:10} {t['days']} {fmt(t['start'])} {fmt(t['end'])}")


if __name__ == "__main__":
    main()
