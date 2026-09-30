#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生鮮IF増分・標準案見積の説明PPT。"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

OUT = Path(__file__).resolve().parent / "【説明】生鮮IF増分_標準案見積_リスク着眼_20260921.pptx"

# Palette
NAVY = RGBColor(0x1F, 0x4E, 0x78)
BLUE = RGBColor(0x2E, 0x75, 0xB6)
LIGHT = RGBColor(0xD6, 0xE3, 0xF0)
ORANGE = RGBColor(0xC6, 0x59, 0x11)
ORANGE_BG = RGBColor(0xFC, 0xE4, 0xD6)
GREEN = RGBColor(0x54, 0x89, 0x35)
GREEN_BG = RGBColor(0xE2, 0xF0, 0xD9)
GRAY = RGBColor(0x59, 0x59, 0x59)
DARK = RGBColor(0x2F, 0x2F, 0x2F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED_SOFT = RGBColor(0xC0, 0x00, 0x00)


def set_run(run, size=14, bold=False, color=DARK, name="Yu Gothic"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def add_text(tf, text, size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT, clear=True):
    if clear:
        tf.clear()
        p = tf.paragraphs[0]
        p.text = ""
    else:
        p = tf.add_paragraph()
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def add_para(tf, text, size=13, bold=False, color=DARK, space_before=4, space_after=2):
    p = tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def blank_slide(prs):
    layout = prs.slide_layouts[6]  # blank
    return prs.slides.add_slide(layout)


def banner(slide, title, subtitle=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.85)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = True
    add_text(tf, title, size=22, bold=True, color=WHITE)
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    shape.text_frame.margin_left = Inches(0.35)
    shape.text_frame.margin_top = Inches(0.18)
    if subtitle:
        bar = slide.shapes.add_textbox(Inches(0.35), Inches(0.92), Inches(12.5), Inches(0.35))
        add_text(bar.text_frame, subtitle, size=12, color=GRAY)


def footer(slide, page, total=11):
    box = slide.shapes.add_textbox(Inches(0.35), Inches(7.15), Inches(10), Inches(0.3))
    add_text(
        box.text_frame,
        "西友MD基幹統合｜生鮮IF増分見積｜2026-09-30",
        size=10,
        color=GRAY,
    )
    num = slide.shapes.add_textbox(Inches(12.2), Inches(7.15), Inches(0.9), Inches(0.3))
    add_text(num.text_frame, f"{page}/{total}", size=10, color=GRAY, align=PP_ALIGN.RIGHT)


def card(slide, left, top, width, height, fill=LIGHT):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = RGBColor(0xB0, 0xC4, 0xDE)
    sh.adjustments[0] = 0.08
    return sh


def table_slide_data(slide, left, top, rows, col_widths, font_size=11):
    """rows: list of list of str; first row = header."""
    nrows = len(rows)
    ncols = len(rows[0])
    table_shape = slide.shapes.add_table(
        nrows, ncols, left, top, sum(col_widths), Inches(0.32 * nrows)
    )
    table = table_shape.table
    for i, w in enumerate(col_widths):
        table.columns[i].width = w
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT if c > 0 or r == 0 else PP_ALIGN.CENTER
                for run in p.runs:
                    set_run(
                        run,
                        size=font_size,
                        bold=(r == 0),
                        color=WHITE if r == 0 else DARK,
                    )
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
            elif r == nrows - 1 and ("合計" in str(row[0]) or "計" == str(row[0])[-1:]):
                cell.fill.solid()
                cell.fill.fore_color.rgb = GREEN_BG
            elif r % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF5, 0xF8, 0xFB)
    return table


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = 11

    # ---- 1 Cover ----
    s = blank_slide(prs)
    bg = s.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY
    bg.line.fill.background()
    accent = s.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.9), Inches(13.333), Inches(1.6)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = BLUE
    accent.line.fill.background()

    t = s.shapes.add_textbox(Inches(0.7), Inches(2.0), Inches(11.5), Inches(1.2))
    add_text(t.text_frame, "生鮮IF増分見積　説明", size=36, bold=True, color=WHITE)
    t2 = s.shapes.add_textbox(Inches(0.7), Inches(3.3), Inches(11.5), Inches(0.6))
    add_text(
        t2.text_frame,
        "標準案｜調査・Mapping〜詳細設計・開発・単体",
        size=18,
        color=LIGHT,
    )
    t3 = s.shapes.add_textbox(Inches(0.7), Inches(4.1), Inches(11.5), Inches(0.5))
    add_text(
        t3.text_frame,
        "160.5人日／¥5,120,000（税抜）",
        size=20,
        bold=True,
        color=ORANGE_BG,
    )
    t4 = s.shapes.add_textbox(Inches(0.7), Inches(6.25), Inches(11.5), Inches(0.8))
    tf = t4.text_frame
    add_text(tf, "西友MD基幹統合　／　Sinops 自動補充 PMI", size=14, color=WHITE)
    add_para(tf, "2026-09-30　出典：マスタデータ判定、標準案Excel、2026-09-18 生鮮IF会議", size=12, color=LIGHT)

    # ---- 2 結論 ----
    s = blank_slide(prs)
    banner(s, "1. 提示値", "生鮮増分のみ。非生鮮Phase2の契約額は据え置き。")
    cards = [
        ("人日", "160.5", "直接145.5（P1 41.5＋P2 104）＋PJ15"),
        ("金額（税抜）", "¥5,120,000", "Phase1 4万／Phase2 2.75万／PJ 4万"),
        ("前提", "流用は未確認", "先方協議の結果を仮置き。B1で突合"),
    ]
    for i, (h, v, n) in enumerate(cards):
        c = card(s, Inches(0.4 + i * 4.2), Inches(1.5), Inches(4.0), Inches(2.0), LIGHT)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.2)
        add_text(tf, h, size=13, color=BLUE, bold=True)
        add_para(tf, v, size=26, bold=True, color=NAVY, space_before=8)
        add_para(tf, n, size=12, color=GRAY, space_before=6)

    box = card(s, Inches(0.4), Inches(3.8), Inches(12.5), Inches(2.9), ORANGE_BG)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_top = Inches(0.15)
    add_text(tf, "先に伝えること", size=15, bold=True, color=ORANGE)
    for line in [
        "・確定請負ではない。突合で覆ったら再見積。",
        "・店舗系12本は先方協議では流用（当社未同席）。勧告2ファイルは1.5本。発注統合変換は新規。",
        "・B2は0。倉庫は対象外。水煮相当（倉庫商品マスタ）は残件。",
        "・非生鮮Phase2の契約額は動かさない。",
    ]:
        add_para(tf, line, size=14, color=DARK, space_before=6)
    footer(s, 2, total)

    # ---- 3 見積前提 ----
    s = blank_slide(prs)
    banner(
        s,
        "2. 前提",
        "店舗系12本は先方協議では流用。当社未同席。確定請負ではない。",
    )
    rows = [
        ["区分", "内容", "見積"],
        ["マスタ", "先方協議では非生鮮と同じ。当社未同席", "本体IFなし。B1で突合"],
        ["波及", "流用が覆ると非生鮮Phase2にも効き得る", "B3に5人日。契約額は据え置き"],
        ["GCS", "生鮮側サーバー→GCS。無い見込み", "確認と必要時のアップロードを含む"],
        ["入荷・日程", "ベンダー流用。PC/CKは別SQLからDBGET", "Aの差分。Layer①に待ち合わせ"],
        ["勧告", "現行の生鮮→統合IFは使わない", "Aで変換IFを新規"],
        ["倉庫", "DC在庫なし。水煮相当は残件", "対象外。要IFなら再見積"],
    ]
    table_slide_data(
        s,
        Inches(0.25),
        Inches(1.3),
        rows,
        [Inches(1.5), Inches(6.2), Inches(5.2)],
        font_size=11,
    )
    note = s.shapes.add_textbox(Inches(0.35), Inches(6.45), Inches(12.5), Inches(0.5))
    add_text(
        note.text_frame,
        "明細は Excel「見積前提条件」シート。",
        size=13,
        bold=True,
        color=ORANGE,
    )
    footer(s, 3, total)

    # ---- 4 構成 ----
    s = blank_slide(prs)
    banner(s, "3. ブロック構成", "Aが本体。B2は0。入荷予定・発注スケジュールにDBGETを加算。")
    rows = [
        ["ブロック", "内訳", "人日", "金額（円）", "備考"],
        ["A（新規／差分IF）", "勧告1.5＋変換1＋入荷2＋日程1", "125.5", "3,782,500", "入荷予定・発注スケジュールにDBGET"],
        ["B1（流用確認・残件）", "突合・倉庫対象外・GCS・水煮", "15", "600,000", "Phase1単価4万"],
        ["B2（マスタ本体IF）", "先方協議では流用。当社未確認", "0", "0", "仮置き"],
        ["B3（非生鮮Phase2波及）", "伝票区分・振替の軽微注記", "5", "137,500", "Phase2直接2.75万"],
        ["", "直接小計（A＋B1＋B2＋B3）", "145.5", "4,520,000", ""],
        ["C（PJ管理按分）", "直接×50／499.5", "15", "600,000", "既存比率"],
        ["", "合計", "160.5", "5,120,000", ""],
    ]
    table_slide_data(
        s,
        Inches(0.4),
        Inches(1.4),
        rows,
        [Inches(3.4), Inches(2.8), Inches(1.1), Inches(1.6), Inches(3.5)],
        font_size=12,
    )
    note = s.shapes.add_textbox(Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.6))
    add_text(
        note.text_frame,
        "B1〜B3は確認・残件・波及。Aだけ見せると残件が見えない。",
        size=13,
        color=DARK,
    )
    footer(s, 4, total)

    # ---- 5 A明細 ----
    s = blank_slide(prs)
    banner(s, "4. A（新規／差分IF）｜勧告は1.5本", "当日1本＋翌日0.5本。発注統合変換は別1本。入荷・日程は差分。S番号は仮採番。")
    rows = [
        ["IF", "Phase1", "Layer①", "Layer③", "計", "備考"],
        ["発注勧告（当日）S24（仮採番）", "5.0", "17.0", "4.5", "26.5", "1.0本"],
        ["発注勧告（翌日以降）S25（仮採番）", "2.5", "8.5", "2.5", "13.5", "当日類似。0.5本"],
        ["勧告→発注統合変換 S26（仮採番）", "5.0", "14.0", "3.0", "22.0", "現行生鮮→統合IFは使わない"],
        ["入荷実績 S11（仮採番）", "4.0", "12.0", "2.5", "18.5", "伝票区分差分＋PC/CK振替"],
        ["入荷予定 S12（仮採番）", "5.0", "15.0", "2.5", "22.5", "PC/CKは別SQL。DBGET＋待ち合わせ"],
        ["発注スケジュール S13（仮採番）", "5.0", "15.0", "2.5", "22.5", "PC/CKは別SQL。DBGET＋待ち合わせ"],
        ["A小計", "26.5", "81.5", "17.5", "125.5", "勧告2ファイル＝1.5本。Layer②＝0"],
    ]
    table_slide_data(
        s,
        Inches(0.3),
        Inches(1.35),
        rows,
        [Inches(4.0), Inches(1.0), Inches(1.0), Inches(1.0), Inches(1.0), Inches(4.6)],
        font_size=11,
    )
    c = card(s, Inches(0.4), Inches(5.5), Inches(12.5), Inches(1.35), LIGHT)
    tf = c.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    add_text(tf, "Aの上振れを大きく載せていない理由", size=13, bold=True, color=NAVY)
    add_para(
        tf,
        "要件の骨格は 9/18 会議。本数は 9/29。"
        "残件は項目精査・伝票区分・GCS、発注統合レイアウト差、入荷予定・発注スケジュールのSQLと作成時刻。覆りは再見積。",
        size=13,
        color=DARK,
        space_before=4,
    )
    footer(s, 5, total)

    # ---- 6 リスク核心 ----
    s = blank_slide(prs)
    banner(s, "5. 流用は未確認", "先方協議では12本流用可。当社未同席。残件はGCS・伝票区分・水煮・SQL作成時刻")
    left = card(s, Inches(0.35), Inches(1.35), Inches(6.2), Inches(5.4), ORANGE_BG)
    tf = left.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    tf.margin_top = Inches(0.15)
    add_text(tf, "判定で決まったこと", size=16, bold=True, color=ORANGE)
    for line in [
        "・店舗系12本は先方協議では非生鮮と同じ（流用）。当社は未同席",
        "・本体IFは置かない（B2＝0）",
        "・勧告ファイル2本は新規。発注統合変換も新規",
        "・入荷・スケジュールはベンダー流用＋差分",
        "・倉庫系はDC在庫なしで対象外",
        "・影響は先行IFに閉じないが、",
        "　流用前提が覆ったときだけ横断が増える",
    ]:
        add_para(tf, line, size=14, color=DARK, space_before=8)

    right = card(s, Inches(6.8), Inches(1.35), Inches(6.1), Inches(5.4), LIGHT)
    tf = right.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    tf.margin_top = Inches(0.15)
    add_text(tf, "見積上どう織り込んだか", size=16, bold=True, color=NAVY)
    for line in [
        "B1（流用確認・残件） 15人日＝突合と残件",
        "　流用12本突合8／倉庫対象外2／",
        "　GCS確認3／水煮残件2",
        "",
        "B2（マスタ本体IF） 0人日",
        "",
        "B3（非生鮮Phase2波及） 5人日＝伝票区分・振替の軽微注記",
        "",
        "※ 覆ったらリスク込み案で再見積",
    ]:
        add_para(tf, line if line else " ", size=14, color=DARK, space_before=5)
    footer(s, 6, total)

    # ---- 7 波及 ----
    s = blank_slide(prs)
    banner(s, "6. B3（非生鮮Phase2波及）", "非生鮮Phase2の契約額は据え置き。手戻りは増分側に計上")
    rows = [
        ["対象", "なぜ効くか", "見積上の扱い"],
        ["店舗系マスタ・実績12本", "先方協議では非生鮮と同じ。当社未確認", "流用仮置き。本体IFなし。B1で突合"],
        ["入荷実績 S11（仮採番）", "伝票区分差分／PC・CK振替", "A 差分"],
        ["入荷予定 S12（仮採番）", "ベンダー同じ／PC・CKは別SQL・DBGET", "A 差分。Layer①に加算"],
        ["発注スケジュール S13（仮採番）", "ベンダー同じ／PC・CKは別SQL・DBGET", "A 差分。Layer①に加算"],
        ["発注勧告 当日・翌日 S24／S25（仮採番）", "2ファイル／1.5本", "当日1本、翌日0.5本"],
        ["勧告→発注統合 S26（仮採番）", "現行生鮮IFは流用不可", "A 新規。参照は非生鮮発注勧告"],
        ["倉庫系", "DC在庫なし", "対象外。水煮相当は午後MT"],
    ]
    table_slide_data(
        s,
        Inches(0.3),
        Inches(1.35),
        rows,
        [Inches(2.8), Inches(5.0), Inches(5.0)],
        font_size=12,
    )
    footer(s, 7, total)

    # ---- 8 ゲート ----
    s = blank_slide(prs)
    banner(s, "7. 進め方と再見積", "提示値は確定請負ではない")
    steps = [
        ("①", "B1（流用確認・残件）", "流用12本の突合\n倉庫対象外の整理"),
        ("②", "A（新規／差分IF）", "差分／新規6本は\nMappingから"),
        ("③", "B2（マスタ本体IF）", "本体IFは0のまま\n覆ったら再見積"),
        ("④", "手戻り", "軽微は B3＝5人日\n超過なら再見積"),
    ]
    for i, (n, t, d) in enumerate(steps):
        c = card(s, Inches(0.35 + i * 3.25), Inches(1.5), Inches(3.05), Inches(2.4), LIGHT)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.15)
        tf.margin_top = Inches(0.15)
        add_text(tf, f"{n} {t}", size=16, bold=True, color=NAVY)
        add_para(tf, d, size=13, color=DARK, space_before=10)

    trig = card(s, Inches(0.35), Inches(4.2), Inches(12.6), Inches(2.5), ORANGE_BG)
    tf = trig.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_top = Inches(0.15)
    add_text(tf, "再見積になるとき", size=15, bold=True, color=ORANGE)
    for line in [
        "・流用12本の一部が差分／新規に戻る",
        "・倉庫商品マスタ（水煮相当）が要IF（午後MT）",
        "・GCSアップロード・接続が想定より重い／伝票区分差分が大きい",
        "・入荷予定・発注スケジュールのSQL接続が重い、生鮮データの作成時刻が待ち合わせでは収まらない",
        "・直接工数が本案想定を大きく超える場合は増分再見積",
    ]:
        add_para(tf, line, size=13, color=DARK, space_before=5)
    footer(s, 8, total)

    # ---- 9 前提除外 ----
    s = blank_slide(prs)
    banner(s, "8. 含むもの／含まないもの", "Phase2は詳細設計・開発・単体まで")
    left = card(s, Inches(0.35), Inches(1.4), Inches(6.2), Inches(5.3), GREEN_BG)
    tf = left.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    add_text(tf, "含む", size=16, bold=True, color=GREEN)
    for line in [
        "・調査用表・①②③ Mapping（Phase1）",
        "・詳細設計・開発・単体（Phase2）",
        "・Aの6本（勧告・変換・入荷・日程）",
        "・流用突合・倉庫対象外整理（B1）",
        "・B2＝0。軽微波及（B3）",
        "・PJ管理按分（C）",
        "・生鮮マスタのGCS確認・作業前提",
    ]:
        add_para(tf, line, size=14, color=DARK, space_before=8)

    right = card(s, Inches(6.8), Inches(1.4), Inches(6.1), Inches(5.3), ORANGE_BG)
    tf = right.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0.05)
    tf.margin_left = Inches(0.22)
    add_text(tf, "含まない", size=16, bold=True, color=ORANGE)
    for line in [
        "・結合・総合・移行・運用",
        "・生鮮基幹／生鮮発注統合本体改修（先方）",
        "・青果@rms／鮮魚市場EOS（自動補充対象外）",
        "・共通基盤の二重構築",
        "・倉庫系（DC在庫なし）",
        "・非生鮮Phase2・共通基盤43・既存PJ50の契約変更",
    ]:
        add_para(tf, line, size=14, color=DARK, space_before=10)
    footer(s, 9, total)

    # ---- 10 話し方 ----
    s = blank_slide(prs)
    banner(s, "9. 話し方", "数字の前に前提。崩れたら再見積。")
    lines = [
        ("導入", "生鮮は非生鮮Phase2の外の増分。提示は160.5人日／512万円（税抜）。"),
        ("前提", "流用は先方協議の仮置き。当社未同席。確定請負ではない。"),
        ("判定", "店舗系12本は流用前提でB1突合。勧告2ファイルは1.5本。発注統合変換は新規。入荷・日程は差分。"),
        ("覆り", "流用が覆ったら本日の数字は無効。再見積。"),
        ("B2", "マスタ本体IFは0。水煮相当は午後MTの残件。"),
        ("進め方", "12本は突合、6本はMappingから。B2は0のまま。"),
        ("範囲", "Phase2は詳細設計・開発・単体まで。結合以降は含まない。"),
    ]
    y = 1.35
    for title_, body in lines:
        lab = s.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.35), Inches(y), Inches(1.7), Inches(0.72)
        )
        lab.fill.solid()
        lab.fill.fore_color.rgb = NAVY
        lab.line.fill.background()
        lab.adjustments[0] = 0.2
        add_text(lab.text_frame, title_, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        lab.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        bx = s.shapes.add_textbox(Inches(2.2), Inches(y), Inches(10.7), Inches(0.72))
        add_text(bx.text_frame, body, size=13, color=DARK)
        y += 0.85
    footer(s, 10, total)

    # ---- 11 まとめ ----
    s = blank_slide(prs)
    banner(s, "10. まとめ", "")
    items = [
        ("① 提示値", "160.5人日／¥5,120,000（税抜）\n非生鮮Phase2は据え置き"),
        ("② 前提", "先方協議の流用は未確認\nB1で突合。覆ったら再見積"),
        ("③ 進め方", "12本突合／6本Mapping／B2＝0\n超過時は再見積"),
    ]
    for i, (h, b) in enumerate(items):
        c = card(s, Inches(0.4 + i * 4.2), Inches(1.5), Inches(4.0), Inches(2.6), LIGHT)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.2)
        tf.margin_top = Inches(0.2)
        add_text(tf, h, size=16, bold=True, color=NAVY)
        add_para(tf, b, size=14, color=DARK, space_before=12)

    end = card(s, Inches(0.4), Inches(4.4), Inches(12.5), Inches(2.3), ORANGE_BG)
    tf = end.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_top = Inches(0.2)
    add_text(tf, "添付・参照", size=14, bold=True, color=ORANGE)
    for line in [
        "・【見積】生鮮IF増分_標準案_…_20260921.xlsx",
        "・20260929_sinops生鮮IF_マスタ部門判定.xlsx",
        "・本フォルダ README.md",
    ]:
        add_para(tf, line, size=13, color=DARK, space_before=4)
    footer(s, 11, total)

    prs.save(OUT)
    print(f"saved: {OUT}")
    return OUT


if __name__ == "__main__":
    build()
