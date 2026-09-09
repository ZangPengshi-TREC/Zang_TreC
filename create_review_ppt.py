#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""発注勧告・受払明細 Review会 PPT 生成（美化版）"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUTPUT = "/Users/treqd/Desktop/Cursor/西友PMI/発注勧告・受払明細_Review会資料.pptx"

# ── カラーパレット ──
C_NAVY      = RGBColor(0x0D, 0x2B, 0x4E)   # 深紺（メイン）
C_BLUE      = RGBColor(0x1A, 0x5C, 0x9E)   # 青（アクセント）
C_TEAL      = RGBColor(0x00, 0x96, 0x88)   # ティール
C_PURPLE    = RGBColor(0x6A, 0x3D, 0x9A)   # 紫（本部系）
C_ORANGE    = RGBColor(0xE8, 0x6C, 0x00)   # オレンジ（警告）
C_RED       = RGBColor(0xD3, 0x2F, 0x2F)   # 赤（確認事項）
C_GREEN     = RGBColor(0x2E, 0x7D, 0x32)   # 緑（確認済）
C_GOLD      = RGBColor(0xF5, 0xA6, 0x23)   # ゴールド（最高優先度）
C_LIGHT     = RGBColor(0xF0, 0xF4, 0xF8)   # 薄グレー背景
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
C_BLACK     = RGBColor(0x21, 0x21, 0x21)
C_GRAY      = RGBColor(0x61, 0x61, 0x61)
C_LGRAY     = RGBColor(0x9E, 0x9E, 0x9E)
C_SKY       = RGBColor(0xB3, 0xD4, 0xFC)
C_ROW_ALT   = RGBColor(0xE8, 0xF0, 0xFE)   # テーブル交互行

FONT = "Meiryo UI"

_slide_counter = [0]


def _next_slide_num():
    _slide_counter[0] += 1
    return _slide_counter[0]


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape(slide, shape_type, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, left, top, width, height, text, size=14,
                bold=False, color=C_BLACK, align=PP_ALIGN.LEFT, font=FONT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.name = font
    p.font.color.rgb = color
    p.alignment = align
    return box


def add_rich_text(slide, left, top, width, height, runs, align=PP_ALIGN.LEFT):
    """runs: list of (text, size, bold, color)"""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    for i, (text, size, bold, color) in enumerate(runs):
        if i == 0 and p.runs:
            run = p.runs[0]
        else:
            run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = FONT
        run.font.color.rgb = color
    return box


def add_bullet_box(slide, left, top, width, height, lines, size=13,
                   color=C_BLACK, bullet="▸", spacing=6):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{bullet}  {line}" if bullet else line
        p.font.size = Pt(size)
        p.font.name = FONT
        p.font.color.rgb = color
        p.space_after = Pt(spacing)
    return box


def add_card(slide, left, top, width, height, title, accent_color=C_BLUE):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height,
              C_WHITE, RGBColor(0xDE, 0xE2, 0xE6))
    add_shape(slide, MSO_SHAPE.RECTANGLE, left, top, Inches(0.06), height, accent_color)
    if title:
        add_textbox(slide, left + Inches(0.2), top + Inches(0.08),
                    width - Inches(0.3), Inches(0.35),
                    title, size=13, bold=True, color=accent_color)


def add_footer(slide, section="", slide_num=None):
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.15), Inches(10), Inches(0.35), C_NAVY)
    if section:
        add_textbox(slide, Inches(0.4), Inches(7.17), Inches(6), Inches(0.3),
                    section, size=9, color=C_SKY)
    if slide_num is not None:
        add_textbox(slide, Inches(8.5), Inches(7.17), Inches(1.2), Inches(0.3),
                    str(slide_num), size=9, color=C_SKY, align=PP_ALIGN.RIGHT)


def add_table(slide, left, top, width, rows_data, col_widths=None,
              header=True, font_size=10, accent_color=C_BLUE, alt_rows=True):
    n_rows = len(rows_data)
    n_cols = len(rows_data[0]) if rows_data else 0
    if n_rows == 0 or n_cols == 0:
        return None
    row_h = Inches(0.34)
    tbl = slide.shapes.add_table(n_rows, n_cols, left, top, width, row_h * n_rows).table
    if col_widths:
        for i, w in enumerate(col_widths):
            tbl.columns[i].width = w
    for r, row in enumerate(rows_data):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(font_size)
                p.font.name = FONT
                if header and r == 0:
                    p.font.bold = True
                    p.font.color.rgb = C_WHITE
                else:
                    p.font.color.rgb = C_BLACK
            if header and r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = accent_color
            elif alt_rows and r % 2 == 0 and r > 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_ROW_ALT
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Pt(6)
            cell.margin_right = Pt(4)
    return tbl


def slide_header(slide, title, subtitle="", accent_color=C_BLUE, section=""):
    set_bg(slide, C_LIGHT)
    # ヘッダーバー
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1.25), C_NAVY)
    add_shape(slide, MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.25), Inches(10), Inches(0.05), accent_color)
    # タイトル
    add_textbox(slide, Inches(0.5), Inches(0.18), Inches(9), Inches(0.55),
                title, size=26, bold=True, color=C_WHITE)
    if subtitle:
        add_textbox(slide, Inches(0.5), Inches(0.72), Inches(9), Inches(0.4),
                    subtitle, size=12, color=C_SKY)
    num = _next_slide_num()
    add_footer(slide, section, num)


def slide_confirm(slide, items, top=Inches(5.6)):
    card_h = Inches(0.35) + Inches(0.32) * len(items) + Inches(0.15)
    add_card(slide, Inches(0.4), top, Inches(9.2), card_h, "本日の確認事項", C_RED)
    y = top + Inches(0.45)
    for i, item in enumerate(items, 1):
        add_rich_text(slide, Inches(0.65), y, Inches(8.7), Inches(0.3), [
            (f"{i}.  ", 11, True, C_RED),
            (item, 11, False, C_GRAY),
        ])
        y += Inches(0.32)


def add_badge(slide, left, top, text, bg_color, text_color=C_WHITE, width=Inches(0.9)):
    h = Inches(0.28)
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, h, bg_color)
    add_textbox(slide, left, top + Inches(0.02), width, h,
                text, size=9, bold=True, color=text_color, align=PP_ALIGN.CENTER)


def add_highlight_box(slide, left, top, width, height, text, icon="✓", bg=C_GREEN):
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height,
              RGBColor(0xE8, 0xF5, 0xE9), bg)
    add_rich_text(slide, left + Inches(0.15), top + Inches(0.08), width - Inches(0.2), height - Inches(0.1), [
        (f"{icon}  ", 12, True, bg),
        (text, 12, False, C_BLACK),
    ])


def add_section_divider(prs, blank, title, subtitle, color):
    s = prs.slides.add_slide(blank)
    set_bg(s, color)
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(3.2), Inches(10), Inches(0.06), C_GOLD)
    add_textbox(s, Inches(0.8), Inches(2.0), Inches(8.5), Inches(1.0),
                title, size=34, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.8), Inches(3.5), Inches(8.5), Inches(0.6),
                subtitle, size=16, color=C_SKY, align=PP_ALIGN.CENTER)
    _next_slide_num()


def build_presentation():
    global _slide_counter
    _slide_counter = [0]

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ══════════ 表紙 ══════════
    s = prs.slides.add_slide(blank)
    set_bg(s, C_NAVY)
    # 装飾ライン
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(2.8), Inches(10), Inches(0.06), C_GOLD)
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.6), Inches(10), Inches(0.04), C_BLUE)
    # 左上アクセント
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.12), Inches(7.5), C_TEAL)
    add_textbox(s, Inches(0.8), Inches(1.5), Inches(8.5), Inches(0.5),
                "TRIAL  ↔  sinops", size=18, color=C_SKY, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.8), Inches(2.0), Inches(8.5), Inches(0.9),
                "IF連携 Review会", size=40, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.8), Inches(3.1), Inches(8.5), Inches(0.6),
                "項目マッピング表・GAP分析書", size=20, color=C_SKY, align=PP_ALIGN.CENTER)
    # 対象IFカード
    if_items = [
        ("①", "受払明細", "uke.txt", C_TEAL),
        ("②", "店舗系発注勧告", "kankoku.txt / subkankoku.txt", C_BLUE),
        ("③", "倉庫系発注勧告", "ka.txt / subka.txt", C_PURPLE),
    ]
    x_start = Inches(0.9)
    for i, (num, name, files, color) in enumerate(if_items):
        x = x_start + Inches(2.95) * i
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(4.0), Inches(2.7), Inches(1.3),
                  RGBColor(0x15, 0x3A, 0x5E))
        add_shape(s, MSO_SHAPE.OVAL, x + Inches(0.15), Inches(4.15), Inches(0.45), Inches(0.45), color)
        add_textbox(s, x + Inches(0.15), Inches(4.2), Inches(0.45), Inches(0.35),
                    num, size=14, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, x + Inches(0.7), Inches(4.15), Inches(1.9), Inches(0.35),
                    name, size=12, bold=True, color=C_WHITE)
        add_textbox(s, x + Inches(0.7), Inches(4.55), Inches(1.9), Inches(0.6),
                    files, size=9, color=C_SKY)
    _next_slide_num()

    # ══════════ アジェンダ ══════════
    s = prs.slides.add_slide(blank)
    slide_header(s, "アジェンダ", section="Review会")
    agenda = [
        ("01", "全体概要", "対象IF・資料構成・判定凡例", C_NAVY),
        ("02", "資料1", "受払明細 項目マッピング表", C_TEAL),
        ("03", "資料2", "受払明細 GAP分析書", C_TEAL),
        ("04", "資料3", "発注勧告 項目マッピング表（店舗系）", C_BLUE),
        ("05", "資料4", "発注勧告 GAP分析書（店舗系）", C_BLUE),
        ("06", "資料5", "発注勧告 項目マッピング表（倉庫系）", C_PURPLE),
        ("07", "資料6", "発注勧告 GAP分析書（倉庫系）", C_PURPLE),
        ("08", "まとめ", "確認事項・次ステップ・Q&A", C_ORANGE),
    ]
    y = Inches(1.55)
    for num, tag, desc, color in agenda:
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), y, Inches(9.0), Inches(0.62),
                  C_WHITE, RGBColor(0xDE, 0xE2, 0xE6))
        add_shape(s, MSO_SHAPE.OVAL, Inches(0.65), y + Inches(0.1), Inches(0.42), Inches(0.42), color)
        add_textbox(s, Inches(0.65), y + Inches(0.13), Inches(0.42), Inches(0.35),
                    num, size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.25), y + Inches(0.08), Inches(1.5), Inches(0.3),
                    tag, size=13, bold=True, color=color)
        add_textbox(s, Inches(2.8), y + Inches(0.1), Inches(6.5), Inches(0.4),
                    desc, size=12, color=C_GRAY)
        y += Inches(0.7)

    # ══════════ 全体概要 ══════════
    s = prs.slides.add_slide(blank)
    slide_header(s, "全体概要", "6資料の位置づけと判定凡例", section="Review会")
    add_card(s, Inches(0.4), Inches(1.45), Inches(4.4), Inches(2.0), "資料構成", C_BLUE)
    add_table(s, Inches(0.55), Inches(1.85), Inches(4.1), [
        ["資料", "内容", "対象"],
        ["① マッピング表", "項目1対1対応", "各IF"],
        ["② GAP分析書", "差異・優先度", "各IF"],
    ], col_widths=[Inches(1.4), Inches(1.5), Inches(1.2)], font_size=10)
    add_card(s, Inches(5.1), Inches(1.45), Inches(4.5), Inches(2.0), "判定凡例", C_TEAL)
    legend = [
        ("○", "直接マッピング可", C_GREEN),
        ("△", "変換・加工が必要", C_ORANGE),
        ("✕", "GAPあり（要対応）", C_RED),
        ("―", "固定値 / NULL", C_LGRAY),
    ]
    ly = Inches(1.9)
    for sym, meaning, color in legend:
        add_textbox(s, Inches(5.3), ly, Inches(0.4), Inches(0.3),
                    sym, size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(5.8), ly, Inches(3.5), Inches(0.3),
                    meaning, size=11, color=C_GRAY)
        ly += Inches(0.38)
    add_card(s, Inches(0.4), Inches(3.65), Inches(9.2), Inches(1.8), "横断的な重要論点", C_RED)
    add_bullet_box(s, Inches(0.6), Inches(4.05), Inches(8.8), Inches(1.3), [
        "商品コード型矛盾：13桁JAN > int32最大値  →  int64変更要",
        "数量変換：sets = バラ数量 × 1000",
        "原単価変換：cost = 単価 × 1000",
        "発注日：TRIAL本体に「発注日」なし  →  登録日で代替？",
    ], size=12, color=C_BLACK, bullet="●")

    # ══════════ セクション：受払明細（資料1・2） ══════════
    add_section_divider(prs, blank, "資料1・2", "受払明細IF（uke.txt）", C_TEAL)

    # 資料1
    s = prs.slides.add_slide(blank)
    slide_header(s, "資料1：項目マッピング表（受払明細）",
                 "【倉庫CD】uke.txt（10項目）",
                 accent_color=C_TEAL, section="受払明細 ① マッピング表")
    add_textbox(s, Inches(0.5), Inches(1.38), Inches(9), Inches(0.3),
                "検証データ：007452uke.txt（2,356行）実データ確認済",
                size=11, color=C_LGRAY)
    add_card(s, Inches(0.4), Inches(1.7), Inches(4.5), Inches(2.3), "サンプル検証の発見", C_TEAL)
    findings = [
        ("1", "正値出力（負値0件）", "×(-1)は不要", C_GREEN),
        ("2", "入庫/売上は非排他", "排他想定は誤り", C_ORANGE),
        ("3", "処理日≠発行日", "43.2%で相異", C_ORANGE),
    ]
    fy = Inches(2.1)
    for num, title, note, color in findings:
        add_shape(s, MSO_SHAPE.OVAL, Inches(0.6), fy, Inches(0.35), Inches(0.35), color)
        add_textbox(s, Inches(0.6), fy + Inches(0.04), Inches(0.35), Inches(0.28),
                    num, size=11, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.1), fy, Inches(3.5), Inches(0.28), title, size=11, bold=True, color=C_BLACK)
        add_textbox(s, Inches(1.1), fy + Inches(0.28), Inches(3.5), Inches(0.25), note, size=9, color=C_GRAY)
        fy += Inches(0.65)
    add_card(s, Inches(5.1), Inches(1.7), Inches(4.5), Inches(2.3), "出力パターン", C_BLUE)
    add_table(s, Inches(5.25), Inches(2.1), Inches(4.2), [
        ["パターン", "行数", "比率"],
        ["出庫のみ → 売上数", "2,233", "94.8%"],
        ["入庫のみ → 入庫数", "14", "0.6%"],
        ["両方設定", "32", "1.4%"],
        ["両方0", "77", "3.3%"],
    ], col_widths=[Inches(2.0), Inches(1.0), Inches(1.2)], font_size=10, accent_color=C_BLUE)
    slide_confirm(s, [
        "正値出力・非排他性の承認",
        "処理日≠発行日時の発行日設定ルール",
        "両方0レコードの出力要否",
    ], top=Inches(4.3))

    # 資料2
    s = prs.slides.add_slide(blank)
    slide_header(s, "資料2：GAP分析書（受払明細）",
                 "【倉庫CD】uke.txt ｜ 全6件",
                 accent_color=C_TEAL, section="受払明細 ② GAP分析書")
    add_table(s, Inches(0.4), Inches(1.55), Inches(9.2), [
        ["優先度", "GAP", "ステータス", "要点"],
        ["★最高", "①符号変換", "確認済", "×(-1)不要、正値出力"],
        ["★高", "②非排他性", "確認済", "32行で両方に値"],
        ["★中", "③日付マッピング", "要確認", "43.2%で処理日≠発行日"],
        ["★中", "④伝票区分フィルタ", "部分確認", "廃棄(19)未検出"],
        ["★低", "⑤固定値項目", "確認済", "全5項目 = 0"],
        ["★中", "⑥両方0レコード", "要確認", "77行、除外対象？"],
    ], col_widths=[Inches(0.85), Inches(2.0), Inches(1.15), Inches(5.2)],
       font_size=10, accent_color=C_TEAL)
    add_card(s, Inches(0.4), Inches(4.25), Inches(9.2), Inches(1.1), "旧資料からの訂正", C_RED)
    add_bullet_box(s, Inches(0.6), Inches(4.65), Inches(8.8), Inches(0.6), [
        "納品数×(-1)  →  正値のまま出力（訂正）",
        "入庫/売上は排他的  →  非排他的に両方設定可（訂正）",
    ], size=11, color=C_BLACK, bullet="⟳")

    # ══════════ セクション：店舗系（資料3・4） ══════════
    add_section_divider(prs, blank, "資料3・4", "発注勧告IF — 店舗系（kankoku / subkankoku）", C_BLUE)

    # 資料3
    s = prs.slides.add_slide(blank)
    slide_header(s, "資料3：項目マッピング表（店舗系）",
                 "kankoku.txt（当日）/ subkankoku.txt（翌日以降）",
                 accent_color=C_BLUE, section="店舗系 ① マッピング表")
    add_textbox(s, Inches(0.5), Inches(1.38), Inches(9), Inches(0.3),
                "sinops R6版（15項目）↔ TRIAL数理発注データの項目対応定義",
                size=11, color=C_LGRAY)
    add_table(s, Inches(0.4), Inches(1.75), Inches(9.2), [
        ["区分", "項目", "判定", "変換ルール"],
        ["Key(5)", "店舗・商品・納品日・便・仕入先", "○", "int(x) で文字→int"],
        ["Key", "発注日", "△", "登録日(reg_date)で代替？"],
        ["数量", "発注バラ数 → sets", "○", "sets = バラ数 × 1000"],
        ["単価", "標準仕入単価 → cost", "○", "cost = 単価 × 1000"],
        ["sinops余剰", "大/中/小/細目分類", "―", "TRIAL側に対応なし"],
        ["TRIAL余剰", "発注種別・納品先等 7項目", "△", "設定方針要決定"],
    ], col_widths=[Inches(1.2), Inches(2.6), Inches(0.65), Inches(4.75)],
       font_size=10, accent_color=C_BLUE)
    slide_confirm(s, [
        "登録日＝発注日としてよいか？",
        "分類コード4項目は連携不要でよいか？",
        "TRIAL余剰7項目の設定方針（固定値 / マスタ参照）",
    ])

    # 資料4
    s = prs.slides.add_slide(blank)
    slide_header(s, "資料4：GAP分析書（店舗系）",
                 "kankoku.txt / subkankoku.txt ｜ 全8件",
                 accent_color=C_BLUE, section="店舗系 ② GAP分析書")
    gap_data = [
        ["★最高", "①商品コード型矛盾", "確認済", "13桁JAN > int32"],
        ["★最高", "④数量合成ロジック", "未確認", "sets = バラ数 × 1000"],
        ["★高", "②発注日データソース", "未確認", "登録日 = 発注日？"],
        ["★中", "③分類コード余剰", "未確認", "sinops側4項目"],
        ["★中", "⑧TRIAL余剰7項目", "未確認", "設定方針要決定"],
        ["★低", "⑤⑥⑦ 型・日付変換", "技術対応可", "実装時対応"],
    ]
    add_table(s, Inches(0.4), Inches(1.55), Inches(9.2), [
        ["優先度", "GAP", "ステータス", "要点"],
        *gap_data,
    ], col_widths=[Inches(0.85), Inches(2.3), Inches(1.05), Inches(5.0)],
       font_size=10, accent_color=C_BLUE)
    slide_confirm(s, [
        "item_id型を int64 へ変更可能か？",
        "sets計算式の業務妥当性",
        "TRIAL余剰7項目の初期値方針",
    ])

    # ══════════ セクション：倉庫系（資料5・6） ══════════
    add_section_divider(prs, blank, "資料5・6", "発注勧告IF — 倉庫系（ka / subka）", C_PURPLE)

    # 資料5
    s = prs.slides.add_slide(blank)
    slide_header(s, "資料5：項目マッピング表（倉庫系）",
                 "【倉庫ｺｰﾄﾞ】ka.txt（当日）/ subka.txt（翌日以降）",
                 accent_color=C_PURPLE, section="倉庫系 ① マッピング表")
    add_textbox(s, Inches(0.5), Inches(1.38), Inches(9), Inches(0.3),
                "sinops W版（使用14項目）↔ TRIAL数理発注データの項目対応定義",
                size=11, color=C_LGRAY)
    add_table(s, Inches(0.4), Inches(1.75), Inches(9.2), [
        ["項目", "店舗系(R6版)", "倉庫系(W版)", "判定"],
        ["Key構成", "店舗+発注日+商品+納品日+便", "仕入先+品番+納品日", "―"],
        ["発注日", "文字(8)", "発行日・整数(8)", "○"],
        ["数量", "発注バラ数+発注単位数", "注文数量+発注ロット+バラ換算", "○"],
        ["店舗・便", "sinops側あり", "sinops側なし", "△"],
        ["発注区分", "固定値 1", "固定値 10", "○"],
    ], col_widths=[Inches(1.2), Inches(2.9), Inches(2.9), Inches(0.7)],
       font_size=10, accent_color=C_PURPLE)
    add_highlight_box(s, Inches(0.4), Inches(4.15), Inches(9.2), Inches(0.55),
                      "バラ換算数量 = 注文数量 × 発注ロット  →  sets = バラ換算数量 × 1000"
                      "　（サンプル348行で100%成立）", "★", C_GREEN)
    slide_confirm(s, [
        "店舗コード → 倉庫コード？固定値？",
        "便区分 → 固定値0でよいか？",
        "発注区分=10（本部発注）で確定してよいか？",
    ], top=Inches(5.0))

    # 資料6
    s = prs.slides.add_slide(blank)
    slide_header(s, "資料6：GAP分析書（倉庫系）",
                 "【倉庫ｺｰﾄﾞ】ka.txt / subka.txt ｜ 全8件",
                 accent_color=C_PURPLE, section="倉庫系 ② GAP分析書")
    add_table(s, Inches(0.4), Inches(1.55), Inches(9.2), [
        ["優先度", "GAP", "店舗系との関係", "ステータス"],
        ["★最高", "①商品コード型矛盾", "同一問題", "確認済"],
        ["★最高", "④数量変換ロジック", "式異なるが×1000共通", "検証済"],
        ["★高", "②発行日→登録日", "W版は整数型", "未確認"],
        ["★中", "⑧TRIAL余剰10項目", "W版固有", "未確認"],
        ["★低", "⑤⑥⑦", "技術対応可", "―"],
    ], col_widths=[Inches(0.85), Inches(2.1), Inches(2.6), Inches(1.0)],
       font_size=10, accent_color=C_PURPLE)
    add_card(s, Inches(0.4), Inches(4.05), Inches(9.2), Inches(1.35), "W版固有の論点", C_ORANGE)
    add_bullet_box(s, Inches(0.6), Inches(4.45), Inches(8.8), Inches(0.9), [
        "sinopsに店舗・便が存在しない → TRIAL PKとして設定要",
        "納品日はsinops側で計算済み（サーバー日+調達期間+休日）",
        "発注区分は 10（本部発注）で設定予定",
    ], size=11, color=C_BLACK)
    slide_confirm(s, [
        "店舗コード・便区分の固定値ルール",
        "納品予定日＝sinops納品日で問題ないか",
    ], top=Inches(5.55))

    # ══════════ 確認事項まとめ ══════════
    s = prs.slides.add_slide(blank)
    slide_header(s, "確認事項まとめ", "本日のReviewで方針確定をお願いしたい項目",
                 accent_color=C_RED, section="まとめ")
    add_table(s, Inches(0.4), Inches(1.55), Inches(9.2), [
        ["#", "区分", "確認事項", "優先度"],
        ["1", "受払", "正値出力・非排他性の承認", "★最高"],
        ["2", "受払", "処理日≠発行日の扱い", "★中"],
        ["3", "受払", "両方0レコードの出力要否", "★中"],
        ["4", "共通", "商品コード型（int32 → int64）", "★最高"],
        ["5", "共通", "登録日＝発注日の可否", "★高"],
        ["6", "共通", "sets / cost 変換ルールの最終確認", "★最高"],
        ["7", "店舗系", "分類コード4項目の要否", "★中"],
        ["8", "店舗系", "TRIAL余剰7項目の設定方針", "★中"],
        ["9", "倉庫系", "店舗コード・便区分の固定値", "★高"],
    ], col_widths=[Inches(0.45), Inches(0.9), Inches(5.7), Inches(0.85)],
       font_size=10, accent_color=C_RED)

    # ══════════ 次ステップ ══════════
    s = prs.slides.add_slide(blank)
    slide_header(s, "次ステップ", section="まとめ")
    steps = [
        ("STEP 1", "本日確認事項の方針確定", "特に★最高・★高のGAP", C_RED),
        ("STEP 2", "マッピング表・GAP分析書の更新", "確認結果の反映、個別シート統一", C_BLUE),
        ("STEP 3", "変換ロジック仕様書の作成", "確定した変換ルールの詳細化", C_TEAL),
        ("STEP 4", "実装設計・開発着手", "商品コード型変更はTRIAL側対応が必要", C_GREEN),
    ]
    y = Inches(1.6)
    for step, title, sub, color in steps:
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), y, Inches(9.0), Inches(1.1),
                  C_WHITE, RGBColor(0xDE, 0xE2, 0xE6))
        add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0.5), y, Inches(0.08), Inches(1.1), color)
        add_textbox(s, Inches(0.75), y + Inches(0.1), Inches(1.2), Inches(0.3),
                    step, size=10, bold=True, color=color)
        add_textbox(s, Inches(0.75), y + Inches(0.4), Inches(8.5), Inches(0.35),
                    title, size=15, bold=True, color=C_BLACK)
        add_textbox(s, Inches(0.75), y + Inches(0.72), Inches(8.5), Inches(0.3),
                    sub, size=11, color=C_GRAY)
        y += Inches(1.25)

    # ══════════ Q&A ══════════
    add_section_divider(prs, blank, "Q&A", "想定質問と回答", C_NAVY)

    qa_pairs = [
        # 受払明細
        ("Q12", "入庫数・売上数は正値でよいか？",
         "全2,356行で負値0件。旧資料の×(-1)は誤り。", C_GREEN),
        ("Q13", "入庫・売上が両方入るケースは？",
         "32行(1.4%)で確認。非排他的に両方設定可能。", C_TEAL),
        ("Q14", "処理日と発行日が異なるケースは？",
         "43.2%で相異。sinops側が期待する発行日の定義を要確認。", C_ORANGE),
        ("Q15", "両方0のレコードは出力すべきか？",
         "在庫変動なし。出力/除外/発行日=00000000の要否を確認。", C_ORANGE),
        # 店舗系・共通
        ("Q1", "登録日を発注日として使ってよいか？",
         "TRIAL本体に「発注日」フィールドなし。登録日を暫定マッピング。\n"
         "発注指示日かDB登録日かの意味確認が必要。", C_ORANGE),
        ("Q2", "13桁JANがint32に入らない問題は？",
         "int64への型変更を推奨。TRIAL側DB仕様修正の可否を確認。", C_RED),
        ("Q3", "sets = バラ数 × 1000 は正しいか？",
         "TRIAL setsは1000倍格納。サンプルデータで検証済み。", C_GREEN),
        ("Q4", "大/中/小/細目の分類コードは必要か？",
         "TRIAL DIFPRODUCTSに対応なし。sinops連携で必須か要確認。", C_ORANGE),
        ("Q5", "TRIAL余剰7項目はどう設定するか？",
         "発注区分=1（確度高）、発注種別・納品先等は要確認。", C_ORANGE),
        # 倉庫系
        ("Q7", "倉庫系発注の店舗コードは？",
         "倉庫コード設定 / 固定値 / マスタ参照のいずれか要決定。", C_PURPLE),
        ("Q8", "便区分（mail）はどう設定するか？",
         "固定値0を想定。業務上の正しい設定値を要確認。", C_PURPLE),
        ("Q9", "W版の数量計算式は店舗系と違うのか？",
         "項目構成は異なるが sets=バラ換算数量×1000 は共通。348行検証済。", C_GREEN),
        ("Q20", "次のステップは？",
         "方針確定 → 資料更新 → 変換ロジック仕様書 → 実装着手", C_BLUE),
    ]
    for qid, q, a, color in qa_pairs:
        s = prs.slides.add_slide(blank)
        slide_header(s, f"Q&A  {qid}", section="Q&A")
        # 質問カード
        add_card(s, Inches(0.4), Inches(1.45), Inches(9.2), Inches(0.85), "", color)
        add_shape(s, MSO_SHAPE.OVAL, Inches(0.6), Inches(1.6), Inches(0.5), Inches(0.5), color)
        add_textbox(s, Inches(0.6), Inches(1.65), Inches(0.5), Inches(0.4),
                    "Q", size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.25), Inches(1.65), Inches(8.2), Inches(0.55),
                    q, size=15, bold=True, color=C_BLACK)
        # 回答カード
        add_card(s, Inches(0.4), Inches(2.5), Inches(9.2), Inches(1.6), "回答", C_TEAL)
        add_textbox(s, Inches(0.65), Inches(2.95), Inches(8.7), Inches(1.0),
                    a, size=14, color=C_BLACK)

    # ══════════ 終了 ══════════
    s = prs.slides.add_slide(blank)
    set_bg(s, C_NAVY)
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(3.4), Inches(10), Inches(0.06), C_GOLD)
    add_textbox(s, Inches(0.8), Inches(2.5), Inches(8.5), Inches(1),
                "ご清聴ありがとうございました", size=34, bold=True,
                color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.8), Inches(3.7), Inches(8.5), Inches(0.6),
                "ご確認・ご指摘をいただければ、次回までに資料を更新いたします。",
                size=15, color=C_SKY, align=PP_ALIGN.CENTER)
    _next_slide_num()

    return prs


if __name__ == "__main__":
    prs = build_presentation()
    prs.save(OUTPUT)
    print(f"生成完了: {OUTPUT}")
    print(f"スライド数: {len(prs.slides)}")
