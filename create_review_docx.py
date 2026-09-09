#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""口播稿 Word 生成（括号假名注音・WPS対応）"""

import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT = "/Users/treqd/Desktop/Cursor/西友PMI/発注勧告・受払明細_Review会_説明话术.docx"
INPUT_MD = "/Users/treqd/Desktop/Cursor/西友PMI/発注勧告・受払明細_Review会_説明话术.md"

FURIGANA = {
    "数理発注システム": "すうりはっちゅうシステム",
    "数理発注データ": "すうりはっちゅうデータ",
    "項目マッピング表": "こうもくマッピングひょう",
    "発注勧告": "はっちゅうかんこく",
    "数理発注": "すうりはっちゅう",
    "受払明細": "うけばらいめいさい",
    "対応付け": "たいおうづけ",
    "未解決事項": "みかいけつじこう",
    "妥当性確認": "だとうせいかくにん",
    "方針決定": "ほうしんけってい",
    "出力タイミング": "しゅつりょくタイミング",
    "固定値設定": "こていちせってい",
    "複合Key": "ふくごうキー",
    "納品予定日": "のうひんよていび",
    "仕入先コード": "しいれさきコード",
    "追加フィールド": "ついかフィールド",
    "登録日": "とうろくび",
    "発注日": "はっちゅうび",
    "発注バラ数": "はっちゅうバラすう",
    "発注単位数": "はっちゅうたんいすう",
    "標準仕入単価": "ひょうじゅんしいれたんか",
    "原単価": "げんたんか",
    "分類コード": "ぶんるいコード",
    "設定方針": "せっていほうしん",
    "発注区分": "はっちゅうくぶん",
    "発注種別": "はっちゅうしゅべつ",
    "納品先センターコード": "のうひんさきセンターコード",
    "納品経路": "のうひんけいろ",
    "余剰項目": "よじょうこうもく",
    "商品コード": "しょうひんコード",
    "型矛盾": "がたむじゅん",
    "合成ロジック": "ごうせいロジック",
    "変換ロジック": "へんかんロジック",
    "データソース": "データソース",
    "確認事項": "かくにんじこう",
    "優先度": "ゆうせんど",
    "業務判断": "ぎょうむはんだん",
    "実装範囲": "じっそうはんい",
    "型定義": "がたていぎ",
    "変換ルール": "へんかんルール",
    "参考情報": "さんこうじょうほう",
    "最終確認": "さいしゅうかくにん",
    "固定値方針": "こていちほうしん",
    "倉庫コード": "そうこコード",
    "店舗コード": "てんぽコード",
    "便区分": "びんくぶん",
    "発行日": "はっこうび",
    "発注ロット": "はっちゅうロット",
    "バラ換算数量": "バラかんさんすうりょう",
    "注文数量": "ちゅうもんすうりょう",
    "調達期間": "ちょうたつきかん",
    "納品日": "のうひんび",
    "数量計算式": "すうりょうけいさんしき",
    "固有の論点": "こゆうのろんてん",
    "設定ルール": "せっていルール",
    "実データ": "じつデータ",
    "出力パターン": "しゅつりょくパターン",
    "符号変換": "ふごうへんかん",
    "非排他性": "ひはいたせい",
    "排他的想定": "はいたてきそうてい",
    "排他的": "はいたてき",
    "非排他的": "ひはいたてき",
    "処理日": "しょりび",
    "日付マッピング": "ひづけマッピング",
    "伝票区分": "でんぴょうくぶん",
    "振替伝票": "ふりかえでんぴょう",
    "仕入伝票": "しいれでんぴょう",
    "物流タイプ": "ぶつりゅうタイプ",
    "固定値項目": "こていちこうもく",
    "在庫変動": "ざいこへんどう",
    "需要予測": "じゅようよそく",
    "除外対象": "じょがいたいしょう",
    "正値出力": "せいちしゅつりょく",
    "得意先コード": "とくいさきコード",
    "旧資料": "きゅうしりょう",
    "成果物": "せいかぶつ",
    "最終版": "さいしゅうばん",
    "会議進行": "かいぎしんこう",
    "所要時間": "しょうようじかん",
    "判定凡例": "はんていはんれい",
    "直接マッピング": "ちょくせつマッピング",
    "変換・加工": "へんかん・かこう",
    "店舗系": "てんぽけい",
    "本部系": "ほんぶけい",
    "倉庫系": "そうこけい",
    "当日分": "とうじつぶん",
    "翌日以降分": "よくじついこうぶん",
    "納品先": "のうひんさき",
    "発注指示日": "はっちゅうしじび",
    "細目": "さいもく",
    "格納": "かくのう",
    "換算": "かんさん",
    "検証済": "けんしょうずみ",
    "検証": "けんしょう",
    "妥当性": "だとうせい",
    "連携": "れんけい",
    "対象": "たいしょう",
    "項目": "こうもく",
    "明細": "めいさい",
    "店舗": "てんぽ",
    "本部": "ほんぶ",
    "倉庫": "そうこ",
    "変換": "へんかん",
    "判定": "はんてい",
    "確認": "かくにん",
    "方針": "ほうしん",
    "決定": "けってい",
    "出力": "しゅつりょく",
    "管理": "かんり",
    "目的": "もくてき",
    "明確": "めいかく",
    "凡例": "はんれい",
    "直接": "ちょくせつ",
    "加工": "かこう",
    "固定値": "こていち",
    "設定": "せってい",
    "複合": "ふくごう",
    "納品": "のうひん",
    "仕入先": "しいれさき",
    "判断": "はんだん",
    "構成": "こうせい",
    "要素": "ようそ",
    "本体": "ほんたい",
    "該当": "がいとう",
    "可能性": "かのうせい",
    "数量": "すうりょう",
    "標準": "ひょうじゅん",
    "共通": "きょうつう",
    "分類": "ぶんるい",
    "余剰": "よじょう",
    "整理": "せいり",
    "不要": "ふよう",
    "差異": "さい",
    "実装": "じっそう",
    "業務": "ぎょうむ",
    "論点": "ろんてん",
    "一覧化": "いちらんか",
    "矛盾": "むじゅん",
    "保持": "ほじ",
    "最大値": "さいだいち",
    "現状": "げんじょう",
    "変更": "へんこう",
    "代替": "だいたい",
    "参照": "さんしょう",
    "技術的": "ぎじゅつてき",
    "洗い出し": "あらいだし",
    "範囲": "はんい",
    "確定": "かくてい",
    "初期化": "しょきか",
    "計算式": "けいさんしき",
    "成立": "せいりつ",
    "相当": "そうとう",
    "段階": "だんかい",
    "注文": "ちゅうもん",
    "調達": "ちょうたつ",
    "期間": "きかん",
    "休日": "きゅうじつ",
    "加味": "かみ",
    "重要": "じゅうよう",
    "固有": "こゆう",
    "再検証": "さいけんしょう",
    "階層": "かいそう",
    "集約": "しゅうやく",
    "発生": "はっせい",
    "比較": "ひかく",
    "最終": "さいしゅう",
    "特徴": "とくちょう",
    "反映": "はんえい",
    "統合": "とうごう",
    "分析": "ぶんせき",
    "伝票": "でんぴょう",
    "区分": "くぶん",
    "符号": "ふごう",
    "適用": "てきよう",
    "想定": "そうてい",
    "負値": "ふち",
    "正値": "せいち",
    "排他": "はいた",
    "同時": "どうじ",
    "内訳": "うちわけ",
    "変動": "へんどう",
    "入荷": "にゅうか",
    "入庫": "にゅうこ",
    "出庫": "しゅっこ",
    "廃棄": "はいき",
    "売上": "うりあげ",
    "補足": "ほそく",
    "個別": "こべつ",
    "統一": "とういつ",
    "予定": "よてい",
    "除外": "じょがい",
    "誤り": "あやまり",
    "未検出": "みけんしゅつ",
    "訂正": "ていせい",
    "指摘": "してき",
    "更新": "こうしん",
    "会議": "かいぎ",
    "進行": "しんこう",
    "目安": "めやす",
    "意味": "いみ",
    "纏め": "まとめ",
    "品番": "ひんばん",
    "発見": "はっけん",
    "全体": "ぜんたい",
    "導入": "どうにゅう",
    "説明": "せつめい",
    "質問": "しつもん",
    "締め": "しめ",
    "付録": "ふろく",
    "作成日": "さくせいび",
}

FURIGANA_KEYS = sorted(FURIGANA.keys(), key=len, reverse=True)


def annotate(text: str) -> str:
    result = []
    pos = 0
    n = len(text)
    while pos < n:
        matched = False
        for key in FURIGANA_KEYS:
            if text.startswith(key, pos):
                after = pos + len(key)
                if after < n and text[after] == "（":
                    result.append(key)
                    pos = after
                    matched = True
                    break
                result.append(f"{key}（{FURIGANA[key]}）")
                pos += len(key)
                matched = True
                break
        if not matched:
            result.append(text[pos])
            pos += 1
    return "".join(result)


def set_run_font(run, size_pt=11, bold=False, color=None):
    run.font.name = "Meiryo UI"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Meiryo UI")
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def add_paragraph_text(doc, text, *, size=11, bold=False, color=None,
                       align=None, quote=False, bullet=False):
    p = doc.add_paragraph()
    if bullet:
        p.style = "List Bullet"
    if align is not None:
        p.alignment = align
    if quote:
        p.paragraph_format.left_indent = Cm(1.0)
        ppr = p._p.get_or_add_pPr()
        pbd = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "12")
        left.set(qn("w:space"), "8")
        left.set(qn("w:color"), "1A5C9E")
        pbd.append(left)
        ppr.append(pbd)

    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        is_bold = part.startswith("**") and part.endswith("**")
        raw = part[2:-2] if is_bold else part
        raw = re.sub(r"`(.+?)`", r"\1", raw)
        annotated = annotate(raw)
        i = 0
        while i < len(annotated):
            m = re.search(r"（[^）]+）", annotated[i:])
            if not m:
                run = p.add_run(annotated[i:])
                set_run_font(run, size, bold or is_bold, color)
                break
            start = i + m.start()
            end = i + m.end()
            if start > i:
                run = p.add_run(annotated[i:start])
                set_run_font(run, size, bold or is_bold, color)
            run = p.add_run(annotated[start:end])
            set_run_font(run, max(8, size - 2), False, RGBColor(0x75, 0x75, 0x75))
            i = end
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.space_before = Pt(2)
    return p


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = annotate(h)
        for run in cell.paragraphs[0].runs:
            set_run_font(run, 10, True, RGBColor(0xFF, 0xFF, 0xFF))
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), "1A5C9E")
        shading.set(qn("w:val"), "clear")
        cell._tc.get_or_add_tcPr().append(shading)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = annotate(val)
            for run in cell.paragraphs[0].runs:
                set_run_font(run, 10)
    doc.add_paragraph()


def parse_and_build(doc):
    with open(INPUT_MD, encoding="utf-8") as f:
        lines = f.readlines()

    i = 0
    in_table = False
    table_rows = []
    table_headers = []

    while i < len(lines):
        line = lines[i].rstrip("\n")

        if line.startswith("|") and not in_table:
            in_table = True
            table_headers = [c.strip() for c in line.strip("|").split("|")]
            i += 1
            if i < len(lines) and lines[i].startswith("|---"):
                i += 1
            continue

        if in_table:
            if line.startswith("|"):
                table_rows.append([c.strip() for c in line.strip("|").split("|")])
                i += 1
                continue
            else:
                add_table(doc, table_headers, table_rows)
                in_table = False
                table_rows = []
                table_headers = []
                continue

        if line.startswith("# ") and not line.startswith("## "):
            add_paragraph_text(doc, line[2:].strip(), size=16, bold=True,
                               align=WD_ALIGN_PARAGRAPH.CENTER)
        elif line.startswith("## "):
            add_paragraph_text(doc, line[3:].strip(), size=14, bold=True,
                               color=RGBColor(0x0D, 0x2B, 0x4E))
        elif line.startswith("### "):
            add_paragraph_text(doc, line[4:].strip(), size=12, bold=True,
                               color=RGBColor(0x1A, 0x5C, 0x9E))
        elif line.startswith("> "):
            add_paragraph_text(doc, line[2:], size=11, quote=True)
        elif line.startswith("- "):
            add_paragraph_text(doc, line[2:], size=11, bullet=True)
        elif line.startswith("**【") and "**" in line:
            add_paragraph_text(doc, line.replace("**", ""), size=10, bold=True,
                               color=RGBColor(0xD3, 0x2F, 0x2F))
        elif line.startswith("**ファイル名"):
            add_paragraph_text(doc, line.replace("**", "").replace("`", ""), size=10,
                               color=RGBColor(0x1A, 0x5C, 0x9E))
        elif line == "---":
            doc.add_paragraph()
        elif line.startswith("*作成日"):
            add_paragraph_text(doc, line.strip("*"), size=9,
                               color=RGBColor(0x75, 0x75, 0x75),
                               align=WD_ALIGN_PARAGRAPH.RIGHT)
        elif line.strip() == "":
            pass
        elif not line.startswith("```"):
            add_paragraph_text(doc, line, size=11)

        i += 1

    if in_table and table_headers:
        add_table(doc, table_headers, table_rows)


def build_docx():
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Meiryo UI"
    style.font.size = Pt(11)

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("※ 複雑な漢字には括弧で読み仮名を付記しています（WPS対応）。")
    set_run_font(run, 9, False, RGBColor(0x75, 0x75, 0x75))
    doc.add_paragraph()

    parse_and_build(doc)

    doc.add_page_break()
    add_paragraph_text(doc, "括弧注音の凡例", size=14, bold=True,
                       color=RGBColor(0x0D, 0x2B, 0x4E))
    for n in [
        "本資料では、ビジネス・技術用語など読みにくい漢字語を「漢字（かな）」形式で注音しています。",
        "括弧内の仮名は灰色で表示されます。WPS / Word / Google Docs いずれでも正常に読めます。",
        "英数字・カタカナ・既に読みが明確な語は注音対象外としています。",
    ]:
        add_paragraph_text(doc, n, size=11, bullet=True)

    doc.save(OUTPUT)
    print(f"生成完了: {OUTPUT}")


if __name__ == "__main__":
    build_docx()
