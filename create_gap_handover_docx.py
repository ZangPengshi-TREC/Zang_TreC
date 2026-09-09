#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GAP引き継ぎ説明话术 — 括弧注音 MD + Word 正式文档生成"""

import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

INPUT_MD = "/Users/treqd/Desktop/Cursor/西友PMI/GAP重点確認_説明话术.md"
OUTPUT_MD = "/Users/treqd/Desktop/Cursor/西友PMI/GAP重点確認_説明话术_注音.md"
OUTPUT_DOCX = "/Users/treqd/Desktop/Cursor/西友PMI/GAP重点確認_説明话术.docx"

FURIGANA = {
    "GAP分析書": "ギャップぶんせきしょ",
    "項目マッピング表": "こうもくマッピングひょう",
    "受払明細": "うけばらいめいさい",
    "入荷実績": "にゅうかじっせき",
    "入荷予定": "にゅうかよてい",
    "便区分": "びんくぶん",
    "伝票区分": "でんぴょうくぶん",
    "振替伝票": "ふりかえでんぴょう",
    "仕入伝票": "しいれでんぴょう",
    "発注データ区分": "はっちゅうデータくぶん",
    "通常発注": "つうじょうはっちゅう",
    "日配発注": "にっぱいはっちゅう",
    "本部発注": "ほんぶはっちゅう",
    "特売送込": "とくばいおくりこみ",
    "新店特売": "しんてんとくばい",
    "顧客発注": "こきゃくはっちゅう",
    "必須項目": "ひっすこうもく",
    "整数値": "せいすうち",
    "同一日": "どういつび",
    "直接対応": "ちょくせつたいおう",
    "影響範囲": "えいきょうはんい",
    "未確定": "みかくてい",
    "未解決": "みかいけつ",
    "未決定": "みけってい",
    "未確認": "みかくにん",
    "物流タイプ": "ぶつりゅうタイプ",
    "納品方法区分": "のうひんほうほうくぶん",
    "業務実態": "ぎょうむじったい",
    "業務妥当性": "ぎょうむだとうせい",
    "変換案": "へんかんあん",
    "変換内容": "へんかんないよう",
    "連携テキスト": "れんけいテキスト",
    "データソース": "データソース",
    "固定値運用": "こていちうんよう",
    "固定値": "こていち",
    "引き継ぎ": "ひきつぎ",
    "責任者": "せきにんしゃ",
    "説明话术": "せつめいこうは",
    "口播稿": "こうはこう",
    "方向性": "ほうこうせい",
    "対応表": "たいおうひょう",
    "暫定対応": "ざんていたいおう",
    "暫定案": "ざんていあん",
    "参考情報": "さんこうじょうほう",
    "サンプル検証": "サンプルけんしょう",
    "サンプルデータ": "サンプルデータ",
    "入庫数": "にゅうこすう",
    "売上数": "うりあげすう",
    "割当ルール": "わりあてルール",
    "割り当て": "わりあて",
    "妥当性確認": "だとうせいかくにん",
    "方針決定": "ほうしんけってい",
    "照合未了": "しょうごうみりょう",
    "導出可能性": "どうしゅつかのうせい",
    "導出": "どうしゅつ",
    "識別": "しきべつ",
    "許容": "きょよう",
    "表記": "ひょうき",
    "空欄": "くうらん",
    "対象候補": "たいしょうこうほ",
    "連携要否": "れんけいようひ",
    "問題点": "もんだいてん",
    "問題の核心": "もんだいのかくしん",
    "次に必要": "つぎにひつよう",
    "引き継ぎ先": "ひきつぎさき",
    "引き継ぎ時": "ひきつぎじ",
    "引き継ぎまとめ": "ひきつぎまとめ",
    "マッピング元": "マッピングもと",
    "方針未決定": "ほうしんみけってい",
    "全体の導入": "ぜんたいのどうにゅう",
    "検討が必要": "けんとうがひつよう",
    "結論": "けつろん",
    "双方": "そうほう",
    "背景": "はいけい",
    "現状": "げんじょう",
    "論点": "ろんてん",
    "整理": "せいり",
    "確認": "かくにん",
    "調整": "ちょうせい",
    "連携": "れんけい",
    "変換": "へんかん",
    "出力": "しゅつりょく",
    "入庫": "にゅうこ",
    "出庫": "しゅっこ",
    "廃棄": "はいき",
    "入荷": "にゅうか",
    "予定": "よてい",
    "実績": "じっせき",
    "発注": "はっちゅう",
    "区分": "くぶん",
    "項目": "こうもく",
    "対象": "たいしょう",
    "要求": "ようきゅう",
    "確定": "かくてい",
    "検討": "けんとう",
    "判断": "はんだん",
    "説明": "せつめい",
    "詳細": "しょうさい",
    "参照": "さんしょう",
    "適用": "てきよう",
    "設定": "せってい",
    "存在": "そんざい",
    "双方": "そうほう",
    "核心": "かくしん",
    "範囲": "はんい",
    "双方": "そうほう",
    "更新": "こうしん",
    "資料": "しりょう",
    "記載": "きさい",
    "状態": "じょうたい",
    "不明": "ふめい",
    "未定": "みてい",
    "別処理": "べつしょり",
    "固有": "こゆう",
    "有無": "うむ",
    "以上": "いじょう",
    "主要": "しゅよう",
    "必要": "ひつよう",
    "今後": "こんご",
    "現時点": "げんじiten",
    "双方": "そうほう",
}

# fix typo
FURIGANA["現時点"] = "げんじてん"

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


def generate_annotated_md():
    with open(INPUT_MD, encoding="utf-8") as f:
        content = f.read()

    lines = content.splitlines()
    out = []
    for idx, line in enumerate(lines):
        if idx == 0 and line.startswith("# "):
            out.append("# GAP未解決論点 — 責任者引き継ぎ用 説明话术（括弧注音版）")
            out.append("")
            out.append("> ※ 読みにくい漢字語には「漢字（かな）」形式で読み仮名を付記しています。")
            continue
        if line.startswith("```"):
            out.append(line)
            continue
        if line.startswith("|") or line.startswith("#") or line.startswith(">"):
            out.append(annotate(line))
        elif line.startswith("**") or line.startswith("- "):
            out.append(annotate(line))
        elif line.strip() == "" or line == "---":
            out.append(line)
        else:
            out.append(annotate(line))

    final = "\n".join(out)
    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write(final)
    print(f"生成完了: {OUTPUT_MD}")


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
        elif line.startswith("**") and line.endswith("**"):
            add_paragraph_text(doc, line.replace("**", ""), size=11, bold=True)
        elif line == "---":
            doc.add_paragraph()
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

    # 表紙
    add_paragraph_text(doc, "GAP未解決論点", size=20, bold=True,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph_text(doc, "責任者引き継ぎ用 説明话术", size=14, bold=True,
                       align=WD_ALIGN_PARAGRAPH.CENTER,
                       color=RGBColor(0x1A, 0x5C, 0x9E))
    doc.add_paragraph()
    meta = [
        "文書種別：引き継ぎ説明資料",
        "対象論点：GAP No.4/No.9（便区分）、No.10（コメント）、受払明細(W)出庫/入庫の参照先",
        "注音方式：括弧注音（WPS / Word 対応）",
    ]
    for m in meta:
        add_paragraph_text(doc, m, size=10, color=RGBColor(0x59, 0x59, 0x59))
    doc.add_page_break()

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

    doc.save(OUTPUT_DOCX)
    print(f"生成完了: {OUTPUT_DOCX}")


if __name__ == "__main__":
    generate_annotated_md()
    build_docx()
