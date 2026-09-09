#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""日立確認会 提問稿 — 括弧注音 MD + Word"""

import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

INPUT_MD = "/Users/treqd/Desktop/Cursor/西友PMI/在庫IF_日立確認会_提问稿.md"
OUTPUT_MD = "/Users/treqd/Desktop/Cursor/西友PMI/在庫IF_日立確認会_提问稿_注音.md"
OUTPUT_DOCX = "/Users/treqd/Desktop/Cursor/西友PMI/在庫IF_日立確認会_提问稿.docx"

FURIGANA = {
    "日立確認会": "ひたちかくにんかい",
    "提問稿": "ていもんこう",
    "項目移送表": "こうもくいそうひょう",
    "在庫調整データ": "ざいこちょうせいデータ",
    "在庫一覧データ": "ざいこいちらんデータ",
    "在庫調整": "ざいこちょうせい",
    "在庫一覧": "ざいこいちらん",
    "適用順序": "てきようじゅんじょ",
    "更新頻度": "こうしんひんど",
    "同一日付": "どういつひづけ",
    "データ日付": "データひづけ",
    "全件更新": "ぜんけんこうしん",
    "全件スナップショット": "ぜんけんスナップショット",
    "差分型": "さぶんがた",
    "随時送信": "ずいじそうしん",
    "日次送信": "にっじそうしん",
    "送信時点": "そうしんじてん",
    "連携タイミング": "れんけいタイミング",
    "同一テーブル": "どういつテーブル",
    "取込設計": "とりこみせっけい",
    "最終値": "さいしゅうち",
    "上書き": "うわがき",
    "反映": "はんえい",
    "抽出結果": "ちゅうしゅつけっか",
    "在庫数量": "ざいこすうりょう",
    "対象範囲": "たいしょうはんい",
    "送信済み実績": "そうしんずみじっせき",
    "抽出時刻": "ちゅうしゅつじこく",
    "補足事項": "ほそくじこう",
    "実際の運用": "じっさいのうんよう",
    "発生の都度": "はっせいのつど",
    "一定の間隔": "いっていのかんかく",
    "切れ目": "きれめ",
    "考え方": "かんがえかた",
    "完全に一致": "かんぜんにいっち",
    "指定した日付": "していしたひづけ",
    "以前お伺い": "いぜんおうかがい",
    "確認": "かくにん",
    "在庫": "ざいこ",
    "調整": "ちょうせい",
    "一覧": "いちらん",
    "差分": "さぶん",
    "全件": "ぜんけん",
    "日付": "ひづけ",
    "件数": "けんすう",
    "数量": "すうりょう",
    "前提": "ぜんてい",
    "見解": "けんかい",
    "記載": "きさい",
    "理解": "りかい",
    "運用": "うんよう",
    "発生": "はっせい",
    "送信": "そうしん",
    "更新": "こうしん",
    "間隔": "かんかく",
    "想定": "そうてい",
    "順番": "じゅんばん",
    "結果": "けっか",
    "最新": "さいしん",
    "日次": "にっじ",
    "前後": "ぜんご",
    "中身": "なかみ",
    "抽出": "ちゅうしゅつ",
    "限定": "げんてい",
    "一致": "いっち",
    "理由": "りゆう",
    "対象": "たいしょう",
    "扱い": "あつかい",
    "含め方": "ふくめかた",
    "設計": "せっけい",
    "回答": "かいとう",
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


def generate_annotated_md():
    with open(INPUT_MD, encoding="utf-8") as f:
        lines = f.read().splitlines()
    out = []
    for idx, line in enumerate(lines):
        if idx == 0 and line.startswith("# "):
            out.append("# 在庫IF 日立確認会 提問稿（括弧注音・照着念）")
            out.append("")
            out.append("> ※ 読みにくい漢字語には「漢字（かな）」を付記しています。")
            continue
        if line.startswith("```") or line.strip() in ("", "---"):
            out.append(line)
            continue
        out.append(annotate(line))
    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(out))
    print("生成完了:", OUTPUT_MD)


def set_run_font(run, size_pt=11, bold=False, color=None):
    run.font.name = "Meiryo UI"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Meiryo UI")
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def add_paragraph_text(doc, text, *, size=11, bold=False, color=None,
                       align=None, quote=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    if quote:
        p.paragraph_format.left_indent = Cm(0.8)
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
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.space_before = Pt(2)
    return p


def parse_and_build(doc):
    with open(INPUT_MD, encoding="utf-8") as f:
        lines = f.readlines()
    for line in lines:
        line = line.rstrip("\n")
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
            add_paragraph_text(doc, line[2:], size=10, quote=True,
                               color=RGBColor(0x4B, 0x55, 0x63))
        elif line == "---":
            doc.add_paragraph()
        elif line.strip() == "":
            pass
        else:
            add_paragraph_text(doc, line, size=12)


def build_docx():
    doc = Document()
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    style = doc.styles["Normal"]
    style.font.name = "Meiryo UI"
    style.font.size = Pt(12)

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run("※ 複雑な漢字には括弧で読み仮名を付記しています（WPS対応）。")
    set_run_font(run, 9, False, RGBColor(0x75, 0x75, 0x75))

    parse_and_build(doc)
    doc.save(OUTPUT_DOCX)
    print("生成完了:", OUTPUT_DOCX)


if __name__ == "__main__":
    generate_annotated_md()
    build_docx()
