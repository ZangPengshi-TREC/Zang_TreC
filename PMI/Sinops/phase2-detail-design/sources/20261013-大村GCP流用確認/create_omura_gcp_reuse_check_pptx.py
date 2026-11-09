#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""大村さんGCP環境の流用確認PPT。"""

from copy import deepcopy
from io import BytesIO
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.util import Emu, Inches, Pt

OUT = Path(__file__).resolve().parent / "大村さんGCP環境_流用可否確認_20261013.pptx"
CURRENT = OUT
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

NAVY = RGBColor(0x1F, 0x4E, 0x78)
NAVY2 = RGBColor(0x2E, 0x5A, 0x88)
ACCENT = RGBColor(0xC4, 0x59, 0x11)
GRAY = RGBColor(0x6B, 0x6B, 0x6B)
DARK = RGBColor(0x2D, 0x2D, 0x2D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xC8, 0xD0, 0xD8)
HEAD = RGBColor(0x1F, 0x4E, 0x78)
ROW = RGBColor(0xF4, 0xF7, 0xFA)
PANEL = RGBColor(0xF4, 0xF7, 0xFA)
NOTE = RGBColor(0xFD, 0xF6, 0xEE)
COL0 = RGBColor(0xE8, 0xEE, 0xF4)


LATIN = "Calibri"
EA = "メイリオ"


def set_run(run, size=14, bold=False, color=DARK, latin=LATIN, ea=EA):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = latin
    rPr = run._r.get_or_add_rPr()
    for tag, face in (("latin", latin), ("ea", ea), ("cs", latin)):
        xml = (
            f'<a:{tag} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'typeface="{face}"/>'
        )
        new = parse_xml(xml)
        old = rPr.find(qn(f"a:{tag}"))
        if old is not None:
            rPr.replace(old, new)
        else:
            rPr.append(new)


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


def add_para(tf, text, size=13, bold=False, color=DARK, space_before=4, space_after=2, align=PP_ALIGN.LEFT):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color)
    return p


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def rect(slide, left, top, width, height, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Emu(6350)
    return sh


def title_block(slide, title):
    rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.08), NAVY)
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.22), Inches(10.5), Inches(0.46))
    add_text(box.text_frame, title, size=22, bold=True, color=NAVY)
    rect(slide, Inches(0.5), Inches(0.72), Inches(1.35), Inches(0.055), ACCENT)
    rect(slide, Inches(0), Inches(7.42), Inches(13.333), Inches(0.08), NAVY)


def page_no(slide, page, total):
    lab = slide.shapes.add_textbox(Inches(0.5), Inches(7.08), Inches(8.5), Inches(0.28))
    add_text(lab.text_frame, "MD基幹統合　流用確認", size=10, color=GRAY)
    num = slide.shapes.add_textbox(Inches(11.6), Inches(7.08), Inches(1.2), Inches(0.28))
    add_text(num.text_frame, f"{page} / {total}", size=10, color=GRAY, align=PP_ALIGN.RIGHT)


def set_cell_border(cell, color="C8D0D8"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith("tcBdr"):
            tcPr.remove(child)
    tcPr.append(
        parse_xml(
            "<a:tcBdr xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>"
            f"<a:left><a:ln w='6350'><a:solidFill><a:srgbClr val='{color}'/></a:solidFill></a:ln></a:left>"
            f"<a:right><a:ln w='6350'><a:solidFill><a:srgbClr val='{color}'/></a:solidFill></a:ln></a:right>"
            f"<a:top><a:ln w='6350'><a:solidFill><a:srgbClr val='{color}'/></a:solidFill></a:ln></a:top>"
            f"<a:bottom><a:ln w='6350'><a:solidFill><a:srgbClr val='{color}'/></a:solidFill></a:ln></a:bottom>"
            "</a:tcBdr>"
        )
    )


def fill_table(table, rows, font_size=13, header=True, label_col=False):
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.16)
            cell.margin_right = Inches(0.12)
            cell.margin_top = Inches(0.1)
            cell.margin_bottom = Inches(0.1)
            is_head = header and r == 0
            is_label = (not is_head) and c == 0 and label_col
            color = WHITE if is_head else (NAVY if is_label else DARK)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    set_run(run, size=font_size, bold=is_head or is_label, color=color)
            if is_head:
                cell.fill.solid()
                cell.fill.fore_color.rgb = HEAD
            elif is_label:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COL0
            elif r % 2 == (1 if header else 0):
                cell.fill.solid()
                cell.fill.fore_color.rgb = ROW
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE
            set_cell_border(cell, "1F4E78" if is_head else "C8D0D8")


def copy_slide(src_prs, dest_prs, src_index):
    src_slide = src_prs.slides[src_index]
    dest_slide = dest_prs.slides.add_slide(dest_prs.slide_layouts[6])
    for shape in list(dest_slide.shapes):
        shape._element.getparent().remove(shape._element)

    rid_map = {}
    for rel in src_slide.part.rels.values():
        if "image" not in rel.reltype:
            continue
        blob = rel.target_part.blob
        _part, new_rid = dest_slide.part.get_or_add_image_part(BytesIO(blob))
        rid_map[rel.rId] = new_rid

    src_tree = src_slide.shapes._spTree
    dest_tree = dest_slide.shapes._spTree
    for child in src_tree:
        tag = child.tag
        if tag.endswith("nvGrpSpPr") or tag.endswith("grpSpPr"):
            continue
        new_el = deepcopy(child)
        for blip in new_el.iter(A_BLIP):
            old = blip.get(R_EMBED)
            if old and old in rid_map:
                blip.set(R_EMBED, rid_map[old])
        dest_tree.append(new_el)

    src_csl = src_slide._element.find(qn("p:cSld"))
    dest_csl = dest_slide._element.find(qn("p:cSld"))
    if src_csl is not None and dest_csl is not None:
        src_bg = src_csl.find(qn("p:bg"))
        if src_bg is not None:
            dest_bg = dest_csl.find(qn("p:bg"))
            if dest_bg is not None:
                dest_csl.remove(dest_bg)
            dest_csl.insert(0, deepcopy(src_bg))


def col_block(slide, left, top, width, heading, lines):
    rect(slide, left, top + Inches(0.06), Inches(0.07), Inches(0.28), NAVY)
    hd = slide.shapes.add_textbox(left + Inches(0.2), top, width - Inches(0.25), Inches(0.4))
    add_text(hd.text_frame, heading, size=15, bold=True, color=NAVY)
    body = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.42), width - Inches(0.3), Inches(4.2))
    tf = body.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            add_text(tf, "・" + line, size=14, color=DARK)
            first = False
        else:
            add_para(tf, "・" + line, size=14, space_before=10, space_after=0)


def confirm_item(slide, top, num, title, sub=None):
    h = Inches(1.42) if sub else Inches(1.15)
    rect(slide, Inches(0.5), top, Inches(12.3), h, PANEL, LINE)
    nb = rect(slide, Inches(0.5), top, Inches(0.82), h, NAVY)
    ntf = nb.text_frame
    ntf.word_wrap = False
    ntf.anchor = MSO_ANCHOR.MIDDLE
    ntf.margin_left = Inches(0)
    ntf.margin_right = Inches(0)
    ntf.margin_top = Inches(0)
    ntf.margin_bottom = Inches(0)
    add_text(ntf, num, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    box = slide.shapes.add_textbox(Inches(1.55), top + Inches(0.22 if sub else 0.32), Inches(10.9), Inches(1.05 if sub else 0.55))
    tf = box.text_frame
    tf.word_wrap = True
    add_text(tf, title, size=18, bold=True, color=NAVY)
    if sub:
        add_para(tf, sub, size=14, color=GRAY, space_before=8)
    return h


def build():
    src = None
    if CURRENT.exists():
        backup = CURRENT.with_name("_slide3_source.pptx")
        backup.write_bytes(CURRENT.read_bytes())
        src = Presentation(str(backup))
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = 6

    # ---- 1 表紙 ----
    s = blank_slide(prs)
    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), WHITE)
    rect(s, Inches(0), Inches(0), Inches(0.16), Inches(7.5), NAVY)
    rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.08), NAVY)
    rect(s, Inches(0), Inches(7.42), Inches(13.333), Inches(0.08), NAVY)

    k = s.shapes.add_textbox(Inches(0.85), Inches(2.05), Inches(11.5), Inches(0.36))
    add_text(k.text_frame, "西友MD基幹統合　Phase2", size=14, color=GRAY)
    rect(s, Inches(0.85), Inches(2.48), Inches(1.35), Inches(0.055), ACCENT)

    t = s.shapes.add_textbox(Inches(0.85), Inches(2.7), Inches(11.5), Inches(0.7))
    add_text(t.text_frame, "GCP環境の流用確認", size=30, bold=True, color=NAVY)

    sub = s.shapes.add_textbox(Inches(0.85), Inches(3.5), Inches(11.5), Inches(0.45))
    add_text(sub.text_frame, "大村さん・田村さん環境（発注転送GCP）", size=16, color=DARK)

    rect(s, Inches(0.85), Inches(5.85), Inches(11.6), Emu(6350), LINE)
    meta = s.shapes.add_textbox(Inches(0.85), Inches(6.0), Inches(11.5), Inches(0.35))
    add_text(meta.text_frame, "2026/10/13", size=13, color=GRAY)

    # ---- 2 背景 ----
    s = blank_slide(prs)
    title_block(s, "背景")
    col_block(
        s,
        Inches(0.5),
        Inches(1.05),
        Inches(6.05),
        "今回",
        [
            "集計SV → 型変換・項目選別補充 → 西友連携GCS",
            "GCS：md-data-integration（暫定）",
            "Cloud Run",
        ],
    )
    col_block(
        s,
        Inches(6.75),
        Inches(1.05),
        Inches(6.05),
        "大村さん・田村さん環境",
        [
            "TRIAL：大村さん　／　西友：田村さん",
            "発注データ。型変換・項目選別補充して転送",
            "GCS：seiyu-freshorder-prod",
            "Cloud Run　4 vCPU / 4 GiB（拡張可）",
            "西友：ods-seiyu-prod / CM0010_IF0079_001",
        ],
    )
    rect(s, Inches(0.5), Inches(5.78), Inches(12.3), Inches(1.05), NOTE, LINE)
    rect(s, Inches(0.5), Inches(5.78), Inches(0.08), Inches(1.05), ACCENT)
    note = s.shapes.add_textbox(Inches(0.85), Inches(6.05), Inches(11.7), Inches(0.55))
    add_text(note.text_frame, "岩瀬さん：既存の発注転送GCPで足りるか、先に確認", size=14, color=DARK)
    page_no(s, 2, total)

    # ---- 3 IFフロー ----
    if src is not None and len(src.slides) >= 3:
        copy_slide(src, prs, 2)
    else:
        s = blank_slide(prs)
        title_block(s, "MD基幹統合 IF連携フロー")
        page_no(s, 3, total)

    # ---- 4 本日確認 ----
    s = blank_slide(prs)
    title_block(s, "本日確認")
    y = Inches(1.08)
    y += confirm_item(s, y, "1", "devの流用可否", "可なら単体テストまで実施") + Inches(0.16)
    y += confirm_item(s, y, "2", "stg / prodの流用可否", "devの後") + Inches(0.16)
    confirm_item(s, y, "3", "使えない環境は新規申請", "不可の環境のみ")
    page_no(s, 4, total)

    # ---- 5 環境比較 ----
    s = blank_slide(prs)
    title_block(s, "環境比較")
    rows = [
        ["項目", "大村さん・田村さん環境", "今回"],
        ["用途", "型変換・項目選別補充の発注データの転送", "集計SV → 型変換・項目選別補充 → 西友連携GCS"],
        ["実行", "Cloud Run　4 vCPU / 4 GiB（拡張可）", "Cloud Run　4 vCPU / 16 GiB"],
        ["GCS", "seiyu-freshorder-prod", "md-data-integration（暫定）"],
        ["西友連携GCS", "ods-seiyu-prod / CM0010_IF0079_001", "最終TXT。フォルダは松尾さんへ新規（CM0010は使わない）"],
    ]
    table_shape = s.shapes.add_table(
        len(rows), 3, Inches(0.5), Inches(1.15), Inches(12.3), Inches(5.15)
    )
    table = table_shape.table
    table.columns[0].width = Inches(1.65)
    table.columns[1].width = Inches(5.325)
    table.columns[2].width = Inches(5.325)
    fill_table(table, rows, font_size=13, header=True, label_col=True)
    page_no(s, 5, total)

    # ---- 6 確認点（4枚目以外。理由を表に出す） ----
    s = blank_slide(prs)
    title_block(s, "確認点")
    qrows = [
        ["項目", "確認", "理由"],
        ["メモリ", "16 GiBまで上げてよいか（現状4 GiB）", "今回の集計変換は16 GiBを見ている"],
        ["西友フォルダ", "松尾さんへ新規。CM0010は使わない", "発注のパスとは分ける"],
        [
            "田村さん",
            "権限・運用に影響がないか",
            "発注転送が今動いている。西友側の窓口は田村さん",
        ],
    ]
    table_shape = s.shapes.add_table(
        len(qrows), 3, Inches(0.5), Inches(1.15), Inches(12.3), Inches(4.85)
    )
    table = table_shape.table
    table.columns[0].width = Inches(2.15)
    table.columns[1].width = Inches(4.85)
    table.columns[2].width = Inches(5.3)
    fill_table(table, qrows, font_size=14, header=True, label_col=True)
    page_no(s, 6, total)

    prs.save(OUT)
    leftover = CURRENT.with_name("_slide3_source.pptx")
    if leftover.exists():
        leftover.unlink()
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
