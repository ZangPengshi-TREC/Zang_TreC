# -*- coding: utf-8 -*-
"""Sinops 自動補充 日次サイクル（業務フロー）の draw.io と PNG を同じ座標から生成する。

出典: IF一覧_Layer1-3.md / SCM-GCP（今回申請）アーキテクチャ.md
実行: python3 create_sinops_business_flow.py   （PNG には Pillow が必要）
"""
from __future__ import annotations

import math
from html import escape
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT_DRAWIO = HERE / "Sinops_業務フロー_日次サイクル.drawio"
OUT_PNG = HERE / "Sinops_業務フロー_日次サイクル.png"
FONT_PATH = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"

W, H = 224, 124
PITCH_X, PITCH_Y = 304, 196
X0, Y0 = 60, 120
CANVAS_W = X0 * 2 + 3 * PITCH_X + W
CANVAS_H = Y0 + 2 * PITCH_Y + H + 64

COLORS = {
    "TRIAL": ("#DAE8FC", "#6C8EBF"),
    "GCP（今回申請）": ("#D5E8D4", "#82B366"),
    "西友 DataSpider": ("#FFE6CC", "#D79B00"),
    "Sinops": ("#E1D5E7", "#9673A6"),
}

# id, 行, 列, 主体(色), 見出し, 明細
STEPS = [
    (1, 0, 0, "TRIAL", "TRIAL 基幹", ["マスタ 01–05・13・14", "実績・在庫 06–12・15", "倉庫系 16–21"]),
    (2, 0, 1, "GCP（今回申請）", "GCP（今回申請）", ["抽出", "型・コード変換", "連携GCSへ配置"]),
    (3, 0, 2, "西友 DataSpider", "西友 DataSpider", ["マスタ: Intake 洗替", "実績: 転送のみ", "全フィールド通過"]),
    (4, 0, 3, "Sinops", "Sinops", ["データ受信", "発注勧告を計算"]),
    (5, 1, 3, "Sinops", "Sinops", ["24 勧告当日分（通常）", "25 翌日以降分（代替）", "22・23 倉庫勧告"]),
    (6, 2, 3, "西友 DataSpider", "西友 DataSpider", ["勧告を転送", "連携GCSへ配置"]),
    (7, 2, 2, "GCP（今回申請）", "GCP（今回申請）", ["勧告を取得", "勧告→発注統合変換", "生鮮増分・新規"]),
    (8, 2, 1, "TRIAL", "TRIAL 自動発注", ["自動発注GCS", "OrdreSV 自動確定", "CoreSaver 同期"]),
    (9, 2, 0, "TRIAL", "TRIAL 各業務", ["発注・物流・納品", "入荷予定 12 を生成", "入荷実績 11 を生成"]),
]
# 起点, 終点, ラベル, 破線か
EDGES = [
    (1, 2, "マスタ・実績", False),
    (2, 3, "連携GCS", False),
    (3, 4, "データ送信", False),
    (4, 5, "勧告計算完了", False),
    (5, 6, "勧告出力", False),
    (6, 7, "連携GCS", False),
    (7, 8, "自動発注GCS", False),
    (8, 9, "発注確定後", False),
    (9, 1, "入荷データ → 次サイクル", True),
]
CORE_TITLE = "日次サイクル"
CORE_LINES = ["往路 1 → 4（データ送信）", "復路 5 → 8（自動確定）", "入荷 9 → 1（次サイクルへ）"]
TITLE = "Sinops 自動補充 日次サイクル"


def pos(row: int, col: int):
    return X0 + col * PITCH_X, Y0 + row * PITCH_Y


def center(step_id: int):
    _, row, col, *_ = STEPS[step_id - 1]
    x, y = pos(row, col)
    return x + W / 2, y + H / 2


def anchors(a: int, b: int):
    """二つのステップをつなぐ線の始点・終点と、drawio の exit/entry。"""
    _, ra, ca, *_ = STEPS[a - 1]
    _, rb, cb, *_ = STEPS[b - 1]
    xa, ya = pos(ra, ca)
    xb, yb = pos(rb, cb)
    if rb == ra and cb > ca:
        return (xa + W, ya + H / 2), (xb, yb + H / 2), (1, 0.5, 0, 0.5)
    if rb == ra and cb < ca:
        return (xa, ya + H / 2), (xb + W, yb + H / 2), (0, 0.5, 1, 0.5)
    if cb == ca and rb > ra:
        return (xa + W / 2, ya + H), (xb + W / 2, yb), (0.5, 1, 0.5, 0)
    return (xa + W / 2, ya), (xb + W / 2, yb + H), (0.5, 0, 0.5, 1)


def build_drawio() -> str:
    cells = ['<mxCell id="0"/>', '<mxCell id="1" parent="0"/>']
    cells.append(
        f'<mxCell id="title" value="{escape(TITLE)}" style="text;html=1;align=left;verticalAlign=middle;fontSize=22;fontStyle=1;" '
        f'vertex="1" parent="1"><mxGeometry x="{X0}" y="28" width="520" height="36" as="geometry"/></mxCell>'
    )
    lx = CANVAS_W - X0 - 4 * 150
    for i, name in enumerate(COLORS):
        fill, stroke = COLORS[name]
        cells.append(
            f'<mxCell id="lg{i}" value="{escape(name)}" style="rounded=1;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};fontSize=12;" '
            f'vertex="1" parent="1"><mxGeometry x="{lx + i * 150}" y="34" width="136" height="26" as="geometry"/></mxCell>'
        )
    for sid, row, col, actor, head, lines in STEPS:
        fill, stroke = COLORS[actor]
        body = "".join(f"<br>・{escape(t)}" for t in lines)
        value = f"<b>{sid}　{escape(head)}</b>{body}"
        x, y = pos(row, col)
        cells.append(
            f'<mxCell id="s{sid}" value="{escape(value, quote=True)}" '
            f'style="rounded=1;whiteSpace=wrap;html=1;align=left;verticalAlign=top;spacingLeft=12;spacingTop=8;fillColor={fill};strokeColor={stroke};strokeWidth=2;fontSize=13;" '
            f'vertex="1" parent="1"><mxGeometry x="{x}" y="{y}" width="{W}" height="{H}" as="geometry"/></mxCell>'
        )
    cx, cy = pos(1, 1)
    core_value = f"<b>{escape(CORE_TITLE)}</b>" + "".join(f"<br>{escape(t)}" for t in CORE_LINES)
    cells.append(
        f'<mxCell id="core" value="{escape(core_value, quote=True)}" '
        f'style="rounded=1;whiteSpace=wrap;html=1;dashed=1;dashPattern=8 6;strokeWidth=2;strokeColor=#666666;fillColor=none;fontSize=14;" '
        f'vertex="1" parent="1"><mxGeometry x="{cx}" y="{cy}" width="{PITCH_X + W}" height="{H}" as="geometry"/></mxCell>'
    )
    for i, (a, b, label, dashed) in enumerate(EDGES, 1):
        _, _, (ex, ey, nx, ny) = anchors(a, b)
        style = (
            "endArrow=block;endFill=1;html=1;strokeWidth=3;strokeColor=#1A202C;fontSize=12;labelBackgroundColor=#FFFFFF;"
            f"exitX={ex};exitY={ey};exitDx=0;exitDy=0;entryX={nx};entryY={ny};entryDx=0;entryDy=0;"
        )
        if dashed:
            style += "dashed=1;dashPattern=8 6;strokeColor=#C53030;"
        cells.append(
            f'<mxCell id="e{i}" value="{escape(label)}" style="{style}" edge="1" parent="1" source="s{a}" target="s{b}">'
            f'<mxGeometry relative="1" as="geometry"/></mxCell>'
        )
    body = "\n        ".join(cells)
    return f'''<mxfile host="app.diagrams.net" version="24.0.0">
  <diagram id="sinops-cycle" name="Sinops-業務フロー_日次サイクル">
    <mxGraphModel dx="{CANVAS_W}" dy="{CANVAS_H}" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="{CANVAS_W}" pageHeight="{CANVAS_H}" math="0" shadow="0">
      <root>
        {body}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
'''


def build_png(scale: int = 2):
    from PIL import Image, ImageDraw, ImageFont

    def f(size):
        return ImageFont.truetype(FONT_PATH, int(size * scale))

    img = Image.new("RGB", (CANVAS_W * scale, CANVAS_H * scale), "white")
    d = ImageDraw.Draw(img)
    S = scale

    def text(xy, s, size, fill="#1A202C", bold=False, anchor="la"):
        font = f(size)
        x, y = xy
        d.text((x * S, y * S), s, font=font, fill=fill, anchor=anchor)
        if bold:
            d.text((x * S + 1.2, y * S), s, font=font, fill=fill, anchor=anchor)

    def arrow(p0, p1, color="#1A202C", width=3, dashed=False, head=14):
        (x0, y0), (x1, y1) = p0, p1
        ang = math.atan2(y1 - y0, x1 - x0)
        bx, by = x1 - head * math.cos(ang), y1 - head * math.sin(ang)
        if dashed:
            length = math.hypot(bx - x0, by - y0)
            n = int(length // 14)
            for k in range(0, n, 2):
                a = k * 14 / length
                b = min((k + 1) * 14 / length, 1)
                d.line([((x0 + (bx - x0) * a) * S, (y0 + (by - y0) * a) * S),
                        ((x0 + (bx - x0) * b) * S, (y0 + (by - y0) * b) * S)], fill=color, width=width * S)
        else:
            d.line([(x0 * S, y0 * S), (bx * S, by * S)], fill=color, width=width * S)
        px, py = -math.sin(ang), math.cos(ang)
        hw = head * 0.55
        d.polygon([(x1 * S, y1 * S), ((bx + px * hw) * S, (by + py * hw) * S), ((bx - px * hw) * S, (by - py * hw) * S)], fill=color)

    text((X0, 30), TITLE, 24, bold=True)
    lx = CANVAS_W - X0 - 4 * 150
    for i, name in enumerate(COLORS):
        fill, stroke = COLORS[name]
        x = lx + i * 150
        d.rounded_rectangle([x * S, 34 * S, (x + 136) * S, 60 * S], radius=6 * S, fill=fill, outline=stroke, width=S)
        text((x + 68, 47), name, 12, anchor="mm")

    # 中央の枠
    cx, cy = pos(1, 1)
    x0, y0, x1, y1 = cx, cy, cx + PITCH_X + W, cy + H
    r = 12 * S
    dash, gap = 10 * S, 7 * S
    def dashed_rect():
        for (ax, ay, bx, by) in [(x0, y0, x1, y0), (x1, y0, x1, y1), (x1, y1, x0, y1), (x0, y1, x0, y0)]:
            L = math.hypot(bx - ax, by - ay) * S
            n = int(L // (dash + gap)) + 1
            ux, uy = (bx - ax) * S / L, (by - ay) * S / L
            for k in range(n):
                s0 = k * (dash + gap)
                s1 = min(s0 + dash, L)
                d.line([(ax * S + ux * s0, ay * S + uy * s0), (ax * S + ux * s1, ay * S + uy * s1)], fill="#666666", width=2 * S)
    dashed_rect()
    mx = (x0 + x1) / 2
    text((mx, cy + 28), CORE_TITLE, 18, bold=True, anchor="mm")
    for i, t in enumerate(CORE_LINES):
        text((mx, cy + 56 + i * 22), t, 13, fill="#4A5568", anchor="mm")

    # カード
    for sid, row, col, actor, head, lines in STEPS:
        fill, stroke = COLORS[actor]
        x, y = pos(row, col)
        d.rounded_rectangle([x * S, y * S, (x + W) * S, (y + H) * S], radius=10 * S, fill=fill, outline=stroke, width=2 * S)
        d.ellipse([(x + 10) * S, (y + 10) * S, (x + 32) * S, (y + 32) * S], fill="white", outline="#1A202C", width=S)
        text((x + 21, y + 21), str(sid), 13, bold=True, anchor="mm")
        text((x + 40, y + 21), head, 15, bold=True, anchor="lm")
        for i, t in enumerate(lines):
            text((x + 14, y + 44 + i * 22), "・" + t, 13)

    # 矢印とラベル
    for a, b, label, dashed in EDGES:
        p0, p1, _ = anchors(a, b)
        color = "#C53030" if dashed else "#1A202C"
        arrow(p0, p1, color=color, dashed=dashed)
        mx, my = (p0[0] + p1[0]) / 2, (p0[1] + p1[1]) / 2
        font = f(12)
        tw = d.textlength(label, font=font) / S
        horizontal = abs(p0[1] - p1[1]) < 1
        if horizontal:
            lab_xy = (mx, my - 16)
        else:
            lab_xy = (mx + 12 + tw / 2, my)
        d.rectangle([(lab_xy[0] - tw / 2 - 3) * S, (lab_xy[1] - 9) * S, (lab_xy[0] + tw / 2 + 3) * S, (lab_xy[1] + 9) * S], fill="white")
        text(lab_xy, label, 12, fill=color, anchor="mm")

    img.save(OUT_PNG)


def main():
    OUT_DRAWIO.write_text(build_drawio(), encoding="utf-8")
    print("drawio", OUT_DRAWIO)
    build_png()
    print("png", OUT_PNG)


if __name__ == "__main__":
    main()
