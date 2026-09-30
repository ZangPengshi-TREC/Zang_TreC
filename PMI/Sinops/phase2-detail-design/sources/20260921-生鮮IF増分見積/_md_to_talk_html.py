#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""話術 md の中文／日文を既存 HTML の枠に流し込む。"""
from __future__ import annotations

import re
from pathlib import Path

from _add_furigana import process_html, process_md

ROOT = Path(__file__).resolve().parent

PAIRS = [
    (
        "標準案説明話術_3分_中日双语_20260924.md",
        "標準案説明話術_3分_中日双语_20260924.html",
    ),
    (
        "標準案Excel説明話術_シート順_中日双语_20260929.md",
        "標準案Excel説明話術_シート順_中日双语_20260929.html",
    ),
    (
        "標準案Excel説明話術_2-3分_中日双语_20260924.md",
        "標準案Excel説明話術_2-3分_中日双语_20260924.html",
    ),
]


def split_sections(md: str) -> list[tuple[str, str, str]]:
    body = md.split("\n---\n", 1)[-1]
    chunks = re.split(r"\n## ", body)
    out: list[tuple[str, str, str]] = []
    for ch in chunks:
        ch = ch.strip()
        if not ch:
            continue
        title, _, body = ch.partition("\n")
        title = title.strip().lstrip("# ").strip()
        m = re.match(r"［(.+?)］(.*)", title)
        if m:
            tail = m.group(2).strip()
            title = f"{m.group(1)}　{tail}" if tail else m.group(1)
        else:
            title = title.replace("［", "").replace("］", "")
        cn_m = re.search(r"### 中文\n(.*?)(?=\n### |\Z)", body, re.S)
        jp_m = re.search(r"### 日文\n(.*?)(?=\n### |\Z)", body, re.S)
        cn = (cn_m.group(1) if cn_m else "").strip()
        jp = (jp_m.group(1) if jp_m else "").strip()
        cn = re.sub(r"\n?---\s*$", "", cn).strip()
        jp = re.sub(r"\n?---\s*$", "", jp).strip()
        out.append((title, cn, jp))
    return out


def to_ps(text: str) -> str:
    lines: list[str] = []
    for raw in text.split("\n"):
        ln = raw.strip()
        if not ln:
            continue
        ln = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", ln)
        lines.append(f"    <p>{ln}</p>")
    return "\n".join(lines)


def rebuild(html_path: Path, sections: list[tuple[str, str, str]]) -> None:
    src = html_path.read_text(encoding="utf-8")
    head, _, _ = src.partition("<main>")
    # keep header; refresh one hint line if present
    if "3分_中日" in html_path.name:
        head = head.replace(
            "160.5人日／¥5,120,000（税抜）。第10页は控えなので読み上げない。",
            "先に前提とリスク（再見積）。160.5人日／¥5,120,000（税抜）。第10页は控えなので読み上げない。",
        )
    if "見積前提条件で前提" not in head and "シート順" in html_path.name:
        head = re.sub(
            r"<p>サマリー →.*?</p>",
            "<p>サマリー → 見積前提条件 → A_新規差分IF → B_横断 → C_PJ管理 → 対既存差分 → 算定根拠</p>\n  <p>見積前提条件で前提とリスクを言い切ってから明細へ。</p>",
            head,
            count=1,
            flags=re.S,
        )
        head = head.replace("Rev.i　2026-09-29", "2026-09-30")
    if "省略しない" not in head and "2-3分" in html_path.name:
        head = re.sub(
            r"<p>サマリー →.*?</p>",
            "<p>サマリー → 見積前提条件 → A → B → C → 対既存差分 → 算定根拠</p>\n  <p>2〜3分でも前提と再見積の線は省略しない。</p>",
            head,
            count=1,
            flags=re.S,
        )
        head = head.replace("2026-09-29 マスタデータ判定反映", "2026-09-30")

    parts = [head.rstrip(), "<main>", ""]
    for title, cn, jp in sections:
        skip = ' class="skip"' if "話し方" in title or "読まない" in title else ""
        h2 = title
        parts.append(f"<section{skip}>")
        parts.append(f"  <h2>{h2}</h2>")
        parts.append('  <div class="cn"><div class="lang">中文</div>')
        parts.append(to_ps(cn))
        parts.append("  </div>")
        parts.append('  <div class="jp"><div class="lang">日文（読み上げ）</div>')
        parts.append(to_ps(jp))
        parts.append("  </div>")
        parts.append("</section>")
        parts.append("")
    parts.append("</main>")
    parts.append("</body>")
    parts.append("</html>")
    html_path.write_text("\n".join(parts) + "\n", encoding="utf-8")


def main() -> None:
    for md_name, html_name in PAIRS:
        md_path = ROOT / md_name
        html_path = ROOT / html_name
        process_md(md_path)
        secs = split_sections(md_path.read_text(encoding="utf-8"))
        rebuild(html_path, secs)
        process_html(html_path)


if __name__ == "__main__":
    main()
