#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""話術 md の中文／日文を左右分割 HTML に流し込む。"""
from __future__ import annotations

import re
from pathlib import Path

from _add_furigana import process_html, process_md

ROOT = Path(__file__).resolve().parent

PAIRS = [
    (
        "標準案説明話術_3分_中日双语_20260924.md",
        "標準案説明話術_3分_中日双语_20260924.html",
        "生鮮IF増分・標準案　話術（PPT順）",
        [
            "第1页から第11页へ進める。戻らない。青い欄（日文）だけ読む。",
            "左＝中文／右＝日文。先に前提とリスク（再見積）。157.5人日／¥5,000,000（税抜）。第10页は控えなので読み上げない。",
        ],
    ),
    (
        "標準案Excel説明話術_シート順_中日双语_20260929.md",
        "標準案Excel説明話術_シート順_中日双语_20260929.html",
        "標準案 Excel　説明話術（シート順）",
        [
            "タブの順に進める。青い欄（日文）だけ読む。シート名は開く合図で、読まない。",
            "左＝中文／右＝日文。サマリー → 見積前提条件 → A_新規差分IF → B_横断 → C_PJ管理 → 対既存差分 → 算定根拠",
            "見積前提条件で前提とリスクを言い切ってから明細へ。157.5人日／¥5,000,000（税抜）　2026-09-30",
        ],
    ),
    (
        "標準案Excel説明話術_2-3分_中日双语_20260924.md",
        "標準案Excel説明話術_2-3分_中日双语_20260924.html",
        "標準案 Excel　簡略話術（2〜3分）",
        [
            "タブの順に進める。青い欄（日文）だけ読む。シート名は開く合図で、読まない。",
            "左＝中文／右＝日文。2〜3分でも前提と再見積の線は省略しない。157.5人日／¥5,000,000（税抜）　2026-09-30",
        ],
    ),
]

CSS = """\
  :root { --navy:#1a365d; --orange:#c05621; --bg:#f7fafc; --line:#e2e8f0; }
  * { box-sizing: border-box; }
  body { margin:0; font-family:"Hiragino Sans","Noto Sans JP","Yu Gothic","Microsoft YaHei",sans-serif; background:var(--bg); color:#1a202c; line-height:1.75; }
  header { background:linear-gradient(135deg,var(--navy),#2c5282); color:#fff; padding:1.2rem 1.1rem; }
  header h1 { margin:0 0 .35rem; font-size:1.25rem; }
  header p { margin:.2rem 0; font-size:.9rem; opacity:.92; }
  main { max-width:1100px; margin:0 auto; padding:1rem 1rem 2.5rem; }
  section { background:#fff; border:1px solid var(--line); border-radius:8px; margin:1rem 0; overflow:hidden; }
  h2 { margin:0; padding:.65rem 1rem; background:var(--navy); color:#fff; font-size:1rem; }
  .cols { display:flex; align-items:stretch; }
  .cn, .jp { flex:1 1 50%; padding:.85rem 1rem; min-width:0; }
  .cn { background:#f0fff4; border-right:1px solid var(--line); }
  .jp { background:#ebf8ff; }
  .lang { font-size:.72rem; font-weight:700; color:#4a5568; margin-bottom:.3rem; }
  p { margin:.35rem 0; }
  rt { font-size:.55em; color:var(--orange); }
  strong { color:var(--navy); }
  .skip h2 { background:#718096; }
  @media (max-width: 800px) {
    .cols { flex-direction: column; }
    .cn { border-right: none; border-bottom:1px solid var(--line); }
  }
"""


def split_sections(md: str) -> list[tuple[str, str, str]]:
    body = md.split("\n---\n", 1)[-1]
    chunks = re.split(r"\n## ", body)
    out: list[tuple[str, str, str]] = []
    for ch in chunks:
        ch = ch.strip()
        if not ch:
            continue
        title, _, sec_body = ch.partition("\n")
        title = title.strip().lstrip("# ").strip()
        m = re.match(r"［(.+?)］(.*)", title)
        if m:
            tail = m.group(2).strip()
            title = f"{m.group(1)}　{tail}" if tail else m.group(1)
        else:
            title = title.replace("［", "").replace("］", "")
        cn_m = re.search(r"### 中文\n(.*?)(?=\n### |\Z)", sec_body, re.S)
        jp_m = re.search(r"### 日文\n(.*?)(?=\n### |\Z)", sec_body, re.S)
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
        lines.append(f"      <p>{ln}</p>")
    return "\n".join(lines)


def build_html(title: str, hints: list[str], sections: list[tuple[str, str, str]]) -> str:
    parts = [
        "<!DOCTYPE html>",
        '<html lang="ja">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        f"<title>{title}</title>",
        "<style>",
        CSS.rstrip(),
        "</style>",
        "</head>",
        "<body>",
        "<header>",
        f"  <h1>{title}</h1>",
    ]
    for h in hints:
        parts.append(f"  <p>{h}</p>")
    parts.extend(["</header>", "<main>", ""])
    for sec_title, cn, jp in sections:
        skip = ' class="skip"' if "話し方" in sec_title or "読まない" in sec_title else ""
        parts.append(f"<section{skip}>")
        parts.append(f"  <h2>{sec_title}</h2>")
        parts.append('  <div class="cols">')
        parts.append('    <div class="cn"><div class="lang">中文</div>')
        parts.append(to_ps(cn))
        parts.append("    </div>")
        parts.append('    <div class="jp"><div class="lang">日文（読み上げ）</div>')
        parts.append(to_ps(jp))
        parts.append("    </div>")
        parts.append("  </div>")
        parts.append("</section>")
        parts.append("")
    parts.extend(["</main>", "</body>", "</html>", ""])
    return "\n".join(parts)


def main() -> None:
    for md_name, html_name, title, hints in PAIRS:
        md_path = ROOT / md_name
        html_path = ROOT / html_name
        process_md(md_path)
        secs = split_sections(md_path.read_text(encoding="utf-8"))
        html_path.write_text(build_html(title, hints, secs), encoding="utf-8")
        process_html(html_path)
        print("built", html_name)


if __name__ == "__main__":
    main()
