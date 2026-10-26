#!/usr/bin/env python3
"""review話術_中日_照着念.md → 照着念 HTML。

注音は本稿に出る語だけ。1 字は意味が一つに決まるものだけ入れる。
共用 RUBY_DICT は「用・先・表」などが衝突するため取り込まない。
"""

from __future__ import annotations

import html
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
MD_PATH = HERE / "review話術_中日_照着念.md"
OUT_PATH = HERE / "review話術_中日_照着念.html"

# (漢字, よみ)。長い語を先に照合する。
_PAIRS: list[tuple[str, str]] = [
    # 書名・人
    ("非生鮮", "ひせいせん"),
    ("詳細設計", "しょうさいせっけい"),
    ("入荷実績", "にゅうかじっせき"),
    ("入荷予定数量", "にゅうかよていすうりょう"),
    ("入荷予定側", "にゅうかよていがわ"),
    ("入荷予定日", "にゅうかよていび"),
    ("入荷予定", "にゅうかよてい"),
    ("入荷数量", "にゅうかすうりょう"),
    ("入荷日", "にゅうかび"),
    ("入荷分", "にゅうかぶん"),
    ("未入荷", "みにゅうか"),
    ("発注勧告", "はっちゅうかんこく"),
    ("翌日以降", "よくじついこう"),
    ("便区分", "びんくぶん"),
    ("岩瀬", "いわせ"),
    ("波多野", "はたの"),
    ("臧鵬仕", "ぞうほうし"),
    ("話術", "わじゅつ"),
    ("承認者", "しょうにんしゃ"),
    ("説明者", "せつめいしゃ"),
    ("着手許可", "ちゃくしゅきょか"),
    # レイヤ・ファイル
    ("西友連携", "せいゆうれんけい"),
    ("自動発注側", "じどうはっちゅうがわ"),
    ("自動発注", "じどうはっちゅう"),
    ("完了印", "かんりょういん"),
    ("差分取込", "さぶんとりこみ"),
    ("二次加工", "にじかこう"),
    ("対象外", "たいしょうがい"),
    ("対象日", "たいしょうび"),
    ("範囲外", "はんいがい"),
    ("型変換", "かたへんかん"),
    ("独自変換", "どくじへんかん"),
    ("物理名", "ぶつりめい"),
    ("論理名", "ろんりめい"),
    ("論理区分", "ろんりくぶん"),
    ("機能名", "きのうめい"),
    ("配置先", "はいちさき"),
    ("抽出条件", "ちゅうしゅつじょうけん"),
    ("抽出失敗", "ちゅうしゅつしっぱい"),
    ("抽出元", "ちゅうしゅつもと"),
    ("送信元", "そうしんもと"),
    ("受信先", "じゅしんさき"),
    ("受信側", "じゅしんがわ"),
    ("設計条件", "せっけいじょうけん"),
    ("改訂履歴", "かいていりれき"),
    ("改訂日", "かいていび"),
    ("指定時刻", "していじこく"),
    ("起動欄", "きどうらん"),
    ("採用締切", "さいようしめきり"),
    # 項目
    ("仕入伝票番号", "しいれでんぴょうばんごう"),
    ("仕入発注", "しいれはっちゅう"),
    ("推奨発注数量", "すいしょうはっちゅうすうりょう"),
    ("発注予定日付", "はっちゅうよていひづけ"),
    ("発注単位数", "はっちゅうたんいすう"),
    ("発注単位", "はっちゅうたんい"),
    ("発注日付", "はっちゅうひづけ"),
    ("発注統合", "はっちゅうとうごう"),
    ("発注残", "はっちゅうざん"),
    ("発注日", "はっちゅうび"),
    ("納品予定日付", "のうひんよていひづけ"),
    ("納品日付", "のうひんひづけ"),
    ("納品日", "のうひんび"),
    ("伝票番号", "でんぴょうばんごう"),
    ("仕入先", "しいれさき"),
    ("店舗", "てんぽ"),
    ("商品", "しょうひん"),
    ("仕入", "しいれ"),
    ("中分類", "ちゅうぶんるい"),
    ("固定長", "こていちょう"),
    ("文字列", "もじれつ"),
    ("小文字", "こもじ"),
    ("大文字", "おおもじ"),
    ("更新日", "こうしんび"),
    ("業務日", "ぎょうむび"),
    ("業務上", "ぎょうむじょう"),
    ("非日配", "ひにっぱい"),
    ("日配", "にっぱい"),
    ("一部確定", "いちぶかくてい"),
    ("未確定", "みかくてい"),
    ("確認事項", "かくにんじこう"),
    ("確認済み", "かくにんずみ"),
    ("確認中", "かくにんちゅう"),
    ("未確認", "みかくにん"),
    ("必須欠落", "ひっすけつらく"),
    ("取得失敗", "しゅとくしっぱい"),
    ("取込失敗", "とりこみしっぱい"),
    ("店間移動", "てんかんいどう"),
    ("符号付き", "ふごうつき"),
    ("整数倍", "せいすうばい"),
    ("勧告計算日", "かんこくけいさんび"),
    ("勧告計算", "かんこくけいさん"),
    ("前日分", "ぜんじつぶん"),
    ("当日分", "とうじつぶん"),
    # 状態・動作の複合
    ("統合結果", "とうごうけっか"),
    ("統合後", "とうごうご"),
    ("全件", "ぜんけん"),
    ("万件", "まんけん"),
    ("件数", "けんすう"),
    ("本共通", "ほんきょうつう"),
    ("両本", "りょうほん"),
    ("両方", "りょうほう"),
    ("本目", "ほんめ"),
    ("点目", "てんめ"),
    ("日後", "にちご"),
    ("後勝ち", "あとがち"),
    ("切捨て", "きりすて"),
    ("非圧縮", "ひあっしゅく"),
    ("圧縮", "あっしゅく"),
    ("投影順", "とうえいじゅん"),
    ("版数", "はんすう"),
    ("表紙", "ひょうし"),
    ("見出し", "みだし"),
    ("区切り", "くぎり"),
    ("一言", "ひとこと"),
    ("一緒", "いっしょ"),
    ("以上", "いじょう"),
    ("以外", "いがい"),
    ("位置", "いち"),
    ("同じ", "おな"),
    ("直接", "ちょくせつ"),
    ("先頭", "せんとう"),
    ("実際", "じっさい"),
    ("具体的", "ぐたいてき"),
    ("二次", "にじ"),
    # よく出る 2 字
    ("本日", "ほんじつ"),
    ("今日", "きょう"),
    ("昨日", "きのう"),
    ("当日", "とうじつ"),
    ("毎日", "まいにち"),
    ("未来", "みらい"),
    ("日常", "にちじょう"),
    ("入荷", "にゅうか"),
    ("実績", "じっせき"),
    ("予定", "よてい"),
    ("発注", "はっちゅう"),
    ("勧告", "かんこく"),
    ("生鮮", "せいせん"),
    ("西友", "せいゆう"),
    ("基幹", "きかん"),
    ("連携", "れんけい"),
    ("共通", "きょうつう"),
    ("対象", "たいしょう"),
    ("変換", "へんかん"),
    ("選別", "せんべつ"),
    ("編集", "へんしゅう"),
    ("項目", "こうもく"),
    ("中身", "なかみ"),
    ("中断", "ちゅうだん"),
    ("固定", "こてい"),
    ("出力", "しゅつりょく"),
    ("採用", "さいよう"),
    ("優先", "ゆうせん"),
    ("数量", "すうりょう"),
    ("確定", "かくてい"),
    ("抽出", "ちゅうしゅつ"),
    ("製造", "せいぞう"),
    ("設計", "せっけい"),
    ("統合", "とうごう"),
    ("想定", "そうてい"),
    ("概要", "がいよう"),
    ("時刻", "じこく"),
    ("空白", "くうはく"),
    ("論理", "ろんり"),
    ("物理", "ぶつり"),
    ("表記", "ひょうき"),
    ("必須", "ひっす"),
    ("差分", "さぶん"),
    ("取込", "とりこみ"),
    ("警告", "けいこく"),
    ("番号", "ばんごう"),
    ("改訂", "かいてい"),
    ("整数", "せいすう"),
    ("未決", "みけつ"),
    ("回答", "かいとう"),
    ("最後", "さいご"),
    ("失敗", "しっぱい"),
    ("独自", "どくじ"),
    ("半分", "はんぶん"),
    ("部分", "ぶぶん"),
    ("原因", "げんいん"),
    ("解消", "かいしょう"),
    ("起動", "きどう"),
    ("集計", "しゅうけい"),
    ("改行", "かいぎょう"),
    ("文字", "もじ"),
    ("根拠", "こんきょ"),
    ("欠品", "けっぱん"),
    ("削除", "さくじょ"),
    ("相手", "あいて"),
    ("理解", "りかい"),
    ("分担", "ぶんたん"),
    ("場所", "ばしょ"),
    ("任意", "にんい"),
    ("区分", "くぶん"),
    ("文言", "もんごん"),
    ("影響", "えいきょう"),
    ("判定", "はんてい"),
    ("障害", "しょうがい"),
    ("倍数", "ばいすう"),
    ("部門", "ぶもん"),
    ("階層", "かいそう"),
    ("候補", "こうほ"),
    ("転送", "てんそう"),
    ("既存", "きそん"),
    ("必要", "ひつよう"),
    ("争点", "そうてん"),
    ("一致", "いっち"),
    ("締切", "しめきり"),
    ("転記", "てんき"),
    ("本文", "ほんぶん"),
    ("反対", "はんたい"),
    ("新規", "しんき"),
    ("方針", "ほうしん"),
    ("担当", "たんとう"),
    ("承認", "しょうにん"),
    ("確認", "かくにん"),
    ("処理", "しょり"),
    ("口頭", "こうとう"),
    ("訂正", "ていせい"),
    ("用途", "ようと"),
    ("所要", "しょよう"),
    ("全文", "ぜんぶん"),
    ("時間", "じかん"),
    ("漢字", "かんじ"),
    ("英語", "えいご"),
    ("数字", "すうじ"),
    ("日文", "にちぶん"),
    ("投影", "とうえい"),
    ("着地", "ちゃくち"),
    ("名前", "なまえ"),
    ("中日", "ちゅうにち"),
    ("審議", "しんぎ"),
    # 1 字。この稿では読みが割れないものだけ。
    ("空", "から"),
    ("桁", "けた"),
    ("正", "せい"),
    ("無", "な"),
    ("見", "み"),
    ("取", "と"),
    ("送", "おく"),
    ("入", "い"),
    ("残", "のこ"),
    ("切", "き"),
    ("置", "お"),
    ("書", "か"),
    ("付", "つ"),
    ("変", "か"),
    ("直", "なお"),
    ("違", "ちが"),
    ("進", "すす"),
    ("閉", "と"),
    ("話", "はな"),
    ("通", "とお"),
    ("次", "つぎ"),
    ("先", "さき"),
    ("層", "そう"),
    ("件", "けん"),
    ("約", "やく"),
    ("時", "じ"),
    ("月", "がつ"),
    ("本", "ほん"),
    ("点", "てん"),
    ("版", "ばん"),
    ("番", "ばん"),
    ("名", "めい"),
    ("後", "あと"),
    ("上", "じょう"),
    ("中", "なか"),
    ("側", "がわ"),
    ("元", "もと"),
    ("表", "ひょう"),
    ("順", "じゅん"),
    ("指", "さ"),
    ("開", "ひら"),
    ("続", "つづ"),
    ("読", "よ"),
    ("替", "か"),
    ("洗", "あら"),
    ("従", "したが"),
    ("渡", "わた"),
    ("店", "てん"),
    ("囲", "かこ"),
    ("揃", "そろ"),
    ("扱", "あつか"),
    ("飛", "と"),
    ("対", "つい"),
    ("向", "む"),
    ("古", "ふる"),
    ("含", "ふく"),
    ("旧", "ふる"),
    ("欄", "らん"),
    ("逆", "ぎゃく"),
    ("掛", "か"),
    ("算", "ざん"),
    ("倍", "ばい"),
    ("丸", "まる"),
    ("選", "えら"),
    ("重", "かさ"),
    ("使", "つか"),
    ("合", "あ"),
    ("移", "うつ"),
    ("引", "ひ"),
    ("黙", "だま"),
    ("作", "つく"),
    ("間", "ま"),
    ("例", "れい"),
    ("足", "た"),
    ("今", "いま"),
    ("言", "い"),
    ("決", "き"),
    ("所", "しょ"),
    ("別", "べつ"),
    ("各", "かく"),
    ("全", "ぜん"),
    ("運", "はこ"),
    ("落", "お"),
    ("戻", "もど"),
    ("身", "み"),
]


def _build_dict(pairs: list[tuple[str, str]]) -> dict[str, str]:
    out: dict[str, str] = {}
    for kanji, yomi in pairs:
        if kanji in out and out[kanji] != yomi:
            raise SystemExit(f"読み衝突: {kanji} {out[kanji]} / {yomi}")
        if len(kanji) < 1:
            raise SystemExit(f"空キー: {kanji}")
        out[kanji] = yomi
    return out


RUBY = _build_dict(_PAIRS)
# 次の1字がこの仮名なら、そのキーは動詞の送り仮名なので振らない。
# 例: 読み上げる の「上」は「じょう」ではない。
_REJECT_NEXT: dict[str, frozenset[str]] = {
    "上": frozenset("げ"),
}
BY_FIRST: dict[str, list[tuple[str, str]]] = {}
for _k, _v in sorted(RUBY.items(), key=lambda kv: len(kv[0]), reverse=True):
    BY_FIRST.setdefault(_k[0], []).append((_k, _v))

_PAREN_RE = re.compile(r"([\u4e00-\u9fff々〆ヵヶ]+)（([ぁ-んァ-ンー]+)）")
_KANJI = frozenset(chr(c) for c in range(0x4E00, 0xA000)) | set("々〆ヵヶ")


def esc(s: str) -> str:
    return html.escape(s, quote=True)


def _ruby(kanji: str, yomi: str) -> str:
    return f"<ruby>{esc(kanji)}<rt>{esc(yomi)}</rt></ruby>"


def _dict_annotate(plain: str) -> str:
    out: list[str] = []
    i = 0
    n = len(plain)
    while i < n:
        ch = plain[i]
        if ch in _KANJI:
            matched = None
            yomi = ""
            for key, reading in BY_FIRST.get(ch, ()):
                if not plain.startswith(key, i):
                    continue
                nxt = plain[i + len(key)] if i + len(key) < n else ""
                if nxt in _REJECT_NEXT.get(key, ()):
                    continue
                matched = key
                yomi = reading
                break
            if matched:
                out.append(_ruby(matched, yomi))
                i += len(matched)
                continue
            out.append(esc(ch))
            i += 1
            continue
        j = i + 1
        while j < n and plain[j] not in _KANJI:
            j += 1
        out.append(esc(plain[i:j]))
        i = j
    return "".join(out)


def furigana_html(s: str) -> str:
    """括弧注音を先に ruby にし、残りを辞典で埋める。"""
    out: list[str] = []
    last = 0
    for m in _PAREN_RE.finditer(s):
        out.append(_dict_annotate(s[last : m.start()]))
        out.append(_ruby(m.group(1), m.group(2)))
        last = m.end()
    out.append(_dict_annotate(s[last:]))
    return "".join(out)


def inline_md(s: str, *, jp: bool) -> str:
    parts = re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", s)
    buf: list[str] = []
    for p in parts:
        if len(p) >= 4 and p.startswith("**") and p.endswith("**"):
            inner = p[2:-2]
            body = furigana_html(inner) if jp else esc(inner)
            buf.append(f"<strong>{body}</strong>")
        elif len(p) >= 2 and p.startswith("`") and p.endswith("`"):
            buf.append(f"<code>{esc(p[1:-1])}</code>")
        else:
            buf.append(furigana_html(p) if jp else esc(p))
    return "".join(buf)


def paras(block: str, *, jp: bool) -> str:
    lines: list[str] = []
    for raw in block.split("\n"):
        line = raw.strip()
        if line.startswith(">"):
            line = line[1:].strip()
        if line:
            lines.append(line)
    return "".join(f"<p>{inline_md(ln, jp=jp)}</p>" for ln in lines)


def visible_text(fragment: str) -> str:
    fragment = re.sub(r"<rt>.*?</rt>", "", fragment)
    fragment = re.sub(r"<[^>]+>", "", fragment)
    return html.unescape(fragment)


def _norm_src(block: str) -> str:
    lines: list[str] = []
    for raw in block.split("\n"):
        line = raw.strip()
        if line.startswith(">"):
            line = line[1:].strip()
        if line:
            line = line.replace("**", "").replace("`", "")
            lines.append(_PAREN_RE.sub(r"\1", line))
    return "".join(lines)


def _check_roundtrip(label: str, block: str, fragment: str, *, jp: bool) -> None:
    got = visible_text(fragment)
    exp = _norm_src(block)
    if got != exp:
        raise SystemExit(f"{label} の本文が原稿と一致しません\nEXP={exp}\nGOT={got}")
    if jp and "（" in got:
        raise SystemExit(f"{label} に括弧注音が残っています")


# 中国語ラベルなど、日本語の読みを振らない語。
_BARE_OK = {"故事线"}


def _bare_compounds(fragment: str) -> list[str]:
    """ruby と code の外に残った 2 字以上の漢字。"""
    text = re.sub(r"<ruby>.*?</ruby>", "\n", fragment, flags=re.S)
    text = re.sub(r"<code>.*?</code>", "\n", text, flags=re.S)
    text = re.sub(r"<[^>]+>", "", text)
    found = re.findall(r"[\u4e00-\u9fff々〆ヵヶ]{2,}", text)
    return [run for run in found if run not in _BARE_OK]


def _md_table(lines: list[str]) -> str:
    rows: list[list[str]] = []
    for line in lines:
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if not cells or set("".join(cells)) <= set("-: "):
            continue
        rows.append(cells)
    if not rows:
        return ""
    head, body = rows[0], rows[1:]
    th = "".join(f"<th>{inline_md(c, jp=True)}</th>" for c in head)
    trs = []
    for row in body:
        tds = "".join(f"<td>{inline_md(c, jp=True)}</td>" for c in row)
        trs.append(f"<tr>{tds}</tr>")
    return (
        "<table><thead><tr>"
        + th
        + "</tr></thead><tbody>"
        + "".join(trs)
        + "</tbody></table>"
    )


def render_preamble(pre: str) -> str:
    lines = pre.splitlines()
    if lines and lines[0].startswith("# "):
        lines = lines[1:]
    chunks: list[str] = []
    i = 0
    buf: list[str] = []
    while i < len(lines):
        line = lines[i]
        if line.startswith("|"):
            if buf:
                chunks.append(_prose_block(buf))
                buf = []
            table_lines = []
            while i < len(lines) and lines[i].startswith("|"):
                table_lines.append(lines[i])
                i += 1
            chunks.append(_md_table(table_lines))
            continue
        if line.strip() == "---":
            i += 1
            continue
        buf.append(line)
        i += 1
    if buf:
        chunks.append(_prose_block(buf))
    return "\n".join(c for c in chunks if c)


def _prose_block(lines: list[str]) -> str:
    paras_out: list[str] = []
    for raw in lines:
        line = raw.strip()
        if not line:
            continue
        paras_out.append(f"<p>{inline_md(line, jp=True)}</p>")
    return "".join(paras_out)


def parse_short(section: str) -> tuple[str, str]:
    m = re.search(
        r"### 中文\n+(.*?)\n### 日本語\n+(.*)\Z",
        section,
        re.S,
    )
    if not m:
        raise SystemExit("2 分版の中日が見つかりません")
    return m.group(1), m.group(2)


def parse_full(section: str) -> list[tuple[str, str, str, str]]:
    intro, *chunks = re.split(r"\n(?=### )", section)
    if not chunks:
        raise SystemExit("全文の節がありません")
    out: list[tuple[str, str, str, str]] = []
    for chunk in chunks:
        m = re.match(
            r"### (.+?)\n+(?:(【[^】]*】)\n+)?#### 中文\n+(.*?)\n#### 日本語\n+(.*)\Z",
            chunk,
            re.S,
        )
        if not m:
            raise SystemExit(f"節を解釈できません: {chunk[:80]!r}")
        out.append((m.group(1).strip(), m.group(2) or "", m.group(3), m.group(4)))
    if len(out) != 6:
        raise SystemExit(f"全文は 6 節のはずです: {len(out)}")
    return out


PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>非生鮮4本　詳細設計レビュー　話術（中日・照着念）</title>
<style>
:root {
  --bg: #f6f3ee;
  --ink: #1c1917;
  --muted: #78716c;
  --cn: #0f766e;
  --jp: #b45309;
  --card: #fffdf9;
  --line: #e7e5e4;
  --accent: #0d9488;
}
* { box-sizing: border-box; }
body {
  margin: 0; font-family: "Hiragino Sans", "Noto Sans JP", "PingFang SC",
    "Microsoft YaHei", sans-serif;
  background: var(--bg); color: var(--ink); line-height: 1.75;
}
header.top {
  position: sticky; top: 0; z-index: 20;
  background: rgba(246,243,238,.94); backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--line); padding: .75rem 1rem;
}
h1 { font-size: 1.15rem; margin: 0 0 .25rem; line-height: 1.45; }
.meta { color: var(--muted); font-size: .85rem; }
.sync { color: var(--accent); font-size: .8rem; margin-top: .2rem; }
.toolbar { display: flex; flex-wrap: wrap; gap: .4rem; margin-top: .55rem; }
.toolbar button, .toc a {
  border: 1px solid var(--line); background: #fff; border-radius: 999px;
  padding: .25rem .7rem; font-size: .8rem; cursor: pointer; color: var(--ink);
  text-decoration: none;
}
.toolbar button:hover, .toc a:hover { border-color: var(--accent); }
.toc { display: flex; flex-wrap: wrap; gap: .35rem; margin-top: .45rem; }
main { max-width: 1100px; margin: 0 auto; padding: 1rem; }
.card {
  background: var(--card); border: 1px solid var(--line); border-radius: 14px;
  padding: 1rem 1.1rem; margin: 0 0 1rem;
}
.card h2 { margin: 0 0 .7rem; font-size: 1.05rem; line-height: 1.5; }
.slide-num {
  display: inline-block; background: var(--accent); color: #fff;
  border-radius: 6px; padding: .05rem .45rem; font-size: .75rem; margin-right: .4rem;
}
.pair { display: grid; gap: .75rem; }
.pair.side { grid-template-columns: 1fr 1fr; }
.panel {
  border: 1px solid var(--line); border-radius: 10px; padding: .75rem .85rem;
  background: #fff;
}
.badge {
  display: inline-block; font-size: .7rem; font-weight: 700;
  border-radius: 999px; padding: .1rem .5rem; margin-bottom: .4rem;
}
.badge.cn { background: #ccfbf1; color: var(--cn); }
.badge.jp { background: #ffedd5; color: var(--jp); }
.cn-panel p, .jp-panel p, .instr p { margin: 0 0 .65rem; }
.cn-panel p:last-child, .jp-panel p:last-child, .instr p:last-child { margin-bottom: 0; }
.jp-panel { font-size: 1.05rem; }
ruby { ruby-align: center; }
rt { font-size: .55em; color: #a16207; font-weight: 500; }
body.big { font-size: 1.14rem; }
body.big .jp-panel { font-size: 1.22rem; }
body.hide-cn .lang-cn-block { display: none !important; }
body.hide-jp .lang-jp-block { display: none !important; }
table { width: 100%; border-collapse: collapse; font-size: .92rem; margin: .4rem 0 .8rem; }
th, td { border: 1px solid var(--line); padding: .45rem .55rem; vertical-align: top; }
th { background: #fafaf9; text-align: left; }
th:nth-child(1), td:nth-child(1),
th:nth-child(3), td:nth-child(3),
th:nth-child(4), td:nth-child(4) { white-space: nowrap; }
.hint { color: var(--muted); font-size: .88rem; margin: 0 0 .7rem; }
.cue {
  background: #f5f5f4; border-left: 3px solid var(--accent);
  padding: .4rem .7rem; margin: 0 0 .75rem; color: var(--muted); font-size: .86rem;
}
.cue strong { color: var(--ink); font-weight: 700; }
.note { font-size: .8rem; color: var(--muted); margin-top: 1rem; text-align: center; }
code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: .88em; }
p, td, h1, h2 { overflow-wrap: anywhere; }
.instr p { font-size: .92rem; }
@media (max-width: 859px) {
  .pair.side { grid-template-columns: 1fr; }
}
@media print {
  header.top { position: static; }
  .toolbar, .toc { display: none; }
  body { background: #fff; }
}
</style>
</head>
<body>
<header class="top">
<h1>__H1__</h1>
<div class="meta">全文约 8 分钟　｜　时间不够只读 2 分版　｜　评审 岩瀬・波多野　｜　承认 岩瀬　｜　说明 臧鵬仕</div>
<div class="sync">【】里的投影提示不念。振假名只标这篇稿子里的读法。文件名和字符码以 07 为准，处理以 02 / 05 Layer① 为准。</div>
<div class="toolbar">
<button type="button" onclick="showBoth()">中日両方</button>
<button type="button" onclick="showCn()">中文のみ</button>
<button type="button" onclick="showJp()">日本語のみ</button>
<button type="button" onclick="toggleBig()">文字大きく</button>
<button type="button" onclick="toggleSide()">左右並び</button>
</div>
<nav class="toc">
<a href="#sec-guide">読み方</a>
<a href="#sec-short">2分</a>
<a href="#sec-1">1 开场</a>
<a href="#sec-2">2 共通</a>
<a href="#sec-3">3 実績</a>
<a href="#sec-4">4 予定</a>
<a href="#sec-5">5 勧告</a>
<a href="#sec-6">6 决定</a>
</nav>
</header>
<main>
__MAIN__
<p class="note">注音：原稿の（かな）＋本稿だけの辞典　｜　源 MD と本文一致を生成時に照合</p>
</main>
<script>
function showBoth(){ document.body.classList.remove('hide-cn','hide-jp'); }
function showCn(){ document.body.classList.add('hide-jp'); document.body.classList.remove('hide-cn'); }
function showJp(){ document.body.classList.add('hide-cn'); document.body.classList.remove('hide-jp'); }
function toggleBig(){ document.body.classList.toggle('big'); }
function toggleSide(){ document.querySelectorAll('.pair').forEach(el => el.classList.toggle('side')); }
if (window.matchMedia('(min-width: 860px)').matches) {
  document.querySelectorAll('.pair').forEach(el => el.classList.add('side'));
}
</script>
</body>
</html>
"""


def pair_html(cn: str, jp: str, label: str) -> str:
    cn_html = paras(cn, jp=False)
    jp_html = paras(jp, jp=True)
    _check_roundtrip(label + "/中文", cn, cn_html, jp=False)
    _check_roundtrip(label + "/日本語", jp, jp_html, jp=True)
    bare = _bare_compounds(jp_html)
    if bare:
        raise SystemExit(f"{label} に未注音の複合漢字: {bare}")
    return (
        '<div class="pair">'
        f'<div class="panel cn-panel lang-cn-block"><span class="badge cn">中文</span>{cn_html}</div>'
        f'<div class="panel jp-panel lang-jp-block"><span class="badge jp">日本語</span>{jp_html}</div>'
        "</div>"
    )


def _self_check() -> None:
    sample = furigana_html("そのまま読み上げる。GCS 上は gz。")
    if sample.count("じょう") != 1 or "み<ruby>上" in sample:
        raise SystemExit("読み上げるの注音が違います: " + sample)


def main() -> None:
    _self_check()
    text = MD_PATH.read_text(encoding="utf-8")
    parts = re.split(r"\n(?=## )", text)
    if len(parts) != 3:
        raise SystemExit(f"節の数が想定外です: {len(parts)}")
    preamble, short, full = parts
    cn_s, jp_s = parse_short(short)
    sections = parse_full(full)

    blocks: list[str] = []
    guide = render_preamble(preamble)
    bare_guide = _bare_compounds(guide)
    if bare_guide:
        raise SystemExit(f"読み方に未注音の複合漢字: {bare_guide}")
    blocks.append(
        '<section class="card" id="sec-guide"><h2>'
        + furigana_html("読み方・投影順")
        + '</h2><div class="instr">'
        + guide
        + "</div></section>"
    )

    blocks.append(
        '<section class="card" id="sec-short"><h2><span class="slide-num">2分</span>'
        + furigana_html("時間が足りないとき")
        + "</h2>"
        + pair_html(cn_s, jp_s, "2分版")
        + "</section>"
    )

    for idx, (title, cue, cn, jp) in enumerate(sections, start=1):
        cue_html = ""
        if cue:
            cue_body = furigana_html(cue)
            bare_cue = _bare_compounds(cue_body)
            if bare_cue:
                raise SystemExit(f"{title} の投影指示に未注音: {bare_cue}")
            cue_html = (
                '<p class="cue"><strong>投影・読まない</strong>　'
                + cue_body
                + "</p>"
            )
        num = title.split(".", 1)[0].strip()
        blocks.append(
            f'<section class="card" id="sec-{idx}">'
            f'<h2><span class="slide-num">{esc(num)}</span>{furigana_html(title)}</h2>'
            f"{cue_html}"
            f"{pair_html(cn, jp, title)}"
            "</section>"
        )

    h1 = furigana_html("非生鮮 4 本　詳細設計レビュー　話術（中日・照着念）")
    page = PAGE.replace("__H1__", h1).replace("__MAIN__", "\n".join(blocks))
    OUT_PATH.write_text(page, encoding="utf-8")
    print(f"wrote {OUT_PATH} ({len(page)} chars)")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as exc:
        print(exc, file=sys.stderr)
        raise
