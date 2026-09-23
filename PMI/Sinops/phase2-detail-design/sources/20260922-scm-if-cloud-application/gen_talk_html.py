#!/usr/bin/env python3
"""从 申請話術_中日_照着念.md 生成带假名注音的 HTML。

注音来源：
1. 原稿括弧 漢字（かな）
2. 词典最长匹配（补全未括弧的汉字）
"""
from __future__ import annotations

from pathlib import Path
import html
import re

ROOT = Path(__file__).resolve().parent
MD_PATH = ROOT / "申請話術_中日_照着念.md"
OUT_PATH = ROOT / "申請話術_中日_照着念.html"

# 词典：长词优先（最长匹配）
RUBY_DICT: dict[str, str] = {

    "元データ": "もとでーた",
    "本PJ": "ほんぴーじぇー",
    "一方向": "いちほうこう",
    "一方向納品": "いちほうこうのうひん",
    "連携経路": "れんけいけいろ",
    "ステートレス": "すてーとれす",
    "付帯": "ふたい",
    "完了マーカー": "かんりょうまーかー",
    "外部ソート領域": "がいぶそーとりょういき",
    "外部ソート": "がいぶそーと",
    "領域": "りょういき",
    "利用": "りよう",
    "用": "よう",
    "管理費あり": "かんりひあり",
    "算定日": "さんていび",
    # 复合／专名
    "自動補充連携": "じどうほじゅうれんけい",
    "統一連携面": "とういつれんけいめん",
    "西友連携": "せいゆうれんけい",
    "別途申請": "べっとしんせい",
    "今回申請": "こんかいしんせい",
    "本機永続": "ほんきえいぞく",
    "実行前提": "じっこうぜんてい",
    "実行基盤": "じっこうきばん",
    "費用初算": "ひようしょさん",
    "申請構成図": "しんせいこうせいず",
    "構成一覧": "こうせいいちらん",
    "費用試算": "ひようしさん",
    "費用比較": "ひようひかく",
    "日次実行": "にちじじっこう",
    "処理方針": "しょりほうしん",
    "想定質問": "そうていしつもん",
    "非拘束": "ひこうそく",
    "議論用": "ぎろんよう",
    "遊休時": "ゆうきゅうじ",
    "固定費": "こていひ",
    "無状態": "むじょうたい",
    "単方向": "たんほうこう",
    "型変換": "かたへんかん",
    "共通化": "きょうつうか",
    "会社間": "かいしゃかん",
    "自動補充": "じどうほじゅう",
    "制御面": "せいぎょめん",
    "管理費込み": "かんりひこみ",
    "込み": "こみ",
    "込": "こ",
    "管理費": "かんりひ",
    "相殺": "そうさい",
    "上乗せ": "うわのせ",
    "切り上げ": "きりあげ",
    "位置づけ": "いちづけ",
    "連携機": "れんけいき",
    "連携面": "れんけいめん",
    "構成図": "こうせいず",
    "枚目": "まいめ",
    "人日": "にんにち",
    "西友側": "せいゆうがわ",
    "一本": "いっぽん",
    "口径": "こうけい",
    "依存": "いぞん",
    "先に": "さきに",
    "先": "さき",
    "一体": "いったい",
    "一覧": "いちらん",
    "一言": "ひとこと",
    "一口": "ひとくち",
    "本体": "ほんたい",
    "本機": "ほんき",
    "本件": "ほんけん",
    "新規": "しんき",
    "新境界": "しんきょうかい",
    "通過": "つうか",
    "添付": "てんぷ",
    "附帯": "ふたい",
    "変換付": "へんかんつき",
    "付き": "つき",
    "含み": "ふくみ",
    "共用": "きょうよう",
    "固定": "こてい",
    "保ち": "たもち",
    "保つ": "たもつ",
    "課金": "かきん",
    "近づ": "ちかづ",
    "重さ": "おもさ",
    "重い": "おもい",
    "机上": "きじょう",
    "実費": "じっぴ",
    "噛み": "かみ",
    "紙": "かみ",
    "表": "おもて",
    "限": "げん",
    "下限": "かげん",
    "出て": "でて",
    "展開": "てんかい",
    "物語": "ものがたり",
    "拡が": "ひろが",
    "収まり": "おさまり",
    "直接的": "ちょくせつてき",
    "監視": "かんし",
    "更新": "こうしん",
    "設計": "せっけい",
    "齟齬": "そご",
    "合い": "あい",
    "合わ": "あわ",
    "同程度": "どうていど",
    "得ます": "えます",
    "主に": "おもに",
    "値段": "ねだん",
    "差": "さ",
    "当て": "あて",
    "分": "ぶん",
    "既に": "すでに",
    "多く": "おおく",
    "枠": "わく",
    "無く": "なく",
    "崩れ": "くずれ",
    "たまたま": "たまたま",
    "自体": "じたい",
    "残り": "のこり",
    "残る": "のこる",
    "大半": "たいはん",
    "遊休": "ゆうきゅう",
    "終了後": "しゅうりょうご",
    "終了": "しゅうりょう",
    "審議": "しんぎ",
    "お願い": "おねがい",
    "要約": "ようやく",
    "代替案": "だいたいあん",
    "閉じ": "とじ",
    "同額": "どうがく",
    "不要": "ふよう",
    "割": "わり",
    "万円": "まんえん",
    "月約": "つきやく",
    "年約": "ねんやく",
    "年間": "ねんかん",
    "初算": "しょさん",
    "明細": "めいさい",
    "未計上": "みけいじょう",
    "可能性": "かのうせい",
    "将来": "しょうらい",
    "変更": "へんこう",
    "起動": "きどう",
    "突合": "とつごう",
    "制御": "せいぎょ",
    "検証": "けんしょう",
    "境界": "きょうかい",
    "入ら": "いら",
    "専用": "せんよう",
    "中央": "ちゅうおう",
    "左": "ひだり",
    "右": "みぎ",
    "東京": "とうきょう",
    "大阪": "おおさか",
    "既存": "きそん",
    "今回": "こんかい",
    "しない": "しない",
    "約": "やく",
    "日": "にち",
    "最終": "さいしゅう",
    "処理": "しょり",
    "源": "げん",
    "自有": "じゆう",
    "暫定": "ざんてい",
    "技術": "ぎじゅつ",
    "判断": "はんだん",
    "常駐": "じょうちゅう",
    "必要": "ひつよう",
    "日次": "にちじ",
    "採用": "さいよう",
    "申請": "しんせい",
    "観点": "かんてん",
    "第一": "だいいち",
    "第二": "だいに",
    "第三": "だいさん",
    "第四": "だいよん",
    "第五": "だいご",
    "第六": "だいろく",
    "動き": "うごき",
    "動く": "うごく",
    "動かす": "うごかす",
    "動": "うご",
    "方": "かた",
    "処理": "しょり",
    "遊休": "ゆうきゅう",
    "時": "じ",
    "止め": "とめ",
    "常設": "じょうせつ",
    "仕事": "しごと",
    "抽出": "ちゅうしゅつ",
    "変換": "へんかん",
    "向き": "むき",
    "向": "む",
    "前提": "ぜんてい",
    "運用": "うんよう",
    "費用": "ひよう",
    "軽い": "かるい",
    "費": "ひ",
    "数字": "すうじ",
    "後": "あと",
    "対比": "たいひ",
    "常時": "じょうじ",
    "基盤": "きばん",
    "見": "み",
    "四": "よ",
    "大": "だい",
    "全量": "ぜんりょう",
    "載せ": "のせ",
    "権限": "けんげん",
    "容量": "ようりょう",
    "転送": "てんそう",
    "円": "えん",
    "万": "まん",
    "月": "つき",
    "年": "ねん",
    "等": "など",
    "割": "わり",
    "系": "けい",
    "側": "がわ",
    "渡す": "わたす",
    "渡": "わた",
    "上": "うえ",
    "台": "だい",
    "取得": "しゅとく",
    "配置": "はいち",
    "納品": "のうひん",
    "並列": "へいれつ",
    "惣菜": "そうざい",
    "新": "あたら",
    "機": "き",
    "別": "べつ",
    "全": "ぜん",
    "持た": "もた",
    "持": "も",
    "永続": "えいぞく",
    "見積": "みつもり",
    "言": "い",
    "回す": "まわす",
    "回": "まわ",
    "要る": "いる",
    "要": "よう",
    "統一": "とういつ",
    "次": "つぎ",
    "作る": "つくる",
    "作": "つく",
    "説明": "せつめい",
    "背景": "はいけい",
    "進": "すす",
    "西友": "せいゆう",
    "基幹": "きかん",
    "統合": "とうごう",
    "第": "だい",
    "段階": "だんかい",
    "集計": "しゅうけい",
    "正本": "せいほん",
    "通路": "つうろ",
    "吸収": "きゅうしゅう",
    "方針": "ほうしん",
    "課題": "かだい",
    "割れ": "われ",
    "割る": "わる",
    "まとめる": "まとめる",
    "連携": "れんけい",
    "位置": "いち",
    "松尾": "まつお",
    "規": "き",
    "付": "つ",
    "体": "たい",
    "将": "しょう",
    "一": "いち",
    "含": "ふく",
    "保管": "ほかん",
    "変わる": "かわる",
    "時刻": "じこく",
    "途中": "とちゅう",
    "再開": "さいかい",
    "書き込み": "かきこみ",
    "通常": "つうじょう",
    "直接": "ちょくせつ",
    "直送": "ちょくそう",
    "二通り": "ふたとおり",
    "保ち": "たもち",
    "保つ": "たもつ",
    "保": "たも",
    "近": "ちか",
    "重": "おも",
    "不": "ふ",
    "願": "ねが",
    "仕事": "しごと",
    "解凍": "かいとう",
    "外排": "がいはい",
    "環境": "かんきょう",
    "本日": "ほんじつ",
    "枚": "まい",
    "理由": "りゆう",
    "構成": "こうせい",
    "最後": "さいご",
    "比較": "ひかく",
    "順": "じゅん",
    "時間": "じかん",
    "使": "つか",
    "理由": "りゆう",
    "主": "おも",
    "得": "え",
    "程度": "ていど",
    "管理": "かんり",
    "既": "すで",
    "多": "おお",
    "無": "な",
    "崩": "くず",
    "収": "おさ",
    "切": "き",
    "噛": "か",
    "選": "えら",
    "安": "やす",
    "話": "はなし",
    "不安定": "ふあんてい",
    "過剰": "かじょう",
    "隔離": "かくり",
    "源隔離": "げんかくり",
    "単方向納品": "たんほうこうのうひん",
    "実行体": "じっこうたい",
    "申請図": "しんせいず",
    "下": "した",
}

# 从 MD 括弧自动并入（覆盖同名）
_paren_re = re.compile(r"([\u4e00-\u9fff々〆ヵヶ]+)（([ぁ-んァ-ンー]+)）")


def _merge_paren_dict(text: str) -> None:
    for kanji, yomi in _paren_re.findall(text):
        RUBY_DICT[kanji] = yomi


def esc(s: str) -> str:
    return html.escape(s)


def _ruby(kanji: str, yomi: str) -> str:
    return f"<ruby>{esc(kanji)}<rt>{esc(yomi)}</rt></ruby>"


def _dict_annotate(plain: str) -> str:
    """对已无括弧／无 ruby 的纯文本做最长匹配注音。"""
    keys = sorted(RUBY_DICT.keys(), key=len, reverse=True)
    if not keys:
        return esc(plain)
    # 跳过已是假名／英数的部分，只扫汉字串
    out: list[str] = []
    i = 0
    n = len(plain)
    while i < n:
        ch = plain[i]
        if "\u4e00" <= ch <= "\u9fff" or ch in "々〆ヵヶ":
            matched = None
            for k in keys:
                if plain.startswith(k, i):
                    matched = k
                    break
            if matched:
                out.append(_ruby(matched, RUBY_DICT[matched]))
                i += len(matched)
                continue
            # 单字未命中：原样输出（避免乱注）
            out.append(esc(ch))
            i += 1
        else:
            # 非汉字连续段
            j = i + 1
            while j < n:
                c = plain[j]
                if "\u4e00" <= c <= "\u9fff" or c in "々〆ヵヶ":
                    break
                j += 1
            out.append(esc(plain[i:j]))
            i = j
    return "".join(out)


def furigana_html(s: str) -> str:
    """先括弧→ruby，再对剩余汉字词典补全。"""
    out: list[str] = []
    last = 0
    for m in _paren_re.finditer(s):
        out.append(_dict_annotate(s[last : m.start()]))
        out.append(_ruby(m.group(1), m.group(2)))
        last = m.end()
    out.append(_dict_annotate(s[last:]))
    return "".join(out)


def inline_md(s: str, *, jp: bool) -> str:
    parts = re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", s)
    buf: list[str] = []
    for p in parts:
        if p.startswith("**") and p.endswith("**"):
            inner = p[2:-2]
            body = furigana_html(inner) if jp else esc(inner)
            buf.append(f"<strong>{body}</strong>")
        elif p.startswith("`") and p.endswith("`"):
            buf.append(f"<code>{esc(p[1:-1])}</code>")
        else:
            buf.append(furigana_html(p) if jp else esc(p))
    return "".join(buf)


def paras(block: str, *, jp: bool) -> str:
    block = block.strip()
    if not block:
        return ""
    # 引用行 > ...
    lines = []
    for raw in block.split("\n"):
        line = raw.strip()
        if line.startswith(">"):
            line = line.lstrip("> ").strip()
        if line:
            lines.append(line)
    return "".join(f"<p>{inline_md(ln, jp=jp)}</p>" for ln in lines)


def jp_title(s: str) -> str:
    """标题：日文段注音，中文「翻到第 N 页」保持原样。"""
    if "／" in s:
        left, right = s.split("／", 1)
        return furigana_html(left) + "／" + esc(right)
    return furigana_html(s)


def main() -> None:
    text = MD_PATH.read_text(encoding="utf-8")
    _merge_paren_dict(text)

    html_parts: list[str] = []
    html_parts.append(
        """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>SCM連携クラウド機　GCP環境申請　話術（中日・照着念）</title>
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
  background: var(--bg); color: var(--ink); line-height: 1.7;
}
header.top {
  position: sticky; top: 0; z-index: 20;
  background: rgba(246,243,238,.92); backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--line); padding: .75rem 1rem;
}
h1 { font-size: 1.15rem; margin: 0 0 .25rem; }
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
.card h2 { margin: 0 0 .7rem; font-size: 1.05rem; }
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
.cn-panel p, .jp-panel p { margin: 0 0 .65rem; }
.cn-panel p:last-child, .jp-panel p:last-child { margin-bottom: 0; }
.jp-panel { font-size: 1.02rem; }
ruby { ruby-align: center; }
rt { font-size: .55em; color: #a16207; }
body.big { font-size: 1.12rem; }
body.big .jp-panel { font-size: 1.18rem; }
body.hide-cn .lang-cn-block { display: none !important; }
body.hide-jp .lang-jp-block { display: none !important; }
table { width: 100%; border-collapse: collapse; font-size: .92rem; }
th, td { border: 1px solid var(--line); padding: .45rem .55rem; vertical-align: top; }
th { background: #fafaf9; text-align: left; }
.qa-expand details { margin-top: .5rem; }
.qa-expand summary { cursor: pointer; font-weight: 600; }
.hint { color: var(--muted); font-size: .85rem; }
.note { font-size: .8rem; color: var(--muted); margin-top: 1rem; text-align: center; }
code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: .9em; }
@media (max-width: 859px) {
  .pair.side { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<header class="top">
<h1>SCM連携クラウド機　GCP環境申請　話術（中日・照着念）</h1>
<div class="meta">故事线：背景→判断→構成→運用→スペック→費用→代替案　｜　6〜7分　｜　PPT 7枚</div>
<div class="sync">日文汉字自动注音（词典＋原稿括弧注音）／与 PPT・MD 同步<br/>用語：<strong>STS</strong>＝Google Cloud Storage Transfer Service（ストレージ転送サービス）</div>
<div class="toolbar">
<button type="button" onclick="showBoth()">中日両方</button>
<button type="button" onclick="showCn()">中文のみ</button>
<button type="button" onclick="showJp()">日本語のみ</button>
<button type="button" onclick="toggleBig()">文字大きく</button>
<button type="button" onclick="toggleSide()">左右並び</button>
</div>
<nav class="toc">
<a href="#sec-15">15秒故事线</a>
<a href="#slide-1">S1</a><a href="#slide-2">S2</a><a href="#slide-3">S3</a>
<a href="#slide-4">S4</a><a href="#slide-5">S5</a><a href="#slide-6">S6</a>
<a href="#slide-7">S7</a><a href="#sec-qa">Q&A</a>
</nav>
</header>
<main>
"""
    )

    # --- 故事线表 ---
    html_parts.append(
        """<section class="card" id="sec-story"><h2>故事线</h2>
<table><thead><tr><th>#</th><th>页</th><th>作用</th></tr></thead><tbody>
<tr><td>1</td><td>背景・なぜ新クラウド機</td><td>为什么要申请</td></tr>
<tr><td>2</td><td>なぜ Cloud Run</td><td>怎么建</td></tr>
<tr><td>3</td><td>申請構成図</td><td>长什么样</td></tr>
<tr><td>4</td><td>日次フロー・権限</td><td>怎么跑</td></tr>
<tr><td>5</td><td>スペック一覧</td><td>请批什么</td></tr>
<tr><td>6</td><td>費用試算</td><td>月／年费用</td></tr>
<tr><td>7</td><td>vs GKE</td><td>关替代案</td></tr>
</tbody></table></section>
"""
    )

    # --- 15 秒版 ---
    m15 = re.search(
        r"## 15 秒版.*?\n#### 中文\n\n(.*?)\n#### 日本語\n\n(.*?)\n---",
        text,
        re.S,
    )
    if not m15:
        raise SystemExit("15秒版 not found")
    cn15, jp15 = m15.group(1).strip().strip("> ").strip(), m15.group(2).strip().strip("> ").strip()
    # blockquotes may be multi-line with >
    def strip_bq(s: str) -> str:
        return "\n".join(ln.lstrip("> ").rstrip() for ln in s.strip().splitlines() if ln.strip())

    html_parts.append('<section class="card" id="sec-15"><h2>15 秒版</h2>\n')
    html_parts.append('<div class="pair"><div class="panel cn-panel lang-cn-block">')
    html_parts.append(f'<span class="badge cn">中文</span>{paras(strip_bq(m15.group(1)), jp=False)}</div>\n')
    html_parts.append('<div class="panel jp-panel lang-jp-block">')
    html_parts.append(f'<span class="badge jp">日本語</span>{paras(strip_bq(m15.group(2)), jp=True)}</div></div></section>\n')

    # --- スライド 1-7 ---
    slides = list(
        re.finditer(
            r"### 【スライド (\d)】([^\n]+)\n\n#### 中文\n\n(.*?)\n#### 日本語\n\n(.*?)(?=\n---\n)",
            text,
            re.S,
        )
    )
    if len(slides) != 7:
        raise SystemExit(f"expected 7 slides, got {len(slides)}")

    for m in slides:
        num, title, cn, jp = m.group(1), m.group(2).strip(), m.group(3), m.group(4)
        html_parts.append(f'<section class="card" id="slide-{num}">\n')
        html_parts.append(f'<h2><span class="slide-num">S{num}/7</span>{jp_title(title)}</h2>\n')
        html_parts.append('<div class="pair">\n')
        html_parts.append(
            f'<div class="panel cn-panel lang-cn-block"><span class="badge cn">中文</span>{paras(cn, jp=False)}</div>\n'
        )
        html_parts.append(
            f'<div class="panel jp-panel lang-jp-block"><span class="badge jp">日本語</span>{paras(jp, jp=True)}</div>\n'
        )
        html_parts.append("</div></section>\n")

    # --- Q&A 表 ---
    qa = re.search(
        r"## 想定質問.*?\n\n(\| 質問 \|.*?\n\|---\|.*?\n)((?:\|.*?\n)+)",
        text,
        re.S,
    )
    if not qa:
        raise SystemExit("Q&A table not found")
    rows_html = []
    for row in qa.group(2).strip().splitlines():
        cols = [c.strip() for c in row.strip("|").split("|")]
        if len(cols) < 3:
            continue
        q, cn, jp = cols[0], cols[1], cols[2]
        rows_html.append(
            "<tr>"
            f"<td>{esc(q)}</td>"
            f"<td class='lang-cn-block'>{inline_md(cn, jp=False)}</td>"
            f"<td class='lang-jp-block'>{inline_md(jp, jp=True)}</td>"
            "</tr>"
        )
    html_parts.append(
        f"""<section class="card" id="sec-qa"><h2>{furigana_html("想定質問")}</h2>
<p class="hint">想定質問先短答，追问再补一句即可。</p>
<table><thead><tr><th>質問</th><th class="lang-cn-block">中文</th><th class="lang-jp-block">日本語</th></tr></thead>
<tbody>
{''.join(rows_html)}
</tbody></table></section>
"""
    )

    # --- Autopilot 展开答（可选；当前 MD 已删则跳过）---
    ap = re.search(
        r"### 展开答：([^\n]+)\n\n> ([^\n]+)\n\n#### 中文（照着念）\n\n(.*?)\n#### 日本語（照着念・注音）\n\n(.*)",
        text,
        re.S,
    )
    if ap:
        html_parts.append(
            f"""<section class="card qa-expand" id="sec-autopilot">
<h2>展开答：不用 Autopilot＋Free tier</h2>
<p class="hint">{esc(ap.group(2).strip())}</p>
<details open>
<summary>核心：不是因为更贵</summary>
<div class="pair" style="margin-top:.75rem">
<div class="panel cn-panel lang-cn-block"><span class="badge cn">中文</span>{paras(ap.group(3), jp=False)}</div>
<div class="panel jp-panel lang-jp-block"><span class="badge jp">日本語</span>{paras(ap.group(4), jp=True)}</div>
</div>
</details>
</section>
"""
        )

    html_parts.append(
        """<p class="note">注音：原稿（かな）＋词典自动补全　｜　源 MD／PPT 对齐</p>
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
    )

    OUT_PATH.write_text("".join(html_parts), encoding="utf-8")
    body = OUT_PATH.read_text(encoding="utf-8")
    assert "218,280" in body
    assert "将来 Hinemos" in body or "将来可能改为 Hinemos" in body
    ruby_n = body.count("<ruby>")
    print(f"OK {OUT_PATH} size={OUT_PATH.stat().st_size} rubies={ruby_n}")

    # 检查日语面板剩余未注音汉字
    from collections import Counter

    jp_parts = re.findall(r'class="panel jp-panel lang-jp-block">(.*?)</div>', body, re.S)
    left = Counter()
    for part in jp_parts:
        part = re.sub(r'<span class="badge[^"]*">.*?</span>', "", part)
        chunks = re.split(r"(<ruby>[\s\S]*?</ruby>)", part)
        for ch in chunks:
            if ch.startswith("<ruby>"):
                continue
            plain = re.sub(r"<[^>]+>", "", ch)
            for m in re.finditer(r"[\u4e00-\u9fff々〆ヵヶ]+", plain):
                left[m.group()] += 1
    if left:
        print("WARN unannotated in JP panels:")
        for w, n in left.most_common(40):
            print(f"  {n:3d}  {w}")
    else:
        print("JP panels: no bare kanji left")


if __name__ == "__main__":
    main()
