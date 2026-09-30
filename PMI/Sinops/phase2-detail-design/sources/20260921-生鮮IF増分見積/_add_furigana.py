#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""日文話術に振り仮名を付ける。プレースホルダ衝突を避ける。"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent

GLOSSARY: list[tuple[str, str]] = [
    ("発注スケジュール", "はっちゅうスケジュール"),
    ("発注統合変換", "はっちゅうとうごうへんかん"),
    ("倉庫商品マスタ", "そうこしょうひんマスタ"),
    ("倉庫商品", "そうこしょうひん"),
    ("前提条件", "ぜんていじょうけん"),
    ("詳細設計", "しょうさいせっけい"),
    ("不確実性", "ふかくじつせい"),
    ("待ち合わせ", "まちあわせ"),
    ("作成時刻", "さくせいじこく"),
    ("伝票区分", "でんぴょうくぶん"),
    ("入荷予定", "にゅうかよてい"),
    ("入荷実績", "にゅうかじっせき"),
    ("発注勧告", "はっちゅうかんこく"),
    ("生鮮勧告", "せいせんかんこく"),
    ("鮮魚市場", "せんぎょいちば"),
    ("連携先", "れんけいさき"),
    ("手戻り", "てもどり"),
    ("読み方", "よみかた"),
    ("進め方", "すすめかた"),
    ("置き方", "おきかた"),
    ("織り込んだ", "おりこんだ"),
    ("成り立つ", "なりたつ"),
    ("覆った", "くつがえった"),
    ("覆る", "くつがえる"),
    ("収まらない", "おさまらない"),
    ("足りなければ", "たりなければ"),
    ("揃えます", "そろえます"),
    ("揃える", "そろえる"),
    ("置きません", "おきません"),
    ("動かしません", "うごかしません"),
    ("変えません", "かえません"),
    ("含みません", "ふくみません"),
    ("含むもの", "ふくむもの"),
    ("含まない", "ふくまない"),
    ("対象外", "たいしょうがい"),
    ("不安定", "ふあんてい"),
    ("顕在化", "けんざいか"),
    ("据え置き", "すえおき"),
    ("再見積", "さいみつもり"),
    ("標準案", "ひょうじゅんあん"),
    ("提示値", "ていじち"),
    ("差分IF", "さぶんIF"),
    ("非生鮮", "ひせいせん"),
    ("請負", "うけおい"),
    ("流用", "りゅうよう"),
    ("突合", "とつごう"),
    ("按分", "あんぶん"),
    ("波及", "はきゅう"),
    ("精査", "せいさ"),
    ("難易度", "なんいど"),
    ("単体", "たんたい"),
    ("結合", "けつごう"),
    ("増分", "ぞうぶん"),
    ("見積", "みつもり"),
    ("残件", "ざんけん"),
    ("判定", "はんてい"),
    ("前提", "ぜんてい"),
    ("新規", "しんき"),
    ("差分", "さぶん"),
    ("勧告", "かんこく"),
    ("実績", "じっせき"),
    ("予定", "よてい"),
    ("変換", "へんかん"),
    ("統合", "とうごう"),
    ("発注", "はっちゅう"),
    ("入荷", "にゅうか"),
    ("店舗", "てんぽ"),
    ("倉庫", "そうこ"),
    ("在庫", "ざいこ"),
    ("確認", "かくにん"),
    ("契約", "けいやく"),
    ("範囲", "はんい"),
    ("調査", "ちょうさ"),
    ("開発", "かいはつ"),
    ("本体", "ほんたい"),
    ("明示", "めいじ"),
    ("比率", "ひりつ"),
    ("既存", "きそん"),
    ("直接", "ちょくせつ"),
    ("小計", "しょうけい"),
    ("合計", "ごうけい"),
    ("内訳", "うちわけ"),
    ("課題", "かだい"),
    ("参照", "さんしょう"),
    ("振替", "ふりかえ"),
    ("伝票", "でんぴょう"),
    ("区分", "くぶん"),
    ("取得", "しゅとく"),
    ("処理", "しょり"),
    ("接続", "せつぞく"),
    ("軽微", "けいび"),
    ("無効", "むこう"),
    ("除外", "じょがい"),
    ("改修", "かいしゅう"),
    ("先方", "せんぽう"),
    ("構築", "こうちく"),
    ("運用", "うんよう"),
    ("移行", "いこう"),
    ("総合", "そうごう"),
    ("質問", "しつもん"),
    ("着手", "ちゃくしゅ"),
    ("類似", "るいじ"),
    ("相当", "そうとう"),
    ("青果", "せいか"),
    ("商品", "しょうひん"),
    ("横断", "おうだん"),
    ("必須", "ひっす"),
    ("確定", "かくてい"),
    ("記号", "きごう"),
    ("意味", "いみ"),
    ("税抜", "ぜいぬき"),
    ("換算", "かんさん"),
    ("本数", "ほんすう"),
    ("工数", "こうすう"),
    ("注記", "ちゅうき"),
    ("明細", "めいさい"),
    ("根拠", "こんきょ"),
    ("水煮", "みずに"),
    ("生鮮", "せいせん"),
    ("含む", "ふくむ"),
    ("以上", "いじょう"),
    ("最後", "さいご"),
    ("説明", "せつめい"),
    ("数字", "すうじ"),
    ("場合", "ばあい"),
    ("上記", "じょうき"),
    ("要点", "ようてん"),
    ("人日", "にんにち"),
    ("万円", "まんえん"),
    ("単価", "たんか"),
    ("結論", "けつろん"),
    ("骨格", "こっかく"),
    ("焦点", "しょうてん"),
    ("反映", "はんえい"),
    ("未知", "みち"),
    ("想定", "そうてい"),
    ("外れる", "はずれる"),
    ("単独", "たんどく"),
    ("本日", "ほんじつ"),
    ("中身", "なかみ"),
    ("提示", "ていじ"),
    ("翌日", "よくじつ"),
    ("当日", "とうじつ"),
    ("以降", "いこう"),
    ("基幹", "きかん"),
    ("基盤", "きばん"),
    ("変更", "へんこう"),
    ("二重", "にじゅう"),
    ("計算", "けいさん"),
    ("内容", "ないよう"),
    ("一緒", "いっしょ"),
    ("必ず", "かならず"),
    ("比較的", "ひかくてき"),
    ("効き得る", "ききうる"),
    ("控え", "ひかえ"),
    ("繰り返し", "くりかえし"),
    ("頁", "ページ"),
    ("鮮魚", "せんぎょ"),
    ("市場", "いちば"),
    ("金額", "きんがく"),
    ("条件", "じょうけん"),
    ("揺れ", "ゆれ"),
    ("伝え", "つたえ"),
    ("管理", "かんり"),
    ("項目", "こうもく"),
    ("書き方", "かきかた"),
]

GLOSSARY.sort(key=lambda x: len(x[0]), reverse=True)

PAREN_YOMI = re.compile(r"（[ぁ-んァ-ンーa-zA-Z／]+）")
RUBY_TAG = re.compile(r"<ruby>(.*?)<rt>.*?</rt></ruby>", re.DOTALL)

REPAIRS = [
    ("生鮮IFPhase2の標準案", "生鮮IF増分の標準案"),
    ("生鮮（せいせん）IFPhase2", "生鮮IF増分"),
    ("IFPhase2の", "IF増分の"),
    ("だけのPhase2", "だけの単独見積"),
    ("5,120,000Phase2", "5,120,000円（税抜）"),
    ("Phase2へのPhase2", "Phase2への波及"),
    ("Phase2へのPhase1", "Phase2への波及"),
    ("へのPhase2", "への波及"),
    ("PJPhase2", "PJ管理按分"),
    ("PJPhase1", "PJ管理按分"),
    ("CのPJ管理按分が15", "CのPJ管理按分が15"),
    ("4Phase2", "4万円"),
    ("これは確定のSQL額", "これは確定の請負額"),
    ("これは確定のPC／CK額", "これは確定の請負額"),
    ("これは確定のPhase2額", "これは確定の請負額"),
    ("12本はPC／CK", "12本は流用"),
    ("12本はSQL", "12本は流用"),
    ("12本のGCS", "12本の突合"),
    ("GCS8、倉庫", "突合8、倉庫"),
    ("項目のSQL", "項目の精査"),
    ("区分のSQL", "区分の精査"),
    ("区分のGCS", "区分の精査"),
    ("Phase2だけです。だからB2", "突合だけです。だからB2"),
    ("Phase2は、Aが", "内訳は、Aが"),
    ("開発・GCS、Aの", "開発・単体、Aの"),
    ("含まないのは、Phase2以降", "含まないのは、結合以降"),
    ("とPhase2で、15人日", "と残件で、15人日"),
    ("**運用**：", "**進め方**："),
    ("<strong>運用</strong>：", "<strong>進め方</strong>："),
]


def protect(text: str, pattern: str, prefix: str) -> tuple[str, list[str]]:
    held: list[str] = []

    def _save(m: re.Match) -> str:
        held.append(m.group(0))
        return f"\x00{prefix}:{len(held) - 1}\x00"

    return re.sub(pattern, _save, text, flags=re.DOTALL), held


def restore(text: str, held: list[str], prefix: str) -> str:
    for i in range(len(held) - 1, -1, -1):
        text = text.replace(f"\x00{prefix}:{i}\x00", held[i])
    return text


def strip_md_yomi(text: str) -> str:
    prev = None
    while prev != text:
        prev = text
        text = PAREN_YOMI.sub("", text)
    return text


def strip_html_ruby(text: str) -> str:
    return RUBY_TAG.sub(r"\1", text)


def repair(text: str) -> str:
    for a, b in REPAIRS:
        text = text.replace(a, b)
    return text


def wrap_md(text: str) -> str:
    text, codes = protect(
        text,
        r"(?:DBGET|SQL|GCS|Layer①|Layer②|Layer③|Phase1|Phase2|PC／CK|PC/CK)",
        "C",
    )
    wrapped: list[str] = []
    for word, yomi in GLOSSARY:
        piece = f"{word}（{yomi}）"

        def repl(_m: re.Match, p: str = piece) -> str:
            wrapped.append(p)
            return f"\x00W:{len(wrapped) - 1}\x00"

        text = re.sub(re.escape(word) + r"(?!（)", repl, text)
    text = restore(text, wrapped, "W")
    text = restore(text, codes, "C")
    return text


def wrap_html(text: str) -> str:
    text, ruby = protect(text, r"<ruby>[\s\S]*?</ruby>", "R")
    text, codes = protect(
        text,
        r"(?:DBGET|SQL|GCS|Layer①|Layer②|Layer③|Phase1|Phase2|PC／CK|PC/CK)",
        "C",
    )
    wrapped: list[str] = []
    for word, yomi in GLOSSARY:
        piece = f"<ruby>{word}<rt>{yomi}</rt></ruby>"

        def repl(_m: re.Match, p: str = piece) -> str:
            wrapped.append(p)
            return f"\x00W:{len(wrapped) - 1}\x00"

        text = re.sub(re.escape(word), repl, text)
    text = restore(text, wrapped, "W")
    text = restore(text, codes, "C")
    text = restore(text, ruby, "R")
    return text


def map_jp_divs(html: str, fn) -> str:
    marker = '<div class="jp">'
    out: list[str] = []
    i = 0
    while True:
        j = html.find(marker, i)
        if j < 0:
            out.append(html[i:])
            break
        out.append(html[i : j + len(marker)])
        k = j + len(marker)
        depth = 1
        p = k
        inner_end = None
        while p < len(html) and depth:
            nxt_open = html.find("<div", p)
            nxt_close = html.find("</div>", p)
            if nxt_close < 0:
                break
            if nxt_open != -1 and nxt_open < nxt_close:
                depth += 1
                p = nxt_open + 4
            else:
                depth -= 1
                if depth == 0:
                    inner_end = nxt_close
                    p = nxt_close + 6
                    break
                p = nxt_close + 6
        if inner_end is None:
            out.append(html[k:])
            break
        out.append(fn(html[k:inner_end]))
        out.append("</div>")
        i = p
    return "".join(out)


def process_html(path: Path) -> None:
    src = path.read_text(encoding="utf-8")

    def jp_inner(inner: str) -> str:
        inner = strip_html_ruby(inner)
        inner = repair(strip_md_yomi(inner))
        return wrap_html(inner)

    out = map_jp_divs(src, jp_inner)
    path.write_text(out, encoding="utf-8")
    print("html", path.name)


def process_md(path: Path) -> None:
    src = path.read_text(encoding="utf-8")

    def jp_sec(m: re.Match) -> str:
        body = repair(strip_md_yomi(m.group(1)))
        return "### 日文" + wrap_md(body)

    out = re.sub(r"### 日文(.*?)(?=\n### |\n## |\Z)", jp_sec, src, flags=re.DOTALL)
    path.write_text(out, encoding="utf-8")
    print("md", path.name)


def main() -> None:
    for p in sorted(ROOT.glob("*話術*.md")):
        process_md(p)
    for p in sorted(ROOT.glob("*話術*.html")):
        process_html(p)


if __name__ == "__main__":
    main()
