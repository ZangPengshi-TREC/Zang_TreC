# -*- coding: utf-8 -*-
"""
在庫IF分析（No.51 在庫調整データ / No.52 在庫一覧データ）の訂正版を生成する。

【訂正の背景】
旧版は TRIAL側テーブルを `current_stock_data`（id/store_code/jan/quantity/amount/
last_updated_at, UNIQUE(store_code,jan) 想定）として整理していたが、これは実在しない
テーブルであることが判明。実際のTRIAL側テーブルは `center_stock` で、カラム構成が
大きく異なる（amount列なし、時刻列が4種、UNIQUE制約未確認）。
本スクリプトは `center_stock` を基準に、マッピング表・GAP分析書・変換ルール定義書・
QA一覧の4種×2IF＝8ファイルを再生成する。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUT_DIR = "/Users/treqd/Desktop/Cursor/西友PMI/在庫IF_訂正版"
os.makedirs(OUT_DIR, exist_ok=True)

HEADER_FILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
BANNER_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
BANNER_FONT = Font(color="9C5700", bold=True, size=10)
NEW_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
CRIT_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BANNER_TEXT = (
    "訂正履歴：旧版はTRIAL側テーブルを『current_stock_data』として整理していましたが、"
    "対象がNo.51/No.52（倉庫コード基準）と不一致であることが判明したため、"
    "実際のテーブル『center_stock』を基準に本資料で全面的に再整理しました。（訂正日：2026年8月11日）\n"
    "【追記】その後の調査で『current_stock_data』は実在するテーブルであることを確認しました"
    "（id/store_code INT4/jan VARCHAR20/quantity/amount(未使用)/last_updated_at, "
    "UNIQUE(store_code,jan)、実データ有）。ただし備考に「仕入れ・売上・廃棄・振替等で更新」と"
    "あり、店舗（店CD）単位の現在庫テーブルと推測されます。No.51/No.52のHitluster側項目は"
    "『倉庫コード』（店舗コードではない）が基準のため、店舗単位のcurrent_stock_dataではなく、"
    "センター/倉庫単位と推測される『center_stock』の方が対象として整合的と考え、本資料では"
    "center_stockを正としています。この対応関係の最終確認は下記QA-G01/QA-S01で改めてお願いします。"
)


def autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_banner(ws, ncols, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value="⚠ " + BANNER_TEXT)
    c.fill = BANNER_FILL
    c.font = BANNER_FONT
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 42


def write_header(ws, headers, row):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def write_rows(ws, rows, start_row, highlight_col=None, highlight_map=None):
    r = start_row
    for row_data in rows:
        tag = row_data[-1] if highlight_col is not None else None
        cells = row_data[:-1] if highlight_col is not None else row_data
        for i, v in enumerate(cells, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BORDER
            if tag == "NEW":
                c.fill = NEW_FILL
            elif tag == "CRIT":
                c.fill = CRIT_FILL
        r += 1
    return r


def make_sheet(wb, title, headers, rows, widths, sheet_name="Sheet1"):
    ws = wb.active if wb.active.title == "Sheet" and not wb.sheetnames[1:] else wb.create_sheet(sheet_name)
    ws.title = sheet_name
    ncols = len(headers)
    write_banner(ws, ncols, row=1)
    write_header(ws, headers, row=3)
    write_rows(ws, rows, start_row=4, highlight_col=True)
    autofit(ws, widths)
    ws.freeze_panes = "A4"
    return ws


# ============================================================
# center_stock 正しい構造（共通で参照）
# ============================================================
CENTER_STOCK_SCHEMA = """center_stock 実際の構造：
  id              BIGSERIAL PK（自動採番）
  stock_date      TIMESTAMP        … （仮説）在庫の業務日
  operation_time  TIMESTAMPTZ      … （仮説）在庫変動/操作発生時刻
  store_code      INT4             … 店舗コード（実データ例:580）
  product_code    VARCHAR(13)      … 商品コード（実データ例:4512754400631, 13桁ちょうど）
  quantity        NUMERIC          … 在庫数量
  transfer_time   TIMESTAMPTZ      … （仮説）IF連携/転送処理時刻
  registered_at   TIMESTAMPTZ      … DB書込時刻
  registered_by   VARCHAR(20)      … データ発生源（実データ例:"Migration"）
  ※ amount（移動平均単価）に相当する列は存在しない
  ※ UNIQUE制約(store_code, product_code)等の有無は未確認（要DB側確認）
"""

# ============================================================
# No.51 在庫調整データ（IFSYLMS320R）
# ============================================================

MAP51_HEADERS = ["Hitluster No.", "Hitluster 項目名", "Hitluster 型/桁", "FMT内容",
                  "TRIAL 項目名\n（center_stock）", "TRIAL 型/桁", "判定", "変換ルール・留意事項"]
MAP51_WIDTHS = [10, 16, 14, 22, 18, 16, 12, 46]

MAP51_ROWS = [
    ["—", "—", "—", "—", "id", "BIGSERIAL\nPRIMARY KEY", "自動生成",
     "TRIAL側で自動採番。Hitluster側に対応項目なし。（旧版と同じ）", None],
    ["2", "倉庫コード", "文字列(6)\n003-008", "西友倉庫コード\n例:\"007452\"",
     "store_code", "INT4",
     "変換必要",
     "【要定義】倉庫コード（文字列6桁）→店舗コード（INT4）へのマッピングテーブル定義が必要。"
     "center_stockの実データではstore_code=580のような小さい数値が入っている。"
     "\"007452\"→???、\"007455\"→???の対応は業務側に確認必要（QA-G01参照、旧版から変更なし）。",
     None],
    ["4", "商品コード", "文字列(14)\n017-030",
     "先頭スペース1byte+13桁JAN\n例:\" 4522646963267\"",
     "product_code", "VARCHAR(13)",
     "変換必要",
     "TRIM()で先頭スペース除去→13桁JAN文字列として格納。"
     "例: \" 4522646963267\" → \"4522646963267\"。"
     "【新規注意】旧版はjan VARCHAR(20)で余裕があったが、center_stockはVARCHAR(13)ちょうど。"
     "13桁以外（英数混在の社内コード等）が来た場合オーバーフローする（QA-G07・新規）。",
     "NEW"],
    ["6,7", "符号区分\n+\n在庫調整数", "文字列(2)\n036-037\n+\n数値(7)\n038-044",
     "01=正(+) / 02=負(−)\n+ 整数7桁\n例:0000001〜0000018",
     "quantity", "NUMERIC",
     "確認必要\n（要件変更）",
     "【要決定・最重要】center_stockに(store_code,product_code)のUNIQUE制約が確認できておらず、"
     "旧版前提の「差分UPSERT」がそのまま成立するか不明。\n"
     "プランA(推奨・仮): 直近の最新quantityを取得し符号で加減算した『調整後の総量』を新規行としてINSERT\n"
     "プランB: 差分量そのものを新規行としてINSERT、参照側でSUM集計\n"
     "プランC: 在庫調整専用の別テーブルへ記録\n"
     "→ center_stockの書込方式（履历追加型/現在値型）と合わせて確定必要（QA-S01・新規★最高）",
     "CRIT"],
    ["8", "移動平均単価", "数値(9)\n045-053", "整数7桁+小数2桁\n例:0001234.56\n=1,234.56円",
     "— （対応列なし）", "—",
     "対応なし\n（GAP拡大）",
     "【要再確認】旧版のcurrent_stock_dataには未使用のamount列が存在したが、"
     "center_stockにはamount相当の列が一切存在しない。恒久的に破棄でよいか改めて確認要"
     "（QA-G04・優先度を低→高へ引き上げ）。",
     "NEW"],
    ["3", "在庫調整日付", "文字列(8)\n009-016", "yyyymmdd\n例:20260625",
     "stock_date /\noperation_time", "TIMESTAMP /\nTIMESTAMPTZ",
     "確認必要",
     "center_stockには時刻系列カラムが4つ（stock_date, operation_time, transfer_time, "
     "registered_at）存在する。現時点の仮説（要確認・QA-S02）：\n"
     "stock_date＝調整対象の業務日\noperation_time＝調整発生時刻"
     "（Hitluster側は日付のみのため、暫定的に当日00:00:00+09を設定）",
     "NEW"],
    ["（新規対応項目）", "IF受信/転送起因", "—", "—",
     "transfer_time", "TIMESTAMPTZ",
     "確認必要\n（新規）",
     "IF受信・連携処理のタイムスタンプと想定。システム側で自動設定される可能性があり、"
     "IF実装側で明示設定が必要かどうか要確認（QA-S02）。",
     "NEW"],
    ["（新規対応項目）", "DB書込起因", "—", "—",
     "registered_at", "TIMESTAMPTZ",
     "確認必要\n（新規）",
     "DB書込時刻。DB側DEFAULT（例:now()）で自動設定される可能性が高く、"
     "IF実装側で明示設定不要の可能性がある（QA-S02）。",
     "NEW"],
    ["（新規対応項目）", "データ発生源識別", "—", "—",
     "registered_by", "VARCHAR(20)",
     "確認必要\n（新規項目）",
     "実データでは\"Migration\"という値が確認されている。本IF（No.51 在庫調整）由来の"
     "レコードにどの値を設定すべきか未定義（例:\"IF_ADJUST\"等）。"
     "命名規則は開発チーム/業務側と要合意（QA-S03・新規）。",
     "NEW"],
    ["1", "データ区分", "文字列(2)\n001-002", "\"32\"固定",
     "—", "—", "不要",
     "変更なし。TRIAL側では使用しない。レコード識別用（Hitluster側固定値\"32\"）。", None],
    ["5", "在庫調整コード", "文字列(5)\n031-035", "ADJ01/ADJ02/DMG05等",
     "—", "—", "確認必要",
     "変更なし。center_stockに対応項目なし。調整理由を別テーブルで記録する要否は"
     "確認必要（QA-A01参照）。", None],
    ["9", "FILLER", "文字列(459)\n054-512", "半角スペース固定",
     "—", "—", "不要",
     "変更なし。パディング領域。TRIAL側では使用しない。", None],
]

GAP51_HEADERS = ["GAPカテゴリ", "優先度", "影響テーブル/ファイル", "対象Hitluster項目",
                  "Hitluster側の状況", "TRIAL側の要件", "対応方針"]
GAP51_WIDTHS = [22, 8, 20, 18, 30, 30, 40]

GAP51_ROWS = [
    ["【訂正】TRIAL側参照テーブルの対象誤り", "★最高", "center_stock\n（全項目）", "全項目",
     "旧資料は`current_stock_data`\n（id/store_code INT4/jan VARCHAR20/quantity/\namount(未使用)/last_updated_at, "
     "UNIQUE(store_code,jan)）\nを参照していた。\n【追記】このテーブルは実在し、実データもあり。"
     "備考に「仕入れ・売上・廃棄・振替等で更新」とあり、店舗（店CD）単位の現在庫テーブルと推測される。"
     "一方、No.51/No.52のHitluster側キー項目は『倉庫コード』であり店舗コードとは概念が異なる",
     "実際に倉庫コード基準のIFが対応すべきは、センター/倉庫単位と推測される`center_stock`と考えられる。\n"
     + CENTER_STOCK_SCHEMA,
     "本資料ではcenter_stockを正として整理。ただし『center_stock.store_code』が真に倉庫/センターコード"
     "を表すか、current_stock_dataとの役割分担（店舗単位 vs 倉庫単位）は未確定のため、TRIAL側に最終"
     "確認をお願いしたい（QA-G01/QA-S01参照）。", "CRIT"],
    ["書き込み方式未確定", "★最高", "center_stock\n（全体）", "符号区分+在庫調整数",
     "Hitlusterは「差分型・随時送信」の\nUPSERT前提で送信してくる",
     "center_stockにUNIQUE制約が確認できず、\n履历追加型（ログ/新規行）の可能性が高い",
     "プランA:最新値を読取り加減算後に新規行INSERT／プランB:差分値をそのまま新規行INSERT"
     "（参照側でSUM）／プランC:調整専用の別テーブルに分離。TRIAL側と方式を確定（QA-S01）", "CRIT"],
    ["マッピング不足：倉庫コード→店舗コード", "★高", "center_stock\n(store_code列)", "倉庫コード(No.2, 003-008)",
     "西友倉庫コード（文字列6桁）\n例:\"007452\"\n店舗コードとの対応関係不明",
     "INT4型のstore_codeが必要\n実データ例: store_code=580",
     "【要確認】倉庫コード→店舗コードのマッピングテーブル定義を業務側に確認（QA-G01、変更なし）", None],
    ["移動平均単価の格納先消失", "★高\n(優先度UP)", "center_stock\n(該当列なし)", "移動平均単価(No.8, 045-053)",
     "数値(9): 整数7桁+小数2桁\n例:0001234.56=1,234.56円\nHitluster側は必ず送信",
     "center_stockには該当列が\n存在しない（旧amount列すら無い）",
     "恒久的に破棄でよいか改めて確認（QA-G04）。旧版は「未使用フィールドあり」だったが、"
     "現在は「格納先自体が無い」ため優先度を引き上げ", "NEW"],
    ["型不一致（オーバーフローリスク上昇）", "★中", "center_stock\n(product_code列)", "商品コード(No.4, 017-030)",
     "固定長14byte\n（先頭スペース1byte+13桁JAN）\n例:\" 4522646963267\"",
     "VARCHAR(13)ちょうど\n（旧VARCHAR(20)から余裕消失）",
     "TRIM()で13桁化して格納。13桁を超えるコード受信時のエラーハンドリングを新たに検討（QA-G07・新規）", "NEW"],
    ["時刻列の割当未確定", "★中", "center_stock\n(stock_date/operation_time/\ntransfer_time/registered_at)",
     "在庫調整日付(No.3, 009-016)",
     "日付のみ（yyyymmdd）\n時刻情報なし\n例:\"20260625\"",
     "4種類の時刻列が存在するが\n用途が未確定",
     "暫定案（stock_date=業務日／operation_time=発生時刻／transfer_time=連携時刻／"
     "registered_at=DB書込時刻）を本資料に記載。TRIAL側に正式定義を確認（QA-S02・新規）", "NEW"],
    ["registered_by設定値未定義", "★中", "center_stock\n(registered_by列)", "（対応Hitluster項目なし・新規論点)",
     "—", "VARCHAR(20)\n実データ例:\"Migration\"",
     "本IF（No.51）由来レコードの識別値を新規定義（例:\"IF_ADJUST\"）。命名規則を開発チームと合意（QA-S03・新規）", "NEW"],
    ["TRIAL側未対応：在庫調整コード", "★中", "（別テーブル要検討）", "在庫調整コード(No.5, 031-035)",
     "調整理由コード:ADJ01/ADJ02/DMG05等\nHitluster側は各調整データに付与",
     "center_stockには該当列なし\nTRIAL側での記録要件不明",
     "別テーブルで記録する要否を確認（QA-A01、変更なし）。不要なら破棄で対応", None],
    ["排他制御・重複防止", "★中", "center_stock\n(store_code+product_code)", "—",
     "同一(store_code,product_code)に対して\n複数調整データが随時送信される可能性",
     "UNIQUE制約未確認のため、\n重複行が無制限に増える恐れ",
     "書込方式（履历追加/現在値）の確定と合わせ、一意性制約の要否・重複防止方式を定義（QA-G06、内容更新）", "NEW"],
    ["データ種別の差異", "★低", "—", "データ区分(No.1) + FILLER(No.9)",
     "データ区分=\"32\"固定（レコード識別用）\nFILLER=半角スペース459byte（パディング）",
     "TRIAL側では使用しない",
     "受信時に破棄。TRIAL側で利用なし（変更なし）", None],
]

CONV51_HEADERS = ["sinops適用データ", "sinops適用項目", "sinops型/桁", "TRIALデータ",
                    "TRIAL項目", "TRIAL型", "変換ロジック", "具体例", "注意事項・エッジケース"]
CONV51_WIDTHS = [14, 14, 14, 12, 14, 16, 34, 22, 34]

CONV51_ROWS = [
    ["IFSYLMS320R\n(在庫調整データ)", "倉庫コード\n(No.2, 003-008)", "文字列(6)", "center_stock", "store_code",
     "INT4",
     "マッピングテーブル参照:\nSTORE_MAPPING = {\n  \"007452\": ???,\n  \"007455\": ???\n}\n"
     "store_code = STORE_MAPPING[倉庫コード]",
     "\"007452\" → ???\n（実データ例: store_code=580）",
     "【要確認】倉庫コード→店舗コードの対応表は業務側から提供必要（QA-G01）。"
     "存在しない倉庫コード受信時のエラーハンドリングも要定義", None],
    ["IFSYLMS320R\n(在庫調整データ)", "商品コード\n(No.4, 017-030)", "文字列(14)\n(先頭スペース1+13桁JAN)",
     "center_stock", "product_code", "VARCHAR(13)",
     "TRIM()で先頭スペース除去→13桁JAN文字列としてそのまま格納\nproduct_code = TRIM(sinops.商品コード)",
     "\" 4522646963267\"\n→ \"4522646963267\"",
     "【新規注意】旧版のVARCHAR(20)と異なり、center_stockはVARCHAR(13)ちょうど。"
     "13桁を超える値（英数字コード等）はINSERT時にエラーとなる（QA-G07）。事前バリデーションの追加を推奨", "NEW"],
    ["IFSYLMS320R\n(在庫調整データ)", "符号区分(No.6)\n+ 在庫調整数(No.7)", "文字列(2)\n+ 数値(7)",
     "center_stock", "quantity", "NUMERIC",
     "【方式未確定・要合意】\n"
     "プランA(推奨・仮): 直近の最新行を取得し、\n"
     "  符号\"01\": new_qty = latest_qty + 調整数\n"
     "  符号\"02\": new_qty = latest_qty - 調整数\n"
     "  → new_qtyを新規行としてINSERT\n"
     "プランB: quantity = ±調整数（差分そのもの）を新規行としてINSERT\n"
     "プランC: 専用テーブルへ記録",
     "【プランA例】\n直近quantity=100、符号01+調整数5\n→ 新規行 quantity=105\n"
     "【プランB例】\n符号02+調整数3\n→ 新規行 quantity=-3",
     "center_stockのUNIQUE制約有無が未確認のため、そもそも「直近の最新行」をどのソートキー"
     "（stock_date? id?）で特定するかも要定義。既存レコードなし時の初期値方針も要確認（QA-A02）", "CRIT"],
    ["IFSYLMS320R\n(在庫調整データ)", "移動平均単価\n(No.8, 045-053)", "数値(9)\n(整数7+小数2)",
     "center_stock", "— (対応列なし)", "—",
     "【破棄（現状）】\ncenter_stockに格納先が存在しないため、受信データは無視して破棄する",
     "単価 0001234.56\n→ 格納先なし・破棄",
     "旧版は「未使用フィールドあり」だったが、center_stockには列自体が存在しない。"
     "将来利用する場合は新規列追加が必要になる点を業務側に共有（QA-G04）", "NEW"],
    ["IFSYLMS320R\n(在庫調整データ)", "在庫調整日付\n(No.3, 009-016)", "文字列(8)\n(yyyymmdd)",
     "center_stock", "stock_date /\noperation_time", "TIMESTAMP /\nTIMESTAMPTZ",
     "【仮説・要確認】\nstock_date = TO_DATE(日付,'YYYYMMDD')\noperation_time = "
     "TO_TIMESTAMP(日付,'YYYYMMDD') AT TIME ZONE 'Asia/Tokyo'\n（日付のみのため00:00:00+09を仮設定）",
     "\"20260625\"\n→ stock_date=2026-06-25\n→ operation_time=2026-06-25 00:00:00+09",
     "4時刻列（stock_date/operation_time/transfer_time/registered_at）の正式な用途定義が必要"
     "（QA-S02）。transfer_time/registered_atはDB側で自動設定される可能性あり", "NEW"],
    ["IFSYLMS320R\n(在庫調整データ)", "（新規：発生源識別）", "—",
     "center_stock", "registered_by", "VARCHAR(20)",
     "【新規要定義】\n本IF由来のレコードであることを示す固定値を設定\n"
     "registered_by = 'IF_ADJUST' （仮称・要合意）",
     "実データ例:\"Migration\"\n（移行データ由来を示す値）",
     "値の命名規則（他IFとの区別方法）を開発チームと合意する必要あり（QA-S03）", "NEW"],
    ["IFSYLMS320R\n(在庫調整データ)", "データ区分\n(No.1, 001-002)", "文字列(2)\n(\"32\"固定)",
     "center_stock", "—", "—",
     "【破棄】\nTRIAL側では使用しない。レコード識別用のみ",
     "—", "内部検証用。\"32\"以外受信時はエラーとする（変更なし）", None],
    ["IFSYLMS320R\n(在庫調整データ)", "在庫調整コード\n(No.5, 031-035)", "文字列(5)\n(ADJ01/ADJ02/DMG05等)",
     "center_stock", "— (対応列なし)", "—",
     "【破棄（現状）】\ncenter_stockに格納先なし。別テーブル記録要否は未確認",
     "ADJ01/ADJ02/DMG05\n→ 記録せず破棄",
     "調整理由の監査ログ必要か確認必要（QA-A01）。必要な場合は別テーブルを新たに設計", None],
    ["IFSYLMS320R\n(在庫調整データ)", "FILLER\n(No.9, 054-512)", "文字列(459)\n(半角スペース固定)",
     "center_stock", "—", "—",
     "【破棄】\nパディング領域、TRIAL側で利用なし",
     "—", "スペース以外が含まれる場合はデータ破損の可能性（変更なし）", None],
]

QA51_HEADERS = ["ID", "課題・質問概要", "詳細説明", "回答"]
QA51_WIDTHS = [10, 26, 50, 20]

QA51_ROWS = [
    ["QA-S01\n★最高\n(新規)", "center_stockの書き込み方式（履历追加型か現在値型か）",
     "center_stockにはUNIQUE制約（store_code,product_code等）が確認できない。"
     "履历追加型（ログ/毎回新規行）なのか、現在値型（1キー1行・UPSERT）なのかで、"
     "No.51/No.52 双方の処理ロジックが根本的に変わる。\n"
     "案A: 履历追加型。最新値は「該当キーの最新行」を都度検索して判定\n"
     "案B: 現在値型。実際はUNIQUE制約が別途存在する（確認漏れ）\n"
     "案C: 用途別に別テーブルが存在し、center_stockはどちらか一方のみ対象", None],
    ["QA-S02\n★中\n(新規)", "4つの時刻列（stock_date/operation_time/transfer_time/registered_at）の用途定義",
     "center_stockには時刻系のカラムが4つ存在するが、旧`current_stock_data`はlast_updated_at "
     "1列のみを想定していた。それぞれの正式な業務的意味・自動設定/IF側設定の要否を確認したい。\n"
     "仮説: stock_date=業務日／operation_time=発生時刻／transfer_time=IF連携時刻／"
     "registered_at=DB書込時刻（DB側DEFAULT想定）", None],
    ["QA-S03\n★中\n(新規)", "registered_byに設定すべき値",
     "実データでは\"Migration\"という値が確認された。本IF（在庫調整データ）由来のレコードに"
     "どの値を設定すべきか（例:\"IF_ADJUST\"）。他IF（No.52等）との区別方法も含め命名規則を確認したい。", None],
    ["QA-G01\n★高\n(既存)", "倉庫コード→store_codeのマッピング定義（center_stock/current_stock_data役割分担含む）",
     "TRIAL側store_codeはINT4型、Hitluster側倉庫コードは文字列6桁。"
     "倉庫コード→store_codeの対応テーブル定義が必要。\n"
     "例: \"007452\"→何番？、\"007455\"→何番？（実データ例:store_code=580）\n"
     "【追記】別途『current_stock_data』という実在テーブル（store_code/jan/quantity, "
     "UNIQUE(store_code,jan)）を確認。備考に「仕入れ・売上・廃棄・振替等で更新」とあり店舗（店CD）"
     "単位の現在庫テーブルと推測される。本IFはcenter_stockを対象としているが、"
     "current_stock_dataとの役割分担（店舗単位 vs 倉庫/センター単位で別管理なのか、"
     "それとも将来的に統合されるのか）を業務側に確認したい。\n"
     "※No.52と共通ルール", None],
    ["QA-G04\n★高\n(優先度UP)", "移動平均単価の値は本当に不要か（格納先が消失）",
     "旧版のcurrent_stock_dataにはamountという未使用フィールドがあったが、実際のcenter_stockには"
     "amount相当の列が存在しない。Hitluster側は移動平均単価（数値9桁）を必ず送信してくる。"
     "将来amountに類する値を利用する予定はあるか、完全に破棄してよいか確認したい。", None],
    ["QA-G07\n★中\n(新規)", "商品コードがVARCHAR(13)ちょうどで桁オーバーフローした場合の扱い",
     "旧版はjan VARCHAR(20)で余裕があったが、center_stock.product_codeはVARCHAR(13)ちょうど。"
     "TRIM後13桁を超えるコード（英数字混在の社内コード等）を受信した場合のエラーハンドリング"
     "方針（拒否/切り捨て/別列退避等）を確認したい。", None],
    ["QA-A02\n★高\n(再構成)", "既存レコードなし・符号02(−)時の扱い（プランにより意味が変わる）",
     "QA-S01の結論（履历追加型かどうか）により本論点の意味合いが変わる。\n"
     "【現在値型の場合】対象キーの既存レコードがない場合の初期quantity方針\n"
     "案A: 負の在庫でINSERT（quantity=−調整数）／案B: INSERTしない／案C: quantity=0でINSERT\n"
     "【履历追加型の場合】「直近の最新行」が存在しない場合の初期値方針（同様の論点）", None],
    ["QA-A01\n★中\n(既存・②回答済)", "在庫調整コードの記録要否",
     "【②回答済・2026-08-20】出典『新WMS 調整コード一覧』。\n"
     "ADJ01＝サイクルカウント差異調整／ADJ02＝問題商品解決の在庫差異／"
     "DMG05＝補充１（昼）作業破損。ほかDMG/RTV/OUT等。\n"
     "2022/8/3：西友MD基幹へ新コードをそのまま連携。在庫調整実績IFに○。\n"
     "【①未決】center_stock / current_stock_data に格納先がない。"
     "TRIAL連携でも必須か、破棄してよいか、別テーブルにするかは未決。", None],
    ["QA-G06\n★中\n(更新)", "排他制御・重複防止方式の定義",
     "「在庫調整データ」は随時送信の差分更新IF。同一(store_code,product_code)に対して"
     "複数調整データが同時に送信される可能性がある。UNIQUE制約の有無（QA-S01）と合わせて、"
     "重複防止・排他制御の方式を定義する必要がある。\n"
     "案A: DB行ロック（SELECT FOR UPDATE）／案B: トランザクション内UPSERT／案C: その他", None],
    ["QA-G05\n★低\n(既存→一部解消)", "在庫調整日付の時刻精度（QA-S02に統合）",
     "本論点はQA-S02（4時刻列の用途定義）に統合済み。個別回答は不要。", None],
]

# ============================================================
# No.52 在庫一覧データ（IFSYLMS310R）
# ============================================================

MAP52_HEADERS = MAP51_HEADERS
MAP52_WIDTHS = MAP51_WIDTHS

MAP52_ROWS = [
    ["—", "—", "—", "—", "id", "BIGSERIAL\nPRIMARY KEY", "自動生成",
     "TRIAL側で自動採番。Hitluster側に対応項目なし。（旧版と同じ）", None],
    ["2", "倉庫コード", "文字列(6)\n003-008", "西友倉庫コード\n例:\"007455\"",
     "store_code", "INT4",
     "変換必要",
     "【要定義】倉庫コード→店舗コードのマッピングテーブル定義が必要。実データ例:store_code=580。"
     "No.51と共通の映射ルール（QA-G01、変更なし）", None],
    ["4", "商品コード", "文字列(14)\n017-030",
     "先頭スペース1byte+13桁JAN\n例:\" 4522646963267\"",
     "product_code", "VARCHAR(13)",
     "変換必要",
     "TRIM()で先頭スペース除去→13桁JANとして格納。"
     "【新規注意】center_stockはVARCHAR(13)ちょうどでオーバーフロー余裕なし（QA-G07・No.51と共通）",
     "NEW"],
    ["5", "在庫数", "数値(7)\n031-037", "商品在庫数+不良品数\n整数(7桁)\n例:0000000〜0000866",
     "quantity", "NUMERIC",
     "確認必要\n（要件変更）",
     "【要決定・最重要】毎日22:00に全件（サンプル3545行/日）を送信する全件スナップショット型IF。\n"
     "center_stockが履历追加型の場合: 当日分をそのまま新規行として一括INSERT"
     "→ 日々レコードが積み上がるためデータ量増大対策（保管期間/パーティション等）の検討が必要"
     "（QA-S04・新規）。\n"
     "center_stockが現在値型の場合: (store_code,product_code)にUNIQUE制約を新設し"
     "DELETE→INSERT or UPSERTで反映（旧版と同様の論点）。\n"
     "→ いずれもQA-S01（書込方式）の結論に依存。在庫数0レコードの扱いは変更なし（QA-B03）",
     "CRIT"],
    ["6", "移動平均単価", "数値(9)\n038-046", "整数7桁+小数2桁\n例:0001234.56\n=1,234.56円",
     "— （対応列なし）", "—",
     "対応なし\n（GAP拡大）",
     "旧版はamount列（未使用）が存在したが、center_stockには対応列が一切ない。"
     "恒久的に破棄でよいか改めて確認（QA-G04・優先度UP、No.51と共通）", "NEW"],
    ["3", "データ抽出日付", "文字列(8)\n009-016", "yyyymmdd\n例:20260625",
     "stock_date /\noperation_time", "TIMESTAMP /\nTIMESTAMPTZ",
     "確認必要",
     "4種の時刻列のうちどれに対応するか未確定（QA-S02）。仮説: stock_date=抽出対象日、"
     "operation_time=抽出発生時刻（日付のみのため暫定00:00:00+09）", "NEW"],
    ["（新規対応項目）", "IF受信/転送起因", "—", "—",
     "transfer_time", "TIMESTAMPTZ",
     "確認必要\n（新規）",
     "No.51と同様、IF受信・連携処理のタイムスタンプと想定。自動設定可否を要確認（QA-S02）", "NEW"],
    ["（新規対応項目）", "DB書込起因", "—", "—",
     "registered_at", "TIMESTAMPTZ",
     "確認必要\n（新規）",
     "DB書込時刻。DB側DEFAULTで自動設定される可能性が高い（QA-S02）", "NEW"],
    ["（新規対応項目）", "データ発生源識別", "—", "—",
     "registered_by", "VARCHAR(20)",
     "確認必要\n（新規項目）",
     "本IF（No.52 在庫一覧）由来レコードの識別値を新規定義（例:\"IF_LIST\"）。"
     "No.51（例:\"IF_ADJUST\"）と区別できる命名規則を合意（QA-S03、No.51と共通）", "NEW"],
    ["1", "データ区分", "文字列(2)\n001-002", "\"31\"固定",
     "—", "—", "不要",
     "変更なし。TRIAL側では使用しない。", None],
    ["7", "FILLER", "文字列(466)\n047-512", "半角スペース固定",
     "—", "—", "不要",
     "変更なし。パディング領域。TRIAL側では使用しない。", None],
]

GAP52_HEADERS = GAP51_HEADERS
GAP52_WIDTHS = GAP51_WIDTHS

GAP52_ROWS = [
    ["【訂正】TRIAL側参照テーブルの対象誤り", "★最高", "center_stock\n（全項目）", "全項目",
     "旧資料は`current_stock_data`を参照していたが、これは実在するものの店舗（店CD）単位の"
     "現在庫テーブルと推測され、倉庫コード基準のNo.51/No.52とは対象が不一致（No.51と同一の訂正内容）",
     "倉庫コード基準のIFはセンター/倉庫単位と推測される`center_stock`が対象と考えられる。\n"
     + CENTER_STOCK_SCHEMA,
     "本資料ではcenter_stockを正として整理。current_stock_dataとの役割分担の最終確認はTRIAL側に"
     "お願いしたい（QA-G01/QA-S01参照）", "CRIT"],
    ["書き込み方式未確定", "★最高", "center_stock\n（全体）", "在庫数（全件スナップショット）",
     "毎日22:00に全件（約3545行）を送信、\n差分ではなく絶対値",
     "center_stockにUNIQUE制約が確認できず、\n履历追加型（ログ/新規行）の可能性が高い",
     "No.51と共通論点（QA-S01）。履历追加型なら「日次スナップショットがそのまま履歴として"
     "積み上がる」設計になり、現在値型なら従来通りDELETE→INSERT/UPSERTが必要", "CRIT"],
    ["マッピング不足：倉庫コード→店舗コード", "★高", "center_stock\n(store_code列)", "倉庫コード(No.2, 003-008)",
     "西友倉庫コード（文字列6桁）\n例:\"007455\"",
     "INT4型のstore_codeが必要\n実データ例: store_code=580",
     "【要確認】倉庫コード→店舗コードのマッピングテーブル定義を業務側に確認（QA-G01、No.51と共通）", None],
    ["履历追加時のデータ量増大", "★高\n(新規)", "center_stock\n（全体）", "在庫数（全件スナップショット）",
     "毎日3545行規模の全件スナップショットを送信",
     "履历追加型の場合、日次で\nレコードが無制限に積み上がる",
     "保管期間・アーカイブ/パーティション方針の検討が必要（QA-S04・新規、QA-S01の結論に依存）", "NEW"],
    ["移動平均単価の格納先消失", "★高\n(優先度UP)", "center_stock\n(該当列なし)", "移動平均単価(No.6, 038-046)",
     "数値(9): 整数7桁+小数2桁\nHitluster側は必ず送信",
     "center_stockには該当列が\n存在しない",
     "恒久的に破棄でよいか改めて確認（QA-G04、No.51と共通）", "NEW"],
    ["型不一致（オーバーフローリスク上昇）", "★中", "center_stock\n(product_code列)", "商品コード(No.4, 017-030)",
     "固定長14byte\n（先頭スペース1byte+13桁JAN）",
     "VARCHAR(13)ちょうど\n（旧VARCHAR(20)から余裕消失）",
     "TRIM()で13桁化して格納。桁超過時のエラーハンドリングを検討（QA-G07、No.51と共通）", "NEW"],
    ["時刻列の割当未確定", "★中", "center_stock\n(4時刻列)", "データ抽出日付(No.3, 009-016)",
     "日付のみ（yyyymmdd）",
     "4種類の時刻列が存在するが\n用途が未確定",
     "暫定案を本資料に記載。TRIAL側に正式定義を確認（QA-S02、No.51と共通）", "NEW"],
    ["registered_by設定値未定義", "★中", "center_stock\n(registered_by列)", "（新規論点）",
     "—", "VARCHAR(20)\n実データ例:\"Migration\"",
     "本IF（No.52）由来レコードの識別値を新規定義（例:\"IF_LIST\"）。No.51と区別できる命名規則を合意（QA-S03）", "NEW"],
    ["IF間整合性：No.51とNo.52の適用順序", "★中\n(内容更新)", "center_stock\n(No.51+No.52両方)", "—",
     "No.51: 随時送信（差分型）\nNo.52: 毎日22:00送信（全件スナップショット）\n両IFが同一テーブルを更新",
     "書込方式（履历追加/現在値）に\より結論が変わる",
     "【現在値型の場合】適用順序の定義が必須（案A/B、旧版と同様）\n"
     "【履历追加型の場合】「順序」ではなく「同一時点の最新行をどちらのIFの行として採用するか」の"
     "定義が必要（QA-B01・内容更新、QA-S01の結論に依存）", None],
    ["在庫数=0レコードの扱い", "★中\n(変更なし)", "center_stock\n(quantity列)", "在庫数(No.5)",
     "サンプルで在庫数0が63.4%\n（2248/3545行）",
     "quantity DEFAULT 0",
     "格納する/しないの方針確認（QA-B03、変更なし）", None],
    ["欠損時の挙動", "★中\n(変更なし)", "center_stock\n（全体）", "—",
     "全件スナップショット型のため\n通信障害等でデータ欠損の可能性",
     "center_stockの整合性を\n維持する必要がある",
     "欠損時のフォールバック方針を確認（QA-B04、変更なし）", None],
    ["データ種別の差異", "★低", "—", "データ区分(No.1) + FILLER(No.7)",
     "データ区分=\"31\"固定\nFILLER=半角スペース466byte",
     "TRIAL側では使用しない",
     "受信時に破棄。TRIAL側で利用なし（変更なし）", None],
]

CONV52_HEADERS = CONV51_HEADERS
CONV52_WIDTHS = CONV51_WIDTHS

CONV52_ROWS = [
    ["IFSYLMS310R\n(在庫一覧データ)", "倉庫コード\n(No.2, 003-008)", "文字列(6)", "center_stock", "store_code",
     "INT4",
     "マッピングテーブル参照:\nSTORE_MAPPING = {\n  \"007452\": ???,\n  \"007455\": ???\n}\n"
     "store_code = STORE_MAPPING[倉庫コード]",
     "\"007455\" → ???\n（実データ例: store_code=580）",
     "【要確認】倉庫コード→店舗コードの対応表は業務側から提供必要（QA-G01、No.51と共通）", None],
    ["IFSYLMS310R\n(在庫一覧データ)", "商品コード\n(No.4, 017-030)", "文字列(14)\n(先頭スペース1+13桁JAN)",
     "center_stock", "product_code", "VARCHAR(13)",
     "TRIM()で先頭スペース除去→13桁JAN文字列としてそのまま格納\nproduct_code = TRIM(sinops.商品コード)",
     "\" 4522646963267\"\n→ \"4522646963267\"",
     "VARCHAR(13)ちょうどでオーバーフロー余裕なし。桁超過時のエラーハンドリングを検討（QA-G07）", "NEW"],
    ["IFSYLMS310R\n(在庫一覧データ)", "在庫数\n(No.5, 031-037)", "数値(7)\n(整数)",
     "center_stock", "quantity", "NUMERIC",
     "【方式未確定・要合意】\n"
     "【履历追加型の場合】\n"
     "  当日分（約3545行）をそのまま新規行として一括INSERT\n"
     "  INSERT INTO center_stock (stock_date, store_code, product_code, quantity, ...)\n"
     "  VALUES (当日, store_code, product_code, 在庫数, ...)  -- 全行分\n"
     "【現在値型の場合】\n"
     "  -- 方式A: DELETE→INSERT\n  DELETE FROM center_stock WHERE store_code=対象店舗;\n  INSERT ...(全行);\n"
     "  -- 方式B: 行ごとUPSERT（ON CONFLICTには事前にUNIQUE制約が必要）",
     "在庫数 0000866\n→ quantity = 866\n\n在庫数 0000000\n→ quantity = 0",
     "書込方式の結論（QA-S01）待ち。履历追加型の場合はデータ量増大対策が別途必要（QA-S04）。"
     "在庫数=0レコードの扱いは変更なし（QA-B03）", "CRIT"],
    ["IFSYLMS310R\n(在庫一覧データ)", "移動平均単価\n(No.6, 038-046)", "数値(9)\n(整数7+小数2)",
     "center_stock", "— (対応列なし)", "—",
     "【破棄（現状）】\ncenter_stockに格納先が存在しないため、受信データは無視して破棄する",
     "単価 0001234.56\n→ 格納先なし・破棄",
     "旧版は「未使用フィールドあり」だったが、列自体が存在しない点を業務側に共有（QA-G04）", "NEW"],
    ["IFSYLMS310R\n(在庫一覧データ)", "データ抽出日付\n(No.3, 009-016)", "文字列(8)\n(yyyymmdd)",
     "center_stock", "stock_date /\noperation_time", "TIMESTAMP /\nTIMESTAMPTZ",
     "【仮説・要確認】\nstock_date = TO_DATE(日付,'YYYYMMDD')\noperation_time = "
     "TO_TIMESTAMP(日付,'YYYYMMDD') AT TIME ZONE 'Asia/Tokyo'",
     "\"20260625\"\n→ stock_date=2026-06-25\n→ operation_time=2026-06-25 00:00:00+09",
     "4時刻列の正式な用途定義が必要（QA-S02、No.51と共通）", "NEW"],
    ["IFSYLMS310R\n(在庫一覧データ)", "（新規：発生源識別）", "—",
     "center_stock", "registered_by", "VARCHAR(20)",
     "【新規要定義】\nregistered_by = 'IF_LIST' （仮称・要合意）",
     "実データ例:\"Migration\"",
     "No.51（例:\"IF_ADJUST\"）と区別できる命名規則を開発チームと合意（QA-S03）", "NEW"],
    ["IFSYLMS310R\n(在庫一覧データ)", "データ区分\n(No.1, 001-002)", "文字列(2)\n(\"31\"固定)",
     "center_stock", "—", "—",
     "【破棄】\nTRIAL側では使用しない。レコード識別用のみ",
     "—", "内部検証用。\"31\"以外受信時はエラーとする（変更なし）", None],
    ["IFSYLMS310R\n(在庫一覧データ)", "FILLER\n(No.7, 047-512)", "文字列(466)\n(半角スペース固定)",
     "center_stock", "—", "—",
     "【破棄】\nパディング領域、TRIAL側で利用なし",
     "—", "スペース以外が含まれる場合はデータ破損の可能性（変更なし）", None],
]

QA52_HEADERS = QA51_HEADERS
QA52_WIDTHS = QA51_WIDTHS

QA52_ROWS = [
    ["QA-S01\n★最高\n(共通・新規)", "center_stockの書き込み方式（履历追加型か現在値型か）",
     "No.51と共通論点。center_stockにUNIQUE制約が確認できない。履历追加型か現在値型かで、"
     "全件スナップショット（約3545行/日）の反映方式が根本的に変わる。\n"
     "案A: 履历追加型。日次スナップショットをそのまま新規行として積み上げる\n"
     "案B: 現在値型。実際は別途UNIQUE制約が存在する（確認漏れ）\n"
     "案C: No.51/No.52で異なるテーブル/方式が使い分けられている", None],
    ["QA-S04\n★高\n(新規)", "履历追加型の場合のデータ量増大対策",
     "全件スナップショットが履历追加型でそのまま蓄積される場合、1日あたり約3545行×365日/年で"
     "急速にテーブルが肥大化する。保管期間、パーティショニング、アーカイブ方針の検討が必要。", None],
    ["QA-G01\n★高\n(既存)", "倉庫コード→store_codeのマッピング定義（center_stock/current_stock_data役割分担含む）",
     "TRIAL側store_codeはINT4型、Hitluster側倉庫コードは文字列6桁。対応テーブル定義が必要。\n"
     "例: \"007455\"→何番？（実データ例:store_code=580）\n"
     "【追記】別途実在する『current_stock_data』（店舗単位の現在庫テーブルと推測）との役割分担も"
     "合わせて確認したい（No.51 QA-G01参照）。\n※No.51と共通ルール", None],
    ["QA-S02\n★中\n(共通・新規)", "4つの時刻列の用途定義",
     "No.51と共通論点。stock_date/operation_time/transfer_time/registered_atの正式な業務的意味と、"
     "IF側で明示的に設定すべきか、DB側で自動設定されるかを確認したい。", None],
    ["QA-S03\n★中\n(共通・新規)", "registered_byに設定すべき値",
     "No.51と共通論点。本IF（在庫一覧データ）由来レコードにどの値を設定すべきか（例:\"IF_LIST\"）。"
     "No.51用の値と区別できる命名規則を確認したい。", None],
    ["QA-G04\n★高\n(優先度UP)", "移動平均単価の値は本当に不要か（格納先が消失）",
     "center_stockにはamount相当の列が存在しない。将来利用予定はあるか、完全に破棄してよいか確認したい。"
     "（No.51と共通）", None],
    ["QA-G07\n★中\n(共通・新規)", "商品コードがVARCHAR(13)ちょうどで桁オーバーフローした場合の扱い",
     "No.51と共通論点。TRIM後13桁を超えるコードを受信した場合のエラーハンドリング方針を確認したい。", None],
    ["QA-B01\n★中\n(内容更新)", "在庫調整データ（差分）と在庫一覧データ（全件）の適用順序",
     "書込方式（QA-S01）の結論により論点が変わる。\n"
     "【現在値型の場合】従来通り適用順序の定義が必須\n"
     "案A: No.51→No.52の順（No.52が最新で上書き）\n案B: No.52→No.51の順（No.51が最新反映）\n"
     "【履历追加型の場合】適用順序ではなく、同一(store_code,product_code)について"
     "「どのIF由来の行を最新として参照するか」の判定ルール（registered_by/stock_date等での判定）が必要", None],
    ["QA-B03\n★中\n(既存)", "在庫数=0レコードの扱い",
     "サンプルデータで在庫数0が63.4%（2248/3545行）を占める。\n"
     "案A: 格納する（quantity=0として保持、「在庫0」に業務的意味あり）\n"
     "案B: 格納しない（除外、レコード数削減）", None],
    ["QA-B04\n★中\n(既存)", "データ欠損時のフォールバック方針",
     "全件スナップショット型のため、通信障害等でデータ欠損時に前回データとの差分補完ができない。\n"
     "案A: 前回データを維持（何もしない）／案B: 在庫全0でクリア／案C: 手動介入待ち", None],
]


def build_all():
    specs = [
        ("_在庫IF分析_No51_在庫調整データ_マッピング表_訂正版.xlsx", MAP51_HEADERS, MAP51_ROWS, MAP51_WIDTHS, "マッピング表"),
        ("_在庫IF分析_No51_在庫調整データ_GAP分析書_訂正版.xlsx", GAP51_HEADERS, GAP51_ROWS, GAP51_WIDTHS, "GAP分析書"),
        ("_在庫IF分析_No51_在庫調整データ_変換ルール定義書_訂正版.xlsx", CONV51_HEADERS, CONV51_ROWS, CONV51_WIDTHS, "変換ルール定義書"),
        ("_在庫IF分析_No51_在庫調整データ_QA一覧_訂正版.xlsx", QA51_HEADERS, QA51_ROWS, QA51_WIDTHS, "QA一覧"),
        ("_在庫IF分析_No52_在庫一覧データ_マッピング表_訂正版.xlsx", MAP52_HEADERS, MAP52_ROWS, MAP52_WIDTHS, "マッピング表"),
        ("_在庫IF分析_No52_在庫一覧データ_GAP分析書_訂正版.xlsx", GAP52_HEADERS, GAP52_ROWS, GAP52_WIDTHS, "GAP分析書"),
        ("_在庫IF分析_No52_在庫一覧データ_変換ルール定義書_訂正版.xlsx", CONV52_HEADERS, CONV52_ROWS, CONV52_WIDTHS, "変換ルール定義書"),
        ("_在庫IF分析_No52_在庫一覧データ_QA一覧_訂正版.xlsx", QA52_HEADERS, QA52_ROWS, QA52_WIDTHS, "QA一覧"),
    ]
    for fname, headers, rows, widths, sheet_name in specs:
        wb = Workbook()
        make_sheet(wb, fname, headers, rows, widths, sheet_name=sheet_name)
        path = os.path.join(OUT_DIR, fname)
        wb.save(path)
        print("saved:", path)


if __name__ == "__main__":
    build_all()
