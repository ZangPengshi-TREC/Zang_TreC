# Phase 2 IF 一覧（Layer1–3）

基準: `sources/IF一覧表_Layer1-3/【Phase2_Sinops・BO】IF一覧表_Layer1-3.xlsx`  
扱い: **詳細設計の IF 一览。** 人日は本表に無い。工数は見積 403 のまま。生鮮増分（2026-09-18 会議）は未掲載。  
シート: `00_凡例・読み方` / `01_IF一覧_自動補充Sinops系` / `02_IF一覧_BO系`（Hulft テンプレが残る） / `BK_02_IF一覧_BO系`

---

## 1. 凡例の層

| 層 | 範囲 | 本表の処理内容 |
|---|---|---|
| Layer1 | TRIAL MD基幹 ↔ GCS | 外部 IF。抽出してファイル生成し GCS へ。受信側は逆 |
| Layer2 | GCS → 西友 DWH | DataSpider Intake。**マスタのみ。** トランは DWH を通さず Layer3 へ |
| Layer3 | DWH・GCS ↔ 業務 | DataSpider。本表は「コード変換・編集して Sinops 連携ファイルを生成」。受信側は取込 |

表題は「Sinops 系 25 本・BO 系 4 本」。Sinops は IF 番号 01–25（10 が店/センター 2 行で 26 データ行）。BO は BK シートが **33/34 の 2 系統**。4 本との差は未整理。

タイミング注記: TRIAL 更新完了時刻・Sinops が欲しい時刻は要確認。表上の既定は往路 **GCS／日次 02:00**。

現行アーキテクチャとの差（一览を優先して IF 有無を書く。経路の細部はアーキテクチャ側）:

- データ元は表上「基幹システム」。抽出は集計 SV／クラウド機、という置き方は `SCM-IFクラウド機アーキテクチャ.md`
- Layer3 の「コード変換・編集」は、申請話術の「DSS は二次加工しない」と張力がある。一览の Layer3 列を変換の正とするか、クラウド機 Layer1 に寄せるかは未凍結
- 22–25 の連携先は表上「TRIAL MD基幹」。連携 GCS → クラウド機 → 自動発注 GCS はアーキテクチャ側の読み

---

## 2. 自動補充 Sinops 系

凡例: L1/L2/L3 は表の「有無」。ファイル名の（仮）は未確定。

| IF | ファイル | 名称 | 区分 | 方向 | L1 | L2 Intake | L3 先 | 主な変換・注 |
|---|---|---|---|---|---|---|---|---|
| 01 | item_mst.txt | 商品マスタ | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_ITEM_MST 洗替 | Sinops-R 差分 | 分類 4+4+4 桁拡張 |
| 02 | store_item_mst.txt | 店舗商品マスタ | マスタ | 基幹→Sinops | **差分**／02:00 | INTAKE_STORE_ITEM_MST 洗替 | Sinops-R 洗替※1 | 店舗コード CV-001 |
| 03 | category_mst.txt | カテゴリマスタ | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_CATEGORY_MST 洗替 | Sinops-R 洗替 | 階層桁マッピング |
| 04 | supplier_mst.txt | 仕入先マスタ | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_SUPPLIER_MST 洗替 | Sinops-R 洗替 | **現行 9 桁のまま** |
| 05 | case_bara.txt | ケースバラ変換 | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_CASE_BARA 洗替 | Sinops-R 洗替 | 四パターン |
| 06 | pos.txt | 販売実績 | トラン | 基幹→Sinops | 勧告計算前日／02:00 | － | Sinops-R 差分 | |
| 07 | hour_pos.txt | 時間帯別販売数量 | トラン | 基幹→Sinops | 同上 | － | Sinops-R 差分 | |
| 08 | kyaku_su.txt | 来客数実績 | トラン | 基幹→Sinops | 同上 | － | Sinops-R 差分 | |
| 09 | haiki.txt | 廃棄情報 | トラン | 基幹→Sinops | 同上 | － | Sinops-R 差分 | |
| 10 店 | zaiko_shusei.txt | 在庫修正（店舗） | トラン | 基幹→Sinops | 同上 | － | Sinops-R 差分 | |
| 10 センター | zaiko_shusei_dc.txt（仮） | 在庫修正（センター） | トラン | 基幹→Sinops | 同上 | － | **－** | Layer3 は店舗分に内包 |
| 11 | nyuka.txt | 入荷実績 | トラン | 基幹→Sinops | 同上 | － | Sinops-R 差分 | 便区分の付与 |
| 12 | nyuka_yotei.txt | 入荷予定 | トラン | 基幹→Sinops | 同上 | － | Sinops-R 差分 | 便区分の付与 |
| 13 | sii.txt | 発注スケジュール | マスタ | 基幹→Sinops | 当日以降全件 21 日／02:00 | INTAKE_SII 洗替 | Sinops-R 洗替 | |
| 14 | location_history.txt | 棚割明細 | マスタ | **双方向** | 当日以降全件／02:00 受送信 | INTAKE_LOCATION_HISTORY 洗替 | Sinops-R 洗替 | |
| 15 | shinshohin_zaiko.txt（仮） | 新商品用在庫 | トラン | 基幹→Sinops | 新規。西友 DWH 棚割＋TRIAL MD マスタ | －（14 の Intake 参照） | Sinops-R 差分 | |
| 16 | item_mst_dc.txt（仮） | 商品マスタ【倉庫】 | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_ITEM_MST_DC 洗替 | Sinops-R（倉庫）差分 | 分類桁、発注ロット＝店別調達の発注単位 |
| 17 | supplier_mst_dc.txt（仮） | 仕入先マスタ【倉庫】 | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_SUPPLIER_MST_DC 洗替 | Sinops-R（倉庫）洗替 | |
| 18 | supplier_holiday_dc.txt（仮） | 仕入先休日【倉庫】 | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_SUPPLIER_HOLIDAY_DC 洗替 | Sinops-R（倉庫）洗替 | |
| 19 | supplier_order_dow_dc.txt（仮） | 仕入先発注曜日【倉庫】 | マスタ | 基幹→Sinops | 全件／02:00 | INTAKE_SUPPLIER_DOW_DC 洗替 | Sinops-R（倉庫）洗替 | |
| 20 | order_result_dc.txt（仮） | 発注実績【倉庫】 | トラン | 基幹→Sinops | 勧告計算前日／02:00 | － | Sinops-R（倉庫）差分 | |
| 21 | ukebarai_dc.txt（仮） | 受払明細【倉庫】 | トラン | 基幹→Sinops | 同上 | － | Sinops-R（倉庫）差分 | 伝票区分マッピング |
| 22 | kankoku_dc.txt（仮） | 発注勧告当日【倉庫】 | トラン | Sinops→基幹 | 受信。GCS／勧告計算後 | － | TRIAL MD基幹 | 勧告数の受信・取込 |
| 23 | subkankoku_dc.txt（仮） | 発注勧告翌日以降【倉庫】 | トラン | Sinops→基幹 | 同上 | － | TRIAL MD基幹 | 障害用の未来分 |
| 24 | kankoku.txt | 発注勧告当日 | トラン | Sinops→基幹 | 同上 | － | TRIAL MD基幹 | 通常採用。勧告数の受信・取込 |
| 25 | subkankoku.txt | 発注勧告翌日以降 | トラン | Sinops→基幹 | 同上 | － | TRIAL MD基幹 | 障害時、または当日未作成時に昨日ファイルで代替。フローは24と同一 |

Layer2 あり: **01–05、13、14、16–19**（11 本）。見積 ② と一致。  
Layer3 なし: **10 センターのみ**。見積「③ は 1 行」と一致。  
倉庫の連携先表記は表上 `Sinops-R（倉庫）`。W 製品名との対応は未注記。

---

## 3. BO 系

`02` シートは Hulft / 流通 BMS の列が、EDI 発注ジョブの値で埋まっている。**IF 定義としては使わない。**

`BK_02` が見直し後の形:

- 商品マスタ（基本・店別）の 2 系統。仕入先・店舗・商品階層は対象外
- **BO-33** 国内流通向け商品マスタ（基本）。L1 全件→差異ファイル。L2/L3 は「BO-15 と一体」
- **BO-34** 国内流通向け商品マスタ（店別）。L1 全件 日次。L2 `INTAKE_BO_STORE_ITEM_MST` 洗替。L3 国内流通、店舗コード CV-001

本作業のプログラム対象（403）は従来どおり自動補充。BO は一览上の記載。

---

## 4. 一览がまだ書いていないこと

- 人日（見出しだけ「工数単位：人日」）
- 生鮮の新規 IF（勧告・入荷実績・入荷予定）
- 22–25 の発注 IF レイアウト名・自動発注 GCS
- SYS ID / 連携用 GCS フォルダ名（`from_to_連番`）
- ※1 の本文
