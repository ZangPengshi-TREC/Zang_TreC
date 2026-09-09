# 発注勧告IF Shinise直接データ版 整理メモ

## 対象

Shinise `auto_orders` テーブルの実データ構造を基準に、
sinops店舗系発注勧告（`kankoku.txt` / `subkankoku.txt`）および
倉庫系発注勧告（`【倉庫CD】ka.txt` / `subka.txt`）との対応を整理する。

本資料は `SIREJAN_RCMDORDER_LAST` 版とは別資料として管理する。

## 物理テーブル

`auto_orders` の物理列は以下の5列。

- `id`：bigserial
- `data`：jsonb
- `registered_by`：varchar(8)。既存サンプル値は継続使用せず、新規採番が必要
- `registered_at`：timestamptz
- `action_type`：int2

発注勧告の業務項目は `data` JSON内に格納される。

## data JSONの項目

- `seq`
- `order_date`
- `order_type`
- `store_code`
- `product_code`
- `delivery_date`（日配対象のみ）
- `order_quantity`
- `order_closing_number`

## 日配／非日配の違い

### 日配対象

- `order_type = 1`
- `delivery_date` を送信する
- `delivery_date` は店舗系No.4／倉庫系No.14 納品日を `yyyy-mm-dd` に変換

### 非日配対象

- `order_type = 0`
- `delivery_date` は送信不要
- JSON内に `delivery_date` キー自体を出力しない
- `null` や空文字としても送信しない

## 主要マッピング（店舗系）

- `order_date` ← No.2 発注日
- `store_code` ← No.1 店舗コード
- `product_code` ← No.3 商品コード
- `delivery_date` ← No.4 納品日（日配のみ）
- `order_quantity` ← **No.11 発注バラ数**
- `order_closing_number`：**発注締番号**。sinopsファイルに直接対応項目なし

`order_closing_number` に No.5 便区分は使用しない。
発注締めマスタ、処理時刻または処理回等からの正式な取得規則を確認する。

`order_quantity` は発注単位調整後のバラ総数であり、次の関係が成立する。

`No.11 発注バラ数 = No.12 発注単位数 × No.15 発注単位`

No.11を整数値のまま直接設定し、1000倍しない。

## 主要マッピング（倉庫系）

- `order_date` ← No.1 発行日
- `store_code` ← **ファイル名の倉庫コード**
- `product_code` ← No.7 品番
- `delivery_date` ← No.14 納品日（日配のみ）
- `order_quantity` ← **No.11 バラ換算数量**
- `order_closing_number`：**発注締番号**。sinopsファイルに直接対応項目なし

倉庫コードは数値化して `store_code` に設定する。
例：ファイル名の `007452` → `store_code: 7452`

倉庫系の数量関係は次のとおり。

`No.11 バラ換算数量 = No.9 注文数量 × No.12 発注ロット`

No.11を整数値のまま直接設定し、1000倍しない。

## サンプルから確認できた事項

- 日配レコードには `delivery_date` が存在する
- 非日配レコードには `delivery_date` が存在しない
- `order_quantity = 0` のレコードも存在する
- `registered_by` には既存値が入っているが、今回のIFでは継続使用できない
- `action_type` のサンプル値は `0`
- `registered_at` はタイムゾーン付き日時

## 未確認事項

1. 日配／非日配を判定する正式なマスタ・項目
2. `seq` の採番単位
3. `action_type` のコード定義
4. `registered_by` の新規8桁コード採番、および利用環境・管理責任者
5. 再送時の重複防止・UPSERT条件
6. `registered_at` の基準タイムゾーン
7. `order_quantity = 0` の業務上の意味
8. `order_closing_number`（発注締番号）の取得元・採番規則

## 成果物

- `発注勧告_①_項目マッピング表.xlsx`
- `発注勧告_②_GAP分析書.xlsx`
- `発注勧告_③_変更ルール定義書.xlsx`
