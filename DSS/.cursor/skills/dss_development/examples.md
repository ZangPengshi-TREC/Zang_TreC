# 样板与培训演示

设计新 IF 时先对齐这些已证实的形态，再按台账改连接和路径。

## 实演样板：`CM0010_IF0079_006` / `01OutputData`

来源：2026-05-13 Teams（田村 / `seiyu_tamura`）。台账页 `SEIYU-TRIALインターフェース管理台帳` → `DSSインターフェース台帳`。Drive 资料夹：`9商品マスタ(惣菜)/DSS`。

| 项 | 值 |
|---|---|
| FROM | CM0010 若菜 新店舗システム |
| TO | IF0079 TRIAL データ連携ストレージ（也有反向 `IF0079_CM0010`、以及 `IF0079_MS0007` 进数据基盤） |
| 业务组 | 生鮮発注 / 惣菜 / 非生鮮発注 / EC |
| 启动 | 多为スケジュール；サイクル 日次・週次・随時 |
| Oracle | 接続 `ORCL_CM0010_wknpsuser`；スキーマ `WKNPSUSER`；例 `WKNST_DAILYORDER`（`COMPANY_CD` / `STORE_CD` / `SALES_DATE=YYYYMMDD`） |
| 工作区 | `/seiyuroot/CM0010_IF0079_006/` |

图标顺序：

| 图标 | 作用 |
|---|---|
| try / catch | 框架用的异常监视；catch 接 `out_*` |
| delete | 清工作区 `/*` |
| `exec_select_CM0010` | 検索系 SQL。`select * from WKNPSUSER…` |
| 現在日付取得 | Mapper 出 `file_date`、`folder_date` |
| `csv_write` | 写出 CSV |
| endfile | 空完了文件 `ORDER_${file_date}.END` |
| `gcs_put` | `GCP_STR_DSS` → 桶 `ods-seiyu-dev`，路径 `/CM0010_IF0079_006/${folder_date}/`，本地 `*` 全上传；无文件夹则建，已有则覆盖 |
| 服务注册 | 项目右键「プロジェクトをサービスとして登録」 |

同族 `CM0010_IF0079_005` 在 GCS 上还有 `deliveryData/日付フォルダ` 与 `deliveryData_bk`。

另外演示过 FTP「リストとデータの読み取り」`ftp_get_data`、连接 `SFTP-cao-test`。Explorer 访问 `/data/Shared_test`（`\\10.10.1.125\Shared_test`）当时失败——**不要把这条挂载写进新脚本**。

## Day1 三个演示（SAMPLE / DEMO 路径）

培训工作目录 `/data/DSSROOT/DEMO/`（OS 上 `F:\DSSROOT\DEMO\`）。新西友 IF **不要**用 DEMO 路径。

| 演示 | 输入 | 要做出的行为 |
|---|---|---|
| ① CSV 转换集计 | `OrderList.csv`（受注日、支店、ケース数、入り数、容量、商品名、受注_ID） | 按支店汇总发注数，写出 `OrderList_total.csv` |
| ② 0 件当错误 | 只有表头的 `OrderList.csv` | 不写输出文件，主动报错（框架下对应警告或 255，按台账） |
| ③ 文件触发器 | 把 CSV 放到输入文件夹 | 触发 ① 的脚本 |

CSV 读：Windows-31J，跳过首行当表头。

## Day2：框架改写 + SQL Server

练习库 `SAMPLE_PROJECT`。SQL 在 `DHUB-DB-DEV` / `DSSROOT`：`if_master`、`param_master`、`DEMO_OrderList`。

把 Day1 脚本改成框架形态：

- 正常 = 跑完，戻り値 0
- 警告 = 0 件，戻り値 1–254
- 异常 = 读或写失败，戻り値 255 + `out_*`

SQL Server 适配器：表读/写、検索SQL、更新SQL（`?{列名}`）、ストアド、ローダ。连接例 `COM_SQLServer2022`。

**DBデモ1**（CSV→表，带事务）：

```
Transaction → DELETE → csv_read → mapping → テーブル書き込み → COMMIT
```

`process_date` 用「現在日時」。组件「データ処理方式」可选「スクリプトの設定を使用する」；大件数不要全进内存。

## 读 / 转 / 写怎么选

| 场景 | 用什么 | 注意 |
|---|---|---|
| DB 全表抽出 | テーブル読み取り | 全局资源可复用 |
| 带条件抽出 / 复杂 SQL | 検索系 SQL 実行 | 向导生成后若手改 SQL，向导就不能再编 |
| 字段变换、过滤、日期、拼接 | Mapper（约 140 个逻辑） | 输出 schema 要先定；CSV 可用表头文件 |
| 文件 JOIN | Merge（APPEND / INTEGRATE / Left-Right JOIN） | DB 侧优先用 SQL JOIN |
| 大量写入 / 主数据洗い替え | テーブル書き込み + 批处理 | 勾选「キー一致行は更新」= upsert |
| 复杂更新 / 外结合 | 更新系 SQL 実行 | 不适配超大件数 |
| 固定长 EDI | 固定长适配器 + 向导 | Group / Record / Field；定义进全局资源 |

## 性能口径（写大件数时）

| 机制 | 行为 | 怎么用 |
|---|---|---|
| PSP（并行流） | 约 1000 件一块，读/转/写多线程；块大小不可改 | 智能编译器默认可开 |
| 大容量データ処理 | 内存只留必要数据，其余落盘 | PSP 优先；PSP 后若 Mapper 接到非 PSP 写入，中间仍会走大容量 |
| 批写入 | テーブル書き込み批量模式 | 86 万件时批量 upsert 约比普通 upsert 快 30 倍 |
| 1 亿件级 | 验证中 PG 7h 仍失败；tmp 可超 10GB | **不要用 DSS 扛** |

卡住时：My Log → 任务管理器结束进程；不行就重启服务。Desktop Studio 模块同步可超 4 小时，西友以 Web 为主。
