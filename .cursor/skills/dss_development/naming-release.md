# 命名、环境、发布

口令不写进本文件。主机/账号见《DSS開発の流れ》和ガイドライン P.103 / P.115–118。

## 命名

| 对象 | 规则 | 例子 |
|---|---|---|
| 接口 ID / 项目名 | `FROM6_TO6_SEQ3` 大写下划线 | `AM0002_MS0007_001` |
| 脚本名 | 两位序号 + 驼峰英文或罗马字 | `01updDelDate` |
| 常量 | 大写下划线 | `FILE_NAME` |
| 变量 | 小写下划线 | `file_name` |
| 参数 | `in_` / `_out` / `io_` 前缀 | `in_file_name` |
| 环境变量 | `ENV_` 前缀大写下划线 | `ENV_RETRY_COUNT` |
| 表名 | `M_` 主数据 / `T_` 事务 | `M_FILE_SEND` |
| ADLS 容器 | 接口 ID 改小写，只用连字符 | `am0002-ms0007-001` |
| 固定长/可变长格式 | `FIX`/`VAR` + `READ`/`WRITE` + 名称 | `FIX_READ_Time` |
| 全局资源（DB/SFTP 等） | 协议_系统6位[_Intake][_枝番] | `SFTP_AM0002_用户ID` |
| 触发器 | 接口ID`_SCHL` 或 `_SCHL_nnn` | `AM0002_MS0007_001_SCHL` |
| 工作文件 | 小写下划线英文 | `file_name` |

同一 FoodCore 源仍要拆全局资源 + 用户 ID，不要共用一个连接给所有 IF。

全局资源前缀常见：`SQL` / `ORCL` / `JDBC` / `ODBC` / `FTP` / `REST` / `MAIL` / `BLOB`。DSS 自用 SQL：`SQL_DSS_SEIYU_DB`。

- 写 Databricks：经 ADLS（MS0007）
- 读 Databricks：JDBC（MS0005），**不要开连接池**

## 路径与备份

- DSS 工作目录：`/seiyuroot/<IF_ID>/`
- SSH 密钥：`/data/SSH/<Intake>/xxxxx.ppk`
- ADLS：一 IF 一容器；外部用 SFTP；文件约 30 天后删除
- 备份原则进 Blob（`AZURE_STR_DSS`）。`/data/backup` 仅备选
- GCS：`ods-seiyu-{dev|stg|prod}`，东京 `asia-northeast1`、Standard、非公开。连接名 `GCP_STR_DSS`
- GCS 目录常见：`/<IF_ID>/${folder_date}/` 或 `/<IF_ID>/deliveryData/YYYYMMDD/`，另有 `deliveryData_bk`

## 三套环境（2026《DSS開発の流れ》）

日常按独立三套 WebStudio，不要沿用 2024 设计书「Staging 里放开发机、不单独建 Development」。

| | DEV | STG | PROD |
|---|---|---|---|
| WebStudio | `https://dhub-dss-dev.seiyu.com:8443/WebStudio/` | `https://dhub-dss-stg.seiyu.com:8443/WebStudio/` | `https://dhub-dss-pro.seiyu.com:8443/WebStudio/` |
| 账号 | 个人账号 | `exec_SY` | `exec_SY` |
| DB | `DHUB-DB-DEV` / `DSSSEIYU`，DBeaver 也可 | `DHUB-DB-STG`，**RDP + SSMS** | `DHUB-DB-PRO`，**RDP + SSMS** |
| GCS | `ods-seiyu-dev` | `ods-seiyu-stg` | `ods-seiyu-prod` |

IP（文档记载，以现网为准）：dev `10.22.214.244`，stg `10.22.214.164`，pro `10.22.212.164`。

2024 インフラ：主环境東日本 本番+Staging；西日本 DR **不用于** DataSpider 切转。当前单机，故障即停服。

资源申请：`mls_ISDDataPlatform@seiyu.com`（用户申请 CC Director）。

## 测试（Day3）

| 手段 | 怎么用 |
|---|---|
| 通常执行 | 绿三角。快，细节日志少 |
| 调试执行 | 虫子。输出执行历史/日志，可逐步看 |
| 断点 | 图标右键「ブレークポイントを設定/解除」 |
| 测试引数 | 脚本右键 → 属性 → テスト実行 |
| 错误日志 | 用消息代码去 Help 搜 |
| 分支覆盖 | 条件分支每条路径都要有期望结果 |

DEBUG 日志在 Designer 里可能被截断，去 My Log 看完整文件。生产日志 **INFO**。

## 发布（リリース手順書 v1.0，HR-EAI，2024-11-14）

用**资源管理员**登录目标环境 WebStudio。顺序：项目 → 非项目资材 → 系统表。STG/PROD 走 SDP 变更；监视提前约 1 周找 ITSM。

### 项目

1. **下载**：マイプロジェクト → 右键 → ローカルへダウンロード。勾选「バージョンを指定してダウンロード」，选**数字最大的最新版**。
2. **上传**：ファイル → ローカルからアップロード。
   - 上書きアップロード = 作为现有项目最新版本提交
   - 新项目名上传 = 另起项目
3. 打开项目 → **服务注册**（服务名 = 项目名）→ マイサービス确认。
4. zip **不带共享/ACL**。STG/PROD 上传后重新共享给 SY 组。

### 非项目（触发器、全局资源、环境变量等）

1. コントロールパネル → DataSpiderServerの設定 → サーバ移行。
2. **导出**：先「全て解除」，只勾本次要搬的资材，确认导出结果无错误。
3. **导入**：保留「プロジェクトの差分インポート」和「環境変数の差分更新」。**去掉会整表覆盖**。
4. 导入后重启 Windows 服务 **DataSpiderServista**。
5. 同名时：服务 / 触发器 / 全局资源等**覆盖**；环境变量和项目文件**可跳过**。全局用户逻辑同名不同内容会自动改成 `name_1`。

### 系统表（手順书 3）

改 `if_master` / `param_master` 前必须备份。

| 方式 | 步骤 |
|---|---|
| CSV 备份 | 新查询 → SELECT 对象表 → 结果右键「結果に名前を付けて保存」 |
| 脚本备份 | 任务 → スクリプトの生成 → 对象表 → 详细设定「スキーマとデータ」→ 存 sql |
| 3.2 追加 | 上位200行の編集 → 末行 `*` → 粘贴或手输 → SQL 再查确认 |
| 3.3 修正 | 上位200行の編集 → 查询设计器加 WHERE → 执行后改值 |
| 3.4 删除 | WHERE 筛出 → 选行删除 → SQL 确认已不存在 |

ガイドライン示例：同 `if_id` DELETE 再 INSERT。`to` 含支持台。
