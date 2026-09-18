# DataSpider Servista 新人使用手册

西友 DSS（DataSpider Servista）入门读本。按 2024 年三日「ツール活用トレーニング」录像整理，并接到现在实际开发方式。

**网页版（推荐转发，含操作截图）：** [DataSpider新人使用手册.html](DataSpider新人使用手册.html)

**日语版：** [DataSpider新人利用マニュアル.md](DataSpider新人利用マニュアル.md) / [DataSpider新人利用マニュアル.html](DataSpider新人利用マニュアル.html)

截图来自セゾン教材 Day1/Day3 PDF，放在 [images/](images/)。

| 项 | 内容 |
|---|---|
| 产品 | DataSpider Servista 4.4 (x64) AP3，SP `A00050` |
| 开发界面 | Studio for Web（西友以 Web 为主，不用 Desktop Studio） |
| 录像 | Day1（2024-11-22，约 1:58）操作与三个演示；Day2（2024-11-28，约 1:57）框架 + SQL Server；Day3（2024-12-05，约 1:45）测试与发布 |
| 讲师 | セゾン / 延吉 |
| 配套资料 | 《DataSpiderツール活用トレーニング》Day1/Day3 PDF、開発ガイドライン v3.2、《DSS開発の流れ》、リリース手順書 v1.0 |

**口令、主机密码不写在本手册里。** 需要登录时打开《DSS開発の流れ》或ガイドライン P.103 / P.115–118，不要抄进脚本、聊天或 Git。

---

## 0. 怎么用这份手册

建议顺序：

```
第 1 章 建立概念     → 对照 Day1 前半（登录、Designer、两条线）
第 2 章 动手三个演示 → 对照 Day1 后半（CSV / 0 件 / 文件触发器）
第 3 章 西友框架     → 对照 Day2（这是和教材最大的差别）
第 4 章 测试         → 对照 Day3 前半
第 5 章 发布         → 对照 Day3 后半 + リリース手順書
第 6 章 真实接口     → 台账取号到上线（日常工作按这一章，不要停在 DEMO）
第 6.5 节 命名规则   → ガイドライン P.12 全表，建项目前先对这一节
```

培训练习库是 `SAMPLE_PROJECT`，工作目录 `/data/DSSROOT/DEMO/`（Windows 上是 `F:\DSSROOT\DEMO\`）。**正式接口不要用 DEMO 路径**，用 `/seiyuroot/<IF_ID>/`。

卡住时：选中组件按 **F1**（帮助里有例外和限制）→ HULFT FAQ → [セゾン技术支持](https://www.hulft.com/tech-support)。产品名填 DataSpider Servista 4.4 (x64) AP3，SP `A00050`。序列号只在ガイドライン P.122，不要抄进 Git。

日常开发顺序以《DSS開発の流れ》（2026-05-12 / 07-22）为准，细则回ガイドライン对应页：

| 要查什么 | ガイドライン | 開発の流れ |
|---|---|---|
| 框架 / 戻り値 / 主数据 | P.22–31、P.50–56 | 「西友フレームワーク」 |
| `if_master` 列与登记 SQL | P.27、P.115–118 | 「if_master登録」 |
| 触发器 / 命名 | P.12、P.31、P.112–114 | 「トリガー作成」 |
| 项目 zip / 服务注册 / 共享 | P.105–111 | 「デプロイ」 |
| 用户 / 全局资源 / Azure / SFTP | P.32–35、P.58–63 | — |
| 三套 URL、账号、GCS 控制台 | P.105 | 第 2 页 |
| 技术支持 | P.122–123 | — |

---

## 1. 产品概念（Day1）

DataSpider 是 GUI 的 EAI：在 Studio 里画脚本，Server 上编译成 Java 再跑。四块能力：

| 块 | 做什么 |
|---|---|
| Adapter | 读/写 CSV、Excel、DB、FTP、REST、固定长等 |
| Mapper / Converter | 字段变换、过滤、日期、拼接（约 140 个逻辑图标） |
| Studio | 开发、运维、看日志 |
| Trigger | 到点、文件落地、JP1 等自动启动 |

西友数据联动里，DSS 是系统间 EAI/ETL（文件 / DB），不是分析仓。选型见ガイドライン P.18：

| 方案 | 主要用途 | 什么时候用 |
|---|---|---|
| DataHub | 文件联动、高 Tier、要冗余 | 基干、Tier 高的联动 |
| **DataSpider** | 系统间文件 / DB 联动 | 中低 Tier、文件或 DB、开发成本中等 |
| Databricks | 积蓄、分析、报表 | 分析数据；写仓走 ADLS（MS0007） |
| Lakehouse Federation | 只参照外部最新状态 | 实时库存/売上等参照。源库负荷要单独考虑 |

文件联动（SFTP / Blob）疏结合、适合大文件、实时性低。DB 联动（JDBC）快、有 ACID，但绑源库结构。高品质事务、要留档时，联动文件可放 DataSpider 用 ADLS（一 IF 一容器，外部用 SFTP 访问，约 30 天后自动删）。

数据质量（ガイドライン P.2）：事先约定文字编码、全角半角、电话号码连字符；要能早发现缺损、重复、异常值。乱码等于数据损坏，按接口约定处理，不要 silently 写进去。

### 1.1 三个名词不要混

| 词 | 含义 |
|---|---|
| **脚本** | Designer 里画的源码 |
| **项目** | 相关脚本的打包；下载、上传、服务注册都以项目为单位 |
| **服务** | 项目登记到 Server 之后，才能被触发器、其他脚本、ScriptRunner 调用 |

开发库和运行库是分开的：只有「作为服务注册」的项目才能挂触发器。**改过脚本必须重新注册服务。**

### 1.2 Studio 里常用窗口

- **Designer**：画脚本和 Mapper
- **Control Panel**：用户、任务、全局资源、挂载
- **Explorer**：DataSpider 逻辑文件系统（脚本里写的是这里的路径，不是 `C:\`）
- **My Service / My Trigger / My Log**：已注册的服务、触发器、运行日志

绿三角 = 通常执行（快，日志少）。虫子图标 = 调试执行（可逐步看、日志全）。

![Studio for Web 登录，在浏览器输入 URL](images/day1-login-url.jpg)

浏览器打开 WebStudio URL 后登录。西友现网地址见第 7 章，不要用教材示例主机。

![登录完成后的 DataSpider 主画面](images/day1-home.jpg)

登录完成：上方是 デザイナ / マイプロジェクト / マイトリガー / マイログ / マイサービス / エクスプローラ / コントロールパネル。

![主画面新建项目、打开项目](images/day1-new-project.jpg)

「新しいプロジェクト」新建，「プロジェクトを開く」打开已有项目。培训练习库是 `SAMPLE_PROJECT`；正式接口项目名 = 台账 IF ID。

![Designer：绿三角通常执行、虫子调试执行、右侧工具箱](images/day1-designer.jpg)

绿三角 = 通常执行，虫子 = 调试执行。右侧工具箱拖图标到画布。进程流必须手拉 Start → End。调试日志出现在画面下方。

### 1.3 两条线

| 线 | 样子 | 规则 |
|---|---|---|
| 数据流 | 橙色虚线 | 指定输入数据后自动出现，可以一对多 |
| 进程流 | 黑色实线 | **必须手拉** Start → … → End；一个图标只能有一条进程出线 |

没拉到 End 的脚本不能当完整处理。

### 1.4 版本、共享、挂载

- 每次保存都留版本，可还原，可出版本比较报告。
- 没有权限就看不到别人的项目。共享：マイプロジェクト → 右键 → プロパティ → 共有。
- 一般用户只能看到自己当「触发器所有者」的触发器。
- OS 目录要先挂到 DataSpider 文件系统（例如 `C:\Document\Sample` → `/data/mount`），脚本才不绑盘符。

![マイプロジェクト里设置项目共有](images/day1-share.jpg)

マイプロジェクト → 右键 → プロパティ → 共有。没有权限就看不到别人的项目。

![把 OS 目录挂到 DataSpider 文件系统](images/day1-mount.jpg)

OS 的 `C:\Document\Sample` 先挂成 `/data/mount`，脚本只写 DataSpider 路径。

![Explorer 管理 DataSpider 文件](images/day1-explorer.jpg)

エクスプローラ：管理逻辑文件系统上的目录和文件。

![マイログ查看执行结果](images/day1-mylog.jpg)

マイログ：按条件查脚本执行结果。DEBUG 被 Designer 截断时来这里看全文。

### 1.5 本环境已确认的适配器（Day1 導入有無）

| 状态 | 适配器 |
|---|---|
| 已购 ● | Salesforce、Microsoft Azure、SQL Server（选择式）、固定长、FTP、REST |
| 标准 〇 | JDBC、CSV、Excel、XML、HTML、文件操作、Mail、Thunderbus、HULFT |

表上没有 ●/〇 的（Oracle、GCP、Box、SAP 等）**不要默认当成可用**，先问现网。HTML 读取走 Web Adapter。可变长走可変長适配器。

CSV 读取：编码 **Windows-31J**；勾选把第一行当表头（不当数据值）。

---

## 2. Day1 三个演示（先会画脚本）

练习目标：会读 CSV、会用 Mapper、会写文件、会注册服务、会挂文件触发器。

### 演示① CSV 转换集计

- 输入：`OrderList.csv`（受注日、支店、ケース数、入り数、容量、商品名、受注_ID）
- 行为：按支店汇总发注数，写出 `OrderList_total.csv`
- 要点：字段变换放 **Mapper**，不要用 IF / 循环逐行加工

### 演示② 0 件当错误

- 输入：只有表头、没有数据行的 `OrderList.csv`
- 行为：不写输出文件，主动结束成错误（后面接到框架时，这就是警告或 255）

### 演示③ 文件触发器

- 把 `OrderList.csv` 放到输入文件夹 → 自动跑①的脚本 → 结果进输出文件夹
- 要点：触发器挂的是**已注册的服务**，不是 Designer 里未登记的脚本

服务注册：项目右键「プロジェクトをサービスとして登録」。画面会比较源项目和目标服务的脚本，确认后再登记。

![项目右键注册为服务](images/day3-service-register.jpg)

改过脚本后必须重新做这一步，否则触发器还在跑旧版。

![マイトリガー列表](images/day1-trigger.jpg)

教材示例是文件触发器；西友正式接口改为调 `SEIYUCOM000X` / `01frameworkBatch`。

![マイサービス确认已注册服务](images/day1-myservice.jpg)

マイサービス：确认服务名 = 项目名，状态可用。

---

## 3. 西友框架（Day2，正式开发必须按这个）

教材可以从 Designer 直接跑脚本。西友正式接口**不是**这样：触发器永远调共通框架，框架再调你的个别项目。

### 3.1 谁干什么

```
调度触发器 / JP1 / ScriptRunner
        ↓
SEIYUCOM000X / 01frameworkBatch     ← 只用，不改
        ↓
SEIYUCOM001X
  01getInterfaceInfo   读 if_master、param_master
  调个别项目（你的 IF）
  02outputLog / 03sendMail
        ↓
个别项目 = 接口 ID，只写业务
```

| 项目 | 角色 | 能改吗 |
|---|---|---|
| `SEIYUCOM000X` | 入口 `01frameworkBatch` | 只用，不改 |
| `SEIYUCOM001X` | 读主数据、调个别脚本、写日志、发邮件 | **禁止修改** |
| `SEIYUCOM002X` | ADLS 路径切割 | 可用 |
| `SEIYUCOM003X` | 备份清理（每天 0 点） | 了解即可 |
| HOUSEKEEP | 本地备份清理 | **非推荐**；备份放 Blob `AZURE_STR_DSS` |
| 接口 ID 项目 | 个别处理 | 这里写业务 |

个别脚本**不要自己发运维邮件、写框架日志**。邮件由 `03sendMail` 发。发信以台账为准，常见 `no-reply-HR-EAI@seiyu.com` 或 `no-reply-Seiyu-EAI@seiyu.com`。

JP1 / ScriptRunner：用共通模板 `C:\Temp\scriptrunner.xml`，第一个参数传 `if_id`。日志进 My Log，和触发器启动一样。

### 3.2 四条收口（Day2 的核心）

框架先读 `if_master`、写开始日志、调你的脚本，再按结束状态分四路。

**戻り値设在 End 组件的属性检查器里**，不是在脚本属性里随便写一个数。

| 分支 | 你的脚本怎么结束 | 框架做什么 |
|---|---|---|
| 1 正常 | End 戻り値 **0** | 记成功日志后结束 |
| 2 警告 | 戻り値 **1–254**（演示：CSV 只有表头 / 0 件） | 警告邮件 + 警告日志 |
| 3 可控异常 | 戻り値 **255**，并且填了 `out_*` | 异常邮件 + 异常日志 |
| 4 未捕获 | Exception / 没 catch | 框架 catch 后发异常邮件。戻り値实际是 **-1** |

警告、异常用「警告变量设定」「异常变量设定」，把 catch 内容塞进：

- `out_component_name` / `out_component_type`
- `out_message_category` / `out_message_code` / `out_message_level`
- `out_error_type` / `out_error_message`

邮件后半段就是这些字段，例如：

```
FILE0001E|java.io.FileNotFoundException|/data/DSSROOT/DEMO/INPUT/OrderList.csv|CSV読取
```

Day2 改写口径：

- 正常 = 跑完，戻り値 0
- 警告 = 0 件，戻り値 1–254
- 异常 = 读或写失败，戻り値 255 + `out_*`

### 3.3 if_master / param_master

库：`DSSSEIYU`。改表前必须按リリース手順书 3.1 **备份**。发版常用同一 `if_id` 先 DELETE 再 INSERT（ガイドライン P.117）。

列名用表上的全名，不要写成 `from` / `to`：

| 列 | 含义 |
|---|---|
| `if_id` | 接口 ID = 项目名 |
| `if_name` | 显示名（不是脚本文件名） |
| `service_name` | 个别项目登记后的服务名（通常 = 项目名） |
| `script_name` | 个别脚本名，如 `01OutputData` |
| `from_mail_address` | 结果通知 From |
| `to_mail_address` | 结果通知 To。**必须含** `isd_support_desk@seiyu.com` |
| `cc_mail_address` / `bcc_mail_address` | 可空 |
| `mail_title_header` | 邮件标题头，常用接口 ID |
| `tier` | 重要度 |

```sql
USE [DSSSEIYU]
GO
IF EXISTS (SELECT 1 FROM [dbo].[if_master] WHERE [if_id] = N'<IF_ID>')
BEGIN
  DELETE FROM [dbo].[if_master] WHERE [if_id] = N'<IF_ID>';
END;
GO
INSERT [dbo].[if_master] (
  [if_id], [if_name], [service_name], [script_name],
  [from_mail_address], [to_mail_address], [cc_mail_address], [bcc_mail_address],
  [mail_title_header], [tier]
) VALUES (
  N'<IF_ID>', N'<显示名>', N'<IF_ID>', N'01xxxx',
  N'no-reply-Seiyu-EAI@seiyu.com',
  N'isd_support_desk@seiyu.com;<业务收件人>',
  NULL, NULL, N'<IF_ID>', NULL
);
```

DEV：`DHUB-DB-DEV` 端口 1433，库 `DSSSEIYU`，DBeaver 也可，账号见ガイドライン P.103（**不要把密码写入手册或脚本**）。STG / PROD：必须 RDP + SSMS；Server 名为 `SeiyuServerStg.seiyu.com` / `SeiyuServerPro.seiyu.com`，账号向 Data 团队确认。

`param_master`：`IF_ID` + `param_no` + `param_value`。`01getInterfaceInfo` 读入后放到 `param1`–`param20`。环境差异（路径、文件名）尽量用这些参数，不要写死在图标属性里。

### 3.4 触发器怎么挂（三套环境相同）

| 字段 | 值 |
|---|---|
| 服务 | `SEIYUCOM000X` |
| 脚本 | `01frameworkBatch` |
| 变量 `if_id` | 本接口 ID |
| 执行用户 | `exec_SY` |
| 状态 | 登録后应为「待機中」 |

**不要**把触发器直接绑到个别脚本上。画面操作见ガイドライン P.112–114：マイトリガー → 选种类（常用スケジュール / ファイル）→ 服务 `SEIYUCOM000X`、脚本 `01frameworkBatch`、变量 `if_id` → 启用后一览应为「待機中」。执行用户一律 `exec_SY`（口令不写在本页）。

### 3.5 SQL Server（Day2 后半）

适配器能做：表读/写、検索 SQL、更新 SQL、ストアド、ローダ。连接例 `COM_SQLServer2022`。更新 SQL 用 `?{列名}` 占位。

练习库：`DHUB-DB-DEV` / `DSSROOT`（`if_master`、`param_master`、`DEMO_OrderList`）。

**DBデモ1**（CSV 进表，带事务）：

```
Transaction → DELETE → csv_read → mapping → テーブル書き込み → COMMIT
```

`process_date` 用「現在日時」。组件「データ処理方式」可选「スクリプトの設定を使用する」。大件数不要全进内存，打开「大容量データ処理」。

默认每个图标各自提交。跨多步要一起成功或一起回滚时，用事务图标或脚本的事务属性。

---

## 4. 测试（Day3）

每条条件分支都要有期望结果（录像里的 ①②③），不要只跑通主路径。

| 手段 | 怎么用 |
|---|---|
| 通常执行 | 绿三角。快，细节日志少 |
| 调试执行 | 虫子。有执行历史和日志，可逐步看 |
| 断点 | 图标右键「ブレークポイントを設定/解除」，停住后再按调试继续 |
| 测试引数 | 脚本右键 → 属性 → テスト実行，给输入变量喂测试值 |
| 调试信息 | 看运行中的脚本变量 |
| 错误日志 | 用消息代码去 Help 搜原因和对策 |

DEBUG 日志在 Designer 里可能被截断，去 **My Log** 看完整文件。生产日志用 **INFO**，不要开 DEBUG / FINEST。

![调试执行：虫子图标和画面下方执行日志](images/day3-debug.jpg)

点虫子做调试执行。下方「実行履歴 / 実行ログ」会出逐步结果。这张里 CSV 读取失败，消息码 FILE0001E。

![图标右键设置断点](images/day3-breakpoint.jpg)

图标右键 → ブレークポイントを設定/解除。调试会停在该组件，再按虫子继续。

![脚本属性里设置测试执行引数](images/day3-test-args.jpg)

脚本右键 → プロパティ → テスト実行。给输入变量喂测试值，不必改图标属性。

---

## 5. 发布（Day3 + リリース手順書 v1.0）

用**资源管理员**登录**目标环境** WebStudio。顺序：项目 → 非项目资材 → 系统表。

STG / PROD 走 SDP 变更；监视提前约 1 周找 ITSM。

### 5.1 项目（zip 下载 / 上传）

1. マイプロジェクト → 右键 → ローカルへダウンロード。勾选「バージョンを指定してダウンロード」，选**数字最大的最新版**。
2. 目标环境：ファイル → ローカルからアップロード。
   - 上書きアップロード = 作为现有项目的最新版本提交
   - 换新项目名上传 = 另起一个项目
3. 打开项目 → **服务注册**（服务名 = 项目名）→ 到マイサービス确认。
4. zip **不带共享权限**。STG / PROD 上传后：マイプロジェクト → 右键 → 共有 → 给 **SY 组读取**，一览里要能看到「共有」。
5. 若脚本写 Explorer 文件夹：给执行用户（`exec_SY` 等）加上写入组权限（ガイドライン P.110）。

![マイプロジェクト右键下载并指定最新版本](images/day3-download.jpg)

右键 → ローカルへダウンロード。选「バージョンを指定してダウンロード」，拿数字最大的最新版（图中 45）。

![目标环境文件菜单本地上传](images/day3-upload-menu.jpg)

目标环境：ファイル → ローカルからアップロード，选刚下载的 zip。

![上書きコミット或作为新项目上传](images/day3-upload-options.jpg)

同名上写选「既存プロジェクトにコミット」；换名选「新しいプロジェクトとしてアップロード」。

![上传后重新服务注册](images/day3-service-register.jpg)

打开上传后的项目 → プロジェクトをサービスとして登録。服务名做成与项目同名。

### 5.2 非项目（触发器、全局资源、环境变量等）

1. コントロールパネル → DataSpiderServerの設定 → サーバ移行。
2. 导出：先点「全て解除」，只勾本次要搬的，确认导出结果没有错误。
3. 导入：**不要去掉**「プロジェクトの差分インポート」和「環境変数の差分更新」。去掉会整表覆盖，删掉目标环境已有内容。
4. 导入后重启 Windows 服务 **DataSpiderServista**。
5. 同名时：服务 / 触发器 / 全局资源等会**覆盖**；环境变量和项目文件**可以跳过**。全局用户逻辑如果同名但内容不同，会自动改成 `name_1`。

![サーバ移行导出先全部解除再勾选](images/day3-export-select.jpg)

导出：先点「全て解除」，再只勾本次要搬的资材。不要整包全选。

![导入时保留差分选项](images/day3-import-keep.jpg)

导入时红框两项**不要去掉**：「プロジェクトの差分インポート」「環境変数の差分更新」。去掉会整表覆盖。

### 5.3 系统表（if_master 等）

只搬项目不够，框架表要单独改 SQL Server。改之前必须备份。

| 方式 | 步骤 |
|---|---|
| CSV 备份 | 新查询 → SELECT 对象表 → 结果右键「結果に名前を付けて保存」 |
| 脚本备份 | 任务 → スクリプトの生成 → 对象表 → 详细设定选「スキーマとデータ」→ 存 sql |
| 追加 | 上位200行の編集 → 末行 `*` → 粘贴或手输 → 再用 SQL 查确认 |
| 修正 / 删除 | 先在查询设计器加 WHERE 筛出行，再改或删，最后用 SQL 确认 |

---

## 6. 真实接口怎么做（日常按这一章）

培训用 DEMO；上线用台账上的 IF ID。一个 IF = 一个处理单位（同一 IF 可以传多个文件）。只做**单体**。ガイドライン把接口分成 単体 / バラマキ / 数珠繋ぎ / 複数，**西友只许単体**。**数珠つなぎ禁止。** 不要从 Designer 直接调生产脚本做联调。

申请与上线节奏（ガイドライン P.4–6）：AppG 提接口申请 → DataG 受付、预审、发 IF ID → 开发测试 → リリース时走 SDP；监视大约提前 1 周找 ITSM。台账要记初回登记日、PRO 登记日、ITSM 运用开始日、稼働开始日。监视需要：启动时刻、L3（Owner 团队与联系人）、重要度/影响。变更或停止也要事先联系 DataG。

Owner：Requester 委托后，**联动源侧 Owner** 负责接口汇总和脚本开发，数据内容也由 Owner 负责。若一组接口有关联，由**起点 IF 的 Owner** 一并统筹。

申请种类：インターフェース申请、IntakeForm（Databricks 联动）、DSS 设定申请、全局资源申请、账号申请、其他个别咨询。共通/共享资源由 Data 团队维持；要专用权限组或 Access Token 时先和 Data 团队商量。

设计书：シーケンス図 + 详细マッピング定义书（ガイドライン P.10–11）。开发前先有这两份，不要只在 Designer 里「画着看」。

### 6.1 取号

从 `SEIYU-TRIALインターフェース管理台帳.xlsx` 的 `DSSインターフェース台帳` 拿 IF ID（西友侧管理，《DSS開発の流れ》第一步）。

格式：`FROM6_TO6_SEQ3`（大写 + 下划线），例如 `AM0002_MS0007_001`。系统 6 位用系统一览里的系统 ID。项目名 = 接口 ID。

### 6.2 DEV → STG → PROD（開発の流れ 8 步）

同一套步骤在三套环境各做一遍。复制现有个别项目来改，不要从空白重写框架调用。

| 步 | 做什么 | 硬性规则 |
|---|---|---|
| 1 | IF ID | 台账已取号 |
| 2 | 建项目 | 项目名 = 接口 ID |
| 3 | 脚本 | 复制现有脚本再改。`01` + 驼峰英文或罗马字，例如 `01updDelDate`。只写个别处理 |
| 4 | Explorer | 建工作目录 `/seiyuroot/<IF_ID>/`，给组加访问权限 |
| 5 | 服务注册 | 服务名 = 项目名。改脚本后必须重新注册 |
| 6 | 触发器 | 只调 `SEIYUCOM000X` / `01frameworkBatch`，参数 `if_id`，用户 `exec_SY` |
| 7 | `if_master` | `to_mail_address` 含支持台；改表前备份。页：P.115–118 |
| 8 | 测试后升环境 | STG/PROD 走 SDP。部署画面：P.105–111 |

资源邮件统一发 `mls_ISDDataPlatform@seiyu.com`。用户申请 **CC 所属 Director**。DEV 全局资源可以自建，但参照权限必须设对；STG/PROD 由 Data 团队按申请书建箱。SFTP **只用密钥**。

用户申请件名：`【Dataspider】ユーザー作成依頼`。正文写：申请区分（新規/変更）、环境（开发/验证/本番）、氏名、邮件、职工番号。

全局资源申请件名：`【Dataspider】グローバルリソース作成依頼`。正文写：DSS 用户 ID、新規/変更、环境、要哪种连接（SFTP / Azure 等）以及理由。管理员会再问细节。

Azure 存储：连 Databricks 走 **SFTP**；连其他系统走 **Azure ストレージコネクタ**。申请件名 `【Dataspider】Azureストレージ接続設定依頼`。

SFTP 密钥由对接系统发给担当，再转给 DSS 管理员放到服务器。路径：`/data/SSH/<Intake号>/xxxxx.ppk`，文件夹按发行方唯一（Intake / ServerName+UserName）。全局资源里只填这个路径，不要把密钥提交到 Git。

### 6.3 个别脚本默认骨架

正式样板是 2026-05-13 实演的 `CM0010_IF0079_006` / `01OutputData`（抽 Oracle → 本地 CSV/END → GCS），不要另发明骨架：

```
try/catch
  → delete 工作区 /*                 # /seiyuroot/<IF_ID>/
  → 源抽出（検索系 SQL / CSV / SFTP …）
  → Mapper（日期、字段、过滤）
  → 写出业务文件 + 完了文件（如 ORDER_${file_date}.END）
  → 交件（GCS / ADLS / SFTP）
  → End 戻り値 0
catch → 异常变量设定 → End 255
```

GCS 连接名 `GCP_STR_DSS`，桶 `ods-seiyu-{dev|stg|prod}`，控制台：

- https://console.cloud.google.com/storage/browser/ods-seiyu-dev
- https://console.cloud.google.com/storage/browser/ods-seiyu-stg
- https://console.cloud.google.com/storage/browser/ods-seiyu-prod

路径常见 `/<IF_ID>/${folder_date}/` 或台账规定的 `deliveryData/YYYYMMDD`。没有文件夹就建，已有则覆盖。

ADLS（ガイドライン P.13）：一 IF 一容器，容器名 = 接口 ID 小写连字符；外部系统用 SFTP 访问该容器。文件约 30 天后自动删除。高 Tier 事务联动建议走这条。

备份原则进 Blob，全局资源名 `AZURE_STR_DSS`。`/data/backup` 只是备选。容器：DEV 开发者可建；STG/PROD 找平台团队。用 Azure 适配器前先点「コンテナ名一覧の更新」，否则列表是空的。

### 6.4 读 / 转 / 写怎么选

| 场景 | 用什么 | 注意 |
|---|---|---|
| DB 全表抽出 | テーブル読み取り | 全局资源可复用 |
| 带条件抽出 / 复杂 SQL | 検索系 SQL 実行 | 向导生成后如果手改 SQL，向导就不能再编 |
| 字段变换、过滤、日期、拼接 | Mapper | 先定输出 schema；CSV 可用表头文件 |
| 文件 JOIN | Merge（APPEND / INTEGRATE / Left-Right JOIN） | DB 侧优先用 SQL JOIN |
| 大量写入 / 主数据洗い替え | テーブル書き込み + 批处理 | 勾选「キー一致行は更新」= upsert。86 万件时批模式大约比普通 upsert 快 30 倍 |
| 复杂更新 / 外结合 | 更新系 SQL 実行 | 不适配超大件数 |
| 固定长 EDI | 固定长适配器 + 向导 | Group / Record / Field；定义进全局资源 |

多脚本共用同一结构时，用**グローバルスキーマ**（テーブルモデル）。给 PSP Mapper 用时不要改成 XML 型；非 PSP 组件即使挂了全局 schema 也不会并行。

リトライ：对连接错误、超时要写清次数和间隔（ガイドライン P.14）。通知按重要度分（クリティカル / 警告），由框架邮件承担，个别脚本不要另搭一套。通信走 SSL/TLS。密码和 API Key 放环境变量或外部设定，不要写进图标。

### 6.5 命名规则（ガイドライン P.12，按这一表抄）

标识符只用**英语或ヘボン式罗马字**，不要用中文 / 日文当名字。风格：アッパー = 全大写；スネーク = 下划线；キャメル = 小写开头驼峰；アッパーキャメル = 大写开头驼峰。

| 对象 | 风格 | 规则 | 例子 |
|---|---|---|---|
| 接口 ID | アッパースネーク | `FROM系统(6位)_TO系统(6位)_序号(3位)`。系统 6 位用系统一览的系统 ID | `AM0002_MS0007_001` |
| 项目名 | 同接口 ID | 必须与接口 ID 完全相同 | `AM0002_MS0007_001` |
| 脚本名 | 两位序号 + キャメル | ① 同一项目内顺序编号 ② 处理概要：英语或ヘボン式 | `01updDelDate` |
| 常量 | アッパースネーク | 英语或ヘボン式 | `FILE_NAME` |
| 变量 | スネーク（小写） | 英语或ヘボン式 | `file_name` |
| 参数 | スネーク + 型前缀 | 输入 `in_`、输出 `_out`、入出力 `io_`。其余用英语 | `in_file_name` |
| 环境变量 | `ENV_` + アッパースネーク | 英语 | `ENV_RETRY_COUNT` |
| 工作文件名 | スネーク（小写） | 脚本里普通文件用英语小写下划线 | `file_name` |
| 表名 | 前缀 + アッパースネーク | 主数据 `M_`、事务 `T_` | `M_FILE_SEND` |
| ADLS 容器 / 文件夹 | 小写 + 连字符 | ADLS 限制：只能小写、数字、`-`。把接口 ID 改成这种 | `am0002-ms0007-001` |
| 固定长 / 可变长格式名 | アッパーキャメル | `FIX` 或 `VAR` + `READ` 或 `WRITE` + 格式名称（英语或ヘボン式） | `FIX_READ_Time` |
| 全局资源（固定长定义） | 同上 + 系统 | 格式名 + `_` + 系统 ID + 名称 + `_01` | `FIX_READ_MS0007_Test_01` |
| 全局资源（其余） | 协议_系统6位[_用户或 Intake] | 见下表。Datalake 时加 Intake 号。需要时加枝番（例如 DB 名） | `SFTP_AM0002_用户ID` |
| 触发器 | 接口 ID + 种类 | 多种スケジュール时：接口 ID + `SCHL_` + 3 位序号 | `AM0002_MS0007_001_SCHL` |

**全局资源协议前缀**

| 前缀 | 用途 | 例子 |
|---|---|---|
| `SQL` | SQL Server | `SQL_DSS_SEIYU_DB`（DSS 自用库） |
| `ORCL` | Oracle | `ORCL_CM0010_wknpsuser` |
| `JDBC` | JDBC（含读 Databricks / MS0005） | `JDBC_AM0002_用户ID` |
| `ODBC` | ODBC | `ODBC_AM0002_用户ID` |
| `FTP` / `SFTP` | 文件传输。SFTP **只用密钥** | `SFTP_AM0002_用户ID`、`SFTP_MS0007_<Intake>` |
| `REST` | REST | `REST_AM0002_用户ID` |
| `MAIL` | 邮件 | `MAIL_AM0002_用户ID` |
| `BLOB` | Azure Blob | `AZURE_STR_DSS` |

同一源系统仍要拆「全局资源 + 用户 ID」，不要所有 IF 共用一个连接。枝番可加在末尾（例：DB 名）。写 Databricks 走 ADLS（MS0007）；读 Databricks 走 JDBC（MS0005），**不要开连接池**。

**触发器种类**

| 识别子 | 种类 | 西友正式接口 |
|---|---|---|
| `SCHL` | スケジュール | 常用。挂 `SEIYUCOM000X` / `01frameworkBatch` |
| `FILE` | ファイル | 培训演示用；正式接口仍挂框架，不挂个别脚本 |
| `HULF` | HULFT | 按台账 |
| `HTTP` | HTTP | 按台账 |
| `WSDL` | WEB 服务 | 按台账 |

路径也按名来：工作区 `/seiyuroot/<IF_ID>/`；SFTP 密钥 `/data/SSH/<Intake>/xxxxx.ppk`；GCS 桶 `ods-seiyu-{dev\|stg\|prod}`，目录常见 `/<IF_ID>/${folder_date}/`。

---

## 7. 三套环境（2026）

日常用独立三套 WebStudio，不要沿用 2024 设计书「Staging 里放开发机」。

| | DEV | STG | PROD |
|---|---|---|---|
| WebStudio | `https://dhub-dss-dev.seiyu.com:8443/WebStudio/` | `https://dhub-dss-stg.seiyu.com:8443/WebStudio/` | `https://dhub-dss-pro.seiyu.com:8443/WebStudio/` |
| 账号 | 个人账号 | `exec_SY` | `exec_SY` |
| 服务器 IP | `10.22.214.244` | `10.22.214.164` | `10.22.212.164` |
| DB | `DHUB-DB-DEV` / `DSSSEIYU`，DBeaver 也可 | `DHUB-DB-STG`，必须 RDP + SSMS | `DHUB-DB-PRO`，必须 RDP + SSMS |
| GCS 桶 | `ods-seiyu-dev` | `ods-seiyu-stg` | `ods-seiyu-prod` |

当前是单机。故障即停服。不定期 14:30–16:00 维护。高 Tier、维护窗口不能停的处理不要放 DSS。维护预告在 Teams「データ利活用ポータル | Dataspider事務局からのお知らせ」。西日本 DR **不用于** DataSpider 切转。

西友以 **Studio for Web** 为主。Desktop Studio 能建测试项目，但三套 Web 环境已经分开，日常用不上；本机 Studio 响应差，Modules 初次同步可能超过 4 小时，不要当开发界面。

---

## 8. 禁则（上线前核对）

会拖垮整台 DSS：

- 用 IF / 循环做行级数据加工（放进 Mapper）
- 单脚本图标超过约 100 个
- 大件数不打开「大容量データ処理」（可能 Java OOM，影响全实例）
- PSP 与大容量同时开时 **PSP 优先**，大容量不跑（读完进 Mapper 再写、且写侧不支持 PSP 时，Mapper 与写入之间仍走大容量）

PSP（スマートコンパイラ默认常开）：按约 1000 件一块、读/转/写多线程。块大小改不了。Designer「表示」里打开「PSPデータフローの表示/非表示」才能看见。大容量：脚本右键 → プロパティ → データ処理方式 → 勾选「大容量データ処理を行う」，最低限度留内存、其余落盘，会比全内存慢。

明确不要做：

- 改 `SEIYUCOM000X` / `SEIYUCOM001X`
- 脚本里硬编码密码、API Key
- 数珠つなぎ / バラマキ / 把多条 IF 串成一条；从 Designer 调生产脚本
- Databricks 全局资源开连接池
- 用 SQL Express 当数据库文件系统（10GB 上限）
- SFTP 密码认证（只用密钥）
- 生产日志开 DEBUG / FINEST（Designer ツール → オプション；升 DEBUG 会涨 CPU 和日志体积）
- 把培训里失败的挂载 `/data/Shared_test` 写进新脚本

性能口径：PSP 大约 1000 件一块并行；テーブル書き込み批模式在 86 万件测试里大约快 30 倍。1 亿件级不要用 DSS 扛（验证时 `server\tmp` 可涨过 10GB，CPU 打满；Postgres 侧应考虑 `psql` / `COPY`）。

卡住：マイログ看是否还在执行 → コントロールパネル **タスクマネージャ** → 选脚本进程 → プロセスの終了。仍不消失就重启服务 `DataSpiderServista`。CLI 的 session 切断见ガイドライン P.89–90，新人先走画面。

---

## 9. 新人练习清单

对照录像自己勾：

**Day1**

- [ ] 能登录 DEV WebStudio，找到 Designer / Explorer / My Log
- [ ] 分得清脚本、项目、服务
- [ ] 分得清橙色数据流和黑色进程流，能手拉到 End
- [ ] CSV 用 Windows-31J、跳过表头，Mapper 做出按支店汇总
- [ ] 项目注册为服务后再挂文件触发器

**Day2**

- [ ] 说得出触发器为什么必须调 `SEIYUCOM000X`
- [ ] End 上会设 0 / 1–254 / 255，catch 会填 `out_*`
- [ ] 能在 `if_master` 用 `if_id` / `service_name` / `script_name` / `to_mail_address` 对上，并知道 TO 必须含支持台
- [ ] 看过带 Transaction 的 CSV→表写入

**Day3 + 真实接口**

- [ ] 用调试 + 断点把正常 / 0 件 / 读失败三条都跑到期望结果
- [ ] 能按手順书下载指定最新版本、上传后重新服务注册
- [ ] 知道 zip 不带共享、STG/PROD 要给 SY 组读权限、导入不能去掉差分选项、改 `if_master` 前要备份
- [ ] 能按第 6.5 节从 IF ID 写出项目名、脚本名、触发器名、ADLS 容器名和全局资源名前缀
- [ ] 能按第 6 章列出骨架，说得出只做单体、Owner 在联动源侧

---

## 10. 附录：样板 IF（看过再画自己的）

2026-05-13 田村实演：`CM0010_IF0079_006` / `01OutputData`。

- FROM：CM0010 若菜 新店舗；TO：IF0079 TRIAL 存储
- Oracle：`ORCL_CM0010_wknpsuser`，schema `WKNPSUSER`，例 `WKNST_DAILYORDER`
- 工作区：`/seiyuroot/CM0010_IF0079_006/`
- 顺序：try/catch → delete `/*` → 検索 SQL → Mapper 出 `file_date`/`folder_date` → csv_write → `ORDER_${file_date}.END` → `gcs_put` 到 `ods-seiyu-dev/<IF_ID>/${folder_date}/`

详细字段见同目录 skill 的 `examples.md`。可视化速查：`.cursor/canvases/DSS-开发速查.canvas.tsx`。

---

## 11. 资料来源

- DataSpider ワークショップ Day1/2/3 会议录像（2024-11/12）
- DataSpiderツール活用トレーニング Day1、Day3 PDF
- 開発ガイドライン v3.2（2024-12-09，DataエンジニアリングG；本版按全文 123 页补了申请、命名、主数据列名、发布与 PSP）
- 《DSS開発の流れ》（2026-05-12 / 2026-07-22，4 页；日常 DEV→STG→PROD 顺序与 URL 以这份为准）
- リリース手順書 v1.0（HR-EAI，2024-11-14）
- 2026-05-13 Teams 实演（CM0010_IF0079_006）

2026-07-22 Teams「DSSについて」约 3GB、无字幕，未能逐句转写；环境 URL 以《DSS開発の流れ》为准。
