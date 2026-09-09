---
name: dss_development
description: >-
  西友 DataSpider Servista（DSS）开发入口。写/改接口脚本、取 IF ID、挂触发器、登记 if_master、DEV→STG→PROD 发布前必须先读本 skill。
  覆盖：SEIYUCOM000X 框架、返回码 0/1–254/255/-1、命名、单体接口、GCS/ADLS/SFTP、リリース手順書。
  关键词：DataSpider、DSS、Servista、WebStudio、接口、IF ID、SEIYUCOM、if_master、param_master、01frameworkBatch、exec_SY、Mapper、触发器、リリース。
---

# dss_development — 西友 DataSpider 开发入口

写或改任何 DSS 接口前按本 skill 执行。**禁止**按通用 ETL / 空项目从零搭调度；**禁止**把密码、密钥抄进脚本、仓库或本 skill 的回复。

凭证与主机口令只在《DSS開発の流れ》、ガイドライン P.103 / P.115–118。需要登录信息时让用户自己打开那两份，不要代抄。

```
Step 1: 确认 IF ID + 单体边界（数珠つなぎ禁止）
Step 2: 复制现有个别项目，只写业务；触发器永远调 SEIYUCOM000X
Step 3: End 戻り値 + if_master；升环境按リリース手順書
```

细则按需再读（一次只读相关文件）：

- 框架 / 返回码 / 系统表 → [framework.md](framework.md)
- 命名 / 三套环境 / 发布步骤 → [naming-release.md](naming-release.md)
- 样板脚本 CM0010_IF0079_006 与培训演示 → [examples.md](examples.md)
- 可视化速查（Cursor Canvas）→ [DSS-开发速查.canvas.tsx](../../canvases/DSS-开发速查.canvas.tsx)

---

## 何时用 / 何时不用

**必须用**：新建或改西友 DSS 接口；画 Designer 脚本；配触发器/服务；改 `DSSSEIYU.if_master` / `param_master`；DEV→STG→PROD 迁移；问 DataSpider 适配器、Mapper、PSP、大容量処理。

**不要用本 skill 替代 SMART 批处理**：SMART `.sh` / CGI / DATAMASTER 走 `smartai_guide`。DSS 是 GUI EAI，产出是 Studio 脚本设计，不是 bash。

产品：DataSpider Servista **4.4 (x64) AP3**，SP `A00050`，Studio for Web。脚本在 Server 上编译成 Java。

---

## Step 1：取号与形态

1. 从 `SEIYU-TRIALインターフェース管理台帳` 的 `DSSインターフェース台帳` 拿 **IF ID**。
2. 格式：`FROM6_TO6_SEQ3`（大写 + 下划线），例 `AM0002_MS0007_001`。
3. **一个 IF = 一个处理单位**（同一 IF 可传多个文件）。只做**单体**。**数珠つなぎ（链式互调）禁止**。
4. 项目名 = IF ID。不要从 Designer 直接调生产脚本做联调。

---

## Step 2：DEV 落地（8 步）

复制现有个别脚本改，不要空白重写框架调用。

| 步 | 做什么 | 硬性规则 |
|---|---|---|
| 1 | IF ID | 台账已取号 |
| 2 | 建项目 | 项目名 = IF ID |
| 3 | 脚本 | `01` + 驼峰英文/罗马字，例 `01updDelDate`。只写个别处理 |
| 4 | 资源 | Explorer 目录、组权限、全局资源。STG/PROD 全局资源由 Data 团队建；SFTP **只用密钥** |
| 5 | 服务注册 | 服务名 = 项目名。改脚本后**必须重新注册** |
| 6 | 触发器 | 只调框架，不调个别脚本（见下表） |
| 7 | `if_master` | TO 必须含系统支持台；改表前按リリース手順书 3.1 **备份** |
| 8 | 测试后升环境 | STG/PROD 走 SDP；监视提前约 1 周找 ITSM |

触发器（三环境相同）：

| 字段 | 值 |
|---|---|
| 服务 | `SEIYUCOM000X` |
| 脚本 | `01frameworkBatch` |
| 变量 `if_id` | 本接口 ID |
| 执行用户 | `exec_SY` |
| 状态 | 登録后应为「待機中」 |

资源申请邮件：`mls_ISDDataPlatform@seiyu.com`（用户申请 CC 所属 Director）。开发全局资源可自建；STG/PROD 走申请。

个别脚本**不要**自己发运维邮件、写框架日志。邮件由 `SEIYUCOM001X` 的 `03sendMail` 发（`no-reply-HR-EAI@seiyu.com`）。

---

## Step 3：收口与发布

个别处理必须：

1. 进程流手拉 **Start → … → End**（一个图标只能有一条进程出线）。
2. 用 try/catch；警告/异常把 catch 内容写入 `out_*`（见 [framework.md](framework.md)）。
3. 在 **End 组件**设戻り値：`0` 正常 / `1–254` 警告 / `255` 可控异常。未捕获 Exception 实际为 `-1`。
4. `if_master` 有对应行（`service_name` / `script_name` 对得上）。

升环境：用资源管理员登录目标 WebStudio。项目 zip **不带共享/ACL**，STG/PROD 上传后重新共享给 SY。系统表先备份再 DELETE+INSERT。细节见 [naming-release.md](naming-release.md)。

---

## 个别脚本默认形态

优先抄 `CM0010_IF0079_006` / `01OutputData`（抽源 → 本地工作区 → 交件），不要发明新骨架：

```
try/catch
  → delete 工作区 /*          # /seiyuroot/<IF_ID>/
  → 源抽出（検索系 SQL / CSV / SFTP …）
  → Mapper（日期、字段、过滤；禁止用 IF/循环逐行加工）
  → 写出业务文件 + 完了文件（如 ORDER_${file_date}.END）
  → 交件（GCS / ADLS / SFTP）
  → End 戻り値 0
catch → 异常变量设定 → End 255
```

工作目录：**`/seiyuroot/<IF_ID>/`**（培训 DEMO 曾用 `/data/DSSROOT/DEMO/`，新 IF 不要用 DEMO 路径）。

GCS 样板：连接 `GCP_STR_DSS`，桶 `ods-seiyu-{dev|stg|prod}`，路径 `/<IF_ID>/${folder_date}/`（或台账规定的 `deliveryData/YYYYMMDD`）。无文件夹则建，已有则覆盖。

CSV：编码 **Windows-31J**；跳过首行当表头。

两条线不要混：

- **数据流**：橙色虚线，指定输入后自动出现，可一对多。
- **进程流**：黑色实线，必须手拉，一个图标一条出线。

---

## 禁则（写设计时逐条核对）

- 改 `SEIYUCOM000X` / `SEIYUCOM001X`
- 脚本硬编码密码、API Key
- 数珠つなぎ；从 Designer 调生产脚本
- 用 IF / 循环做行级数据加工（放 Mapper）
- 单脚本图标 > 100
- 大件数不打开「大容量データ処理」（可能 OOM 拖垮整实例）
- Databricks 全局资源开连接池
- SQL Express 当数据库文件系统
- SFTP 密码认证
- 生产日志 DEBUG / FINEST（用 INFO）
- 高 Tier、维护窗口（不定期 14:30–16:00）不可停的处理放 DSS
- 未在西友授权表上的适配器当默认可用（Oracle/GCP/Box/SAP 等先确认）
- 把 `/data/Shared_test` 一类不通挂载写进脚本

已确认授权：已购 Salesforce / Azure / SQL Server / 固定长 / FTP / REST；标准 JDBC / CSV / Excel / XML / HTML / 文件操作 / Mail / Thunderbus / HULFT。HTML 读取走 Web Adapter。

---

## 产品操作要点

- **脚本** = Designer 源码；**项目** = 打包单位；**服务** = 登记到 Server 后才能被触发器 / 其他脚本 / ScriptRunner 调用。
- 绿三角通常执行；虫子调试。断点：图标右键。测试引数：脚本属性 → テスト実行。
- 每次保存留版本。无权限看不到别人的项目。共享：マイプロジェクト → プロパティ → 共有。
- OS 目录先挂到 DataSpider 文件系统，脚本不绑盘符。
- DB 全表：テーブル読み取り。条件/复杂 SQL：検索系 SQL（向导生成后若手改 SQL，向导不能再编）。大量写入：テーブル書き込み批处理 +「キー一致行は更新」。SQL Server 更新 SQL 用 `?{列名}`。
- 默认每图标各自提交；跨步回滚用事务图标或脚本事务属性。
- 卡住：组件 F1（含例外和限制）→ HULFT FAQ → DMS Cube。支持 hulft.com，产品 4.4 x64 AP3 / SP A00050。

---

## 回答时怎么写

用户要的是可执行的接口设计，不是产品科普。默认给出：

1. IF ID / 项目名 / 脚本名
2. 图标顺序（含 try/catch 与 End 戻り値）
3. 全局资源名、路径、`if_master` 字段
4. 触发器仍指向 `SEIYUCOM000X` / `01frameworkBatch`
5. 本条禁则里碰了哪几条

不确定的环境差异（桶路径、schema、是否已有全局资源）标成「需台账/现网确认」，不要编造连接串。
