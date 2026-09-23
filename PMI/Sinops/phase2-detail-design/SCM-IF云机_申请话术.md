# SCM連携云机　申请话术（中文）

用途: Spec 申请时，用约 **2～2.5 分钟** 说明 SCM連携クラウド機 范围、执行基盤选择、往路复路与边界，并带上近期计划。短时用 15 秒版。  
日文对照: `SCM-IFクラウド機_申請話術.md`  
依据: `SCM-IFクラウド機アーキテクチャ.md` / 同名 draw.io  
名称: 口头用 **SCM連携クラウド機**。见积 Layer①。不只是文件 IF，共通化、型转换和一部分生成逻辑也在本机。  
假称／ID: **GCP 项目 ID（暂定）=`md-data-integration-prod`。** 旧资料里的 ProjectD / Seiyu_Order 是指向该 ID 的说明用别名。口头可说暂定 ID，或直接说「SCM連携クラウド機」。  
备注／用語: **STS**＝Google Cloud 的 Storage Transfer Service（ストレージ転送サービス）。

---

## 申请什么（口头摘要）

| 申请 | 不申请（本次云机 Spec 外） |
|---|---|
| TRIAL GCP **新建 1 个项目**（SCM連携クラウド機。暂定 ID: **`md-data-integration-prod`**） | **GKE**（不需要常设集群） |
| 执行基盤用 **Cloud Run Jobs**（东京。4 vCPU / 16 GiB 案） | 本机（VM）**持久化磁盘／ステージング** |
| **`md-data-integration-prod` 附带 GCS（集计保存）** 与 **西友连携 GCS（处理结果）**（均在东京） | DataSpider（西友 Azure。②③） |
| （桶契约需对账） | 自动发注 GCS（TRIAL 既存。不是云机项目） |

要点: 本机是 **文件 IF 批处理转换机**。集计数据保存在 **`md-data-integration-prod` 附带 GCS** → Cloud Run（SMART）处理后 → **仅最终 TXT 交付到西友连携 GCS** → 西友DS。**源数据不进西友连携 GCS。无二次プッシュ。** 申请构成图上控制流为 Scheduler+Workflows；与 Hinemos 的分工 TBD。不申请 GKE。

### 为什么用 Cloud Run、不用 GKE（对比·口头用）

| 観点 | **Cloud Run Jobs（采用）** | **GKE（不采用）** |
|---|---|---|
| 运行方式 | 按时启动 → 处理 → 结束；空闲可停 | 集群／节点容易常设 |
| 本机工作 | 抽出・转换・GCS プッシュ／プル的**批处理** | 更适合微服务／常驻 API |
| 存储 | 无状态；landing/output・正本在 **GCS**；与设计一致 | 往往要节点／持久卷运维 |
| 运维・费用 | 申请与运维轻；按批处理计费 | 运维重、固定成本高；对本案过重 |
| 结论 | **适合 IF 连携机** | **本次范围不需要** |

口头一句: 「不是常开的服务底座，而是需要时才跑的批处理，所以用 Cloud Run Jobs。GKE 以常设集群为前提，对本 IF 机过重。」

---

## 时间不够时（15 秒）

> 申请的是 TRIAL GCP 上新建的 1 个项目，SCM連携クラウド機。暂定 ID 是 **`md-data-integration-prod`**。执行用 **东京的 Cloud Run Jobs**。从 **集计SV GCS（大阪）** 同步到 **本项目附带 GCS**，处理后 **仅把最终 TXT 交付到西友连携 GCS**。**源数据不进西友侧。无二次プッシュ。** 不申请 GKE。

---

## 正文（约 2 分～2 分半）

今天说明 SCM連携クラウド機 的申请范围。分四点：申请什么、为什么选这个执行基盤、往路和复路做什么、哪些不进本次 Spec。

先说申请对象。在 TRIAL GCP 上新建 **1 个项目**。**暂定项目 ID 是 `md-data-integration-prod`。** 资料里的 ProjectD、Seiyu_Order 是指向这个暂定 ID 的说明用叫法。定位上，和现有的惣菜、MD-Link、Inunaki 并列，是一台 **SCM 連携机**。如果按课题拆成多个项目，启动对象和文件落点会变成两套，所以 **收敛为 1 个项目**。环境设计图（gcp環境設計／current-architecture-flow）就是 **这个项目的申请构成图**。

再说执行基盤。采用 **Cloud Run Jobs**（东京 `asia-northeast1`，申请图为 4 vCPU / 16 GiB）。流程：**集计SV GCS（大阪）** 输入确认 → Storage Transfer 同步到 **本项目附带 GCS（`source-landing/`，源数据约 7 日）** → Cloud Run（SMART）转换 → **仅最终 TXT upload 到西友连携 GCS（`result/` + `_SUCCESS`）** → 西友 DataSpider。**源数据不进西友连携 GCS；Seiyu 侧权限仅为 objectCreator。无二次プッシュ。** 申请图上的控制流是 **Scheduler → Workflows**；**将来也有可能改成 Hinemos 启动**（执行体仍是 Cloud Run）。与 Hinemos（尤其 DSS②③）的分工仍要再对一次。不申请 GKE。不做 VM／本机持久化ステージング。

负责范围是见积的 **Layer①**。不只是文件进出，还要做吸收西友与 TRIAL 差异的 **共通化・型转换**，以及一部分生成逻辑。  
往路：从 **集计SV GCS（大阪）** 同步到 **`md-data-integration-prod` 附带 GCS**，经 Cloud Run 处理后把结果反映到 **西友连携 GCS**。TRIAL 与西友之间的唯一通道就是西友连携 GCS。  
复路发注劝告：西友 Azure 的 **Layer③ DataSpider** 写到同 GCS；**Layer① Cloud Run** プル后转成发注 IF，再 **プッシュ到自动发注 GCS**。

Layer②、Layer③ 在西友 Azure 的 **既存 DataSpider**。② 是主数据 Intake；③ 往路送到 Sinops／BO，复路把劝告写到西友连携 GCS。**不新申请 DataSpider，用既存即可。** 西友连携 GCS 本体不进本次云机 Spec，但仍要向松尾 **另行申请**。BO 程序本阶段不做，但会共用同一套云机与 DataSpider 面。

最后是计划。本周内继续固化 SCM連携クラウド機 的功能范围。到下周为止，由中国侧团队确认架构；**GCP 费用已有 Calculator 初算（月约 ¥1.8 万，大半为大阪→东京転送）**，再交给岩濑先生提出。

以上。收束：「1 个项目的連携机」「东京 Cloud Run Jobs、不上 GKE」「源数据在本项目附带 GCS、仅最终 TXT 到西友连携 GCS、无本机磁盘」「源数据隔离・单方向纳品」「往路交付西友连携 GCS、复路プル后到自动发注 GCS」「Layer②③ 用既存 DataSpider／西友连携 GCS 另行申请」。

---

## 被问到再补（正文里不念）

**为什么不叫 IF 云机。**  
见积 Layer① 写的是外部 IF，但本机还做共通化、型转换和一部分生成，所以叫 SCM連携クラウド機。

**为什么用 Cloud Run、不用 GKE／Autopilot。**  
见第 2・7 页。本件是日次文件批，Cloud Run Jobs 对口；GKE（含 Autopilot）仍有集群管理费常驻、运维更重，且 Spec 口径不上 GKE。第 7 页对照的是 **Cloud Run ≈ ¥18,190** vs **Autopilot（管理费込み）≈ ¥29,000** vs **Standard 常驻 ≈ ¥36,000〜48,000**（讨论用概算）。

**DataSpider 会做类型转换或字段筛选吗。**  
原则上不做。DSS（Layer②／③）只做 **全字段通过／转发**。类型转换、字段筛选、代码转换、编辑都在 **Layer①（本云机）**。

**往路都从集计 SV 取吗。**  
不是。多数往路从集计 SV。新商品在库在云机用西友 DWH 棚割 × TRMD マスタ生成。仕入先休日的源是 TRIAL 自动发注连携。发注劝告是 Layer③ DataSpider 从 Sinops 处理到西友连携 GCS，再由 Layer① 云机プル、转换后 TRIAL 自动发注 GCS へプッシュ。

**本机要带磁盘做落地吗。**  
不做 VM／本机持久化。**集计数据在 `md-data-integration-prod` 附带 GCS、处理结果在西友连携 GCS（均在东京）。** 自动发注 GCS 是复路专用、另一套。

**本项目附带 GCS 和西友连携 GCS 是两套吗。**  
**是两套。** 集计／源数据放在 **`md-data-integration-prod` 附带 GCS（`source-landing/`）**；处理后 **仅最终 TXT** 交付到 **西友连携 GCS（`result/`）**。**源数据不进西友侧。** 无二次プッシュ。

**启动是 Hinemos 还是 Scheduler。**  
申请构成图上是 **Cloud Scheduler + Workflows**（READY→STS→Job→验证）。**将来有可能改为由 Hinemos 启动 Cloud Run Job**（执行体仍是 Cloud Run，不变）。与 Hinemos／DSS②③ 的分工仍要再对。DSS 侧暂仍按 Hinemos。

**GCS 不进 Spec，是不是不用申请。**  
这次 Spec 申请云机（`md-data-integration-prod`：Cloud Run 与 **本项目附带 GCS** 等）。**西友连携 GCS** 不是云机项目，但仍要向松尾另行申请。DataSpider **用既存**，不进这次 Spec。

**西友连携 GCS 和自动发注 GCS 是两套吗。**  
是两套。西友连携 GCS 是 TRIAL 侧和西友侧之间的唯一通道。自动发注 GCS 是 TRIAL 内的既存功能，用来把劝告交给 OrdreSV。本项目附带 GCS（集计保存）也要分开写。

**为什么必须是 1 个项目。**  
为了不把 SCM 统一联动面拆开。按课题各开一个项目，启动对象和文件落点就会变成两套。

**资料里的 ProjectD 或 md-data-integration-prod 是正式名吗。**  
**`md-data-integration-prod` 是暂定项目 ID**，正式采番仍可能调整。ProjectD / Seiyu_Order 是旧资料说明用别名，现在都指同一台机。环境设计 pptx 当作本项目申请图用。

**BO 和共通基盤呢。**  
BO 程序本阶段不做，但会用同一套云机和 DataSpider。共通基盤的启动是枢纽侧，不是云机上的連携・转换本身。

**Inunaki 的 PSC 呢。**  
那是网络前提。不是这次申请的云机功能范围。和桶契约的最终对应还没对上。

**费用大概多少。**  
GCP Pricing Calculator（`SPEC.xlsx`、2026-09-22）初算月约 **¥18,190**、年约 **¥218,280**（×12，讨论用、非拘束）。其中约 **¥15,300／月** 是亚洲区域内大阪→东京転送；Cloud Run 与东京 Standard Storage 相对较小。Scheduler／Workflows／STS 等本表未计。
