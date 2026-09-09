import {
  Callout,
  Card,
  CardBody,
  CardHeader,
  Grid,
  H1,
  H2,
  H3,
  Pill,
  Row,
  Stack,
  Stat,
  Table,
  Text,
  useCanvasState,
} from "cursor/canvas";

type TabId = "flow" | "framework" | "naming" | "product" | "donts" | "workshop";

const TABS: { id: TabId; label: string }[] = [
  { id: "flow", label: "开发流水线" },
  { id: "framework", label: "西友框架" },
  { id: "naming", label: "命名规则" },
  { id: "product", label: "产品概念" },
  { id: "donts", label: "禁则与性能" },
  { id: "workshop", label: "培训与实演" },
];

export default function DssDevCheatSheet() {
  const [tab, setTab] = useCanvasState<TabId>("dss-tab", "flow");

  return (
    <Stack gap={20} style={{ padding: 24, maxWidth: 980 }}>
      <Stack gap={6}>
        <H1>DSS 开发速查</H1>
        <Text tone="secondary">
        西友实际开发以框架 + if_master 为中心。2026-05-13 实演的样板是 CM0010_IF0079_006：抽 Oracle → 本地 CSV/END → GCS。
        </Text>
      </Stack>

      <Row gap={24} align="end">
        <Stat value="DEV → STG → PROD" label="必须按环境顺序落地" />
        <Stat value="SEIYUCOM000X" label="触发器只调框架入口" />
        <Stat value="0 / 5 / 255" label="正常 / 警告 / 异常" />
      </Row>

      <Callout tone="warning" title="接口形态">
        只做「单体」接口。数珠つなぎ（链式）禁止。一个 IF 对应一个处理单位；同一 IF 内可以传多个文件。
      </Callout>

      <Row gap={8} wrap>
        {TABS.map((item) => (
          <span key={item.id}>
            <Pill active={tab === item.id} onClick={() => setTab(item.id)}>
              {item.label}
            </Pill>
          </span>
        ))}
      </Row>

      {tab === "flow" && <FlowSection />}
      {tab === "framework" && <FrameworkSection />}
      {tab === "naming" && <NamingSection />}
      {tab === "product" && <ProductSection />}
      {tab === "donts" && <DontsSection />}
      {tab === "workshop" && <WorkshopSection />}

      <Text tone="tertiary" size="small">
        来源：产品介绍、DSSK v44、ガイドライン v3.2、DSS 開発の流れ、三日 WS 录像（Day1–3）、リリース手順書、2026-05-13 Teams 实演。2026-07-22 录像约 3GB、无字幕，未能转写。凭证不转载。
      </Text>
    </Stack>
  );
}

function FlowSection() {
  return (
    <Stack gap={16}>
      <H2>一条接口从取号到上线</H2>
      <Text tone="secondary">
        先在西友侧接口台账拿到 IF ID，再按下面 8 步在 DEV 做完，原样推到 STG、PROD。
      </Text>

      <Table
        headers={["步骤", "做什么", "要点"]}
        rows={[
          ["1. IF ID", "从《SEIYU-TRIAL インターフェース管理台帳》取得", "FROM6_TO6_SEQ3，例 AM0002_MS0007_001"],
          ["2. 项目", "项目名 = 接口 ID", "从现有脚本复制，不要从空白重写框架调用"],
          ["3. 脚本", "脚本号 + 驼峰英文/罗马字", "例 01updDelDate；只写个别处理"],
          ["4. 资源", "Explorer 目录、组权限、全局资源", "STG/PROD 全局资源由 Data 团队建；SFTP 只用密钥"],
          ["5. 服务注册", "项目注册为服务，服务名 = 项目名", "改过脚本必须重新注册服务"],
          ["6. 触发器", "My Trigger → 调框架，不调个别脚本", "执行用户 exec_SY；参数 if_id"],
          ["7. if_master", "向 DSSSEIYU.if_master 插记录", "TO 必须含系统支持台；STG/PROD 需 RDP + SSMS"],
          ["8. 测试", "跑通后再升环境", "STG/PROD 走 SDP 变更管理；监视提前约 1 周找 ITSM"],
        ]}
        striped
        stickyHeader
      />

      <Grid columns={3} gap={12}>
        <Card>
          <CardHeader>DEV</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>WebStudio：dhub-dss-dev.seiyu.com:8443</Text>
              <Text>账号：个人账号</Text>
              <Text>DB：DHUB-DB-DEV / DSSSEIYU，DBeaver 也可</Text>
              <Text>GCS：ods-seiyu-dev</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>STG</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>WebStudio：dhub-dss-stg.seiyu.com:8443</Text>
              <Text>账号：exec_SY</Text>
              <Text>DB：DHUB-DB-STG，必须 RDP + SSMS</Text>
              <Text>GCS：ods-seiyu-stg</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>PROD</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>WebStudio：dhub-dss-pro.seiyu.com:8443</Text>
              <Text>账号：exec_SY</Text>
              <Text>DB：DHUB-DB-PRO，必须 RDP + SSMS</Text>
              <Text>GCS：ods-seiyu-prod</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H3>触发器绑定（STG/PROD 一样）</H3>
      <Table
        headers={["字段", "值"]}
        rows={[
          ["服务", "SEIYUCOM000X"],
          ["脚本", "01frameworkBatch"],
          ["变量 if_id", "if_master 里登记的接口 ID"],
          ["执行用户", "exec_SY"],
          ["状态", "登録后应为「待機中」"],
        ]}
        framed
      />

      <Callout tone="info" title="申请走邮件，不自己建生产资源">
        用户、Azure 存储、全局资源一律发 mls_ISDDataPlatform@seiyu.com（用户申请 CC 所属 Director）。开发环境全局资源可自建，STG/PROD 走申请。
      </Callout>

      <H3>测试（Day3）</H3>
      <Table
        headers={["手段", "怎么用"]}
        rows={[
          ["通常执行", "绿三角。快，细节日志少"],
          ["调试执行", "虫子图标。输出执行历史/日志，可逐步看"],
          ["断点", "图标右键「ブレークポイントを設定/解除」。停在该组件，再按调试继续"],
          ["测试引数", "脚本右键 → 属性 → テスト実行。给脚本输入变量喂测试值"],
          ["调试信息", "看运行中的脚本变量"],
          ["错误日志", "用消息代码去 Help 搜原因和对策"],
          ["分支覆盖", "条件分支每条路径（①②③）都要有期望结果"],
        ]}
        striped
      />
      <Text tone="secondary">
        DEBUG 日志在 Designer 里可能被截断，去 My Log 看完整文件。生产日志用 INFO。
      </Text>

      <H3>发布（リリース手順書 v1.0）</H3>
      <Text tone="secondary">
        HR-EAI 发布步骤书（2024-11-14）。用资源管理员账号登录目标环境 WebStudio。项目走下载/上传；触发器等走サーバ移行；if_master 走 SQL Server，且必须先备份。
      </Text>
      <Table
        headers={["搬什么", "方法", "硬性规则"]}
        rows={[
          ["1.1 项目下载", "マイプロジェクト → 右键 → ローカルへダウンロード", "勾选「バージョンを指定してダウンロード」，选数字最大的最新版"],
          ["2.1 项目上传", "ファイル → ローカルからアップロード", "上書きアップロード＝作为现有项目最新版本提交；新项目名上传＝另起项目。然后打开 → 服务注册（服务名=项目名）→ マイサービス确认"],
          ["权限", "下载/上传不带共享", "STG/PROD 上传后重新共享给 SY 组（ガイドライン）"],
          ["1.2 非项目导出", "コントロールパネル → DataSpiderServerの設定 → サーバ移行 → エクスポート", "先「全て解除」，只勾本次要搬的资材，确认导出结果无错误"],
          ["2.2 非项目导入", "同上 → インポート", "不要去掉「プロジェクトの差分インポート」「環境変数の差分更新」。去掉会整表覆盖。导入后重启 DataSpiderServista。同名时：服务/触发器/全局资源等覆盖；环境变量和项目文件可跳过。全局用户逻辑同名不同内容会自动改成 name_1"],
          ["3. 框架表", "SSMS 改 DSSSEIYU", "先 3.1 备份。ガイドライン示例是同 if_id DELETE 再 INSERT；TO 含支持台"],
        ]}
        striped
      />

      <H3>系统表改数之前（手順书 3.1）</H3>
      <Table
        headers={["方式", "步骤"]}
        rows={[
          ["CSV 备份", "新查询 → SELECT 对象表 → 结果右键「結果に名前を付けて保存」"],
          ["脚本备份", "任务 → スクリプトの生成 → 对象表 → 详细设定「スキーマとデータ」→ 存 sql"],
          ["3.2 追加", "上位200行の編集 → 末行 * → 从 Excel 粘贴或手输 → SQL 再查确认"],
          ["3.3 修正", "上位200行の編集 → 查询设计器 SQL 窗加 WHERE → 执行后改值/粘贴 → Enter"],
          ["3.4 删除", "同样用 WHERE 筛出 → 选行删除 → SQL 确认已不存在"],
        ]}
        striped
      />

      <Text tone="tertiary" size="small">
        主机、账号口令写在《DSS 開発の流れ》和ガイドライン P.103 / P.115–118，不要抄进脚本或仓库。
      </Text>
    </Stack>
  );
}

function FrameworkSection() {
  return (
    <Stack gap={16}>
      <H2>西友共通框架</H2>
      <Text tone="secondary">
        触发器 → 服务控制器 → 读 if_master / param_master → 调个别项目 → 写日志 / 发邮件。个别处理只负责业务，返回码交给框架收口。
      </Text>

      <Table
        headers={["项目", "角色", "能改吗"]}
        rows={[
          ["SEIYUCOM000X", "入口。01frameworkBatch 由调度触发器、JP1、ScriptRunner 调用", "只用，不改"],
          ["SEIYUCOM001X", "共通：01getInterfaceInfo / 02outputLog / 03sendMail / 调个别脚本", "全项目共用，禁止修改"],
          ["SEIYUCOM002X", "ADLS 路径切割库", "可用"],
          ["SEIYUCOM003X", "备份清理（每天 0 点）", "了解即可"],
          ["HOUSEKEEP", "DSS 本地备份清理", "非推荐，备份放 Blob"],
          ["接口 ID 项目", "个别处理，被 001X 调用", "这里写业务"],
        ]}
        striped
        rowTone={["info", "warning", undefined, undefined, "warning", "success"]}
      />

      <Grid columns={2} gap={12}>
        <Card>
          <CardHeader>if_master 必填列</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>if_id / if_name / service_name / script_name</Text>
              <Text>from / to / cc / bcc / mail_title_header / tier</Text>
              <Text tone="secondary">
                to 必须含 isd_support_desk@seiyu.com。改表前按リリース手順书 3.1 备份；发版常用同 if_id DELETE 再 INSERT。
              </Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>param_master</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>IF_ID + param_no + param_value</Text>
              <Text>01getInterfaceInfo 读入后塞到 param1–param20</Text>
              <Text tone="secondary">个别脚本用这些变量，不要把环境差异写死在图标属性里。</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H3>返回码与四条收口（Day2）</H3>
      <Text tone="secondary">
        框架先读 if_master、写开始日志、调个别处理，再按结束状态分四路。个别脚本不要自己发邮件、写运维日志。
      </Text>
      <Table
        headers={["分支", "个别处理怎么结束", "框架做什么"]}
        rows={[
          ["1 正常", "End 戻り値 0（默认）", "正常日志后结束"],
          ["2 警告", "戻り値 1–254（演示用 0 件数据）", "警告邮件 + 警告日志。例：文件只有表头"],
          ["3 可控异常", "戻り値 255 等，脚本自己设了 out_*", "异常邮件 + 异常日志"],
          ["4 未捕获", "Exception / 未 catch 的错误", "catch 后异常邮件。戻り値实际为 -1"],
        ]}
        striped
      />
      <Text>
        警告/异常时用「警告变量设定」「异常变量设定」把 catch 内容塞进 out_component_name/type、out_message_*、out_error_*。邮件后半段就是这些字段，例如 FILE0001E|java.io.FileNotFoundException|/data/DSSROOT/DEMO/INPUT/OrderList.csv|CSV読取。发信 no-reply-HR-EAI@seiyu.com。
      </Text>
      <Text tone="secondary">
        个别处理必须：登记 if_master、设 End 戻り値。可选：param_master 入参、输出 out_*。Day2 把 Day1 脚本改成框架形态：正常=跑完；警告=0 件；异常=读或写失败。
      </Text>

      <Callout tone="info" title="JP1 / ScriptRunner">
        用共通模板 C:\Temp\scriptrunner.xml，第一参数传 if_id 即可。日志进 My Log，和触发器启动一样。
      </Callout>
    </Stack>
  );
}

function NamingSection() {
  return (
    <Stack gap={16}>
      <H2>命名与资源约定</H2>
      <Table
        headers={["对象", "规则", "例子"]}
        rows={[
          ["接口 ID / 项目名", "FROM6_TO6_SEQ3 大写下划线", "AM0002_MS0007_001"],
          ["脚本名", "两位序号 + 驼峰英文或罗马字", "01updDelDate"],
          ["常量", "大写下划线", "FILE_NAME"],
          ["变量", "小写下划线", "file_name"],
          ["参数", "in_ / _out / io_ 前缀", "in_file_name"],
          ["环境变量", "ENV_ 前缀大写下划线", "ENV_RETRY_COUNT"],
          ["表名", "M_ 主数据 / T_ 事务", "M_FILE_SEND"],
          ["ADLS 容器", "接口 ID 改小写，只用连字符", "am0002-ms0007-001"],
          ["固定长/可变长格式", "FIX|VAR + READ|WRITE + 名称", "FIX_READ_Time"],
          ["全局资源（DB/SFTP 等）", "协议_系统6位[_Intake][_枝番]", "SFTP_AM0002_用户ID"],
          ["触发器", "接口ID_SCHL 或 _SCHL_nnn", "AM0002_MS0007_001_SCHL"],
          ["工作文件", "小写下划线英文", "file_name"],
        ]}
        striped
        stickyHeader
      />

      <Grid columns={2} gap={12}>
        <Card>
          <CardHeader>ADLS / 备份</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>接口 ID 一个容器；外部用 SFTP 访问。</Text>
              <Text>文件约 30 天后自动删除。</Text>
              <Text>备份原则进 Blob（AZURE_STR_DSS）。/data/backup 仅备选。</Text>
              <Text>DSS 工作目录：/seiyuroot/&lt;IF_ID&gt;/（实演）</Text>
              <Text>GCS：ods-seiyu-&#123;env&#125; / &lt;IF_ID&gt; / deliveryData / YYYYMMDD/ ，另有 deliveryData_bk</Text>
              <Text>GCS 适配器连接名：GCP_STR_DSS；路径可用 $&#123;folder_date&#125;</Text>
              <Text>SSH 密钥：/data/SSH/&lt;Intake&gt;/xxxxx.ppk</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>全局资源前缀</CardHeader>
          <CardBody>
            <Stack gap={6}>
              <Text>SQL / ORCL / JDBC / ODBC</Text>
              <Text>FTP / REST / MAIL / BLOB</Text>
              <Text>写 Databricks：经 ADLS（MS0007）</Text>
              <Text>读 Databricks：JDBC（MS0005）</Text>
              <Text>DSS 自用 SQL：SQL_DSS_SEIYU_DB</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>
    </Stack>
  );
}

function ProductSection() {
  return (
    <Stack gap={16}>
      <H2>DataSpider 产品怎么工作</H2>
      <Text tone="secondary">
        GUI 脚本在 Server 上编译成 Java。四块：Adapter 入出、Converter/Mapper 转换、Studio 开发运维、Trigger 自动启动。西友用 Studio for Web，版本 4.4 x64 AP3。
      </Text>

      <Grid columns={2} gap={12}>
        <Card>
          <CardHeader>Studio 工具</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>Designer：画脚本和 Mapper</Text>
              <Text>Control Panel：用户、任务、全局资源、挂载</Text>
              <Text>Explorer：逻辑文件系统</Text>
              <Text>My Service / My Trigger / My Log</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>两条线不要混</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>数据流：橙色虚线，指定输入数据后自动出现，可一对多。</Text>
              <Text>进程流：黑色实线，必须手拉 Start→…→End，一个图标只能有一条出线。</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H3>读、转、写</H3>
      <Table
        headers={["场景", "用什么", "注意"]}
        rows={[
          ["DB 全表抽出", "テーブル読み取り", "全局资源可复用"],
          ["带条件抽出 / 复杂 SQL", "検索系 SQL 実行", "向导生成后若手改 SQL，向导就不能再编"],
          ["字段变换、过滤、日期、拼接", "Mapper（约 140 个逻辑图标）", "输出 schema 要先定；CSV 可用表头文件"],
          ["文件 JOIN", "Merge（APPEND / INTEGRATE / Left-Right JOIN）", "DB 侧优先用 SQL JOIN"],
          ["大量写入 / 主数据洗い替え", "テーブル書き込み + 批处理", "勾选「キー一致行は更新」= upsert"],
          ["SQL Server 适配器", "表读/写、検索SQL、更新SQL、ストアド、ローダ", "连接例 COM_SQLServer2022；更新SQL 用 ?{列名} 占位"],
          ["带事务的 CSV→表", "Transaction → DELETE → csv_read → mapping → テーブル書き込み → COMMIT", "Day2 DBデモ1。process_date 用「現在日時」"],
          ["复杂更新 / 外结合", "更新系 SQL 実行", "不适配超大件数"],
          ["固定长 EDI", "固定长适配器 + 向导", "Group / Record / Field；定义进全局资源"],
        ]}
        striped
      />

      <H3>触发器种类（教材）与西友常用</H3>
      <Text>
        教材覆盖：スケジュール / ファイル / HTTP / Web 服务 / DB / FTP / SAP / ScriptRunner / HULFT。西友主用スケジュール、ファイル，以及 JP1→ScriptRunner。
      </Text>
      <Text tone="secondary">
        开发库与运行库是分开的：只有「作为服务注册」的项目才能挂触发器。改脚本后必须重新注册。
      </Text>

      <H3>本环境已确认的适配器（Day1 導入有無）</H3>
      <Table
        headers={["状态", "适配器"]}
        rows={[
          ["已购 ●", "Salesforce、Microsoft Azure、SQL Server（选择式）、固定长、FTP、REST"],
          ["标准 〇", "JDBC、CSV、Excel、XML、HTML、文件操作、Mail、Thunderbus、HULFT"],
        ]}
        framed
      />
      <Text tone="secondary">
        表上未标 ●/〇 的（Oracle、GCP、Box、SAP 等）不要默认当成可用。可变长走可変長适配器；HTML 读取走 Web Adapter。
      </Text>

      <H3>事务</H3>
      <Text>
        默认每个图标各自提交。用事务图标或脚本「事务属性」可以把多步（含子脚本）包成同一提交单位；任一步失败则回滚。
      </Text>
    </Stack>
  );
}

function DontsSection() {
  return (
    <Stack gap={16}>
      <H2>禁则、性能、可用性</H2>

      <Callout tone="danger" title="会拖垮整台 DSS 的写法">
        脚本里用 IF / 循环做数据加工会极慢，尽量放进 Mapper。单脚本图标建议不超过 100 个。大件数务必打开「大容量データ処理」，否则可能 Java OOM，影响全实例。
      </Callout>

      <Table
        headers={["机制", "行为", "怎么用"]}
        rows={[
          ["PSP（并行流）", "按约 1000 件一块，读/转/写多线程；块大小不可改", "智能编译器默认可开；Designer 打开 PSP 数据流显示"],
          ["大容量データ処理", "内存只留必要数据，其余落盘", "PSP 优先；PSP 后若 Mapper 接到非 PSP 写入，中间仍会走大容量"],
          ["批写入", "テーブル書き込み批量模式", "86 万件时批量 upsert 约比普通 upsert 快 30 倍"],
          ["1 亿件级", "验证中 PG 7h 仍失败；tmp 可超 10GB", "不要用 DSS 扛；Postgres 用 psql /copy 等"],
        ]}
        striped
      />

      <Grid columns={2} gap={12}>
        <Card>
          <CardHeader>明确禁止 / 不要做</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>改 SEIYUCOM001X 共通脚本</Text>
              <Text>脚本里硬编码密码、API Key</Text>
              <Text>数珠つなぎ接口</Text>
              <Text>Databricks 全局资源开连接池</Text>
              <Text>用 SQL Express 数据库文件系统（10GB 上限）</Text>
              <Text>SFTP 密码认证（只用密钥）</Text>
              <Text>生产日志开 DEBUG / FINEST（用 INFO）</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>运维边界</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>当前单机。故障即停服。</Text>
              <Text>不定期 14:30–16:00 维护停机。</Text>
              <Text>高 Tier、维护窗口不可停的处理不要放 DSS。</Text>
              <Text>卡住时：My Log → 任务管理器结束进程；不行就重启。</Text>
              <Text>Desktop Studio 模块同步可超 4 小时，西友以 Web 为主。</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H3>数据质量（指南基本方针）</H3>
      <Text>
        要考虑字符编码/乱码、缺失/重复/异常值的早期检测，以及全角半角、电话号码连字符等输入标准化。
      </Text>

      <H3>厂商支持</H3>
      <Text>
        hulft.com/tech-support · 产品 DataSpider Servista 4.4 (x64) AP3 · SP コード A00050。序列号在ガイドライン末页，按环境区分本番/开发/验证机。
      </Text>
      <Text tone="secondary">
        卡住时：选中组件按 F1 跳对应帮助（含例外和限制）→ FAQ → DMS Cube。工作坊 QA 窗口已过期（当时最多 30 件，走 seiyu-hr-eai-admin@saison-technology.com）。
      </Text>
    </Stack>
  );
}

function WorkshopSection() {
  return (
    <Stack gap={16}>
      <H2>三日工作坊在教什么</H2>
      <Text tone="secondary">
        2024-11/12 セゾン「ツール活用トレーニング」。Day1 操作与演示，Day2 框架 + SQL Server，Day3 测试与发布。
      </Text>

      <Table
        headers={["天", "主题", "材料状态"]}
        rows={[
          ["Day1（11/22 录像 1:58）", "Studio 操作、服务注册比较画面、CSV 读（Windows-31J、跳过首行当表头）", "MP4 已抽帧"],
          ["Day2（11/28 录像 1:57）", "框架四分支、个别处理改写、if_master/SQL Server、大容量データ処理", "MP4 已抽帧（此前 Clipchamp 为空）"],
          ["Day3（12/05 录像 1:45）", "调试、导入覆盖表、现场翻リリース手順書", "MP4 已抽帧"],
        ]}
        striped
        rowTone={[undefined, "warning", undefined]}
      />

      <Callout tone="info" title="Day2 录像已看（11/28 延吉）">
        框架调用个别处理后分四路结束。练习库 SAMPLE_PROJECT，工作目录 /data/DSSROOT/DEMO/（OS 上 F:\DSSROOT\DEMO\）。SQL 在 DHUB-DB-DEV / DSSROOT：if_master、param_master、DEMO_OrderList。SQL 适配器还能跑ストアド和ローダ。组件「データ処理方式」可选「スクリプトの設定を使用する」，大件数不要全进内存。
      </Callout>

      <Callout tone="info" title="2026-07-22 Teams「DSSについて」">
        zip 里只有一条约 3.07GB 的 mp4（15:18 开始），没有 VTT/纪要。日期与《DSS 開発の流れ》修订日一致。本机未能转写音频。
      </Callout>

      <H2>2026-05-13 实演（田村 / seiyu_tamura）</H2>
      <Text tone="secondary">
        Chrome 录的 Teams：前 3 段共约 7 分钟（流程 PPT、Drive 资料夹、My Project），主会 52 分 52 秒。资料夹在 Google Drive「9商品マスタ(惣菜) / DSS」。登录 dhub-dss-dev，有リポジトリ DB。
      </Text>

      <H3>台账上的这条 IF</H3>
      <Table
        headers={["项", "值"]}
        rows={[
          ["演示项目", "CM0010_IF0079_006（脚本 01OutputData）"],
          ["台账页", "SEIYU-TRIALインターフェース管理台帳 → DSSインターフェース台帳"],
          ["FROM", "CM0010 若菜 新店舗システム"],
          ["TO", "IF0079 TRIAL データ連携ストレージ（也有反向 IF0079_CM0010、以及 IF0079_MS0007 进数据基盤）"],
          ["业务组", "生鮮発注 / 惣菜 / 非生鮮発注 / EC"],
          ["启动", "多为スケジュール；サイクル 日次・週次・随時"],
          ["Oracle", "接続 ORCL_CM0010_wknpsuser；スキーマ WKNPSUSER；例 WKNST_DAILYORDER（COMPANY_CD / STORE_CD / SALES_DATE=YYYYMMDD）"],
        ]}
        striped
      />

      <H3>01OutputData 实际图标顺序</H3>
      <Table
        headers={["图标", "作用"]}
        rows={[
          ["try / catch", "框架用的异常监视；catch 接 out_* 错误变量"],
          ["delete", "清工作区 /seiyuroot/CM0010_IF0079_006/*"],
          ["exec_select_CM0010", "検索系 SQL。连接 ORCL_CM0010_wknpsuser，select * from WKNPSUSER…"],
          ["現在日付取得", "Mapper 出 file_date、folder_date"],
          ["csv_write", "写出 CSV"],
          ["endfile", "空完了文件 ORDER_${file_date}.END，目录同上"],
          ["gcs_put", "GCP_STR_DSS → 桶 ods-seiyu-dev，路径 /CM0010_IF0079_006/${folder_date}/，本地 * 全上传；无文件夹则建，已有则覆盖"],
        ]}
        striped
      />

      <Text>
        GCS 上已有同族接口 CM0010_IF0079_005：ods-seiyu-dev / IF_ID / deliveryData / 日付フォルダ，以及 deliveryData_bk。东京 asia-northeast1、Standard、非公开。
      </Text>
      <Text tone="secondary">
        另外演示了 FTP「リストとデータの読み取り」ftp_get_data、连接 SFTP-cao-test。做完脚本后右键「プロジェクトをサービスとして登録」。Explorer 访问 /data/Shared_test（\\10.10.1.125\Shared_test）当时弹出错误，挂载不通不要硬写这条路径。
      </Text>

      <H3>Day1 三个演示（典型脚本模式）</H3>
      <Table
        headers={["演示", "输入", "要做出的行为"]}
        rows={[
          ["① CSV 转换集计", "OrderList.csv（受注日、支店、ケース数、入り数、容量、商品名、受注_ID）", "按支店汇总发注数，写出 OrderList_total.csv"],
          ["② 0 件当错误", "只有表头的 OrderList.csv", "不写输出文件，主动报错"],
          ["③ 文件触发器", "OrderList.csv 放到输入文件夹", "触发 ① 的脚本，结果进输出文件夹"],
        ]}
        striped
      />

      <H3>Studio 操作要点（Day1）</H3>
      <Grid columns={2} gap={12}>
        <Card>
          <CardHeader>脚本 vs 项目 vs 服务</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>脚本 = 源码，Designer 里画。</Text>
              <Text>项目 = 相关脚本的打包；读写、服务注册都以项目为单位。</Text>
              <Text>服务 = 项目登记到 Server 后，才能被触发器、其他脚本、ScriptRunner 调用。</Text>
              <Text>绿三角普通跑，虫子调试跑。</Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>版本、共享、挂载</CardHeader>
          <CardBody>
            <Stack gap={4}>
              <Text>每次保存都留版本；可还原，可出版本比较报告。</Text>
              <Text>无权限看不到别人的项目。共享：マイプロジェクト → 右键 → プロパティ → 共有。</Text>
              <Text>一般用户只能看自己当「触发器所有者」的触发器。</Text>
              <Text>OS 目录要先挂到 DataSpider 文件系统（例 C:\Document\Sample → /data/mount）脚本才不绑盘符。</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H3>环境架构摘录（Day3 / インフラ基本设计）</H3>
      <Text>
        2024 年 HR-EAI 设计：主环境在東日本，本番 + Staging；西日本 DR 明确不用于 DataSpider 切转。当时写的是「Staging 里放开发机，不单独建 Development」。2026 年《DSS 開発の流れ》已经有独立的 dhub-dss-dev / stg / pro 三套 WebStudio，日常按这三套走。
      </Text>
      <Table
        headers={["环境", "用途（设计书）"]}
        rows={[
          ["Production", "本番业务，24H/365D，商用许可，本番数据"],
          ["Staging 验证专用机", "外部连接、结合/综合、运用验证，构成与本番相同，开发许可"],
          ["验证兼开发机", "单体～结合，开发数据，构成不必与本番相同"],
        ]}
        framed
      />

      <H3>帮助入口</H3>
      <Text>
        选中组件按 F1 → 逆引きリファレンス（带样本项目，可上传后对照做）→ HULFT FAQ → DMS Cube。Help 里能查例外和规格限制。
      </Text>
    </Stack>
  );
}
