import {
  BarChart,
  Callout,
  Card,
  CardBody,
  CardHeader,
  Divider,
  Grid,
  H1,
  H2,
  H3,
  PieChart,
  Pill,
  Row,
  Stack,
  Stat,
  Table,
  Text,
  useCanvasState,
} from "cursor/canvas";

type Family = "master" | "plan" | "actual" | "stock";
type Filter = "all" | Family;

type IfRow = {
  id: string;
  name: string;
  family: Family;
  dir: string;
  format: string;
  mode: string;
  fields: number;
  recLen: string;
  ui08: string;
  ss06: string;
  sample: string;
  note: string;
  tone?: "success" | "warning" | "danger" | "info" | "neutral";
};

const IFS: IfRow[] = [
  {
    id: "010S（商品）",
    name: "商品マスタ",
    family: "master",
    dir: "MD → WMS",
    format: "SJIS 固定長",
    mode: "差分",
    fields: 48,
    recLen: "512",
    ui08: "有",
    ss06: "有",
    sample: "10,280 行",
    note: "SFTP GET。UI08 多出温度帯詳細／仕入区分",
    tone: "warning",
  },
  {
    id: "020S（部门）",
    name: "部門マスタ",
    family: "master",
    dir: "MD → HUB → WMS",
    format: "UTF-8 CSV",
    mode: "全件",
    fields: 15,
    recLen: "—",
    ui08: "无",
    ss06: "有",
    sample: "4,052 行",
    note: "规格写 BOM 付，样本无 BOM",
    tone: "warning",
  },
  {
    id: "030S（仕入先）",
    name: "取引先マスタ",
    family: "master",
    dir: "MD → HUB → WMS",
    format: "UTF-8 CSV",
    mode: "全件",
    fields: 15,
    recLen: "—",
    ui08: "无",
    ss06: "有",
    sample: "CSV 在 SS06 夹",
    note: "9.6MB，未放进サンプルデータ",
    tone: "warning",
  },
  {
    id: "040S（出荷先）",
    name: "出荷先マスタ",
    family: "master",
    dir: "MD → HUB → WMS",
    format: "UTF-8 CSV",
    mode: "全件",
    fields: 14,
    recLen: "—",
    ui08: "无",
    ss06: "有",
    sample: "879 行",
    note: "规格 UTF-8，样本实为 Shift_JIS",
    tone: "danger",
  },
  {
    id: "110S（DC入荷予定）",
    name: "DC入荷予定",
    family: "plan",
    dir: "MD → WMS",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 26,
    recLen: "512",
    ui08: "有",
    ss06: "有",
    sample: "455 行",
    note: "納品形態 30：DC総量入荷(買取)",
    tone: "neutral",
  },
  {
    id: "120S（DC出荷指示）",
    name: "DC出荷指示",
    family: "plan",
    dir: "MD → WMS",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 31,
    recLen: "512",
    ui08: "有",
    ss06: "有",
    sample: "34,981 行",
    note: "納品形態 51；含便区分",
    tone: "info",
  },
  {
    id: "130S（TC総量予定）",
    name: "TC総量入荷予定",
    family: "plan",
    dir: "MD → WMS",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 26,
    recLen: "512",
    ui08: "有",
    ss06: "有",
    sample: "94,978 行 / 4 文件",
    note: "形態 12/10；MD 布局另有 IFSHDLMS130S",
    tone: "warning",
  },
  {
    id: "140S（TC店别予定）",
    name: "TC店舗振替予定",
    family: "plan",
    dir: "MD → WMS",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 32,
    recLen: "512",
    ui08: "有",
    ss06: "有",
    sample: "359,647 行 / 4 文件",
    note: "与 130S（TC総量予定） 伝票番号成对；同日 4 批",
    tone: "info",
  },
  {
    id: "210R（DC入荷実績）",
    name: "DC入荷実績",
    family: "actual",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 23,
    recLen: "512",
    ui08: "有",
    ss06: "无",
    sample: "231 行",
    note: "欠品理由 00001 取引先責",
    tone: "neutral",
  },
  {
    id: "220R（DC出荷実績）",
    name: "DC出荷実績",
    family: "actual",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 29,
    recLen: "512",
    ui08: "有",
    ss06: "无",
    sample: "8,700 行",
    note: "欠品 00002 センター責；含移动平均单价",
    tone: "info",
  },
  {
    id: "230R（TC総量実績）",
    name: "TC総量入荷実績",
    family: "actual",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 25,
    recLen: "512",
    ui08: "有*",
    ss06: "无",
    sample: "2,798 行",
    note: "UI08 电文 ID 写成 IFSHDLMS230R",
    tone: "danger",
  },
  {
    id: "240R（TC店别実績）",
    name: "TC店舗振替実績",
    family: "actual",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 31,
    recLen: "512",
    ui08: "有",
    ss06: "无",
    sample: "16,848 行",
    note: "欠品理由 DMG12；Hitluster 启动マスタ未单列",
    tone: "warning",
  },
  {
    id: "310R（在库一览）",
    name: "在庫一覧",
    family: "stock",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "全件",
    fields: 7,
    recLen: "512",
    ui08: "无",
    ss06: "无",
    sample: "3,545 行",
    note: "PMI 已做 No.52；补足写快照口径",
    tone: "success",
  },
  {
    id: "320R（在库调整）",
    name: "在庫調整",
    family: "stock",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "差分",
    fields: 9,
    recLen: "512",
    ui08: "无",
    ss06: "无",
    sample: "13 行",
    note: "PMI 已做 No.51；调整码不在本包",
    tone: "success",
  },
  {
    id: "330R（単価）",
    name: "単価データ",
    family: "stock",
    dir: "WMS → MD",
    format: "SJIS 固定长",
    mode: "全件",
    fields: 6,
    recLen: "512",
    ui08: "无",
    ss06: "无",
    sample: "15 行",
    note: "No.154：TRIAL 不直接用此单价",
    tone: "info",
  },
];

export default function WmsHitlusterSourcePack() {
  const [filter, setFilter] = useCanvasState<Filter>("wms-pack-filter", "all");
  const rows = IFS.filter((r) => filter === "all" || r.family === filter);

  return (
    <Stack gap={24}>
      <Stack gap={8}>
        <H1>Hitluster 原始资料包</H1>
        <Text tone="secondary">
          `03. WMS (Hitluster)` · Google Drive 导出 2026-09-02 · 压缩 15MB /
          解压约 296MB · 57 个文件
        </Text>
        <Row gap={8} wrap>
          <Pill size="sm" active>
            AS-IS：TC ↔ Hitluster
          </Pill>
          <Pill size="sm">规格主体 2022</Pill>
          <Pill size="sm">样本 2026-06/07</Pill>
          <Pill size="sm">不是 Shinise To-Be</Pill>
        </Row>
      </Stack>

      <Callout tone="info" title="这包是什么">
        西友现行 WMS（Hitluster）与 MD 基干（Tomorrow Chain）之间的接口原件：15
        本 `IFSYLMS*` 的項目移送表、TC 侧 UI08 / SS06、Hitluster
        内部レイアウト，以及 2026 年 6–7 月生产样本。它描述的是「现在怎么联」，不是
        PMI 方案 3 里要接到 Shinise / 物流控制器之后的 To-Be。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="15" label="正式 IF（移送表）" />
        <Stat value="6" label="资料夹" />
        <Stat value="~53.8 万" label="样本记录（514B/行）" />
        <Stat value="1" label="旁路 IF（ASN 900S（ASN））" tone="warning" />
      </Grid>

      <H2>资料夹怎么读</H2>
      <Text>
        四层规格叠加，不要当成同一份文档的副本。項目移送表定义 WMS
        文件长什么样；SS06 定义 TC 从哪张表填进去；UI08 是 TC
        对外接口定义；テーブルレイアウト是 Hitluster 收进来之后的内部表。
      </Text>
      <Table
        headers={["夹名", "文件", "角色", "年代", "覆盖"]}
        rows={[
          [
            "項目移送表",
            "15",
            "WMS 文件布局（权威）",
            "2022-04 起，最晚 2023-08",
            "15/15 IF 齐全",
          ],
          [
            "TC IF仕様書",
            "10",
            "TC 外部 IF 定义 UI08 + DCロケーション",
            "文件日期 2026-06～08",
            "缺 020S/030S/040S（部门/仕入先/出荷先）、310R/320R/330R",
          ],
          [
            "TC_WMSテーブル項目転送仕様",
            "9",
            "SS06：TC 表 → WMS 文件",
            "规格 2022–23；文件 2026-07～08",
            "只覆盖 S（发送）8 本",
          ],
          [
            "テーブルレイアウト",
            "2",
            "Hitluster 内部レイアウト + EDI ASN",
            "2023-02 注记",
            "15 IF + IFSYLMS900S",
          ],
          [
            "サンプルデータ",
            "18+",
            "生产固定长 / CSV",
            "2026-06-17 予定、06-25 実績",
            "缺 030S（跑到 SS06 夹）",
          ],
          [
            "共有ファイル",
            "1",
            "中島：入荷と出荷(振替)予定と実績",
            "2026-07-16",
            "DC / TC総量 / TC店別 三页",
          ],
        ]}
        rowTone={["success", "warning", "info", "neutral", "info", "neutral"]}
      />
      <Text tone="tertiary" size="small">
        源：zip 目录 + 各书改版履歴。SS06 / UI08
        的内容仍是 2022 次期 WMS 转换期规格，2026
        的时间戳多半是 Drive 再导出。
      </Text>

      <Grid columns="1.2fr 1fr" gap={20}>
        <Stack gap={8}>
          <H3>解压后体积构成（MB）</H3>
          <PieChart
            donut
            size={220}
            data={[
              { label: "140S（TC店别予定） 振替予定", value: 185, tone: "info" },
              { label: "130S（TC総量予定） 総量入荷予定", value: 49 },
              { label: "120S（DC出荷指示） DC出荷指示", value: 18 },
              { label: "030S（仕入先） 取引先 CSV", value: 9.6 },
              { label: "240R（TC店别実績） 振替実績", value: 8.7 },
              { label: "其余样本 + 规格", value: 26 },
            ]}
          />
          <Text tone="tertiary" size="small">
            源：zip 未压缩字节。140S（TC店别予定）
            单日四批占了大半，是 TC 店别明细的体量特征，不是规格文件。
          </Text>
        </Stack>
        <Stack gap={8}>
          <H3>编号规律</H3>
          <Table
            headers={["段", "方向", "内容"]}
            rows={[
              ["0xxS", "基干 → WMS", "マスタ"],
              ["1xxS", "基干 → WMS", "予定 / 指示"],
              ["2xxR", "WMS → 基干", "入出荷 / 振替実績"],
              ["3xxR", "WMS → 基干", "在庫一覧 / 調整 / 単価"],
              ["900S（ASN）", "EDI → WMS", "ASN，仅レイアウト"],
            ]}
          />
          <Text size="small" tone="secondary">
            S=基干→WMS，R=WMS→基干。020S/030S/040S（部门/仕入先/出荷先） 走データ
            HUB 的 SFTP PUT，其余移送表一律写 SFTP GET。
          </Text>
        </Stack>
      </Grid>

      <H2>15 本 IF 目录（S=基干→WMS，R=WMS→基干）</H2>
      <Row gap={8} wrap>
        <Pill active={filter === "all"} onClick={() => setFilter("all")}>
          全部 15
        </Pill>
        <Pill active={filter === "master"} onClick={() => setFilter("master")}>
          マスタ 4
        </Pill>
        <Pill active={filter === "plan"} onClick={() => setFilter("plan")}>
          予定 4
        </Pill>
        <Pill active={filter === "actual"} onClick={() => setFilter("actual")}>
          実績 4
        </Pill>
        <Pill active={filter === "stock"} onClick={() => setFilter("stock")}>
          在庫 / 単価 3
        </Pill>
      </Row>
      <Table
        headers={[
          "IF",
          "名称",
          "方向",
          "格式",
          "更新",
          "项数",
          "UI08",
          "SS06",
          "样本",
          "要点",
        ]}
        rows={rows.map((r) => [
          r.id,
          r.name,
          r.dir,
          r.format,
          r.mode,
          String(r.fields),
          r.ui08,
          r.ss06,
          r.sample,
          r.note,
        ])}
        rowTone={rows.map((r) => r.tone)}
        striped
        stickyHeader
      />
      <Text tone="tertiary" size="small">
        记录长 512 字节 + CRLF（样本按 514
        字节计行）。字段数含 FILLER。UI08 / SS06
        「无」表示本包没有对应书，不代表生产没有。
      </Text>

      <H2>三条物流路径（中島 PPT）</H2>
      <Text>
        共有ファイル把 Hitluster 业务收成三页，和 PMI
        总览里的物流四模式能对上：只有 DC 路径真正管中心在库；TC
        是入荷 + 店舗振替成对，不在 DC 持有库存。
      </Text>
      <Table
        headers={["路径", "对应模式", "予定（S）", "実績（R）", "納品形態", "样本仓库例"]}
        rows={[
          [
            "入荷と出荷 (DC)",
            "ストック",
            "110S（DC入荷予定） 入荷予定 → 120S（DC出荷指示） 出荷指示",
            "210R（DC入荷実績） 入荷実績 → 220R（DC出荷実績） 出荷実績",
            "30 / 51（買取）",
            "三郷 007452、昭島 007453、007455",
          ],
          [
            "入荷と振替 (TC総量)",
            "スルー / グロス（総量）",
            "130S（TC総量予定） 総量入荷予定 → 140S（TC店别予定） 振替予定",
            "230R（TC総量実績） 総量入荷実績 → 240R（TC店别実績） 振替実績",
            "12 TC総量入荷",
            "小牧 007013、長野 007450、007451",
          ],
          [
            "入荷と振替 (TC店別)",
            "スルー（店別）",
            "仍是 130S（TC総量予定） + 140S（TC店别予定），発注已是店舗単位",
            "仍是 230R（TC総量実績） + 240R（TC店别実績）",
            "10 TC店別入荷",
            "同上；140S（TC店别予定） 行数 ≫ 130S（TC総量予定）",
          ],
        ]}
        rowTone={["info", "neutral", "neutral"]}
      />
      <Text tone="tertiary" size="small">
        源：`入荷と出荷(振替)予定と実績_中島.pptx` 3 页 +
        項目移送表納品形態区分。店直不进本包。
      </Text>

      <H2>样本：2026-06-17 予定、06-25 実績</H2>
      <BarChart
        categories={["140S（TC店别予定）", "130S（TC総量予定）", "120S（DC出荷指示）", "010S（商品）", "240R（TC店别実績）", "220R（DC出荷実績）", "310R（在库一览）"]}
        series={[
          {
            name: "样本行数（固定长按 514B 计）",
            data: [359647, 94978, 34981, 10280, 16848, 8700, 3545],
            tone: "info",
          },
        ]}
        height={220}
        valueSuffix=""
      />
      <Text tone="tertiary" size="small">
        横轴：IF 编号。纵轴：记录行数。源：zip 内 DAT
        字节数 / 514。140S（TC店别予定） 同日 05:55 / 11:04 / 12:04 / 15:02
        四批，是差分连发而不是全量重送。310R（在库一览） 与先前 Review
        会「3,545 行」一致。
      </Text>

      <H3>样本里出现的仓库</H3>
      <Table
        headers={["コード", "名称（记录抬头）", "出现 IF"]}
        rows={[
          ["007452", "三郷ＤＣ", "110S（DC入荷予定） / 220R（DC出荷実績） / 320R（在库调整） / 330R（単価）"],
          ["007453", "昭島ＤＣ", "120S（DC出荷指示）"],
          ["007455", "（DC，抬头未展开）", "210R（DC入荷実績） / 310R（在库一览）"],
          ["007013", "小牧ＴＣ", "130S（TC総量予定） / 140S（TC店别予定）"],
          ["007450", "長野ＴＣ", "130S（TC総量予定） / 140S（TC店别予定）"],
          ["007451", "（TC）", "230R（TC総量実績） / 240R（TC店别実績）"],
          ["007402", "（商品マスタ仓库码）", "010S（商品）"],
        ]}
      />

      <H2>和本工作区已做部分的关系</H2>
      <Grid columns={2} gap={16}>
        <Card>
          <CardHeader trailing={<Pill size="sm" active>已推进</Pill>}>
            310R（在库一览） / 320R（在库调整） 在庫
          </CardHeader>
          <CardBody>
            <Text>
              移送表与 06-25 样本就是先前 No.51 / No.52、日立确认会
              No.153 / 157 / 164 的底本。310R（在库一览）
              补足事项写清了快照口径：入荷未送信不加库存、出荷未送信仍算库存；用途写「月末在库金额」和「日次自动补充」。
            </Text>
            <Text size="small" tone="secondary">
              调整代码清单、22:00 vs 20
              时、TRIAL 是否另表，都不在本 zip 里。
            </Text>
          </CardBody>
        </Card>
        <Card>
          <CardHeader trailing={<Pill size="sm">尚未用这包做</Pill>}>
            予定 / 実績 8 本 + マスタ
          </CardHeader>
          <CardBody>
            <Text>
              110S–140S / 210R–240R
              才是物流控制器和伝票番号转换真正要吃的面。SS06
              显示予定来自 TC「仕入伝票(ワーク)」和「店別明細(ワーク)」——方案
              3 切到 Shinise 后，这两张表不再是源。
            </Text>
            <Text size="small" tone="secondary">
              330R（単価） 单价全件与 220R（DC出荷実績） / 310R（在库一览） / 320R（在库调整）
              上的移动平均单价，对应课题 No.154。
            </Text>
          </CardBody>
        </Card>
      </Grid>

      <H2>规格裂缝（本包内部就能看见）</H2>
      <Table
        headers={["编号", "现象", "为什么要紧"]}
        rows={[
          [
            "G1",
            "040S（出荷先） 规格 UTF-8 BOM，样本是 Shift_JIS",
            "出荷先マスタ解析会 mojibake；020S（部门） 同样无 BOM",
          ],
          [
            "G2",
            "UI08 230R（TC総量実績） 电文 ID = IFSHDLMS230R，不是 IFSYLMS230R",
            "MD 布局也残存 IFSHDLMS130S。旧 ID 与现 ID 并存",
          ],
          [
            "G3",
            "010S（商品） UI08 比移送表多「温度帯詳細区分」「仕入区分」",
            "商品マスタ项位可能已漂，不能只信一份",
          ],
          [
            "G4",
            "移送表没有処理周期 / 送信时刻",
            "所以 310R（在库一览）「22:00」不能从这包得出；样本文件名是 22:00 落地",
          ],
          [
            "G5",
            "SS06 只覆盖 S，R 没有 TC 侧映射",
            "実績是 WMS 自生成；To-Be 要另画 Shinise 怎么收",
          ],
          [
            "G6",
            "Hitluster 启动マスタ未单列 240R（TC店别実績）",
            "振替実績可能挂在 230R（TC総量実績） 同一处理链，切系统时容易漏",
          ],
          [
            "G7",
            "便区分：010S（商品） 1 桁，予定/実績 2 桁（01/02）",
            "与 Sinops 便课题同一字段家族，不能各 IF 各定默认值",
          ],
          [
            "G8",
            "ASN IFSYLMS900S 仅 EDI レイアウト",
            "取引先マスタ ASN対応区分=2 时 EDI→WMS 直送，本包无样本",
          ],
        ]}
        rowTone={[
          "danger",
          "danger",
          "warning",
          "warning",
          "info",
          "warning",
          "info",
          "neutral",
        ]}
      />

      <H2>PMI 切替时这包解决不了的事</H2>
      <Callout tone="warning" title="边界">
        方案 3 的物流控制器、Shinise 采番、未 CV / 已 CV
        混在、Hitluster 是否继续服务未 CV 店——这些决策不在原件里。原件能锁定的是：WMS
        仍按这 15 个文件说话，字段和コード体系是西友的（仓库 6 桁、仕入先
        6+3、商品 14 桁 JAN、DCロケーション 3 桁）。
      </Callout>
      <Table
        headers={["西友 Hitluster 侧（本包）", "TRIAL / PMI 要另定"]}
        rows={[
          ["伝票番号 10 桁，来自 TC 発注", "Shinise 采番 + 控制器转换（方案 1 因此否决）"],
          ["仕入先 6 桁 + 枝番 3 桁", "TRIAL 4/6 桁，登录不合并"],
          ["商品コード 14 桁（左空格填充）", "TRIAL JAN 20 桁"],
          ["出荷先 = 店舗コード 6/10 桁", "TRIAL 店番桁扩展中"],
          ["発注区分 0/2/7 ↔ WMS 01/22/21", "Shinise 売上/発注区分是否同一套"],
          ["欠品理由 00001 / 00002 / DMG12", "TRIAL 欠品・调整码是否另表（No.157）"],
          ["移動平均単価 7.2", "TRIAL 侧重算（No.154 方向已定）"],
          ["DCロケーション 3 桁 marshalling", "温度帯 / デパ，TRIAL 无对等物则要丢或映射"],
        ]}
      />

      <Divider />

      <H2>建议的下一步</H2>
      <Text>
        原件已经够做「AS-IS IF
        清单」和字段级对照，不够做工数或 To-Be 设计结论。比较有杠杆的三刀：
      </Text>
      <Table
        headers={["顺序", "动作", "用这包的哪一层"]}
        rows={[
          [
            "1",
            "把 15 本 IF 做成一张 PMI 用目录（方向、混在期归属、谁改 WMS / 谁改控制器 / 谁改 Shinise）",
            "移送表封面 + 中島三路径",
          ],
          [
            "2",
            "先拆 130S/140S/230R/240R 成对规则（伝票番号、行番号、枝番、形態 10 vs 12）",
            "移送表 + 同日四批样本",
          ],
          [
            "3",
            "マスタ 020S/030S/040S（部门/仕入先/出荷先） 的编码与 HUB 路径单独开一题，不要和固定长 010S（商品） 混谈",
            "样本 vs 规格 G1",
          ],
        ]}
        rowTone={["info", "info", "warning"]}
      />
      <Text size="small" tone="tertiary">
        规格 Excel 已解到工作区 `.wms_raw_extract/`（未解超大
        DAT）。需要某一本的逐字段对照时，直接点 IF 编号即可往下拆。
      </Text>
    </Stack>
  );
}
