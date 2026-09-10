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

type View = "catalog" | "pair" | "hub";

export default function WmsPmiIfCatalog() {
  const [view, setView] = useCanvasState<View>("wms-pmi-view", "catalog");

  return (
    <Stack gap={24}>
      <Stack gap={8}>
        <H1>Hitluster PMI 目录 · 成对规则 · マスタ HUB</H1>
        <Text tone="secondary">
          方案 3（发订 / EDI / 会计按店舗 CV 切）下的工作目录。AS-IS
          来自項目移送表 + SS06 + 2026-06-17 / 06-25 生产样本。已 CV
          列是归属假设，不是已签字的 To-Be。
        </Text>
        <Row gap={8} wrap>
          <Pill active={view === "catalog"} onClick={() => setView("catalog")}>
            15 IF 归属
          </Pill>
          <Pill active={view === "pair"} onClick={() => setView("pair")}>
            130S/140S · 230R/240R 成对
          </Pill>
          <Pill active={view === "hub"} onClick={() => setView("hub")}>
            020S/030S/040S（部门/仕入先/出荷先） HUB
          </Pill>
        </Row>
      </Stack>

      {view === "catalog" ? <CatalogView /> : null}
      {view === "pair" ? <PairView /> : null}
      {view === "hub" ? <HubView /> : null}
    </Stack>
  );
}

function CatalogView() {
  return (
    <Stack gap={20}>
      <Callout tone="info" title="怎么用这张表">
        混在期 Hitluster 文件布局按冻结看（日立 800 万是验证改修，不是重做 15
        本）。编号：S=基干→WMS，R=WMS→基干。未 CV 继续 TC 产予定、TC 吃実績。已 CV
        的伝票要经物流控制器换号后再写成同一套 `IFSYLMS*`。マスタ不走控制器，走
        HUB 或直送。WMS To-Be（是否换 Himalaya）9 月末前仍是覆盖风险。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="15" label="正式 IF" />
        <Stat value="8" label="伝票系要过控制器" />
        <Stat value="3" label="HUB CSV マスタ" />
        <Stat value="2" label="TRE 已对照（310R/320R）" tone="success" />
      </Grid>

      <H2>归属目录（S=基干→WMS，R=WMS→基干）</H2>
      <Table
        headers={[
          "IF",
          "名称",
          "路径",
          "未 CV 产 → 吃",
          "已 CV 产 → 吃",
          "日立 WMS",
          "控制器",
          "Shinise / TRMD",
          "TRE",
        ]}
        rows={[
          [
            "010S（商品）",
            "商品マスタ",
            "直送 GET",
            "TC → Hitluster",
            "TRMD 商品 → Hitluster",
            "布局不改",
            "不经过",
            "JAN/桁/阶层要对齐",
            "未开始",
          ],
          [
            "020S（部门）",
            "部門マスタ",
            "HUB PUT CSV",
            "TC → HUB → WMS",
            "谁喂 HUB 未决",
            "已吃荷主 0000",
            "不经过",
            "分类桁未决",
            "未开始",
          ],
          [
            "030S（仕入先）",
            "取引先マスタ",
            "HUB PUT CSV",
            "TC → HUB → WMS",
            "谁喂 HUB 未决",
            "已吃 9 桁",
            "不经过",
            "4/6 桁对 6+3",
            "未开始",
          ],
          [
            "040S（出荷先）",
            "出荷先マスタ",
            "HUB PUT CSV",
            "TC → HUB → WMS",
            "已 CV 店也要在 WMS 有出荷先",
            "店码 6 桁",
            "不经过",
            "店番桁扩展",
            "未开始",
          ],
          [
            "110S（DC入荷予定）",
            "DC入荷予定",
            "ストック",
            "TC 仕入ワーク → WMS",
            "Shinise → 控制器 → WMS",
            "不改",
            "采番 + 形态 30",
            "仕入/発注源",
            "未开始",
          ],
          [
            "120S（DC出荷指示）",
            "DC出荷指示",
            "ストック店别",
            "TC 店别明细 → WMS",
            "同一 DC 可混未 CV / 已 CV 店",
            "不改",
            "按店切 + 采番",
            "已 CV 行的源",
            "未开始",
          ],
          [
            "130S（TC総量予定）",
            "TC総量入荷予定",
            "スルー/グロス",
            "TC 仕入ヘッダ/明細",
            "Shinise → 控制器",
            "不改",
            "必须与 140S（TC店别予定） 同号成对",
            "仕入伝票源",
            "未开始",
          ],
          [
            "140S（TC店别予定）",
            "TC店舗振替予定",
            "スルー/グロス",
            "TC 店别明细",
            "已 CV 店行走 TRIAL 采番",
            "不改",
            "必须与 130S（TC総量予定） 同号成对",
            "店别数量源",
            "未开始",
          ],
          [
            "210R（DC入荷実績）",
            "DC入荷実績",
            "ストック",
            "WMS → TC → GLOVIA",
            "WMS → 控制器 → Shinise → Biz",
            "输出不改",
            "回写新伝票号",
            "收実績",
            "未开始",
          ],
          [
            "220R（DC出荷実績）",
            "DC出荷実績",
            "ストック店别",
            "WMS → TC",
            "按店拆给 TC 或 Shinise",
            "输出不改",
            "按店切 + 回写号",
            "不直接用移动平均单价",
            "未开始",
          ],
          [
            "230R（TC総量実績）",
            "TC総量入荷実績",
            "スルー/グロス",
            "WMS → TC",
            "WMS → 控制器 → Shinise",
            "输出不改",
            "与 240R（TC店别実績） 成对回写",
            "收実績",
            "未开始",
          ],
          [
            "240R（TC店别実績）",
            "TC店舗振替実績",
            "スルー/グロス",
            "WMS → TC",
            "按店拆",
            "启动マスタ未单列",
            "与 230R（TC総量実績） 成对回写",
            "欠品码 DMG12",
            "未开始",
          ],
          [
            "310R（在库一览）",
            "在庫一覧",
            "DC ストック",
            "WMS → TC",
            "WMS → TRMD；同一 DC 两套账未决",
            "快照口径已写",
            "不经过（无伝票号）",
            "日次补充 / 月末评价",
            "已对照 No.52",
          ],
          [
            "320R（在库调整）",
            "在庫調整",
            "DC ストック",
            "WMS → TC",
            "WMS → TRMD",
            "差分；调整码另表",
            "不经过",
            "是否另表未决 No.157",
            "已对照 No.51",
          ],
          [
            "330R（単価）",
            "単価",
            "DC ストック",
            "WMS → TC",
            "TRIAL 不直接采用",
            "全件",
            "不经过",
            "No.154 自算",
            "方向已定",
          ],
        ]}
        rowTone={[
          "warning",
          "warning",
          "warning",
          "warning",
          "neutral",
          "danger",
          "info",
          "info",
          "neutral",
          "danger",
          "info",
          "info",
          "success",
          "success",
          "neutral",
        ]}
        striped
        stickyHeader
      />
      <Text tone="tertiary" size="small">
        红行：同一 DC / 同一文件里会按店切开，混在期最容易漏。蓝行：130S/140S
        与 230R/240R 必须成对换号，样本已 100%
        验证。绿行：本工作区已做映射，详细设计仍有 No.153/156/160/161。
      </Text>

      <H3>控制器真正要动的面</H3>
      <Table
        headers={["对象", "为什么必须进控制器", "不能只改一边"]}
        rows={[
          [
            "110S（DC入荷予定） / 210R（DC入荷実績）",
            "DC 入荷予定与実績用同一 10 桁 伝票番号",
            "只改予定不改実績，TC/Shinise 对不上入荷",
          ],
          [
            "120S（DC出荷指示） / 220R（DC出荷実績）",
            "店别出荷指示与実績；便区分 01/02 在文件里",
            "已 CV 店行会计走 Biz，未 CV 走 GLOVIA",
          ],
          [
            "130S（TC総量予定） / 140S（TC店别予定） 成对",
            "ヘッダ行与店别枝番共用 伝票番号 + 行番号",
            "只换 130S（TC総量予定） 不换 140S（TC店别予定），WMS 振替会断",
          ],
          [
            "230R（TC総量実績） / 240R（TC店别実績） 成对",
            "実績回写同一对键，確定数量店别加总 = 総量",
            "只回收 230R（TC総量実績），店别欠品 DMG12 进不了已 CV 账",
          ],
        ]}
        rowTone={["neutral", "danger", "info", "info"]}
      />

      <H3>明确不进控制器的</H3>
      <Text>
        010S（商品） 直送、020S/030S/040S（部门/仕入先/出荷先） 走データ HUB、310R/320R/330R
        无伝票号。这些要另开主数据 / 在库课题，不要塞进「采番转换」同一张工单。
      </Text>

      <Callout tone="warning" title="9 月末前仍覆盖本目录的两件事">
        ① WMS To-Be：Hitluster 继续还是切 Himalaya。② 同一 DC 同时服务未 CV /
        已 CV 店时，120S（DC出荷指示） / 220R（DC出荷実績） / 140S（TC店别予定） / 240R（TC店别実績）
        是按店拆给两套 MD，还是 WMS 仍只认西友店码、由控制器翻译。目录里「已 CV
        产→吃」两列随这两项会整列改写。
      </Callout>
    </Stack>
  );
}

function PairView() {
  return (
    <Stack gap={20}>
      <Callout tone="success" title="样本结论：成对键 100% 成立">
        配对键是 検品倉庫コード + 伝票番号(10) +
        伝票行番号(3)。2026-06-17 四批 130S/140S（94,978 対 359,647
        行）与 06-25 007451 的 230R/240R（2,798 対 16,848
        行）全部命中、无孤儿、无键重复。入力連番在父子行上完全相同，可当同批校验。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="100%" label="130S（TC総量予定） 键被 140S（TC店别予定） 覆盖" tone="success" />
        <Stat value="100%" label="140S（TC店别予定） 行能回到 130S（TC総量予定）" tone="success" />
        <Stat value="0" label="伝票+行 跨仓碰撞（晨批）" />
        <Stat value="53" label="单行最多店数（枝番）" />
      </Grid>

      <H2>成对规则（可写进详细设计）</H2>
      <Table
        headers={["规则", "内容", "样本证据"]}
        rows={[
          [
            "P1 键",
            "倉庫 + 伝票番号 + 伝票行番号。140S/240R 再加 伝票行枝番（店别 001…n）",
            "四批 0 孤儿；枝番无重复、恰为 001..n",
          ],
          [
            "P2 同批",
            "同一 入力連番 贯穿 130S（TC総量予定） 行与其全部 140S（TC店别予定） 子行（実績同理）",
            "seq_diff = 0",
          ],
          [
            "P3 バラ加总",
            "SUM(140S（TC店别予定）.出荷数量バラ) = 130S（TC総量予定）.発注数量バラ；実績 SUM(確定数量) 同样成立",
            "qty_equal = 全部键；差额 0",
          ],
          [
            "P4 単位数不加总",
            "発注単位数 / 出荷単位数是店别行自己的包装数。多店时 SUM(子) ≠ 父",
            "晨批 1 店键 39,894 単位相等；多店 37,689 不相等",
          ],
          [
            "P5 形态不是 1:1",
            "10＝仕入先「発注/ピッキング」，12＝「発注のみ」。两种都会拆成多店 140S（TC店别予定） 行",
            "晨批 10 的 55.8% 才是 1 店；12 仅 35.1% 是 1 店",
          ],
          [
            "P6 同号字段",
            "JAN、取引先、便区分、納品形態 在父子行必须一致",
            "jan/vend/keitai mismatch = 0",
          ],
          [
            "P7 文件形态",
            "予定：一文件多仓。実績：文件名带 6 桁仓库码，一仓一文件",
            "130S（TC総量予定） 晨批 8+ 仓；230R（TC総量実績） 文件名 007451",
          ],
        ]}
        rowTone={[
          "success",
          "info",
          "success",
          "warning",
          "danger",
          "success",
          "neutral",
        ]}
      />
      <Text tone="tertiary" size="small">
        源：項目移送表「TC総量入荷予定と対になる」+ SS06
        仕入伝票番号/伝票行/納品方法別枝番 + zip 内 DAT（512 字节 + CRLF，按
        Shift_JIS 字节切，不能按 Unicode 下标切）。
      </Text>

      <H3>四批 130S（TC総量予定） → 140S（TC店别予定） 规模（2026-06-17）</H3>
      <BarChart
        categories={["05:55", "11:04", "12:04", "15:02"]}
        series={[
          { name: "130S（TC総量予定） 父行（商品×仓）", data: [77583, 5306, 9809, 2280] },
          { name: "140S（TC店别予定） 子行（店别枝番）", data: [211057, 68434, 48719, 31437] },
        ]}
        height={220}
      />
      <Text tone="tertiary" size="small">
        横轴：同日四批文件时刻。纵轴：记录行数。140S（TC店别予定） / 130S（TC総量予定） ≈ 2.7、12.9、5.0、13.8
        倍，差分连发不是全量重送。源：DAT 字节 / 514。
      </Text>

      <Grid columns="1fr 1fr" gap={16}>
        <Stack gap={8}>
          <H3>納品形態 10 vs 12（晨批 05:55）</H3>
          <PieChart
            donut
            size={200}
            data={[
              { label: "10 発注/ピッキング 父行", value: 61334, tone: "info" },
              { label: "12 発注のみ 父行", value: 16249, tone: "neutral" },
            ]}
          />
          <Text tone="tertiary" size="small">
            源：SS06 仕入先.発注/ピッキングメッセージ対象区分 → 10 / 12。不要把
            10 设计成「一店一行」：该批 10 仍有最多 50 店。
          </Text>
        </Stack>
        <Stack gap={8}>
          <H3>実績 007451（06-25 13:04）</H3>
          <Table
            headers={["项", "值"]}
            rows={[
              ["230R（TC総量実績） 父行", "2,798（10：1,193 / 12：1,605）"],
              ["240R（TC店别実績） 子行", "16,848（10：9,244 / 12：7,604）"],
              ["键命中", "100%，seq 相同"],
              ["確定数量加总", "2,798 / 2,798 相等"],
              ["単位数加总", "仅 680 个 1 店键相等（= store hist 1）"],
              ["欠品理由", "全部 00000（本文件无欠品）"],
              ["最大店数", "20"],
            ]}
          />
        </Stack>
      </Grid>

      <H3>控制器换号时必须一起改的字段</H3>
      <Table
        headers={["字段", "130S（TC総量予定） / 230R（TC総量実績）", "140S（TC店别予定） / 240R（TC店别実績）", "换号约束"]}
        rows={[
          ["伝票番号", "10 桁，来自仕入伝票番号 13→10", "同一值", "父子必须同一新号"],
          ["伝票行番号", "商品别 3 桁", "同一值", "不要重排行"],
          ["伝票行枝番", "无", "店别 001…n", "不要重排枝番；店码另译"],
          ["入力連番", "8 桁", "同一值", "可保留作同批校验"],
          ["出荷先コード", "无", "6 桁店码", "已 CV 店要译 TRIAL 店番"],
          ["発注数量バラ", "父行合计", "子行，加总还原父行", "不要改数量关系"],
          ["発注単位数", "父行包装", "子行各自包装，不可加总", "跟 9/1 入数说明一致"],
        ]}
      />

      <Card>
        <CardHeader>样本对照一行</CardHeader>
        <CardBody>
          <Text>
            15:02 批：仓 007013 小牧、伝票 0008462934 行 001、JAN 4901820444806、形态
            10、便 01、連番 00157169。130S（TC総量予定） バラ 32。140S（TC店别予定） 枝番 015 店 005282 バラ
            2（该父行共 16 店加总为 32）。実績侧同构：230R（TC総量実績） バラ = Σ 240R（TC店别実績） 確定数量。
          </Text>
        </CardBody>
      </Card>
    </Stack>
  );
}

function HubView() {
  return (
    <Stack gap={20}>
      <Callout tone="warning" title="020S（部门） / 030S（仕入先） / 040S（出荷先） 不要和 010S（商品） 放在同一条设计里">
        010S（商品） 是 Shift_JIS 固定长、SFTP GET、差分、带仓库码。020S/030S/040S（部门/仕入先/出荷先） 是 UTF-8
        CSV、データ HUB PUT、全件、无仓库码。2022-12
        改版才从 DAT/GET/SJIS 切到 CSV/PUT/UTF-8。混谈会把编码、荷主码和更新区分三条规则一起写错。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="无 BOM" label="三份样本 vs 规格「BOM付」" tone="warning" />
        <Stat value="0000 / 00" label="荷主 / 企業（WMS 侧）" />
        <Stat value="02" label="全件文件里每行的更新区分" />
        <Stat value="62,585" label="030S（仕入先） 取引先行" />
      </Grid>

      <H2>规格写的 vs 样本落地</H2>
      <Table
        headers={["项", "移送表 / SS06", "020S（部门） 部门 4,052 行", "030S（仕入先） 取引先 62,585 行", "040S（出荷先） 出荷先 879 行"]}
        rows={[
          ["编码", "UTF-8 BOM 付", "UTF-8 无 BOM", "UTF-8 无 BOM", "UTF-8 无 BOM"],
          ["路径", "MD → HUB → WMS PUT", "文件名 IFSYLMS020S_*.csv", "放在 SS06 夹，不在サンプル", "IFSYLMS040S_*.csv"],
          ["更新区分", "00/02/03/04 按时间戳", "全部 02 訂正", "全部 02", "全部 02"],
          ["荷主 / 企業", "SS06 写死 'SEIYU'", "0000 / 00", "0000 / 00", "0000 / 00"],
          ["列数", "15 / 15 / 14（CSV 不垫 住所3/4）", "15，大分类/细分类空", "15，GLN 常空、电话可为 0", "14，英字空 213 行"],
          ["代码宽度", "部门 2/2/10/2/4/10；取引先 9；店 10", "与规格列对齐", "取引先一律 9 桁", "店码一律 6 桁，无空格垫到 10"],
        ]}
        striped
      />
      <Text tone="tertiary" size="small">
        源：2026-07-09 040S（出荷先）、07-16 020S（部门）、07-17 030S（仕入先）。前一轮把 040S（出荷先） 判成
        Shift_JIS 是错的：整文件 UTF-8 替换符为 0，CP932 无法解码。
      </Text>

      <H3>HUB 实际做了什么（从两侧对得上的部分）</H3>
      <Table
        headers={["层", "荷主コード", "含义"]}
        rows={[
          ["SS06（TC 出）", "区分管理マスタ info1 = 'SEIYU'", "西友业务码"],
          ["Hitluster 启动マスタ SYS_EXIFM101", "OwnerCode 0000", "WMS 荷主"],
          ["CSV 样本", "0000 / 企業 00", "已经是 WMS 收件形态"],
        ]}
        rowTone={["neutral", "info", "success"]}
      />
      <Text>
        样本是 HUB 之后、WMS 入口的样子，不是 TC
        出口的样子。所以「SEIYU → 0000」是 HUB 的职责，不是 TRE 在 IF
        里再转一次。已 CV 之后要先定：HUB 是否仍在、谁给它喂部门/取引先/店。
      </Text>

      <H2>和 010S（商品） 的隔离清单</H2>
      <Table
        headers={["", "010S（商品） 商品", "020S（部门） / 030S（仕入先） / 040S（出荷先）"]}
        rows={[
          ["协议", "SFTP GET", "SFTP PUT（HUB → WMS）"],
          ["格式", "SJIS 固定长 512", "UTF-8 CSV，无引号、无定宽空格"],
          ["更新", "差分 00/02/03", "全件文件，行上却标 02"],
          ["粒度", "按倉庫コード（样本 007402）", "无仓库，全企业一份"],
          ["控制器", "不经过", "不经过"],
          ["PMI 风险", "JAN 14 vs 20、阶层、ITF、便 1 桁", "编码 BOM、荷主 0000、店 6 vs 桁扩展、仕入先 9 vs 4/6"],
        ]}
      />

      <Callout tone="danger" title="混在期不要假设「全件 + 更新区分」能当增量">
        三份全件样本每一行都是 02（訂正），没有 00 新規、没有 04
        无变更。WMS 若按补足事项用时间戳算更新区分，那是 TC/HUB
        生成逻辑；落地文件不能用来推断增量协议。详细设计要单独问 HUB：全件覆盖，还是真差分。
      </Callout>

      <H3>建议的单独课题（不要并进 310R/320R）</H3>
      <Table
        headers={["课题", "要拍板的话", "卡住谁"]}
        rows={[
          [
            "H1 编码",
            "WMS 入口是否接受无 BOM UTF-8（样本如此）还是必须 BOM（规格如此）",
            "020S/030S/040S（部门/仕入先/出荷先） 解析",
          ],
          [
            "H2 荷主",
            "已 CV 后 HUB 是否继续把任何来源都写成 0000/00",
            "010S（商品） 以外的マスタ",
          ],
          [
            "H3 谁喂 HUB",
            "未 CV 仍 TC；已 CV 是 Shinise 直连 HUB、还是 TRMD→HUB、还是两路合并",
            "部门 / 取引先 / 店",
          ],
          [
            "H4 店码",
            "040S（出荷先） 样本 6 桁；店番桁扩展后 CSV 是 6、10 还是变长",
            "140S/240R 出荷先、120S/220R",
          ],
          [
            "H5 仕入先",
            "030S（仕入先） 9 桁（6+枝番3）对 TRIAL 4/6 桁；ASN 区分 0/1/2 是否仍由这张マスタ驱动 900S（ASN）",
            "EDI ASN、欠品责任方",
          ],
        ]}
        rowTone={["warning", "info", "danger", "warning", "info"]}
      />
    </Stack>
  );
}
