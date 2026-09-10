import {
  Callout,
  Card,
  CardBody,
  CardHeader,
  Grid,
  H1,
  H2,
  Pill,
  Row,
  Stack,
  Stat,
  Table,
  Text,
  useCanvasState,
} from "cursor/canvas";

type View = "overview" | "catalog" | "conflict" | "quality";

export default function WmsTrialMappingReview() {
  const [view, setView] = useCanvasState<View>("wms-map-view", "overview");

  return (
    <Stack gap={24}>
      <Stack gap={8}>
        <H1>Hitluster × TRIAL mapping 审查</H1>
        <Text tone="secondary">
          来源：mapping_03. WMS (Hitluster) 2026-09-02 zip（53
          文件，约 7 月 8 日–8 月 25 日）。对照对象：項目移送表 15 IF、SS06
          納品形態、2026-06-17 / 06-25 生产样本成对、HUB CSV 实态。这是 Phase 1
          对照稿，不是已签字的 To-Be。
        </Text>
        <Row gap={8} wrap>
          <Pill active={view === "overview"} onClick={() => setView("overview")}>
            总览
          </Pill>
          <Pill active={view === "catalog"} onClick={() => setView("catalog")}>
            15 IF 对照
          </Pill>
          <Pill active={view === "conflict"} onClick={() => setView("conflict")}>
            与样本 / 方案3 冲突
          </Pill>
          <Pill active={view === "quality"} onClick={() => setView("quality")}>
            稿件质量
          </Pill>
        </Row>
      </Stack>

      {view === "overview" ? <OverviewView /> : null}
      {view === "catalog" ? <CatalogView /> : null}
      {view === "conflict" ? <ConflictView /> : null}
      {view === "quality" ? <QualityView /> : null}
    </Stack>
  );
}

function OverviewView() {
  return (
    <Stack gap={20}>
      <Callout tone="warning" title="先读这一句">
        15 本 IF 都有文件夹，但只有商品 / 部門 / 入荷実績 / 在库
        四块真正对着 TRIAL 对象写完了。仕入先・店マスタ是空模板；予定 4
        本是字段清单而不是 GAP；240R（TC店别実績） 被同时映射成仕入伝票和振替伝票。物流控制器换号、130S/140S
        成对键、単位数不加算，这套 mapping 里都没有。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="15" label="IF 文件夹（含 330R（単価））" />
        <Stat value="4" label="可当对照用的块" tone="success" />
        <Stat value="2" label="空模板（030S（仕入先） / 040S（出荷先））" tone="danger" />
        <Stat value="★最高 ×8" label="未决：伝票区分 / 单价" tone="warning" />
      </Grid>

      <H2>这套资料在画什么</H2>
      <Table
        headers={["方向", "Hitluster", "mapping 认定的 TRIAL 对象", "方案3 还缺什么"]}
        rows={[
          [
            "MD→WMS マスタ",
            "010S（商品） / 020S（部门） / 030S（仕入先） / 040S（出荷先）",
            "F053+F023、分类 6 表 JOIN、MST_TORIHIKI（空）、店マスタ（空）",
            "HUB 谁喂、店码 6 桁、仕入先 9 桁",
          ],
          [
            "MD→WMS 予定",
            "110S（DC入荷予定） / 120S（DC出荷指示） / 130S（TC総量予定） / 140S（TC店别予定）",
            "発注伝票；120S/140S 另用 発注振分伝票",
            "控制器采番；130S（TC総量予定） 与 140S（TC店别予定） 必须同号成对",
          ],
          [
            "WMS→MD 実績",
            "210R（DC入荷実績） / 230R（TC総量実績） / 220R（DC出荷実績） / 240R（TC店别実績）",
            "仕入伝票 TMDLI0110；振替伝票；240R（TC店别実績） 两份都写",
            "1 文件是否拆成 2 种伝票；混在期按店切开",
          ],
          [
            "WMS→MD 在库",
            "310R（在库一览） / 320R（在库调整） / 330R（単価）",
            "center_stock（v2）；330R（単価） 只有調査",
            "仓库 007452→INT4 仍是 ???；330R（単価） 不进仕入",
          ],
        ]}
      />

      <Grid columns={2} gap={16}>
        <Card>
          <CardHeader>写完、可以对的部分</CardHeader>
          <CardBody>
            <Stack gap={8}>
              <Text>
                010S（商品）：F053 / F023 / PRODUCTS → 47 项，10 条真 GAP（阶层、DC
                ロケ、通過在庫、取引先桁）。
              </Text>
              <Text>
                020S（部门）：6 表非正规化 → CSV 15 项；荷主 0000 / 企业 00
                与生产样本一致。更新区分要 HUB 记上次送信时刻。
              </Text>
              <Text>
                210R（DC入荷実績） / 230R（TC総量実績）：对着 仕入伝票 35 项。判定サマリ自称 ○ 仅
                6%、△ 46%、✕ 3%、― 43%。单价和伝票区分标 ★最高。
              </Text>
              <Text>
                310R（在库一览） / 320R（在库调整） v2：目标表从 current_stock_data 改成
                center_stock，jan→product_code(13)，丢掉 amount。
              </Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>还不能当设计输入的部分</CardHeader>
          <CardBody>
            <Stack gap={8}>
              <Text>
                030S（仕入先） / 040S（出荷先）：GAP 和変換ルール标题仍是「仕入先リマスタ」，表体空白。
              </Text>
              <Text>
                110S/120S/130S/140S 的 GAP 几乎是「处理日、連番、仓库码…」逐字段备忘，优先度全是 ★中，没有成对 / 入数规则。
              </Text>
              <Text>
                業務ルール合意只有一条 leftover：sinops 的 item_mst.txt（IF
                63/64）。論点整理空。WMS 专用合意未填。
              </Text>
              <Text>
                コード変換テーブル混了 sinops 店码（001987→785）和 Hitluster
                形态码；仓库 007452 对照表仍是空。
              </Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <H2>伝票パターン（写在 230R（TC総量実績） GAP 里）</H2>
      <Text tone="secondary">
        这是 mapping 对「一个 WMS 文件进 MD 后变成哪种伝票」的主张。①③
        并没有全部落地。
      </Text>
      <Table
        headers={["WMS 文件", "主张生成的 MD 伝票", "①③ 实际写了什么"]}
        rows={[
          ["240R（TC店别実績） TC店舗振替実績", "TC店别仕入 ＋ TC総量振替", "两张表都写了"],
          ["230R（TC総量実績） TC総量入荷実績", "センター仕入 ＋ 店舗振替", "只写了仕入 TMDLI0110"],
          ["210R（DC入荷実績） DC入荷実績", "センター仕入", "仕入 TMDLI0110"],
          ["220R（DC出荷実績） DC出荷実績", "店舗振替", "振替伝票"],
          ["センター間 A→B", "A 振替、B 无票 / 或 A→B 再 B→店", "无对应 IF 文件夹"],
        ]}
        rowTone={["warning", "danger", "success", "success", "warning"]}
      />

      <Callout tone="info" title="和先前 PMI 目录怎么叠">
        未 CV：TC 继续产予定、吃実績。已
        CV：Shinise 发订经物流控制器换号后再写成同一套 IFSYLMS*。mapping
        默认「伝票番号直接映射」，等于还没把控制器放进转换规则。
      </Callout>
    </Stack>
  );
}

function CatalogView() {
  return (
    <Stack gap={20}>
      <H2>15 IF：成熟度 × TRIAL 对象</H2>
      <Table
        headers={[
          "IF",
          "①②③",
          "成熟度",
          "TRIAL 源 / 宿",
          "mapping 主张的关键转换",
        ]}
        rows={[
          [
            "010S（商品） 商品",
            "②③ 有（① 3.4MB 未拆）",
            "可对照",
            "F053 PRODUCTS + F023 店别調達 / DIFBRANCHPURCHASEINFO",
            "阶层 4→6；納品方法→経路/通過在庫；更新区分 00/02/03",
          ],
          [
            "020S（部门） 部門",
            "②③ 有",
            "可对照",
            "F032 DIV / F022 部門 / F068 ミニ / F078 品種 / F038 品目",
            "荷主 0000、企业 00；大分类/细分类空格固定；HUB 记前回送信",
          ],
          [
            "030S（仕入先） 仕入先",
            "②③ 空模板",
            "不可用",
            "标题写 MST_TORIHIKI，表体空",
            "无",
          ],
          [
            "040S（出荷先） 出荷先",
            "②③ 空，标题仍是仕入先",
            "不可用",
            "未指定店マスタ字段",
            "无",
          ],
          [
            "110S（DC入荷予定） DC入荷予定",
            "②③ 字段清单",
            "备忘",
            "発注伝票 + 店别調達センターコード",
            "ストック→30；伝票番号前补零 10 桁",
          ],
          [
            "120S（DC出荷指示） DC出荷指示",
            "②③ 字段清单",
            "备忘",
            "発注振分伝票（振替元センター）",
            "物流タイプ→納品形態 51 等",
          ],
          [
            "130S（TC総量予定） TC総量予定",
            "①②③ 有，① 含 06-17 样本",
            "备忘+",
            "発注伝票",
            "物流タイプ 2-TC総量；副标题误写成 140S（TC店别予定）",
          ],
          [
            "140S（TC店别予定） TC店别予定",
            "②③ 双 sheet",
            "备忘+",
            "発注振分（総量 sheet）和 発注伝票（店别 sheet）",
            "1-TC店别→形態 10；伝票行枝番「001？」",
          ],
          [
            "210R（DC入荷実績） DC入荷実績",
            "①②③ → TMDLI0110",
            "可对照",
            "仕入伝票 35 项",
            "数量×1000；伝票区分固定值未决；单价从商品マスタ补",
          ],
          [
            "220R（DC出荷実績） DC出荷実績",
            "②③ → 振替伝票",
            "可对照",
            "振替伝票",
            "形態 51 为起点拆物流タイプ；单价 ★最高",
          ],
          [
            "230R（TC総量実績） TC総量実績",
            "①②③ → TMDLI0110",
            "可对照",
            "仕入伝票；パターン还想再出振替",
            "形態 12→納品形式 1 総量；判定サマリ ○ 仅 2 项",
          ],
          [
            "240R（TC店别実績） TC店别実績",
            "①②③ 双目标",
            "主张过宽",
            "仕入伝票（店别）+ 振替伝票（総量）",
            "出荷先→仕入.店舗；管轄倉庫→振替元",
          ],
          [
            "310R（在库一览） 在库一览",
            "v1+v2 各 ①②③",
            "可对照（用 v2）",
            "center_stock 全件快照",
            "仓库码→store_code ???；忽略移动平均单价",
          ],
          [
            "320R（在库调整） 在库调整",
            "v1+v2 各 ①②③",
            "可对照（用 v2）",
            "center_stock UPSERT",
            "符号 01/02 ±数量；调整码 ADJ01 无列",
          ],
          [
            "330R（単価） 単価",
            "調査 .xls only",
            "结论：不直接用",
            "对照 TRIAL 総平均法 / F006 月别単品会计在库",
            "センター现用総平均；WMS 送移动平均。与 No.154 一致",
          ],
        ]}
        rowTone={[
          "success",
          "success",
          "danger",
          "danger",
          "warning",
          "warning",
          "info",
          "warning",
          "success",
          "info",
          "success",
          "danger",
          "success",
          "success",
          "info",
        ]}
      />

      <H2>実績侧反复出现的 ★最高 GAP</H2>
      <Table
        headers={["GAP", "出现在", "mapping 的对应方针", "对 PMI 的含义"]}
        rows={[
          [
            "伝票区分 / 采番未决",
            "210R（DC入荷実績） / 220R（DC出荷実績） / 230R（TC総量実績） / 240R（TC店别実績）",
            "10 仕入 / 11 本部仕入 / 30 振出 / 31 振入，要西友/MD 定固定值",
            "这就是物流控制器工作单的核心，不能「直接映射伝票番号」",
          ],
          [
            "原単価・売単価缺失",
            "同上四本",
            "从 DIFBRANCHPURCHASEINFO 发注原価、DIFBRANCHSALESINFO 基准売価 ×1000 补，金额=単価×確定数量",
            "不要用 330R（単価） 填仕入。会计正确性取决于マスタ时点，不取决于 WMS",
          ],
          [
            "数量 1000 倍 / 行番号桁",
            "同上",
            "WMS バラ確定数量 ×1000 进 MD",
            "与 9/1 入数 standup 是另一层：单位数仍不能对 130S（TC総量予定） 做 SUM",
          ],
          [
            "仓库 6 桁 → 店 INT4",
            "310R（在库一览） / 320R（在库调整） QA-G01",
            "STORE_MAPPING = {\"007452\": ???, \"007455\": ???}",
            "コード変換里的店舗シート是 sinops 001987，对不上 Hitluster 仓库",
          ],
        ]}
      />

      <Text tone="secondary">
        010S（商品） 独有的高优先级：商品分类粒度、DCロケーション未取得、通過在庫判定。020S（部门）
        独有：8 层对 6 层、4 桁对 2 桁、「04 変更なし」需要前回連携日時。
      </Text>
    </Stack>
  );
}

function ConflictView() {
  return (
    <Stack gap={20}>
      <Callout tone="danger" title="不能直接按 mapping 开工的五条">
        納品形態 10/12 的语义、240R（TC店别実績） 一份两票、伝票号直通、130S/140S
        不成对、030S/040S 空白。这五条任一落地都会让混在期串票或 WMS 拒文件。
      </Callout>

      <H2>与样本 / SS06 / 方案3 的冲突</H2>
      <Table
        headers={["#", "mapping 怎么写", "资料 + 样本怎么证", "建议"]}
        rows={[
          [
            "C1",
            "コード変換「納品形式」：12=TC総量入荷→1 総量，10=TC店别入荷→2 店别。140S（TC店别予定） GAP：1-TC店别→10",
            "SS06：10=仕入先「発注/ピッキング」，12=「発注のみ」。06-17 晨批：形態 10 仅 55.8% 单店（最多 50 店），12 仅 35.1% 单店（最多 53 店）。10/12 不是 1 店 vs N 店",
            "停用这张对照。形態是作业指示，店别明细在 140S/240R 行上",
          ],
          [
            "C2",
            "240R（TC店别実績） 同时生成仕入伝票（店别）和振替伝票（総量）。230R（TC総量実績） パターン还想再出店舗振替",
            "240R（TC店别実績） 是店别行（検品倉庫+伝票番号+行番号+枝番）。230R（TC総量実績） 是総量行，没有出荷先。06-25 007451：230R/240R 100% 成对，確定数量加算",
            "一票一文件。総量入荷→仕入用 230R（TC総量実績）；店别着荷→仕入或振替用 240R（TC店别実績），先定一种，不要从 230R（TC総量実績） 爆炸出店行",
          ],
          [
            "C3",
            "130S/140S/210R/230R 変換：伝票番号「直接マッピング可」、前补零 10 桁",
            "TC 仕入伝票番号 VARCHAR2(13) 截 10 位。方案3：已 CV 店必须经物流控制器换号；130S（TC総量予定） 与 140S（TC店别予定） 必须同号",
            "S 侧两本一起换，R 侧按换后的号吃。mapping 的「直接」只适用于未 CV",
          ],
          [
            "C4",
            "140S（TC店别予定） 行枝番写「001？」。単位数只写前补零 7 桁、小数截断",
            "成对键=検品倉庫+伝票番号(10)+行番号(3)，枝番 001…n。SUM(140S（TC店别予定）.出荷数量バラ)=130S（TC総量予定）.発注数量バラ。発注単位数在 N 店>1 时不能加总",
            "把成对键和「単位数不加总」写入 130S/140S/230R/240R 的 ③，不要只做桁处理",
          ],
          [
            "C5",
            "110S（DC入荷予定） 納品形態：0 ストック→30，1 スルー→…；120S（DC出荷指示）：0-DC→51 DC买取出荷店别",
            "三条物流路径：DC 用 110S/120S↔210R/220R（三郷 007452、昭岛 007453）；TC 用 130S/140S↔230R/240R（小牧 007013 等）。形态码是路径开关，不是「店直」那张 1/2 表",
            "按路径分文件，不要把 010S（商品） 的納品経路 01/02 套到予定文件的 2 桁形態",
          ],
          [
            "C6",
            "020S（部门） 更新区分 00/02/03/04，要 HUB 持前回成功时刻。030S/040S 未写",
            "生产 CSV：UTF-8 无 BOM（规格写 BOM 付）；三文件全行 更新区分=02；荷主 0000 / 企业 00——020S（部门） ③ 的固定值与样本一致",
            "030S/040S 按 020S（部门） 同一 HUB 契约补完。BOM 与「全件却全是 02」要业务确认，不能当已决",
          ],
          [
            "C7",
            "仕入数量 = WMS 確定バラ ×1000。単位数小数截断",
            "9/1 standup：西友入数是店别包装回数，TRIAL 入数是箱入数。mapping 的 ×1000 是 MD 内部小数，解决不了入数语义",
            "バラ走数量；入数另开对照。截断単位数会在不定贯/拆零上丢数",
          ],
          [
            "C8",
            "310R/320R v2 丢弃移动平均単価。仕入单价从商品マスタ补",
            "330R（単価） 調査：TRIAL センター用総平均法；WMS 送抽出時点の移动平均。样本 007452 単価如 2299.92",
            "与 Hitachi QA No.154 一致：330R（単価） 不作为仕入原価。不要两套单价并行进会计",
          ],
        ]}
        rowTone={[
          "danger",
          "danger",
          "danger",
          "warning",
          "warning",
          "info",
          "warning",
          "success",
        ]}
      />

      <H2>140S（TC店别予定） 的双源模型</H2>
      <Text>
        ③ 给 140S（TC店别予定） 准备了两张转换表：総量 sheet 从 発注振分伝票（振替元センター）出，店别
        sheet 从 発注伝票（纳品先店舗）出。这是在用 TRIAL
        的两张票去近似 Hitluster 的「親 130S（TC総量予定） + 子 140S（TC店别予定）」。样本里亲子是同一入力連番、同一伝票番号，形态码也相同——不是两套源。
      </Text>
    </Stack>
  );
}

function QualityView() {
  return (
    <Stack gap={20}>
      <H2>模板污染（整包重复出现）</H2>
      <Table
        headers={["位置", "写着什么", "实际是什么"]}
        rows={[
          [
            "几乎所有 ③ 的第 2、第 3 sheet",
            "分类コード変換 / NULL・固定値",
            "sinops item_mst.txt / store_item_mst.txt 的 leftover（细目 0、销售允许日数、発注开始日）",
          ],
          [
            "310R（在库一览） ① 的 sheet 名",
            "【倉庫コード】ka.txt 当日 / 翌日 subka.txt / 受払 uke.txt",
            "sinops 在库・発注勧告・废弃 IF，不是 IFSYLMS310R",
          ],
          [
            "320R（在库调整） ① 第二 sheet",
            "受払明細_废弃 uke.txt ↔ ALLSIRE_MEISAI 伝票区分 19",
            "相邻财务 IF，不是 320R（在库调整） 调整码",
          ],
          [
            "210R（DC入荷実績） 判定サマリ标题",
            "仍写 IFSYLMS230",
            "从 230R（TC総量実績） 整表复制后改入出力仕様 ID",
          ],
          [
            "040S（出荷先） GAP / ③ 标题",
            "TRIAL→WMS 仕入先マスタ",
            "文件夹是店マスタ，内容未改",
          ],
          [
            "130S（TC総量予定） GAP 副标题",
            "↔ TC店舗振替予定データ（IFSYLMS140S）32 项目",
            "本文是 130S（TC総量予定）",
          ],
          [
            "110S（DC入荷予定） GAP 表头",
            "对象 sinops 项目 / sinops 侧要件",
            "本文是 Hitluster 110S（DC入荷予定）",
          ],
          [
            "310R（在库一览） v2 ① 字段行",
            "符号区分 + 在库调整数",
            "那是 320R（在库调整） 的列；一览 IF 没有符号",
          ],
          [
            "業務ルール合意 No.0.0",
            "文件名 item_mst.txt / store_item_mst.txt，IF 63/64",
            "sinops 商品 IF，不是 WMS",
          ],
          [
            "コード変換 店舗コード",
            "sinops 001987 ＳＴ花小金井 → TRIAL 785",
            "WMS 要的是 007452 三郷 这类 6 桁仓库码",
          ],
        ]}
        rowTone={[
          "warning",
          "danger",
          "warning",
          "warning",
          "danger",
          "warning",
          "warning",
          "danger",
          "warning",
          "danger",
        ]}
      />

      <Grid columns={2} gap={16}>
        <Card>
          <CardHeader>v1 / v2 在库表</CardHeader>
          <CardBody>
            <Stack gap={8}>
              <Text>
                v1 目标 current_stock_data：store_code INT4 + jan
                VARCHAR(20) + amount。v2 改为 center_stock：product_code
                VARCHAR(13)，删除 amount，增加 stock_date / operation_time /
                transfer_time / registered_at / registered_by。
              </Text>
              <Text>
                两套 ①②③ 都还在包里。设计只应引用带「Center_Stock」后缀的
                v2。QA-G01 仓库对照在 v2 里仍然是 ???。
              </Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>建议的清理顺序</CardHeader>
          <CardBody>
            <Stack gap={8}>
              <Text>
                1. 删掉每本 ③ 里的 sinops 分类 / NULL
                sheet，以及 310R/320R 的 ka.txt、uke.txt。
              </Text>
              <Text>
                2. 填 030S（仕入先） / 040S（出荷先），按 020S（部门） 的 HUB 契约（UTF-8、荷主
                0000、更新区分）。
              </Text>
              <Text>
                3. 把 C1–C4 写进 コード変換和 130S/140S/230R/240R
                ③，替换「10=店别 / 12=総量」和「伝票番号直接」。
              </Text>
              <Text>
                4. 業務ルール合意改成 WMS
                专用：形态语义、控制器采番、单价来源、仓库对照、BOM、全件却全是
                02。
              </Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <Callout tone="info" title="这包仍然有用">
        它第一次把「CV 之后 TRIAL 哪张表喂哪本 IF、吃哪张伝票」画出来了：予定吃発注/振分，入荷実績进
        TMDLI0110，出荷実績进振替，在库进 center_stock，330R（単価）
        不用。后续工作是去掉 sinops 残页、把样本成对规则和物流控制器嵌进去，而不是重做 15
        本文件夹。
      </Callout>
    </Stack>
  );
}
