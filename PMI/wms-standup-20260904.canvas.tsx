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

type View = "snap" | "closed" | "qty" | "map" | "next";

export default function WmsStandup20260904() {
  const [view, setView] = useCanvasState<View>("wms-su0904-view", "snap");

  return (
    <Stack gap={24}>
      <Stack gap={8}>
        <H1>WMS要件定義 Daily Stand Up · 2026-09-04</H1>
        <Text tone="secondary">
          13:58 JST 起约 80 分。Gemini 两份：无「(1)」是全文转写（902
          段），「(1)」是摘要（41 段）。下文以转写为准，摘要里有三处说过头。主线仍是入数 mapping，前半顺手关了ロケ
          / ASN / 単価 / 返品 / 棚卸。
        </Text>
        <Row gap={8} wrap>
          <Pill active={view === "snap"} onClick={() => setView("snap")}>
            当场结论
          </Pill>
          <Pill active={view === "closed"} onClick={() => setView("closed")}>
            已关 / 半关
          </Pill>
          <Pill active={view === "qty"} onClick={() => setView("qty")}>
            入数（未关）
          </Pill>
          <Pill active={view === "map"} onClick={() => setView("map")}>
            mapping 陷阱
          </Pill>
          <Pill active={view === "next"} onClick={() => setView("next")}>
            作业 / 摘要勘误
          </Pill>
        </Row>
      </Stack>

      {view === "snap" ? <SnapView /> : null}
      {view === "closed" ? <ClosedView /> : null}
      {view === "qty" ? <QtyView /> : null}
      {view === "map" ? <MapView /> : null}
      {view === "next" ? <NextView /> : null}
    </Stack>
  );
}

function SnapView() {
  return (
    <Stack gap={20}>
      <Callout tone="info" title="一场会真正对齐的事">
        Hitluster 信伝票上的出荷単位做除法，不信 WMS
        マスタ里的出荷単位；ケース入数现场还能改，出荷単位错了现场改不了。TRIAL
        仓几乎没有ケース JAN：饮料是単品 JAN，啤酒 6 罐パック才是ボール
        JAN，お菓子按ボール拣但业务口头叫「バラ出荷」。LOMOS/TRIAL
        発注単位=1 是バラ/ケース旗标，不是入数。
      </Callout>

      <Grid columns={4} gap={12}>
        <Stat value="~80 分" label="13:58 起；主会约 76 分" />
        <Stat value="入数未关" label="三田要对照现行 WMS 实值" tone="warning" />
        <Stat value="棚卸已确认" label="三田：无独立功能，走调整" tone="success" />
        <Stat value="来周验证" label="白鳥+第2 全パターン样本" />
      </Grid>

      <H2>谁在场（发言）</H2>
      <Table
        headers={["人", "角色", "这场干什么"]}
        rows={[
          ["張艶", "TRE mapping", "带着课题表问；后半解释 TRIAL 発注単位=1"],
          ["中島孝司", "TRIAL 物流", "控场、DCロケ方针、返品作业、要全パターン样本"],
          ["井上弘隆", "西友/TRIAL 仓运", "TRIAL 无ケース JAN；ビール特例；バラ出荷=ボール拣"],
          ["日立三田", "Hitluster", "讲①カルピス图；ロケ=温度帯；棚卸；要回去对实值"],
          ["李見見", "TRE 010S（商品） mapping", "准备逐项讲商品マスタ，被拉回入数"],
          ["李青廷", "TRE", "催 mapping 说明会；问样本"],
          ["清水 / 柳沼", "日立 通信", "主会后另开 Teams，不属本场 WMS 要件"],
        ]}
      />

      <Callout tone="warning" title="不要把 Gemini 摘要当合同">
        「ASN 全部设 2」「商品阶层整棵树进控制器」「WMS 设定值也一起信」这三句摘要都过宽。转写里的限定见「摘要勘误」。
      </Callout>
    </Stack>
  );
}

function ClosedView() {
  return (
    <Stack gap={20}>
      <H2>DCロケーション ≠ 棚番</H2>
      <Table
        headers={["对象", "是什么", "不是什么"]}
        rows={[
          ["010S（商品） / 予定的 DCロケーション", "常温/低温（畜产、チルド等）分区码，用来把一张大仓拆成逻辑仓", "物流中心内部棚レイアウト"],
          ["WMS エリアロケ / ロケーションマスタ", "仓内自己维护，不从基干收", "MD 对账字段"],
          ["TRIAL 侧", "没有同名项。中島：用商品阶层在物流控制器里落成接近 Hitluster 的ロケ", "把棚番号从 TRIAL 仓系统抄过来"],
        ]}
        rowTone={["success", "neutral", "info"]}
      />
      <Text>
        张问「没有就空白？」三田：这个必须从上位来，否则温度帯分不了。中島当场定方针：阶层 →
        控制器转换 → 填 Hitluster ロケ。对应 QA 101/103，不是 8 层分类投影（118）。
      </Text>

      <H2>ASN 区分（本场只定 WMS 这条）</H2>
      <Table
        headers={["码", "含义（西友現行）", "本场结论"]}
        rows={[
          ["0", "不对应 ASN", "不动"],
          ["1", "MD 基干向け ASN", "店直今后若用 ASN，归取引先マスタ PJ，本会不答"],
          ["2", "WMS 向け ASN", "凡要进 Hitluster 的，设 2"],
        ]}
      />
      <Text>
        井上：电文都先过 MD，再按旗标决定要不要投 WMS。中島一度当成 TRIAL
        全量 ASN 进基干，后更正「这是西友 030S（仕入先）」。摘要写「全部 ASN=2」会把店直也卷进来。
      </Text>

      <H2>其余当场能写进设计的</H2>
      <Table
        headers={["题", "结论", "残"]}
        rows={[
          ["110S（DC入荷予定） 発注単価", "税抜。中島当场答", "无"],
          ["返品 / 废弃", "WMS 只做在库增减，走 320R（在库调整）。有的调整码不回 MD。对象=買い取り在库，通過不算（井上）", "仕入先返品的会计仕訳：中島作业（196）"],
          ["分类桁", "三田再答「拡張できる、影响大概不大」。TRIAL 码会原样过来还是控制器转，三田追问未钉", "208 仍不是设计约束"],
          ["棚卸", "三田确认：无独立功能。出在库一览，现场对实物，差异走在库调整；清单+ハンディ", "209 可写进在库详细设计"],
          ["ITF", "ハンディ能扫。三田先否认后改口「できます」", "画面メンテ项，不挡 IF"],
        ]}
        rowTone={["success", "warning", "info", "success", "success"]}
      />
    </Stack>
  );
}

function QtyView() {
  return (
    <Stack gap={20}>
      <Callout tone="warning" title="公式没翻，荷姿地图变细了">
        三田当场把①カルピス再讲了一遍：72÷24=3，マスタケース入数=1；出荷 48÷24=2。和日立幻灯①一致。未关的是：TRIAL
        没有ケース JAN、啤酒指示数可能已经是ボール数、お菓子的ケース入数当场口说
        80——和幻灯③的 8 打架。三田要拿现行 WMS 实值核对后再画完成版。
      </Callout>

      <H2>井上：TRIAL 仓怎么管 JAN</H2>
      <Table
        headers={["荷姿", "JAN", "库存", "拣货 / 口头"]}
        rows={[
          ["饮料等（カルピス类）", "単品 JAN，无ケース JAN", "バラ", "整箱出时仍按ピース倍数"],
          ["啤酒 6 罐パック", "ボール JAN", "ボール", "2 ケース = 8 ボール（入数 4）"],
          ["お菓子（きのこの山）", "単品 JAN，通常无ボール JAN", "単品", "按ボール拣，业务叫「バラ出荷」"],
        ]}
      />

      <H2>三田现场算式 vs 幻灯③</H2>
      <Table
        headers={["例子", "指示バラ", "出荷単位", "ケース入数", "来源"]}
        rows={[
          ["① カルピス 3 箱入", "72", "24", "1", "幻灯① + 本场 39:50，一致"],
          ["③ きのこの山（幻灯）", "160 / 60", "10", "8（箱内ボール）", "日立资料③"],
          ["きのこの山（本场 57:59）", "ピース", "10（WHPK）", "口说 80", "三田+井上；与③冲突"],
          ["ビール 2 ケース", "8（ボール数，不是 48 ピース）", "待核", "4", "井上；三田说形成数会对不上"],
        ]}
        rowTone={["success", "info", "danger", "danger"]}
      />
      <Text>
        啤酒若指示数已经是 8，再按「MD 永远送ピース」去除法会错。お菓子若ケース入数写
        80，第二步会把ボール再除成箱内ピース。设计暂仍按幻灯③（入数=8），但
        189 不能标关闭——三田自己说「形成数ちぐはぐ，要对实值」。
      </Text>

      <H2>谁信哪个数</H2>
      <Table
        headers={["字段", "信谁", "现场能不能改"]}
        rows={[
          ["伝票 出荷単位（WHPK）", "上位トランザクション。井上确认 WMS マスタ不拿来换算", "不能。错了整箱黑盒就错"],
          ["010S（商品） ケース入数", "商品マスタ；WMS 画面/ハンディ検品可改", "能"],
          ["TRIAL 発注単位=1", "LOMOS 01 旗标（バラ化するか），不是入数", "不能拿去填出荷単位"],
        ]}
        rowTone={["danger", "info", "danger"]}
      />
    </Stack>
  );
}

function MapView() {
  return (
    <Stack gap={20}>
      <Callout tone="danger" title="mapping 两个选择支进一个 Hitluster 列 = 不合格">
        中島：商品マスタ入数把「ケース」和「小分け/内箱」两个 TRIAL
        项塞进同一条 ケース入数，选项会打架。李見見原先打算按荷姿切换内箱/外箱——中島/井上当场否：ケース入数必须是外箱语义。
      </Callout>

      <H2>発注単位=1 不是 24，也不是 10</H2>
      <Text>
        張：样本里発注単位全是
        1，这个 1 表示「バラ」，不是数量。若 WMS 要出荷単位，必须另找マスタ（店别調達等），不能抄这一列。井上：LOMOS
        送到仓的 1 是「バラ化するか」的 01 旗标，和 8、10、24、80 都不是同一个东西。中島：昨天有人把発注単位 map
        进入数，完全看不懂——现在明白了。
      </Text>

      <H2>样本质量</H2>
      <Table
        headers={["现状", "中島要求", "答应"]}
        rows={[
          ["上周到本周样本有错；只有 3 个 JAN 级例子", "通过+在库+ケース出荷+バラ出荷+ビール；食品不够还要医療", "張：白鳥 + 白鳥第2"],
          ["Himalaya 与 LOMOS 文件不同？", "張更正：下载是全量，仓内再按旗标分系统，列相同", "不要当两套入数语义"],
          ["只对文言就签字", "中島拒绝。要实数对实物，否则切数据会出事", "下一 round：三田完成版图 × 中島对样本"],
        ]}
        rowTone={["danger", "info", "warning"]}
      />
    </Stack>
  );
}

function NextView() {
  return (
    <Stack gap={20}>
      <H2>作业（主会内）</H2>
      <Table
        headers={["谁", "做什么", "卡住什么"]}
        rows={[
          ["中島", "買い取り在库 → 仕入先返品的会计仕訳（西友現行）", "196"],
          ["三田", "邮件列出还要确认的 mapping 项；用现行 WMS 实值核对①③；出完成版换算图", "189/198；ビール形成数"],
          ["張", "白鳥+第2 抽全パターン（通過/在库/ケース/バラ/ビール/医療）", "日历；量大会慢"],
          ["張 / 李見見", "另约商品マスタ mapping 说明（本场没讲完）", "不要再从第 1 行讲起，先切入数列"],
        ]}
      />
      <Text tone="secondary">
        主会后清水×三田×柳沼改去 Teams 谈通信要件，与本场 IF
        无关。中島希望来周（含验证）把入数收干净；本周原计划没关上。
      </Text>

      <H2>Gemini 摘要勘误</H2>
      <Table
        headers={["摘要原句", "转写实际"]}
        rows={[
          ["ASN 区分全部按 WMS 向け=2 运用", "只有要进 Hitluster 的设 2。店直 ASN 归取引先マスタ，本会明确不管"],
          ["TRIAL 商品阶层转到物流控制器", "只定了 DCロケーション从阶层转换。不是 8→4 分类投影，也不是整棵树"],
          ["信用基干トランザクション和 WMS 设定值", "出荷単位只信伝票。ケース入数才是 WMS 可改的设定值。两句不能并成一句"],
          ["白鳥中心含饮料的全パターン", "白鳥 + 第2；ビール单列；食品不够还要医療（不是饮料）"],
        ]}
        rowTone={["danger", "warning", "danger", "info"]}
      />

      <H2>对已有分册</H2>
      <Table
        headers={["分册", "本场要改的"]}
        rows={[
          ["数量换算", "加 TRIAL 无ケース JAN、ビール第四例、③的 8 vs 当场 80 未决"],
          ["Overview", "DCロケ=温度帯不是棚番；棚卸三田已确认；ASN=2 仅 WMS"],
          ["TRE QA", "209 三田口头关闭功能缺口；208 仍是「能扩」；101/103 ロケ方针钉了"],
          ["工数", "包络不动。189/198 门闩更硬；啤酒特例是验证项不是新 IF"],
        ]}
      />
    </Stack>
  );
}
