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

type View = "formula" | "case" | "ball" | "design";

export default function WmsQtyConversion() {
  const [view, setView] = useCanvasState<View>("qty-case-view", "formula");

  return (
    <Stack gap={24}>
      <Stack gap={8}>
        <H1>数量换算：MD 用ピース，WMS 用管理单位</H1>
        <Text tone="secondary">
          日立① ケース管理（カルピスソーダ）+ ③ ボール管理（きのこの山）。公式相同：WMS
          管理数 = 指示ピース ÷ 出荷単位。ケース入数只回答「一箱里有几个管理单位」，从来不是箱内ピース。
        </Text>
        <Row gap={8} wrap>
          <Pill active={view === "formula"} onClick={() => setView("formula")}>
            同一公式
          </Pill>
          <Pill active={view === "case"} onClick={() => setView("case")}>
            ① ケース
          </Pill>
          <Pill active={view === "ball"} onClick={() => setView("ball")}>
            ③ ボール
          </Pill>
          <Pill active={view === "design"} onClick={() => setView("design")}>
            设计 / 单测
          </Pill>
        </Row>
      </Stack>

      {view === "formula" ? <FormulaView /> : null}
      {view === "case" ? <CaseView /> : null}
      {view === "ball" ? <BallView /> : null}
      {view === "design" ? <DesignView /> : null}
    </Stack>
  );
}

function FormulaView() {
  return (
    <Stack gap={20}>
      <Callout tone="info" title="两步，不要合成一个数">
        第一步（IF 上发生）：ピース ÷ 出荷単位 = WMS 管理数（ケース商品得到ケース，ボール商品得到ボール）。第二步（仓内可选）：管理数 ÷
        010S（商品） ケース入数 = 物理箱数。出荷可以不是整箱，所以出荷常常停在第一步。
      </Callout>

      <Grid columns={3} gap={12}>
        <Stat value="ピース ÷ 出荷単位" label="WMS 管理数（始终）" />
        <Stat value="管理数 ÷ ケース入数" label="物理箱数（入荷展示用）" />
        <Stat value="出荷単位 × ケース入数" label="一箱里的ピース（验算）" />
      </Grid>

      <H2>对照</H2>
      <Table
        headers={["", "① カルピス（ケース管理）", "③ きのこの山（ボール管理）"]}
        rows={[
          ["WMS 管理单位", "1 ケース", "1 ボール"],
          ["1 ボール", "无（或不单独管）", "10 ピース"],
          ["1 ケース", "24 ピース", "8 ボール = 80 ピース"],
          ["出荷単位（伝票）", "24", "10"],
          ["010S（商品） ケース入数", "1", "8"],
          ["验算 出荷単位×入数", "24×1=24 ピース/箱", "10×8=80 ピース/箱"],
          ["入荷例", "72÷24=3 ケース", "160÷10=16 ボール → 16÷8=2 箱"],
          ["出荷例", "48÷24=2 ケース", "60÷10=6 ボール（不足一箱）"],
        ]}
        rowTone={[
          "success",
          "neutral",
          "neutral",
          "info",
          "info",
          "warning",
          "neutral",
          "neutral",
        ]}
      />

      <Callout tone="danger" title="QA 189 不能送「箱内バラ」当ケース入数">
        カルピス若送 24、きのこの山若送 80，第二步会把管理数再除（或现场再乘）错一档。ケース管理入数必须是
        1；ボール管理入数必须是箱内ボール数 8。箱内ピース只允许出现在出荷単位，或用 出荷単位×入数 验算。
      </Callout>

      <Callout tone="warning" title="9/4 standup：公式还在，荷姿地图未关">
        三田当场把①再讲了一遍（72÷24=3，入数=1），信伝票出荷単位，不信 WMS
        マスタ出荷単位。井上：TRIAL 仓没有ケース JAN；啤酒 6 罐是ボール JAN（2 箱=8
        ボール）；お菓子按ボール拣但叫「バラ出荷」。57:59 三田口说きのこの山ケース入数=80，与幻灯③的
        8 冲突——设计暂按③，等三田对现行 WMS 实值。LOMOS 発注単位=1 是バラ/ケース旗标，不是入数。
      </Callout>

      <Text>
        ③ 页脚：出荷単位来自商品マスタ的発注属性。TRIAL 内部発注単位=1
        不能原样出 IF。出 110S（DC入荷予定）/120S（DC出荷指示） 时出荷単位要写成管理单位的ピース数（24 或
        10）。必須チェック仍说 010S（商品） 発注単位数列 WMS 不看、全 1——真正的除数在伝票出荷単位数。
      </Text>
    </Stack>
  );
}

function CaseView() {
  return (
    <Stack gap={20}>
      <H2>① ケース管理 · カルピスソーダ</H2>
      <Table
        headers={["IF", "列", "值", "含义"]}
        rows={[
          ["010S（商品）", "ケース入数", "1", "1 物理箱 = 1 个 WMS 单位"],
          ["110S（DC入荷予定）", "発注数量バラ", "72", "3 箱 × 24"],
          ["110S（DC入荷予定）", "出荷単位数", "24", "1 ケース的ピース"],
          ["120S（DC出荷指示）", "出荷数量バラ", "48", "2 箱 × 24"],
          ["120S（DC出荷指示）", "出荷単位数", "24", "同上"],
        ]}
      />
      <Text>
        管理数已经是ケース，第二步 ÷1 不变。回 MD 的確定数量仍是 72 / 48
        ピース，不是 3 / 2。
      </Text>
    </Stack>
  );
}

function BallView() {
  return (
    <Stack gap={20}>
      <H2>③ ボール管理 · 明治 きのこの山</H2>
      <Text>
        幻灯③：TRIAL 按単品 JAN 管理，WMS 按ボール作业，可出不足一箱。9/1
        曾把「入数 80」当成ケース出荷。9/4
        井上补充：お菓子通常没有ボール JAN，拣货仍按ボール，口头叫「バラ出荷」；三田当场又口说入数=80。在三田完成版图出来之前，010S（商品）
        仍按 8（箱内ボール），80 只作验算。
      </Text>
      <Table
        headers={["IF", "列", "值", "含义"]}
        rows={[
          ["010S（商品）", "ケース入数", "8", "一箱里 8 个 WMS 单位（ボール）"],
          ["110S（DC入荷予定）", "発注数量バラ", "160", "2 箱 × 80 ピース"],
          ["110S（DC入荷予定）", "出荷単位数", "10", "1 ボール = 10 ピース"],
          ["—", "仓内再算", "160÷10=16，16÷8=2", "先得ボール，再得物理箱"],
          ["120S（DC出荷指示）", "出荷数量バラ", "60", "6 ボール × 10"],
          ["120S（DC出荷指示）", "出荷単位数", "10", "出 6 ボール，不凑整箱"],
        ]}
      />
      <Callout tone="info" title="入荷按箱、出荷按ボール">
        入荷指示 2 ケース仍先换成 160 个再除 10。出荷 6
        ボール只除一次，不再 ÷8。同一 JAN 入侧看到箱、出侧看到ボール，除数都是出荷単位
        10，变的是要不要用ケース入数做第二步。
      </Callout>
    </Stack>
  );
}

function DesignView() {
  return (
    <Stack gap={20}>
      <H2>生成规则（Shinise / 控制器）</H2>
      <Table
        headers={["要写的", "ケース管理", "ボール管理"]}
        rows={[
          ["判断荷姿", "上游指定ケース", "上游指定ボール（WMS 不猜）"],
          ["010S（商品） ケース入数", "1", "箱内ボール数"],
          ["伝票 出荷単位数", "箱内ピース", "ボール内ピース"],
          ["伝票 数量バラ", "箱数 × 箱内ピース", "ボール数 × ボール内ピース"],
          ["禁止", "入数=箱内ピース", "入数=80 这种箱内ピース"],
        ]}
      />
      <Text>
        第四例（9/4 井上，未闭环）：啤酒 6 罐ボール JAN。2 ケース的指示数可能已经是 8（ボール数）而不是 48
        ピース，ケース入数=4。若仍按「MD 永远送ピース」去除法会错。三田说形成数会对不上，要查西友現行怎么收。
      </Text>

      <H2>单测</H2>
      <Table
        headers={["用例", "期望"]}
        rows={[
          ["カルピス 72 / 24，入数 1", "WMS 3 ケース"],
          ["カルピス入数误填 24", "失败（会当成 24 个管理单位/箱）"],
          ["きのこの山 160 / 10，入数 8", "16 ボール = 2 箱"],
          ["きのこの山 60 / 10", "6 ボール，允许非 8 倍数"],
          ["きのこの山入数误填 80", "失败"],
          ["出荷単位误为 1（抄 LOMOS 旗标）", "管理数=ピース，库存放大，失败"],
          ["ビール 2 箱指示=8 再当ピース除", "待三田；先当失败用例"],
          ["50 ÷ 24", "余数拒收"],
        ]}
        rowTone={[
          "success",
          "danger",
          "success",
          "success",
          "danger",
          "danger",
          "danger",
          "warning",
        ]}
      />

      <Text tone="secondary">
        ハイチュウ（9/1：DC 単位 1、出荷単位 8）是「入荷除数和出荷除数可以不同」的第三例，110S（DC入荷予定）
        才同时要発注単位数和出荷単位数。①③ 两页入出荷用的是同一个出荷単位。
      </Text>
    </Stack>
  );
}
