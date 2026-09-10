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

type View = "pyramids" | "wms" | "names" | "design";

export default function HierarchySeiyuTrial() {
  const [view, setView] = useCanvasState<View>("hier-view", "pyramids");

  return (
    <Stack gap={24}>
      <Stack gap={8}>
        <H1>商品阶层：西友金字塔 × TRIAL 8 层 × WMS 实际 4 码</H1>
        <Text tone="secondary">
          来源：マスタ階層比較（西友 vs トライアル）。红虚线框把西友 5 层对齐到
          TRIAL 的 DIV～サブカテゴリー。这是商品企画用的粒度带，不是 Hitluster
          010S/020S 的字段对照表。WMS 文件虽有 6 个分类槽，作业只用 4 个。
        </Text>
        <Row gap={8} wrap>
          <Pill active={view === "pyramids"} onClick={() => setView("pyramids")}>
            两套金字塔
          </Pill>
          <Pill active={view === "wms"} onClick={() => setView("wms")}>
            接到 WMS 4 码
          </Pill>
          <Pill active={view === "names"} onClick={() => setView("names")}>
            同名不同层
          </Pill>
          <Pill active={view === "design"} onClick={() => setView("design")}>
            对详细设计
          </Pill>
        </Row>
      </Stack>

      {view === "pyramids" ? <PyramidsView /> : null}
      {view === "wms" ? <WmsView /> : null}
      {view === "names" ? <NamesView /> : null}
      {view === "design" ? <DesignView /> : null}
    </Stack>
  );
}

function PyramidsView() {
  return (
    <Stack gap={20}>
      <Callout tone="warning" title="先读红框的含义">
        红框说的是「西友整棵 5 层树，大约落在 TRIAL 8 层里从 DIV 到サブカテゴリー这一段」。不是
        Division=DIV、デバ=ライン 这种 1:1。节点数量对不上（7 vs 27、49 vs
        59、391 vs 184、2,655 vs 3,837），只能当粒度带宽，不能当码表。
      </Callout>

      <Grid columns={2} gap={16}>
        <Stack gap={8}>
          <H2>西友 5 层（POS / 商品）</H2>
          <Table
            headers={["层（图上名称）", "图上的「分类」叫法", "节点数"]}
            rows={[
              ["部門", "大分類", "4（不含仕入先管理部门）"],
              ["Division", "中分類", "7"],
              ["デバ", "小分類", "49"],
              ["西友ライン", "細分類", "391"],
              ["ファインライン", "—", "2,655"],
            ]}
          />
          <Text tone="secondary">
            现行未进 POS 的对象也画进去了。除外：dept_nbr 49 / 71 / 78 /
            99（49=垃圾袋、99=テナント等）；fineline_nbr ≥ 9900（如 POS
            部门登录）。用语是西友/Walmart 系：dept、fineline。
          </Text>
        </Stack>
        <Stack gap={8}>
          <H2>TRIAL 8 层全階層</H2>
          <Table
            headers={["层", "节点数", "相对红框"]}
            rows={[
              ["事業部", "10", "框外（更粗）"],
              ["DIV", "27", "框上沿"],
              ["ライン", "59", "框内"],
              ["部門", "184", "框内"],
              ["カテゴリー", "903", "框内"],
              ["サブカテゴリー", "3,837", "框下沿 ≈ 西友ファインライン"],
              ["セグメント", "16,233", "框外（更细）"],
              ["サブセグメント", "52,343", "框外（更细）"],
            ]}
            rowTone={[
              "neutral",
              "info",
              "info",
              "info",
              "info",
              "info",
              "warning",
              "warning",
            ]}
          />
          <Text tone="secondary">
            图注写明约一年前，不是当前冻结版。統合后节点数会变。
          </Text>
        </Stack>
      </Grid>

      <Grid columns={4} gap={12}>
        <Stat value="5 vs 8" label="层数" />
        <Stat value="2,655" label="西友最细：ファインライン" />
        <Stat value="3,837" label="红框底：TRIAL サブカテ" />
        <Stat value="52,343" label="TRIAL 最细：サブセグ" tone="warning" />
      </Grid>

      <H2>红框对齐（粒度带，非码）</H2>
      <Table
        headers={["西友", "约对应 TRIAL 带宽", "数量级"]}
        rows={[
          ["部門 4（大分類）", "比 DIV 更粗；TRIAL 用事業部 10 才到这个量级", "4 vs 10"],
          ["Division 7", "DIV 27", "7 vs 27，一对多"],
          ["デバ 49", "ライン 59", "最接近的个数，仍不是同名层"],
          ["西友ライン 391", "部門 184 或 カテゴリー 903 之间", "个数对不上，必有拆合"],
          ["ファインライン 2,655", "サブカテゴリー 3,837", "红框底；仍细于西友"],
        ]}
      />
    </Stack>
  );
}

function WmsView() {
  return (
    <Stack gap={20}>
      <Callout tone="info" title="WMS 不吃 8 层，也不吃西友 5 层全套">
        Hitluster 020S（部门） 布局有 6 个分类槽，必須チェック规定：大分类现行
        00、细分类不使用。作业主键是 部門グループ + 部門 + 中分類 +
        小分類。010S（商品） バッチ区分只看前两码。予定文件里的中/小分类 WMS 只收不看。
      </Callout>

      <H2>图上的西友层 → 移送表槽位</H2>
      <Table
        headers={["图上西友层", "节点", "WMS 字段", "桁", "必須チェック"]}
        rows={[
          ["部門（大分類）", "4", "大分類コード", "10", "- 现行 00 或空格，不作业"],
          ["Division（中分類）", "7", "部門グループコード", "2", "○ 備考：Division。020S（部门） PK、010S（商品） バッチ"],
          ["デバ（小分類）", "49", "部門コード", "2", "○ 備考：Dept"],
          ["西友ライン（細分類）", "391", "中分類コード", "2", "○ 備考：西友ライン。予定 IF 只收不看"],
          ["ファインライン", "2,655", "小分類コード", "4", "○ 備考：Fineline。予定 IF 只收不看"],
          ["（无对应最细）", "—", "細分類コード", "10", "- 不使用"],
        ]}
        rowTone={["neutral", "success", "success", "info", "info", "neutral"]}
      />
      <Text>
        2 桁装不下全局唯一的 391 条ライン，所以中分类码是「父层下的局部码」，020S（部门）
        必须带齐 4 码才能当主键。生产样本 020S（部门） 约 4,052
        行，粒度接近ファインライン组合行，不是 4 个大分类。
      </Text>

      <H2>mapping 里的 TRIAL 表 ≠ 图上 8 层表头</H2>
      <Table
        headers={["WMS 槽", "mapping 取的 TRIAL 源", "图上 8 层里可能落在哪"]}
        rows={[
          ["部門グループ", "F032 DIVマスタ DivisionCD", "DIV（框上沿）"],
          ["部門", "F022 部門マスタ DepartmentCD", "部門——但图上 184，不是西友的「部門 4」"],
          ["大分類", "F049 関連組織 → 固定空/00", "不要用事业部去填这个槽"],
          ["中分類", "F068 ミニ部門 SubDeptCD", "カテゴリー一带，待码表确认"],
          ["小分類", "F078 品種 VarietyCD", "サブカテゴリー一带，待码表确认"],
          ["細分類", "F038 品目 → 不送", "セグメント / サブセグ 本来就送不进 WMS"],
        ]}
      />

      <Grid columns={2} gap={16}>
        <Card>
          <CardHeader>WMS 要的投影</CardHeader>
          <CardBody>
            <Text>
              每个 JAN 在 010S（商品） 上带 4 个分类码；020S（部门） 按这 4
              码出名称行。事業部、セグメント、サブセグメント、西友最顶的 4
              个大分類，都不进 Hitluster 作业。
            </Text>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>谁做 8→4 的码表</CardHeader>
          <CardBody>
            <Text>
              商品マスタ統合 PJ（mapping 010S（商品） GAP 写的课题 118）。物流控制器和
              IF 开发只消费已经投影好的 4
              码，不要在伝票转换里临时拼階層。
            </Text>
          </CardBody>
        </Card>
      </Grid>
    </Stack>
  );
}

function NamesView() {
  return (
    <Stack gap={20}>
      <H2>同名陷阱（禁止按汉字对层）</H2>
      <Table
        headers={["汉字", "西友图上", "WMS 移送表", "TRIAL 图上", "mapping F 表"]}
        rows={[
          [
            "部門",
            "最顶，4 个大分類",
            "第 2 作业层，Dept，约 49 デバ",
            "第 4 层，184",
            "F022，4 桁",
          ],
          [
            "Division / DIV",
            "第 2 层，7，图上叫中分類",
            "部門グループ，备考 Division",
            "第 2 层，27",
            "F032 DivisionCD",
          ],
          [
            "ライン",
            "西友ライン = 第 4 层，391，图上叫細分類",
            "中分類，备考西友ライン",
            "第 3 层，59",
            "不单独对应 F 表「ライン」",
          ],
          [
            "小分類 / 細分類",
            "デバ=小分類；西友ライン=細分類",
            "小分類=Fineline；細分類槽空着",
            "不用这两个词",
            "品種→WMS小分類；品目→WMS細分類（不送）",
          ],
        ]}
        rowTone={["danger", "warning", "warning", "danger"]}
      />
      <Callout tone="danger" title="最容易写错的一句">
        不能写「TRIAL 部門码填进 010S（商品） 部門码」而不核对粒度。图上 TRIAL 部門是
        184 节点；WMS 部門码对应的是西友デバ约 49。方向反了，バッチ区分和 020S（部门）
        PK 会全错。
      </Callout>

      <H2>POS 除外 ≠ WMS 除外</H2>
      <Text>
        图上 POS 不联动 dept 49/71/78/99、fineline ≥
        9900。物流中心仍可能收到这些 JAN（废弃袋、テナント、部门登录用）。010S（商品）
        是否过滤要单独定，不能因为 POS
        除外就把它们从仓主数据里删掉。
      </Text>
    </Stack>
  );
}

function DesignView() {
  return (
    <Stack gap={20}>
      <H2>详细设计要写成的契约</H2>
      <Table
        headers={["产出", "内容", "不做什么"]}
        rows={[
          [
            "4 码投影表（商品統合 PJ）",
            "每个现行 JAN：WMS 部門G / 部門 / 中分類 / 小分類 + 名称。来源可以是西友ファインライン，也可以是 TRIAL サブカテ投影，但一行只能有一套",
            "不要试图把 サブセグ 52,343 压进 020S（部门）",
          ],
          [
            "020S（部门） 生成规则",
            "按 4 码去重出 CSV；荷主 0000、大分类 00、细分类空；更新区分走 HUB 前回送信",
            "不要做 mapping 里的 4 桁→10 桁大分类扩展",
          ],
          [
            "010S（商品）",
            "每个 JAN 带同一套 4 码；バッチ只用部門G+部門",
            "不要把納品路径/通過在库当分类替代",
          ],
          [
            "110S–140S",
            "部門G（110S（DC入荷予定） ○）/ 部門 ○；中小分类可填投影码，WMS 不读",
            "不要在予定转换里再做一层階層 JOIN",
          ],
          [
            "210R–240R 仕入/振替",
            "MD 部門码用 TRIAL 自己的部门体系，不要用 WMS 2 桁 Dept 去填 Shinise 8 桁部门",
            "mapping 已写「カテゴリマスタ変換」——这是第二张对照，和 020S（部门） 不是同一张",
          ],
        ]}
      />

      <Grid columns={2} gap={16}>
        <Card>
          <CardHeader>对工数的影响</CardHeader>
          <CardBody>
            <Stack gap={8}>
              <Text>
                8 层对 5 层的继承关系是商品マスタ統合的活，不应摊进 WMS IF
                的 010S/020S 人日。
              </Text>
              <Text>
                WMS 侧只估「投影表 → 4 槽」：020S（部门） 保持标准档（约 10–14
                人日），010S（商品） 复杂点在赏味/ケース/取引先，不在 8 层展开。
              </Text>
              <Text>
                若投影表不按期交付，010S/020S/全伝票分类码会停工——这是日历等待，不是把
                380 人日再加 8 层对照。
              </Text>
            </Stack>
          </CardBody>
        </Card>
        <Card>
          <CardHeader>还缺的三张表</CardHeader>
          <CardBody>
            <Stack gap={8}>
              <Text>1. 西友ファインライン → WMS 4 码（现行 Hitluster 已在跑，应用生产 020S（部门） 反推）。</Text>
              <Text>2. TRIAL サブカテ（或品種）→ 同一套 WMS 4 码（CV 后 JAN 用）。</Text>
              <Text>3. WMS 部門 2 桁 → Shinise 仕入伝票部门码（実績回写用，另一粒度）。</Text>
            </Stack>
          </CardBody>
        </Card>
      </Grid>

      <Callout tone="info" title="和先前 GAP 的修正">
        mapping 010S/020S 把「TRIAL 4 层 vs 西友 6 层」写成 ★高。用这张图校正：真正的差是「TRIAL
        8 层企画树 vs 西友 5 层 POS 树」；落到 WMS 都是 4
        个作业码，其中两大槽（大分类、细分类）本来就空着。阶层 GAP 的主战场在商品統合，不在
        Hitluster 改布局。
      </Callout>
    </Stack>
  );
}
