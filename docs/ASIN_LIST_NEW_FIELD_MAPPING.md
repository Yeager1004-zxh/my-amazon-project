# AsinListNew 与新品推广周报字段映射

本映射依据用户上传的领星「查询产品表现（新版）AsinListNew」Markdown。云端对 `apidoc.lingxing.com` 的请求曾在受管代理连接阶段被拒绝（403），因此本文件不声称已在线重新读取官方页面。上传文档列出的 API 为 `POST /bd/productPerformance/openApi/asinList`。

## 可安全执行的 SKU 单日只读请求

单店铺本地 SKU 汇总应使用：

```json
{
  "offset": 0,
  "length": 10000,
  "sort_field": "volume",
  "sort_type": "desc",
  "search_field": "local_sku",
  "search_value": ["<用户输入的 local SKU>"],
  "sid": "<所选店铺的 SID>",
  "start_date": "YYYY-MM-DD",
  "end_date": "相同的 YYYY-MM-DD",
  "summary_field": "sku",
  "is_recently_enum": false,
  "purchase_status": 0
}
```

日期按所选站点日期解释。文档说明日期范围双闭区间、最多 92 天；`length` 最大 10000；`search_value` 最多 50 个值；按 `offset` 分页；单店铺 `sid` 使用字符串；`purchase_status=0` 按退款发生时间统计退货退款。请求不传 `currency_code`，保留原币种。单日请求可用于核对该日期筛选结果，但响应记录没有日期字段，必须在 Windows 与领星同日后台数据对照后，才能认定字段是单日统计。

请求成功必须同时满足 HTTP 2xx、`code=0`、`data` 为对象、`data.total` 为非负整数、`data.list` 为对象数组。SKU 结果使用 `sku`、`sids`、`price_list.local_sku/sid/seller_sku` 与所选 SID 复核；未返回记录、字段未返回、`null` 和明确的 `0` 分开处理。

## 广告费用日期与汇总维度诊断

用户提供的官方 Markdown 明确说明 `start_date`、`end_date` 为必填、双闭区间、最多 92 天；`summary_field` 支持 `sku` 和 `msku`，并说明 SKU 维度键为 `sku + sid`、MSKU 维度键为 `seller_sku + sid`。现有只读诊断按钮要求输入一个 SKU 和连续三天，将串行发送四组请求：三天各一次 `summary_field=sku`，以及中间日一次 `summary_field=msku`。每一日请求都令 `start_date=end_date`，沿用 `search_field=local_sku` 与同一个 local SKU 条件；MSKU 结果只在 `price_list` 能唯一映射回所选 SID 和请求 SKU 时显示。

诊断展示 `spend`、SP、SB、SBV、SD 和差异分摊字段及币种，但不求和、不转换口径、不写周报。日期筛选参数确实发送到 API，但 AsinListNew 响应没有行级报表日期；所以跨日数值变化只能提供筛选行为的线索，不能单凭该结果证明费用字段采用了请求日期。官方 Markdown 将 `spend` 描述为广告花费组成项总计，并列出 `shared_ads_sb_cost`/`shared_ads_sbv_cost` 为 SB/SBV 广告费，但未说明它们的统计来源、分摊算法、广告活动关联、历史覆盖或费用日期时点。若单 SKU 数值仍高于同日全店活动实际花费，必须继续以领星“产品表现”和“广告管理”页面分开比对，不得把该费用字段当作活动实际花费。

## 模板字段映射与当前状态

“已接入”只表示现有生成流程已经处理该字段，不代表新 AsinListNew 字段已完成 Windows 验收。AsinListNew 候选字段在单日实测及口径确认前不写入正式周报。

| 模板字段 | 当前数据源 / AsinListNew 候选字段 | 状态 | 说明 |
|---|---|---|---|
| 周次 | D01 时间轴 | 已接入 | 由程序按连续自然日和每 7 天一周生成。 |
| 天数 | 用户确认的 D01 | 已接入 | D01、D02……；未手动确认 D01 时阻止正式生成。 |
| 周开始日期 | D01 与时间轴 | 已接入 | 每个 SKU 使用独立的人工确认日期。 |
| 售价 | 每日销售额 / 每日销量；候选 `price_list.price` | 已接入 | 现有 API 路径可按日销额与销量计算均价；AsinListNew 的价格行没有足以确认逐日口径的日期字段。 |
| 日销 | `asinDailyLists` 销量接口 | 已接入 | 已有 Windows 真实验证；明确返回 0 与没有记录分开。 |
| 自然单 | 无直接已核实字段；`order_items` 是订单总量，`ad_order_quantity` 是广告订单量 | 无法获取 | 不做总订单减广告订单推算；当前 AsinListNew Schema 没有自然订单字段。 |
| 广告单 | 候选 `ad_order_quantity` | 待验证 | 需以单日范围核对广告后台及日期/归因口径，后续才接 Excel。 |
| 广告订单占比 | 候选 `adv_rate` | 待验证 | Schema 说明为广告订单量占比；需确认单日时点和其分母与模板一致。 |
| 类目 | 候选 `rank_category`、`small_cate_rank[].category` | 待验证 | 需核实哪个字段与模板类目列的业务定义一致。 |
| 小类排名 | 候选 `small_cate_rank[].rank` | 待验证 | 排名记录缺少可用的历史有效日期；不得用当前排名回填过去日期。 |
| 评分 | 候选 `avg_star` | 待验证 | Schema 没有证明该值是指定历史日期快照。 |
| 评论数 | 候选 `reviews_count` | 待验证 | Schema 没有提供其历史快照时间口径。 |
| 日目标销量 | 人工输入 | 人工填写 | 更新已有周报时保护原值。 |
| 目标达成率 | 每日销量 / 日目标销量 | 已接入 | 只有两项输入都存在时计算；不改写人工目标。 |
| Sessions | 候选 `sessions_total` | 待验证 | 单日请求结果需与后台同店铺、同 SKU 对照。 |
| 自然点击 | 无直接字段；`clicks` 的流量归属未由文档证明 | 无法获取 | 不用总点击减广告点击推算。 |
| 广告曝光 | 候选 `impressions` | 待验证 | 需核实其广告类型覆盖范围、日期口径及与 SP/SB/SD 明细是否一致。 |
| 广告点击 | 候选 `clicks` | 待验证 | 需验证是否仅广告点击及其广告类型覆盖范围。 |
| 广告 CTR | 候选 `ctr` | 待验证 | Schema 定义为 clicks / impressions；需核对是否与模板广告流量口径一致。 |
| 广告 CVR | 候选 `ad_cvr` | 待验证 | 需核实订单归因窗口与该周报使用口径。 |
| 广告花费 | V1.0 正式范围 `SP_SD_ONLY_CONFIRMED`；只来自 SP/SD 历史商品报表 | 已接入待端到端验收 | AsinListNew 的 `spend`、`ads_sp_cost`、SB/SBV 分摊字段不替代 SP+SD 活动/商品报表核验，不写入正式花费。 |
| 广告销售额 | 候选 `ad_sales_amount` | 待验证 | 需核对归因窗口、广告类型覆盖范围及单日口径。 |
| 整体 CVR | CSV 流量计算为销量 / Sessions；候选 `cvr` | 已接入 | 仅当两个来源均存在且口径一致才计算；AsinListNew `cvr` 需先验证。 |
| ACOS | SP+SD 完整已核验花费 / 同范围、可比归因口径广告销售额 | 待验证 | 目前没有已验证的逐日 SP+SD 广告销售额来源，故正式 OpenAPI 生成留空；不得用整体 `ad_sales_amount` 直接计算。 |
| 推广阶段（表头） | 用户输入 | 人工填写 | 生成和更新均保护人工内容。 |
| 发货时间（表头） | 用户输入 | 人工填写 | 生成和更新均保护人工内容。 |

## 需 Windows 核验的重点字段

API Schema 直接列出 `sessions_total`、`ad_order_quantity`、`ad_sales_amount`、`impressions`、`clicks`、`ctr`、`ad_cvr`、`avg_star`、`reviews_count`、`small_cate_rank`、`ads_sp_cost`、`shared_ads_sb_cost`、`shared_ads_sbv_cost`、`ads_sd_cost`、`spend` 及 `currency_code`。

其中 `shared_ads_sb_cost` 与 `shared_ads_sbv_cost` 出现在 SKU 汇总响应中，是判断领星是否已输出 SKU 维度 SB/SBV 费用的强候选，但上传的 Schema 没有说明它们的分摊算法、活动归属规则和历史覆盖范围。只有在 EX-shop-US / ex3114 / 2026-08-29 单日请求中确认返回字段、币种为 USD，并与领星同日商品级数据核对后，才能判断是否可替代另行开发的 SB 分摊。程序不会内置或比对任何预设金额。

目前 AsinListNew 页面仍缺少或没有说明：

- 广告花费各分项与 `spend` 的完整求和公式，以及字段是否包含全部广告类型。
- `shared_ads_sb_cost`、`shared_ads_sbv_cost` 的确切 SKU 分摊规则、历史推广商品范围及是否覆盖视频广告。
- `avg_star`、`reviews_count`、`small_cate_rank` 的历史快照时间字段或可按日期查询的保证。
- `clicks`、`impressions`、`ad_sales_amount` 的广告产品类型覆盖及归因窗口。
- `small_cate_rank` 在 Schema 中为数组对象，但成功示例为 `null`；数值字段 Schema 标为 number/integer，成功示例部分又以数字字符串返回。

在以上事项经单日只读响应和领星后台核验前，新增字段仅用于本机预览，不进入 Excel 或校验日志。SP、SD、每日销售的既有实现不因本次预览而改变。
