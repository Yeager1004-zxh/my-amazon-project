# SP / SB / SD 广告费用自动归属

本模块把“识别推广商品”和“核对实际花费”分成两步。活动名称、广告归因销售额、平均比例和人工逐条绑定都不参与自动归属。

## 当前数据流

1. 以所选店铺 SID、站点日期、`profile_id` 和 `campaign_id` 建立活动键。
2. 活动报表 cost 只作核对总额。SP/SD 使用商品广告报表行；SB 只有在报表日推广身份和费用分摊口径都经过验证后，才可归属商品或广告组实际花费。
3. 商品标识只接受本地产品表现映射中的 MSKU 和 ASIN 精确匹配。Ad ID、Ad Group ID、Campaign ID 用于限制记录关系，不能单独推断 SKU。
4. 若同一报表行的 MSKU 与 ASIN 指向不同 local SKU、任一标识重复指向多个 SKU、记录跨 SID、成本重复或成本来源不明确，该行不自动归属。
5. 活动总额与其唯一明细粒度的 cost 逐活动核平。活动报表不加到商品/广告组费用上。某活动不平时只隔离该活动，其他已核平活动仍可在预览中显示。
6. 缺少商品记录不是 0。只有接口明确返回商品级 `cost=0` 才显示 0。未知归属、查询失败、成本不平和未映射分别显示。

金额运算使用 `Decimal`。一个 local SKU 可具有多个 MSKU、ASIN；历史别名仅在其有效日期经过核验时使用。当前产品表现接口返回的别名集合没有提供逐日有效期，因此 UI 会将自动匹配标为预览，不能作为正式周报广告输入。正式写入还需验证日期时区、币种、覆盖范围和真实后台总额。

## SP

使用 SP 商品广告报表的商品行 cost，先按 `profile_id + campaign_id + report_date` 与活动报表核对，再按真实 MSKU/ASIN 映射 local SKU。活动名称和广告成交归因销售额不会改变费用归属。活动行不重复加到商品 cost。

## SB

现有 SB 活动和广告组报表没有商品标识，只能核对活动/广告组总额。已接入 `hsaProductAds` 当前创意目录，可按 `profile_id + campaign_id + ad_group_id + ad_creative_id` 读取顶层 `asin` 候选，并在预览中显示其与本地 SKU 的精确 ASIN 映射。该接口没有报表日期或历史生效区间；`creation_date` / `last_updated_date` 仅提供 number，文档没有说明数值时间单位和时区。因此当前目录不能证明 ASIN 在历史花费日期有效，不能把广告组成本自动记入候选 SKU。预览会显示当前商品目录，但成本留在未归属状态。

未来只在以下两条路径通过官方口径核验后自动归属：

- **独占广告组路径：** 需有覆盖完整的历史推广身份，且明确其在报表日有效；每条记录按活动/广告组/创意 ID 关联到被推广 ASIN；组内所有目标唯一映射到同一个 local SKU；广告组 cost 与活动 cost 核平。当前 `hsaProductAds` 目录不满足历史有效期条件。
- **官方已分摊金额路径：** `sbDivideAsinReports` 的用户提供 Schema 表格与示例字段错位，`report_date`、`profile_id`、`campaign_id`、`ad_group_id`、`asin`、`sku` 的类型/样例无法相互对应；`spends`、`percent` 等数组没有元素结构或金额语义。客户端在本地直接阻止此接口，不发送请求。必须有修正后的官方 Schema/SDK 结构说明 `spends` 是金额还是列表、每项结构、SKU/ASIN 关系、percent 用法、舍入、币种/日期口径及如何与活动 cost 核平，才可接入。

已接入但只用于当前目录诊断的接口：

- `POST /pb/openapi/newad/hsaProductAds`：游标或 offset 分页；读取顶层 ASIN 和真实广告关系 ID，不用活动名、creative/landing page 内的 ASIN 或视频素材 URL 猜测商品归属。
- `POST /pb/openapi/newad/sdProductAds`：带 `X-API-VERSION: 2`，读取当前 MSKU/ASIN/ad ID 目录作交叉核对；历史费用仍以 SD 商品报表按日返回的 SKU/ASIN/cost 为准。

SB 分摊接口 `POST /pb/openapi/newad/sbDivideAsinReports` 当前被阻止。上述目录接口只用于展示当前商品标识，不能代替历史快照。所有分页必须按 `total` 完整读取、检查重复与空页；分页游标仅在内存中传递。

## SD

现有 SD 商品报表本身包含 `ad_id`、`campaign_id`、`ad_group_id`、MSKU、ASIN 和 cost，可按精确商品身份匹配并与活动 cost 核对。已接入 `sdProductAds` 当前目录用作商品身份 ID 交叉核验，不覆盖历史商品报表中矛盾的数据，也不把当前状态当作历史有效期证明。

## 当前 Windows 预览验收

使用 `EX-shop-US`，输入 `ex3114`，日期 `2026-08-29`：

1. 更新本机分支至最新提交，启动 GUI，刷新店铺，选择 EX-shop-US。
2. 录入 `ex3114`，日期填 `2026-08-29` 至 `2026-08-29`，执行广告费用预览。
3. 分开记录本地参数校验、每个 OpenAPI HTTP/业务 code 状态、SB/SD 当前目录记录数、是否返回真实活动/商品/广告组行，以及 SP/SB/SD 每日活动总额和唯一明细总额是否核平。
4. 比对领星后台 SB 专属花费 `$19.21`。当前创意目录仅能展示 ASIN/SKU 候选，不会自动把历史 SB 广告组 cost 记入 `ex3114`；必须先补充修正后的分摊 Schema 或报表日历史身份证据。
5. 不用 synthetic test 的金额证明真实数据。只反馈脱敏的字段结构、接口名、分页总数及本地/接口/业务阶段成功状态，不要发送凭证、Token、Seller ID、广告 ID、订单明细或完整响应。

真实 ex3114 验收之前，当前费用仍只用于预览，不进入 Excel。验证完成后还需在正式接入前确认当前 ASIN/MSKU 映射覆盖历史广告商品变化，并完成 SB 视频广告及日期/币种核对。
