# SP / SB / SD 广告费用自动归属

本模块把“识别推广商品”和“核对实际花费”分成两步。活动名称、广告归因销售额、平均比例和人工逐条绑定都不参与自动归属。

## 当前数据流

1. 以所选店铺 SID、站点日期、`profile_id` 和 `campaign_id` 建立活动键。
2. 活动报表 cost 只作核对总额。SP/SD 使用商品广告报表行；SB 只有在报表日推广身份和费用分摊口径都经过验证后，才可归属商品或广告组实际花费。
3. 商品标识只接受本地产品表现映射中的 MSKU 和 ASIN 精确匹配。Ad ID、Ad Group ID、Campaign ID 用于限制记录关系，不能单独推断 SKU。
4. 若同一报表行的 MSKU 与 ASIN 指向不同 local SKU、任一标识重复指向多个 SKU、记录跨 SID、成本重复或成本来源不明确，该行不自动归属。
5. 活动总额与其唯一明细粒度的 cost 逐活动核平。活动报表不加到商品/广告组费用上。某活动不平时只隔离该活动，其他已核平活动仍可在预览中显示。
6. 缺少商品记录不是 0。只有接口明确返回商品级 `cost=0` 才显示 0。未知归属、查询失败、成本不平和未映射分别显示。

## Windows 实测反馈（用户报告）

EX-shop-US / `ex3114`，2026-08-28 至 2026-08-29：SP 商品花费可按 MSKU/ASIN 匹配；SB 当前创意目录返回 12 条，SD 当前商品目录返回 39 条；SP 和 SB 活动/明细报表差异均为 0。用户已在领星后台核实 2026-08-29 的 SKU 级 SB 花费为 `$19.21`，店铺级 SB 为 `$23.08`，差额来自其他商品。该结果是真实业务验收反馈，但程序当前还不能从返回关系自动得到 `$19.21`，没有在代码中硬编码该金额。

用户另报告 SD 当前目录有记录但所选日期的历史广告报表没有历史明细。当前商品目录 `sdProductAds` 没有 `report_date` 或 `cost`，所以这不能证明指定日期花费为 0。预览会把“成功响应但无历史行”显示为未知；若活动报表有行而商品报表为空，会明确显示商品接口无记录，不能把活动总额复制到 SKU。

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

SB 分摊接口 `POST /pb/openapi/newad/sbDivideAsinReports` 当前被阻止。公开页面正文虽列出请求字段 `profile_id`、`report_date`、`offset`、`length`、`next_token`，但返回字段表和成功示例明显错位。实现自动归属前还缺：

- 修正的官方 `data` 记录结构与字段类型，特别是 `report_date`、`profile_id`、`campaign_id`、`ad_group_id`、`sku`、`asin`。
- `spends` 和 `percent` 的每项结构、单位、金额含义、舍入规则。必须确认 `spends` 是否直接表示 ASIN/SKU 级实际花费，或只是列表/归因值/待分摊输入。
- 该接口是否覆盖 SP/SB 视频和所有 SB campaign 类型，以及 `sum(SKU spends)` 应如何与 SB 商品报表、广告组报表、活动报表核平。
- 该报表数据的历史保留范围、结算延迟、站点日期时区和币种。

Codex 云端用普通 HTTPS 请求访问领星文档时得到代理 `403 Forbidden`，不尝试绕过。已上传的 2024 Java SDK 包含通用签名/HTTP 客户端和订单示例，没有 SB/SD 广告报表实现可用于核对该响应语义。

如果 `sbDivideAsinReports` 不返回 SKU/ASIN 级实际花费，就需要领星官方提供报表日商品推广身份历史或 SB 商品级实际花费接口；当前 `hsaProductAds` 的实时目录不足以把 `$23.08` 店铺费用拆到 `ex3114`。

所有分页必须按 `total` 完整读取、检查重复与空页；分页游标仅在内存中传递。

## SD

历史费用请求使用 `POST /pb/openapi/newad/sdProductAdReports`，每个站点日期发送 `sid`、`report_date`、`show_detail=0`、`offset`、`length`，并带 `X-API-VERSION: 2`。预期历史商品行含 `ad_id`、`campaign_id`、`ad_group_id`、MSKU、ASIN、`cost` 和 `report_date`。只有这些商品行完整返回、日期/活动费用核平后才按 MSKU/ASIN 归属。已接入 `sdProductAds` 当前目录用作身份交叉核验，不覆盖历史商品报表中矛盾的数据，也不把当前状态当作历史有效期证明。

本次 Windows 反馈指出历史 SD 明细无记录，但没有说明响应 envelope 的 `code`、`total` 以及同日活动报表行数。需要确认具体属于 `code=0,total=0`、活动有行但商品无行，还是业务响应/请求失败；这些状态不能混为零花费。若 `sdProductAdReports` 对该日成功但返回 0 行，关键缺口就是领星对该历史日期是否有该类型商品级 spend 行、报表保留范围/生成延迟，或是否有另一个按日返回 `ad_id + SKU/ASIN + cost` 的官方接口。

## 当前 Windows 预览验收

使用 `EX-shop-US`，输入 `ex3114`，日期 `2026-08-29`：

1. 更新本机分支至最新提交，启动 GUI，刷新店铺，选择 EX-shop-US。
2. 录入 `ex3114`，日期填 `2026-08-29` 至 `2026-08-29`，执行广告费用预览。
3. 分开记录本地参数校验、每个 OpenAPI HTTP/业务 code 状态、SB/SD 当前目录记录数、是否返回真实活动/商品/广告组行，以及 SP/SB/SD 每日活动总额和唯一明细总额是否核平。
4. 比对领星后台 SB 专属花费 `$19.21`。当前创意目录仅能展示 ASIN/SKU 候选，不会自动把历史 SB 广告组 cost 记入 `ex3114`；必须先补充修正后的分摊 Schema 或报表日历史身份证据。
5. 不用 synthetic test 的金额证明真实数据。只反馈脱敏的字段结构、接口名、分页总数及本地/接口/业务阶段成功状态，不要发送凭证、Token、Seller ID、广告 ID、订单明细或完整响应。

程序能自动归属 ex3114 之前，当前广告费用仍只用于预览，不进入 Excel。还需确认历史 ASIN/MSKU 映射有效期，并完成 SB 视频广告及日期/币种核对。
