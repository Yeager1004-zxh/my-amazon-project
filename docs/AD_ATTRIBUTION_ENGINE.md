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

用户报告 SP 商品花费可按 MSKU/ASIN 匹配，SP/SB 活动与明细报表在其 Windows 测试中核平；SB 的 SKU 级和店铺级花费均已与后台核对，但程序仍不能从已返回关系自动完成 SKU 归属。真实店铺、SKU、日期和金额只在用户本机用于验收，不保存在仓库或自动化测试。

用户已报告若干 SD 历史商品报表日期与领星后台核对一致。历史报表本身包含日期、SKU、ASIN、ad_id、campaign_id 和 cost，应直接用这些日期范围内的标识匹配本地 SKU。当前 `sdProductAds` 目录只作补充诊断，不是历史归属前提；如果目录查询失败，不能阻挡已成功核平的历史商品报表。其他日期仍需逐日扩展真实验收，不能把无历史商品行视为零花费。具体店铺、SKU、日期和金额不保存在仓库。

金额运算使用 `Decimal`。一个 local SKU 可具有多个 MSKU、ASIN；以广告历史商品报表内同一条报表日记录提供的 SKU/ASIN 做逐行归属，若两个标识已知且冲突、标识映射到多个 local SKU、活动未核平，则该行未归属。产品表现别名集合缺少逐日有效期时，不可单独作为历史有效关系。总广告花费正式写入还要求 SP、SB、SD 覆盖完整且已核验币种、站点日期和后台金额；不完整的 SP 或 SD 子金额不作为总花费。

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

SB 分摊接口 `POST /pb/openapi/newad/sbDivideAsinReports` 当前被阻止。公开页面正文虽列出请求字段 `profile_id`、`report_date`、`offset`、`length`、`next_token`，但返回字段表和成功示例明显错位。向领星官方需要确认的最少问题是：

1. 请提供真实脱敏成功响应的 `data` 结构及每个字段/数组元素的数据类型，尤其是 `report_date`、`profile_id`、`campaign_id`、`ad_group_id`、`sku`、`asin`、`spends`、`percent`。
2. `spends` 的实际含义和单位是什么？它是最终 SKU/ASIN 级广告实际花费金额，还是花费明细列表、归因金额或分摊输入？每个元素如何关联 SKU、ASIN 和活动？
3. `percent` 的计算公式、分母和精度/舍入规则是什么？其结果是否乘以活动/广告组 cost 才得到 SKU 花费，尾差如何处理？
4. `campaign_id` 如何关联到同一 `profile_id`、报表日期的 SB 活动与广告组记录？一个活动多个广告组/ASIN 时，哪些维度唯一标识一次分摊？
5. 接口可查询的最早历史日期、数据更新/结算延迟以及是否覆盖所有 SB 活动类型和 SB 视频广告是什么？
6. 金额返回币种及日期按哪个站点时区统计？SKU `spends` 合计应如何与哪个层级（广告组还是活动）的实际花费对账，允许何种精度差？

收到以上资料并用 Windows 对真实日期核验前，不调用该接口、不按比例猜测、也不把活动级费用写入 SKU。

Codex 云端用普通 HTTPS 请求访问领星文档时得到代理 `403 Forbidden`，不尝试绕过。已上传的 2024 Java SDK 包含通用签名/HTTP 客户端和订单示例，没有 SB/SD 广告报表实现可用于核对该响应语义。

如果 `sbDivideAsinReports` 不返回可核平的 SKU/ASIN 级实际花费，就需要领星官方提供报表日商品推广身份历史或 SB 商品级实际花费接口；当前 `hsaProductAds` 的实时目录不足以把店铺级 SB 花费拆到 SKU。用户提供的真实 SKU 金额仅在 Windows 本地作为对照目标，不在代码、自动化测试或仓库文档中硬编码。

所有分页必须按 `total` 完整读取、检查重复与空页；分页游标仅在内存中传递。

## SD

历史费用请求使用 `POST /pb/openapi/newad/sdProductAdReports`，每个站点日期发送 `sid`、`report_date`、`show_detail=0`、`offset`、`length`，并带 `X-API-VERSION: 2`。历史商品行含 `ad_id`、`campaign_id`、`ad_group_id`、SKU/MSKU、ASIN、`cost` 和 `report_date`。先按 `profile_id + campaign_id + report_date` 与活动报表核平，再使用同条历史行的 MSKU/ASIN 匹配本地 SKU；两项指向不同 SKU 或重复映射时隔离。`sdProductAds` 当前目录是补充诊断，不是历史映射前置条件，失败不能覆盖或阻断历史商品报表结果。

用户已在 Windows 验证指定日期的 SD 历史商品花费与领星后台一致。其他日期仍需扩展真实验收。若某日期活动报表有花费但历史商品报表成功返回 0 行，仍按“SKU 商品归属未知”处理，不能把活动总额复制到 SKU，也不能据此推断零花费。真实店铺、SKU、日期和金额由用户本地保管。

## 当前 Windows 预览验收

使用用户指定的店铺和测试 SKU、单日站点日期：

1. 更新本机分支至最新提交，启动 GUI，刷新店铺，选择 EX-shop-US。
2. 录入测试 SKU 和指定日期，执行广告费用预览。
3. 分开记录本地参数校验、每个 OpenAPI HTTP/业务 code 状态、SB/SD 当前目录记录数、是否返回真实活动/商品/广告组行，以及 SP/SB/SD 每日活动总额和唯一明细总额是否核平。
4. 比对领星后台 SKU 级与店铺级 SB 实际花费。当前创意目录仅能展示 ASIN/SKU 候选，不会自动把历史 SB 广告组 cost 记入 SKU；必须先补充修正后的分摊 Schema 或报表日历史身份证据。
5. 不用 synthetic test 的金额证明真实数据。只反馈脱敏的字段结构、接口名、分页总数及本地/接口/业务阶段成功状态，不要发送凭证、Token、Seller ID、广告 ID、订单明细或完整响应。

程序能自动归属 SB 费用之前，SB 费用仍只用于预览，不进入 Excel。还需确认历史 ASIN/MSKU 映射有效期，并完成 SB 视频广告及日期/币种核对。
