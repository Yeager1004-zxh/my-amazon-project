# 亚马逊新品推广周报工具

面向 Windows 10/11 的本地新品推广周报工具。V1.0 数据路线为：领星 OpenAPI 作为主数据接口、官方 MCP 作为辅助查询/交叉核验、领星 Excel/CSV 导入作为应急备用。Windows 动态店铺列表已通过真实验收；用户报告每日销售、订单量和销售额查询与领星后台一致。当前分支支持用户按店铺和 SKU 手动确认并持久化 D01，且 OpenAPI 销售数据可生成模板工作簿；完整工作簿生成仍待 Windows 本地验收。用户已报告 SP 实际花费以及 SD 历史商品报表在指定日期的结果与后台核对一致。SB 自动 SKU 归属仍待官方分摊口径。上述真实验收为用户反馈，仓库未保存原始响应或业务明细。

V1.0 分阶段任务、模块输入/输出、依赖、测试和验收门槛见 [正式实施清单](docs/V1_IMPLEMENTATION_PLAN.md)。当前版本仅聚焦新品推广周报。

## 当前实现

- 延续原模板七天一周的样式、周次、边框和条件格式，按实际来源覆盖日期扩展 D61、D62……；自然日连续排列，来源缺失的日期指标保持空白。
- 支持多个 SKU、多个来源文件和多个店铺；字段必须在本机 JSON 映射中明确指定，不猜报表表头。
- 新 SKU 页签会清除模板样例指标。增量更新时，实际来源有值的指标会更新；来源没有记录的旧指标保留。`A2` 推广阶段、`G2` 发货时间和 `M` 列日目标销量由用户维护，程序不会覆盖。
- D01 必须由用户按每个店铺/SKU 明确输入并确认，格式为 `YYYY-MM-DD`，按所选店铺站点日期解释。Windows 本机设置保存在 `%APPDATA%\MyAmazonProject\d01_dates.json`，键为领星 SID 和 SKU；程序不会从商品创建时间、订单或首个有销量日期推断。已有周报的 D01 若被修改，会先提示日期错位风险并要求确认，之后生成仍会再次确认更新风险。
- 销量、销售金额、广告实际花费、广告表现和流量可按真实字段映射。广告花费必须使用实际花费字段；`sku_scope=SKU_LEVEL_CONFIRMED` 只表示单行费用可唯一归属 SKU。只有再明确确认 `coverage_scope=ALL_AD_TYPES_TOTAL_CONFIRMED`，程序才把费用写进模板的总广告花费列；否则部分 SP/SD 金额只记录为来源金额和完整性状态，总广告花费与 ACOS 留空。增量更新时也会清除覆盖不完整的旧总花费值。重复、共享活动或冲突映射会被排除并记入校验日志。
- 校验日志为 JSONL，记录来源文件名、SKU 的 D01、最新数据覆盖日、销量/销售额/广告花费汇总和异常提示。关闭日期后的实际花费只有在来源提供并映射 `campaign_disabled_at` 时才提示。
- Windows 本地 MCP 客户端支持 MCP 初始化连接、工具列表分页、工具目录发现，以及按业务 Schema 核验的只读 action 查询。请求带 `X-Mcp-Key`、MCP session 和协议版本；常规工具仍要求 `readOnlyHint=true`。
- OpenAPI Token 请求使用 multipart/form-data；Access Token 仅在本机内存中按响应明确提供的有效期缓存，并在有效期临近时刷新。未返回可识别有效期的 Token 不缓存。
- 周报页可使用当前动态选择的店铺 SID，按 local SKU 查询产品表现映射，再按站点自然日读取每日销售额、销量和订单量。点击“领星 OpenAPI 查询并生成”会直接生成销售周报；查询日期按最多 92 天连续分段。每段映射单独核验，任何映射、分页、业务码或限流失败都会停止，不生成部分历史结果。广告总花费在跨 SP、SB、SD 完整性未核验前留空，并在结果及校验日志标记为“部分完成”；销售工作簿不会因 SB 暂缺而被阻止。
- 每日 API 的无记录指标保持空白；仅显式返回的 `map_value: 0` 显示为零。销量和订单量是不同指标。商品表现返回的 ASIN 候选不会被假设为与 MSKU 一一对应；表格中的 ASIN 只取自相同每日 MSKU 记录。
- 每日统计字段说明没有单独定义取消订单/退款处理；用户已报告 ex3106 每日销量、订单量和销售额与领星后台一致。D01 固定为用户手动确认，订单列表不再是 V1.0 的 D01 依赖。
- 新增 `ad_attribution.py` 自动归属层：按 SID、站点日期、profile_id、campaign_id 核平，再用精确 MSKU/ASIN 映射归属商品级成本；MSKU 与 ASIN 冲突、映射重复、活动差异、重复数据和未识别商品分别隔离。SB 当前创意目录可显示 ASIN 与广告组/活动候选关系，但没有历史有效区间；SB 分摊接口 Schema 错位，客户端在发送前阻止。活动级 cost 仅作核对基准，不与商品或广告组明细重复相加。SB 归属仍仅供预览，未核验的总广告花费不会进入周报。
- SP/SD 历史商品报表按其自身报表日期提供的 MSKU、ASIN、广告 ID、活动 ID 和 cost 进行 SKU 归属，并按 campaign 核对；SD 历史归属不依赖当前 `sdProductAds` 目录。用户报告若干指定日期的 SD 商品花费已与后台核对一致。若历史报表中的 MSKU 与 ASIN 冲突、多重映射或活动总额未核平，对应行仍未归属。当前费用仍不能单独作为模板总花费，因为 SB 覆盖和跨类型总额完整性尚未通过。
- `hsaProductAds` 和 `sdProductAds` 的用户提供 Schema 已接入只读当前目录分页查询；SB 顶层 ASIN、SD 的 MSKU/ASIN 与广告 ID 仅用于预览候选及交叉核验，不被误当作历史目录。`sbDivideAsinReports` 的表格与示例字段错位，`spends`/`percent` 数组没有元素结构或金额语义，因此本地在发请求前阻止调用。继续自动归属 SB 所缺资料列在 `docs/AD_ATTRIBUTION_ENGINE.md`。
- 销量预览中可提示所选窗口内最早的明确正销量日，但这不会参与 D01 设置。产品创建日期和订单接口也不会覆盖用户的手动确认。查询日期表示请求已成功执行，不证明空白指标为零，也不证明最近一天已经结算。

模板来自用户上传的《新品推广周报.xlsx》：第 5 行表头，D01–D60 在第 6–65 行，24 列，无公式。样例 SKU 为 ex3107；生成新 SKU 时会清空其样例数据。模板没有“日销售额”列，销售金额会写入校验日志，模板中的“售价”只写来源售价或根据销售额/销量计算的单价，不会把销售金额误填进售价列。原模板没有历史类目、排名、评分、评论数的日期字段，这些值会保持空白，避免把当前快照伪装成历史数据。

## Windows 安装与启动

安装 Python 3.12（勾选 Tcl/Tk 和 Python Launcher），然后在项目目录双击 `run_gui.bat`。脚本会创建 `.venv`、安装 `requirements.txt` 并启动中文界面。测试可运行 `run_tests.bat`。

也可以在 PowerShell 执行：

```powershell
py -3.12 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m amazon_weekly_report gui
```

命令行查看能力：

```powershell
.venv\Scripts\python -m amazon_weekly_report --help
.venv\Scripts\python -m amazon_weekly_report inspect "C:\领星导出文件夹"
.venv\Scripts\python -m amazon_weekly_report inspect-template "C:\新品推广周报.xlsx"
.venv\Scripts\python -m unittest discover -s tests -v
```

## 领星 MCP 本机配置

在领星官方帮助文档取得你账号可用的 Streamable HTTP 地址和 `X-Mcp-Key`。程序不预设未知的服务 URL，也不使用社区 MCP 服务冒充领星官方服务。

打开界面的“领星 MCP 连接”页，填写官方地址和密钥，点击“安全保存本机凭证”。密钥使用 Windows Credential Manager 保存；`%APPDATA%\MyAmazonProject\settings.json` 只存服务地址。密钥不会写入仓库、命令行参数或程序日志。重置密钥时可在 Windows“凭据管理器”中删除 `my-amazon-project.lingxing-mcp` 下的 `X-Mcp-Key`。

之后可依次使用“测试连接”和“获取工具列表”。工具列表显示每个工具的名称、描述、annotations 与完整 inputSchema，不再只展示被 `readOnlyHint=true` 标注的工具。

“搜索业务工具”只会调用名称为 `search` 且工具说明明确表明它用于搜索业务工具目录/Schema 的发现工具；发现调用不会被当作业务查询。填写该工具 inputSchema 要求的 JSON 参数后，程序显示其返回的业务工具 ID、描述和 Schema。选择一项后可查看完整参数，并点击“核验本次只读范围”。只有描述明确属于查询且没有明显写入操作的候选工具可以进入人工核验。许可只在当前程序会话内有效，并绑定当次搜索返回的业务工具 ID、描述与 Schema 指纹。

“执行已验证的只读查询”会重新调用 `search`，确认 ID 和 Schema 未改变，按业务 Schema 校验查询参数，再通过 `action` 的 inputSchema 构造调用。缺少 search 目录说明、Schema 不支持校验、工具描述含有写入操作、许可不匹配或返回 Schema 变化时，程序会拒绝执行。未经核验或无法确认安全性的 action 不会发送。普通 MCP 工具仍仅允许调用 `readOnlyHint=true` 的工具。所有查询结果仅在本机窗口显示，并对凭证字段做隐藏；不会写入日志或工作簿。

更新并验证本机 MCP 时：先运行 `git switch feature/weekly-report-v1` 和 `git pull --ff-only origin feature/weekly-report-v1`，再启动 `run_gui.bat`。在“领星 MCP 连接”页获取工具列表并检查元数据；按官方 Schema 搜索业务工具，并仅在确认只读后核验。验证 EX 店铺的 `ex3106` 时，按该业务 Schema 查询 2026-08-04 至 2026-08-06。若结果没有按日期和 SKU 唯一归属的“活动实际花费”，或可能把共享活动费用重复计到 SKU，不要把该字段用于周报；需与领星原始报表交叉核对。

命令行连接测试也只从 Windows Credential Manager 读取凭证：

```powershell
.venv\Scripts\python -m amazon_weekly_report mcp-test
.venv\Scripts\python -m amazon_weekly_report mcp-tools
.venv\Scripts\python -m amazon_weekly_report mcp-query 工具名称 --arguments "{}"
```

Codex 云端没有 Windows 本机 MCP 会话、endpoint 或密钥，未连接你的真实账号，也没有验证 ex3106 数据。请在 Windows 中查看领星返回的三项基础工具及搜索结果，依据领星官方说明核验只读工具后执行查询。实际销量、销售额和广告活动实际花费只有在 MCP 返回字段与日期/SKU 归属得到验证后才能用于正式周报；尤其不能将共享活动总花费重复归入多个 SKU。完成本机实测前，周报仍使用下述 Excel/CSV 导入流程。

## 领星 OpenAPI 本机测试

OpenAPI 使用固定官方域名 `https://openapi.lingxing.com`。Token 请求为 `POST /api/auth-server/oauth/access-token`，以 multipart/form-data 发送 `appId` 和 `appSecret`。在“领星 OpenAPI 连接”页安全保存凭证后，界面只显示身份验证结果和响应中可识别的有效期，不显示 Access Token。

已根据本次上传的领星 **Java** SDK 源码实现签名和只读店铺列表请求：`GET /erp/sc/data/seller/lists`。签名按 ASCII 参数名排序、使用 SDK 的字符串转换/去空格规则拼接 `key=value`、计算大写 MD5，再用 AppID 执行 AES/ECB/PKCS5Padding，密文使用标准 Base64，并按 Java `URLEncoder` 的 UTF-8 规则编码查询参数。Python SDK 并未包含在压缩包内；本项目只移植与签名相关的 SDK 行为，没有复制其会输出签名输入的日志代码。签名向量已用独立 OpenSSL 结果核对，但当前环境没有 Java 编译器，未执行官方 Java SDK。

店铺列表仅在业务 `code` 为整数 `0`、`data` 为直接数组且信封和每条店铺记录类型有效时才视为成功。名称、SID、站点会动态显示在 GUI 的店铺列表和周报页下拉框中；下拉框选项在内存中关联接口返回的 SID，不硬编码店铺。`status` 和 `has_ads_setting` 只显示接口返回的原始整数，不推断其业务含义。用户已确认 Windows 动态店铺列表真实验收通过；自动化测试仍只使用合成数据。

### Windows 验收记录

用户确认最新版在 Windows 本地成功刷新真实店铺列表，并在 OpenAPI 页和周报页下拉框正确显示。云端未访问用户凭证或业务数据。不要将响应正文、Seller ID、SID、Token 或请求记录回传到聊天。

云端自动化测试使用 HTTP 模拟夹具，仅验证请求构造、Token 生命周期和错误脱敏；不代表真实 API 或业务数据验收。不要将 AppID、AppSecret、Access Token、完整签名 URL 或真实报表发到聊天或 GitHub。

### SKU 识别与每日销量接口

根据用户提供的领星官方 `AsinListNew` 与 `AsinDailyLists` 字段说明实现：

- 产品表现：`POST /bd/productPerformance/openApi/asinList`。以 `search_field=local_sku` 搜索 1–50 个本地 SKU，`summary_field=sku`，单店铺 `sid` 按文档作为字符串发送，日期为 `YYYY-MM-DD` 双闭区间且最多 92 天。`offset`/`length` 分页，单页长度不超过 10000，按 `data.total` 完成分页。只有 `code=0`、`data.total` 与 `data.list` 结构有效才接受结果。
- 映射字段：顶层 `sku` 匹配请求的 local SKU；`price_list.local_sku/seller_sku/sid` 提供此店铺的 MSKU；`asins.asin/sid` 提供此店铺的 ASIN 候选。文档示例中嵌套 SID 为数字字符串，解析时仅接受整数字符串或整数。ASIN 数组和 MSKU 数组没有定义一一关系，所以不自行配对。
- 每日销量统计：`POST /erp/sc/data/sales_report/asinDailyLists`，以 `sid` 整数、`event_date` 站点日期、`asin_type=2`（MSKU）查询。`type=1/2/3` 分别读取销售额、销量、订单量；`offset` 从 0 开始，`length` 不超过 1000，逐页读到短页结束。只接受 `code=0` 和 `data` 直接数组，并核对返回 SID、站点日期和 MSKU。
- 用户提供的两个页面没有完整历史首单日期字段/覆盖声明。`product_create_time` 表示商品创建时间，不是首单时间。因此界面只列出本次输入范围内最早的明确正销量日，不会据此填写 D01。
- 用户提供的 `AsinDailyLists` 字段说明未定义取消/退款状态筛选或销售额净额算法；这些口径必须通过 Windows 与领星同店铺同日期的报表对照确认，完成前预览数字不能进入正式周报。
- 两个接口均只读。查询金额使用 `Decimal`，请求和响应不写日志或工作簿；错误信息不包含 Token、签名或完整响应。Product Performance 的单店查询间隔按官方说明至少 1 秒。每日统计页面只给出令牌桶容量 5、未给出补充速率；客户端串行执行，收到 HTTP 429 时停止且不自动重发。若响应带整数 `Retry-After`，错误提示会显示建议等待秒数。
- 合成 HTTP 测试覆盖 POST 签名、参数结构、SKU 映射、站点日期、显式零/缺失、分页及限流停止；这些测试不能证明领星真实业务请求成功。

云端网络策略没有放行 `apidoc.lingxing.com`；以上参数来自用户提供的官方页面内容，未尝试绕过网络限制，也未使用真实凭证或企业记录。

### Windows SKU/销量预览测试

1. 在项目目录更新分支并启动：

   ```powershell
   git switch feature/weekly-report-v1
   git pull --ff-only origin feature/weekly-report-v1
   run_gui.bat
   ```

2. 确认原有 OpenAPI 凭证仍保存在 Windows Credential Manager；切到“生成周报”页，点击“刷新店铺列表”，选择要查询的店铺。
3. 输入本地 SKU 和 `YYYY-MM-DD` 站点日期范围。首次实测建议先查 1 天，以便验证映射和三种每日统计；当前程序不会因 429 自动重放请求。
4. 点击“查询并预览真实数据”，对照领星相同店铺/站点日期的原始报表，检查 local SKU、MSKU、ASIN、销量、订单量、销售额和币种。表格空白表示接口没有返回该指标；显式零值显示为 `0`。
5. 只反馈“本地校验通过/失败、MCP/OpenAPI 调用是否成功、业务查询是否成功、是否返回数据及脱敏字段结构/是否与后台一致”。不要回传密钥、Token、完整响应、Seller ID 或订单/商品明细。查询结果只供预览，尚未写入正式周报。

### Windows OpenAPI 销售数据生成测试

当前分支可将领星每日销售结果送入原模板。完成一次真实数据测试前，生成文件只作为验收样本，不作为已经验收的正式周报。

1. 选择已核实的店铺、模板和输出文件夹；输入 SKU。
2. 在“D01 起算日期（手动确认）”中按每行 `SKU=YYYY-MM-DD` 填入该 SKU 的 D01，点击“确认并保存 D01”。日期按所选店铺站点日期解释，并保存在本机。切换店铺或 SKU 后会载入对应设置；没有保存 D01 的 SKU 不能生成。
3. 填查询开始/结束日（站点日期）。开始日须不晚于最早 D01；结束日须晚于或等于所有 SKU 的 D01。超过 92 天的日期范围会分段查询，接口限流、映射不全或任何分段失败都会中止。
4. 点击“领星 OpenAPI 查询并生成”。对照领星同店铺/站点日期的真实记录，检查日销、单价及校验日志里的每日销售额/总订单量；无记录仍为空白，只有接口明确返回的零才显示 `0`。总订单量和销售额在现有模板没有同义字段时只写入校验日志。
5. 检查 D61+、周次、原模板样式，以及 `A2` 推广阶段、`G2` 发货时间、`M` 列日目标销量和已有人工内容。广告预览当前不会写入周报，本阶段生成状态仍为“部分完成”。
6. 更新已有周报时选择原周报，输出必须是新文件。真实验收只回报每个数据阶段成功/失败、是否与后台一致及脱敏错误，不发送凭证或原始业务明细。

### SP、SB、SD 广告费用预览

选择店铺和 SKU，输入查询日期后点击“查询 SP / SB / SD 广告费用并预览”。查询按天读取各广告类型的活动报表及商品/广告组报表；广告报表使用 `X-API-VERSION: 2` 和 offset 分页，SB 创意目录按官方 next_token/offset 分页。SP、SD 商品行按报表日的 MSKU 和/或 ASIN 做忽略大小写的精确映射；两项标识指向不同本地 SKU、共享标识非唯一或活动成本不相等时，相关行不会归属。SD 历史 `sdProductAdReports` 是指定日期的商品费用依据，直接使用其 SKU、ASIN、广告 ID、活动 ID 和 cost 做匹配；`sdProductAds` 当前目录只是补充诊断，不会决定历史归属，也不会阻断历史报表查询。广告活动名称和广告归因销售额不参与映射。SB 当前创意目录的 ASIN 和广告组/活动/创意 ID 会显示成候选关系，但接口不提供报表日历史快照，所以 SB 活动费用仍保持未归属。

活动报表 cost 是核对基准；汇总不会把活动 cost 再加到商品/广告组 cost。未归属的 SP、SB、SD 金额按类型分开显示。官方 Schema 未说明统一 cost 币种与 `report_date` 站点日期对账方式，因此未经完整核验时，广告总花费留空；SB 视频广告覆盖及历史创意映射仍待官方确认。`sbDivideAsinReports` 由于返回字段结构不可靠，不发送请求或猜测金额。空数据、查询失败、费用差异、冲突映射、未映射和行内明确 cost=0 有不同状态。用户已在 Windows 对 SP 与部分日期的 SD 历史费用进行后台核对；SB 的 SKU 级数值也已确认，但软件尚未自动归属。具体店铺、SKU、日期和金额不记录在仓库。广告预览不写正式周报；销售数据可独立生成周报。详情设计见 [广告自动归属规则](docs/AD_ATTRIBUTION_ENGINE.md)。

## 配置 Excel/CSV 字段映射

复制 `config/mapping.example.json` 到 `config/mapping.local.json`。本地映射文件已加入 `.gitignore`。根据领星真实导出文件填写：

- `header_row`：表头实际行号；`sheet_name`：XLSX 工作表（可留空使用首张表）。
- `file_patterns`：相对数据文件夹的文件名模式，例如 `*销售*.xlsx`；不要使用 `..` 或绝对路径。
- `source_headers`：源文件的完整列名。CSV 日期若不是 ISO 格式，填写真实日期格式，例如 `%m/%d/%Y`。
- `actual_ad_spend.source_headers.sku` 必须映射到报表中能唯一对应 SKU 的字段；活动级总花费重复出现在多个 SKU 的报表不得标为 `SKU_LEVEL_CONFIRMED`。
- 如果还没有广告实际花费导出，可先把 `status` 设为 `SALES_ONLY_READY` 并保留广告花费映射空白。对应花费单元格会留空，日志会标明数据缺失；这不代表花费为 0。
- 样本数据确认后，有广告实际花费映射的配置才可设为 `READY_FOR_REVIEW`，且 `sku_scope` 必须是 `SKU_LEVEL_CONFIRMED`。若源文件只含已核实的部分广告类型，设置 `coverage_scope=PARTIAL_OR_UNKNOWN`：来源金额保留在本地校验记录中，但模板总广告花费/ACOS 必须空白。只有人工核对源文件对相同日期和 SKU 包含所需全部广告类型（SP、SB、SD，及所需视频广告范围）后，才可设置 `coverage_scope=ALL_AD_TYPES_TOTAL_CONFIRMED`。这是用户对源报表粒度和覆盖范围的确认，不是程序自动推断。

CSV 支持 UTF-8、UTF-8 BOM 和 GB18030；XLSX/XLSM 支持明确指定工作表和表头行。旧版 XLS 请先另存为 XLSX。

## 生成周报

图形界面进行 Excel/CSV 备用流程时，选择格式模板、导出目录、映射文件、SKU、D01 规则和输出目录后点击“开始生成周报”。OpenAPI 主流程见上面的 Windows OpenAPI 销售数据生成测试。也可通过命令行运行离线备用流程：

```powershell
.venv\Scripts\python -m amazon_weekly_report generate `
  --template "C:\新品推广周报.xlsx" `
  --data "C:\领星导出数据" `
  --mapping "C:\配置\mapping.local.json" `
  --skus "SKU_A,SKU_B" `
  --d01 SKU_A=2026-08-01 --d01 SKU_B=2026-08-08 --confirm-d01 `
  --output "C:\输出\新品推广周报_SKU_A_SKU_B.xlsx"
```

若 `--existing` 工作簿中的 D01 与本次输入不同，还必须增加 `--confirm-d01-change`，明确接受日期轴变化风险。

增量更新时另选已有周报，输出仍必须是新文件，不覆盖来源：

```powershell
.venv\Scripts\python -m amazon_weekly_report generate `
  --template "C:\新品推广周报.xlsx" --existing "C:\上周报.xlsx" `
  --data "C:\新增领星导出" --mapping "C:\配置\mapping.local.json" `
  --skus "SKU_A" --d01 SKU_A=2026-08-01 --confirm-d01 `
  --output "C:\输出\新品推广周报_更新.xlsx"
```

若更新时 D01 与已有周报的 D01 不同，须再加入 `--confirm-d01-change`。命令行不保存 D01 设置；持久化设置通过 GUI 按店铺和 SKU 保存。

销量与销售金额来自真实销售来源；销量写入“日销”，售价写来源单价或销售额除以销量。广告订单、广告销售、曝光、点击和流量只有在相应报表及字段已映射时才填写。比率在分母为零或数据缺失时留空。模板中日销售额没有独立列，所以销售金额总和记录在校验日志中。

没有记录不等于 0；空白不会被填成 0。D01 与覆盖日期之间仍会生成连续自然日，缺少来源记录的日期指标留空。多店铺合计时，如果某 SKU 的日期行未覆盖该 SKU 在来源中出现过的所有店铺，相关日汇总会留空并提示。相同店铺/SKU/日期的重复记录不会默认累加；广告记录可用活动 ID 和广告组 ID 区分，无法区分的重复会被排除并提示。广告关闭后异常花费只会在源记录提供关闭日期时检测，实际花费值不会被程序擅自删除。

生成文件旁会写入 `*_数据校验.jsonl`。该日志不是原始业务明细，不包含 MCP 凭证或查询结果。输入数据和输出目录建议放在本机私有目录。

## 在 Windows 打包 EXE

在 Windows 项目目录运行：

```powershell
.venv\Scripts\python -m pip install pyinstaller
.venv\Scripts\python -m PyInstaller --clean --noconfirm --windowed `
  --name AmazonWeeklyReport --collect-all keyring `
  --hidden-import keyring.backends.Windows windows_app.py
```

生成目录为 `dist\AmazonWeeklyReport`。当前交付是源代码和打包命令，并未在 Linux 云环境中声称生成 Windows EXE。

## 验证范围与待实测项目

`tests/` 使用合成夹具验证 D61 扩行、手动字段保留、店铺/SKU 级 D01 保存、D01 日期变更确认、SP/SB/SD 广告费用核对、分页、错误与数据缺失状态、MCP 协议、OpenAPI 请求签名、商品映射、每日查询及 API→Excel 适配。用户报告 Windows 动态店铺列表、每日销售查询以及 SP 和指定日期 SD 历史商品花费真实验收通过；本分支 D01 持久化、完整的 API→Excel Windows 操作、SB 自动归属和跨类型广告总额仍待验收。广告报表币种、日期时区、SB 视频覆盖待确认；广告正式写入、EXE、可信版本清单及第二台 Windows 电脑验收尚未完成。云端模拟测试不代表真实数据验收。
