# 新品推广周报 V1.0 实施清单

V1.0 只实现新品推广周报。当前分支为 `feature/weekly-report-v1`，所有开发在此分支进行；`main` 保持稳定。

## 数据接口路线

1. **领星 OpenAPI：长期主数据接口。** 负责授权后的店铺、SKU 映射、销售与广告数据读取。
2. **领星官方 MCP：辅助查询与交叉核验。** 保留现有 Streamable HTTP、Schema 发现、只读授权、工具 ID 与 Schema 指纹绑定；不要求日常用户填写 Tool ID/JSON。
3. **领星 Excel/CSV：应急备用。** 继续复用现有导入映射与验证流程。

三类来源应接入同一标准日数据模型，再调用已有周报生成器。只有官方接口口径和实体归属均已验证的数据才能成为正式输入。共享广告活动费用不可分摊或重复计入 SKU。

## 当前基线

- 已有：OpenAPI multipart Token 请求、Windows Credential Manager 凭证保存、连接测试 UI、按用户上传的领星 Java SDK 源码实现的签名和店铺列表只读请求；MCP 客户端与安全 Schema 校验；Excel/CSV 导入、D01、D61+、数据校验日志和模板生成。
- 本阶段新增：在用户提供的 `AsinListNew` 与 `AsinDailyLists` 官方 Schema 后，实现 local SKU→MSKU/ASIN 候选识别、每日销售额/销量/订单量只读查询、偏移分页、业务码/字段校验、站点日期检查、精确金额解析以及缺失值与零值区分。Windows GUI 的“生成周报”页可显示日期排序的预览，但预览结果不写入工作簿。
- 用户已确认最新版在 Windows 完成真实店铺列表验证：授权店铺动态显示在 OpenAPI 列表及周报页下拉框中。SKU/销量/销售额尚未完成 Windows 真实业务对照；产品创建时间不是历史首单日期，当前查询范围内最早正销量日不得冒充真实 D01。尚未完成实际广告花费查询、统一数据层、真实接口到 Excel 集成。当前 SDK 是 Java，不是 Python；签名相关源码已检查但云端未安装 Java 编译器，没有执行官方 SDK。
- 云端夹具测试只验证程序行为，不表示真实接口或真实数据验收。不得用模拟业务结果填正式周报。

## 阶段和验收标准

| 阶段 | 范围与输入 → 输出 | 依赖 | 验收方式与完成条件 |
|---|---|---|---|
| 1. OpenAPI 基础连接 | 官方认证/签名规范与凭证 → Token 生命周期、已签名只读请求、店铺列表 | 用户提供的 Java SDK 源码/POM；店铺列表响应字段结构 | 单测覆盖 SDK 源码定义的序列化、加密和 URL 编码并与独立 OpenSSL 向量比较；合成响应覆盖业务码和字段校验；用户已确认 Windows 真实店铺列表及动态下拉框通过。 |
| 2. 商品与销售 | 动态店铺、local SKU、站点日期范围 → MSKU/ASIN 候选和每日销售额/销量/订单量预览 | 用户提供的 AsinListNew 与 AsinDailyLists Schema；首单日期仍需覆盖范围定义 | 已实现 API 解析、请求、分页、严格归属检查和 Tkinter 预览；合成测试通过。待 Windows 对照真实销售报表。完整历史首单需另有官方接口或确认历史覆盖。 |
| 3. 广告实际花费 | 广告报表/活动记录 → 日期、活动、广告组、SKU/ASIN 唯一归属的实际花费 | 官方广告接口 Schema 与口径；ex3106 原始记录 | 核对 2026-08-04 至 2026-08-10；活动关闭后的记录仍按事实核验；共享活动无法唯一归属时留空，不分摊、不重复累计。 |
| 4. 统一数据层 | OpenAPI、MCP、Excel 输入 → 同一 SKU 日记录及来源/口径/日期/完整性/归属元数据 | 阶段 1–3 字段定义 | Decimal 金额；零值与缺失区分；重复、跨店覆盖、时区、活动归属均有诊断；MCP 仅辅助核验，不混用不同统计口径。 |
| 5. 周报集成 | 已校验的标准日记录 + 模板/已有周报 → 新工作簿与脱敏日志 | 阶段 4；现有 `report_generation.py`、模板扩展 | ex3106 后 ex3106+ex3114 本机对照；D01、连续日期、D61+、周次、格式、公式、合并单元格和人工字段保护通过回归。缺失值空白并标注部分完成。 |
| 6. Windows 流程与打包 | 店铺、SKU → 进度、输出、诊断和 EXE | 阶段 5；Credential Manager、Tk/Tcl、PyInstaller | 普通界面无需 JSON/Schema/Tool ID；Windows 10/11 干净环境和真实数据端到端验收通过后才发布 V1.0。 |

## 第一阶段执行任务

Token 请求为 `POST /api/auth-server/oauth/access-token`，使用 multipart/form-data 发送 `appId` 与 `appSecret`；用户已在 Windows 确认 Token 获取成功。店铺列表 GET 请求和签名已实现：Java SDK 的 `ApiSign`/AES 源码定义 ASCII 排序、`key=value` 连接、大写 MD5、AES/ECB/PKCS5Padding、标准 Base64；SDK 请求层使用 UTF-8 `URLEncoder`。本地签名测试向量与 OpenSSL 独立计算结果一致。SDK 样例把 AppID 传入 AES 密钥参数。当前云端未执行 Java SDK（没有 `javac`），也未能访问官方文档站点；SellerLists 字段结构来自用户在 Windows 的实际脱敏检查。HTTP 2xx JSON 本身不视为业务成功，仍需检查业务码和响应结构。

实现顺序：

1. 从用户上传的 Java SDK 源码移植并测试签名；覆盖 ASCII 排序、字符串 trim、MD5 大写、AES 填充、Base64 密文和 Java URL 编码。云端已通过固定测试向量及独立 OpenSSL 对照；该测试不等于官方 Java SDK 执行或真实请求验收。
2. 按用户提供的脱敏真实字段结构解析 SellerLists：成功码必须是整数 `0`，`data` 必须是直接店铺数组，店铺字段严格按确认类型解析；HTTP 2xx 不会单独被视为业务成功。通过全合成数据测试正常、空列表、业务失败和错误类型。
3. Windows 刷新店铺列表后，GUI 动态显示名称、SID、站点和两个状态字段的原始值；周报页下拉框绑定实际返回 SID。用户已确认最新版 Windows 真实验收通过。
4. 阶段 2 接口参数来自用户提供的官方文档内容；真实业务请求仍只在用户 Windows 授权环境执行。

## 第二阶段：SKU 识别与每日销量

已依据用户提供的官方 Schema 实现：

- 商品表现：[AsinListNew](https://apidoc.lingxing.com/new/api-docs/Statistics/AsinListNew/)，POST `/bd/productPerformance/openApi/asinList`。用 `search_field=local_sku` 搜索 1–50 个本地 SKU，`summary_field=sku`，单店铺 `sid` 按文档作为字符串发送，查询日期为双闭区间且最多 92 天。使用 `offset`/`length` 分页，单页长度不超过 10000，按 `data.total` 完成分页。只有 `code=0`、`data.total` 和 `data.list` 结构有效才接受结果。
- 映射：顶层 `sku` 匹配 local SKU；`price_list.local_sku/seller_sku/sid` 提供本店 MSKU；`asins.asin/sid` 提供本店 ASIN 候选。文档没有说明 ASIN 与 MSKU 一一对应，客户端不会自行配对。
- 每日销量：[AsinDailyLists](https://apidoc.lingxing.com/new/api-docs/Statistics/AsinDailyLists/)，POST `/erp/sc/data/sales_report/asinDailyLists`。使用 `asin_type=2` 按 MSKU 查询；`type=1/2/3` 分别获取销售额、销量、订单量。日期按站点 `YYYY-MM-DD` 原样传递，不做 UTC 偏移；响应按业务 `code=0`、直接 `data` 数组、SID、日期及 MSKU 校验。
- 产品表现日期区间最多 92 日、搜索最多 50 个值、单页 `length` 上限 10000；每日统计 `length` 上限 1000。金额使用 Decimal。没有记录的指标为空，只有显式 `map_value=0` 才是零。分页以短页结束。每日接口只公布桶容量 5，未公布 refill 速率；请求串行，遇到 429 即停止且不自动重放，避免猜测等待时间。
- 用户提供的每日接口片段没有定义取消/退款筛选和销售额净额算法；须在 Windows 用相同店铺、日期和后台报表核对后，才可把预览列纳入正式周报映射。
- GUI 在“生成周报”页显示日期排序的 SKU/MSKU/ASIN 和每日指标预览，不写入工作簿。范围内最早明确正销量日期会注明所查范围，不会冒充完整历史首单。

合成测试覆盖 POST 签名、参数结构、映射解析、业务码、站点日期、显式零与缺失、两种分页和限流停止。用户 Windows 本地的 OpenAPI 店铺列表已验收；SKU/销量/销售额仍待 Windows 对照真实报表。初次 API 测试建议使用一天范围；核对同一店铺、站点日期、MSKU 与销售统计口径后，只回报成功/失败、各调用阶段是否成功、是否返回数据、脱敏字段结构和是否与后台一致。不要回传原始业务明细或密钥。

当前两页未定义完整历史首单日期或历史覆盖范围。`product_create_time` 是商品创建时间，不是首单日期；范围内最早正销量日不能直接作为 D01。要自动填真实首单 D01，仍需领星官方首单接口或能确认完整覆盖的官方数据来源说明。

云端网络策略不允许访问 `apidoc.lingxing.com`，本阶段只依据用户提供的官方页面内容，未绕过网络限制，也没有使用真实凭证或发送真实业务查询。真实验收必须在用户 Windows 环境完成。
