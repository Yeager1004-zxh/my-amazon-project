# 新品推广周报 V1.0 实施清单

V1.0 只实现新品推广周报。当前分支为 `feature/weekly-report-v1`，所有开发在此分支进行；`main` 保持稳定。

## 数据接口路线

1. **领星 OpenAPI：长期主数据接口。** 负责授权后的店铺、SKU 映射、销售与广告数据读取。
2. **领星官方 MCP：辅助查询与交叉核验。** 保留现有 Streamable HTTP、Schema 发现、只读授权、工具 ID 与 Schema 指纹绑定；不要求日常用户填写 Tool ID/JSON。
3. **领星 Excel/CSV：应急备用。** 继续复用现有导入映射与验证流程。

三类来源应接入同一标准日数据模型，再调用已有周报生成器。只有官方接口口径和实体归属均已验证的数据才能成为正式输入。共享广告活动费用不可分摊或重复计入 SKU。

## 当前基线

- 已有：OpenAPI multipart Token 请求、Windows Credential Manager 凭证保存、连接测试 UI、按用户上传的领星 Java SDK 源码实现的签名和店铺列表只读请求；MCP 客户端与安全 Schema 校验；Excel/CSV 导入、D01、D61+、数据校验日志和模板生成。
- 本阶段新增：OpenAPI 客户端仅在 Token 响应给出明确有效期时进行内存缓存；SellerLists 严格检查 `code=0`、响应信封和 `data` 数组；校验确认的店铺字段类型并动态映射到 GUI 店铺表与周报页下拉框。`status`、`has_ads_setting` 保持原始整数，不推断业务含义。
- 用户已确认最新版在 Windows 完成真实店铺列表验证：授权店铺动态显示在 OpenAPI 列表及周报页下拉框中。尚未完成：SKU/ASIN/MSKU 映射、首单日期/每日销量/销售额查询、实际广告花费查询、OpenAPI/MCP 统一数据层以及真实接口到 Excel 的集成。当前 SDK 是 Java，不是 Python；源代码已检查但由于云端未安装 Java 编译器，没有运行官方 SDK。
- 云端夹具测试只验证程序行为，不表示真实接口或真实数据验收。不得用模拟业务结果填正式周报。

## 阶段和验收标准

| 阶段 | 范围与输入 → 输出 | 依赖 | 验收方式与完成条件 |
|---|---|---|---|
| 1. OpenAPI 基础连接 | 官方认证/签名规范与凭证 → Token 生命周期、已签名只读请求、店铺列表 | 用户提供的 Java SDK 源码/POM；店铺列表响应字段结构 | 单测覆盖 SDK 源码定义的序列化、加密和 URL 编码并与独立 OpenSSL 向量比较；合成响应覆盖业务码和字段校验；用户已确认 Windows 真实店铺列表及动态下拉框通过。 |
| 2. 商品与销售 | 店铺、local SKU → ASIN/MSKU → 历史首单日期和每日销量/销售额 | 官方端点参数、分页、日期/时区、订单状态及响应字段 | mock 覆盖分页/重复/缺失/限流；Windows 使用 ex3106 对照领星真实报表，确认首单日期、每日销量和销售额一致。 |
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
4. 阶段 2 等待商品表现、每日销量和首单日期接口 Schema 核验后实施；云端不尝试发送未核验的业务请求。

## 第二阶段：SKU 识别与每日销量

- 商品表现（新版 / ASIN 列表）：[AsinListNew](https://apidoc.lingxing.com/new/api-docs/Statistics/AsinListNew/)，核对 local SKU/MSKU/ASIN 关系、店铺过滤参数及分页。
- 亚马逊 ASIN 每日销量统计：[AsinDailyLists](https://apidoc.lingxing.com/new/api-docs/Statistics/AsinDailyLists/)，核对日期过滤、销量/销售额字段、时区、分页及业务状态/退款口径。
- 若以上接口不提供完整历史首单日期与历史覆盖范围，还需要领星官方首单/销量明细查询接口文档，不能用查询窗口中最早一条记录代替真实 D01。

请提供官方页面 PDF/HTML 或粘贴脱敏的参数、分页、响应字段、JSON 示例和字段定义。必须能确认 method/path、参数名与类型、分页终止条件、业务成功码、SKU/MSKU/ASIN 归属、销量与订单量区别、销售额口径、时区、退款/取消处理及历史可查询范围。不要提交真实 SID、Seller ID、商品/订单记录、AppID、AppSecret 或 Token。

Codex 云端当前的受限网络允许列表不含 `apidoc.lingxing.com`，无法直接读取上述页面；未绕过网络策略，也未对未核验的请求发起调用。

## Windows 商品与销售本地验收（待 Schema 核验并实现后）

完成 Schema 核验和本阶段实现后，再用已配置的 Windows Credential Manager 凭证，在动态店铺列表中选择店铺并输入测试 SKU。先查看日期排序的身份识别与每日指标结果，对照领星后台相同店铺、相同日期范围和统计口径；暂不写入正式周报。只反馈成功/失败、脱敏字段结构及是否一致，不回传原始业务记录或任何密钥。

真实验收必须在用户 Windows 环境完成。云端模拟测试、Token 单测或 SDK 代码审阅均不能替代真实业务数据对照。
