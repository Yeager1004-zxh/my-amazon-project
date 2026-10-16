# 亚马逊新品推广周报工具

面向 Windows 10/11 的本地新品推广周报工具。V1.0 数据路线为：领星 OpenAPI 作为主数据接口、官方 MCP 作为辅助查询/交叉核验、领星 Excel/CSV 导入作为应急备用；三种来源最终进入统一标准数据层，再复用原周报模板生成器。当前离线 Excel/CSV 周报功能可用，OpenAPI 业务查询和自动写入尚未通过真实接口验收。

V1.0 分阶段任务、模块输入/输出、依赖、测试和验收门槛见 [正式实施清单](docs/V1_IMPLEMENTATION_PLAN.md)。当前版本仅聚焦新品推广周报。

## 当前实现

- 延续原模板七天一周的样式、周次、边框和条件格式，按实际来源覆盖日期扩展 D61、D62……；自然日连续排列，来源缺失的日期指标保持空白。
- 支持多个 SKU、多个来源文件和多个店铺；字段必须在本机 JSON 映射中明确指定，不猜报表表头。
- 新 SKU 页签会清除模板样例指标。增量更新时，实际来源有值的指标会更新；来源没有记录的旧指标保留。`A2` 推广阶段、`G2` 发货时间和 `M` 列日目标销量由用户维护，程序不会覆盖。
- `first_sale` 模式以真实销售来源中首个正销量日期作为 D01；`launch_date` 模式要求按 SKU 填正式开售日期。
- 销量、销售金额、广告实际花费、广告表现和流量可按真实字段映射。广告花费必须使用实际花费字段；只有在用户将 `sku_scope` 明确设为 `SKU_LEVEL_CONFIRMED` 后才接受该映射。重复、跨 SKU 共享的同一活动花费或店铺覆盖不完整的记录会从相关日汇总中排除并写入校验日志。
- 校验日志为 JSONL，记录来源文件名、SKU 的 D01、最新数据覆盖日、销量/销售额/广告花费汇总和异常提示。关闭日期后的实际花费只有在来源提供并映射 `campaign_disabled_at` 时才提示。
- Windows 本地 MCP 客户端支持 MCP 初始化连接、工具列表分页、工具目录发现，以及按业务 Schema 核验的只读 action 查询。请求带 `X-Mcp-Key`、MCP session 和协议版本；常规工具仍要求 `readOnlyHint=true`。
- OpenAPI Token 请求使用 multipart/form-data；Access Token 仅在本机内存中按响应明确提供的有效期缓存，并在有效期临近时刷新。未返回可识别有效期的 Token 不缓存。

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

OpenAPI 使用固定官方域名 `https://openapi.lingxing.com`。目前已实现 Token 请求：`POST /api/auth-server/oauth/access-token`，以 multipart/form-data 发送 `appId` 和 `appSecret`。在“领星 OpenAPI 连接”页分别填写两项凭证并点击“安全保存凭证”；凭证保存在 Windows Credential Manager 的 `my-amazon-project.lingxing-openapi` 服务项中，不写入配置文件、命令行或日志。之后点击“测试 API 身份验证”，界面只显示成功状态和响应中可识别的有效期，不显示 Access Token。

当前店铺列表接口 `GET /erp/sc/data/seller/lists` 仍保持 fail-closed：虽已知接口路径和鉴权参数名，但签名输出编码尚未从官方 SDK 核实，店铺响应字段也待确认。SKU→ASIN/MSKU 映射、首单日期、每日销量/销售额的接口路径、参数、分页和字段口径尚未拿到官方定义。因此“获取店铺列表”和“测试 SKU 数据”不会发送未核验请求，也不会读模拟结果或写入正式周报。广告接口未接入。

### 第一阶段 Windows 首测

1. 更新开发分支并启动：`git switch feature/weekly-report-v1`、`git pull --ff-only origin feature/weekly-report-v1`、`run_gui.bat`。
2. 在“领星 OpenAPI 连接”页安全保存 AppID/AppSecret（只写入 Windows Credential Manager）。
3. 点击“测试 API 身份验证”，确认结果只显示成功/错误和有效期，不显示 Token。请求需从已加入领星 IP 白名单的 Windows 公网出口发出。
4. 店铺列表按钮仍会拒绝发送，直到签名 SDK 的准确编码和店铺列表响应定义完成核验。此提示是安全阻断，不表示真实店铺查询失败。

云端自动化测试使用 HTTP 模拟夹具，仅验证请求构造、Token 生命周期和错误脱敏；不代表真实 API 或业务数据验收。不要将 AppID、AppSecret、Access Token 或真实报表发到聊天或 GitHub。

### 仍需的领星官方资料

- 官方 Python SDK 下载/代码地址或签名函数原文，尤其是 AES 密文的输出编码、URL 编码及参与签名参数范围。
- 店铺列表接口完整参数和响应字段定义（店铺名称、SID、站点、授权状态及分页规则）。
- local SKU→ASIN/MSKU 的官方查询接口/报表定义。
- 历史首单日期、每日销量和销售额接口的路径、参数、分页、日期时区、订单状态/退款口径及响应字段。

资料可用官方文档链接、SDK 文件/链接，或来自 Windows 本机授权只读 MCP 的脱敏 Schema 与响应字段结构补充。未经核对不使用社区 SDK 或猜测参数作为生产请求依据。

## 配置 Excel/CSV 字段映射

复制 `config/mapping.example.json` 到 `config/mapping.local.json`。本地映射文件已加入 `.gitignore`。根据领星真实导出文件填写：

- `header_row`：表头实际行号；`sheet_name`：XLSX 工作表（可留空使用首张表）。
- `file_patterns`：相对数据文件夹的文件名模式，例如 `*销售*.xlsx`；不要使用 `..` 或绝对路径。
- `source_headers`：源文件的完整列名。CSV 日期若不是 ISO 格式，填写真实日期格式，例如 `%m/%d/%Y`。
- `actual_ad_spend.source_headers.sku` 必须映射到报表中能唯一对应 SKU 的字段；活动级总花费重复出现在多个 SKU 的报表不得标为 `SKU_LEVEL_CONFIRMED`。
- 如果还没有广告实际花费导出，可先把 `status` 设为 `SALES_ONLY_READY` 并保留广告花费映射空白。对应花费单元格会留空，日志会标明数据缺失；这不代表花费为 0。
- 样本数据确认后，有广告实际花费映射的配置才可设为 `READY_FOR_REVIEW`，且 `sku_scope` 必须是 `SKU_LEVEL_CONFIRMED`。这是用户对源报表粒度的确认，不是程序自动验证出的结论。

CSV 支持 UTF-8、UTF-8 BOM 和 GB18030；XLSX/XLSM 支持明确指定工作表和表头行。旧版 XLS 请先另存为 XLSX。

## 生成周报

图形界面选择格式模板、导出目录、映射文件、SKU、D01 规则和输出目录后点击“开始生成周报”。也可通过命令行：

```powershell
.venv\Scripts\python -m amazon_weekly_report generate `
  --template "C:\新品推广周报.xlsx" `
  --data "C:\领星导出数据" `
  --mapping "C:\配置\mapping.local.json" `
  --skus "ex3106,ex3114" `
  --output "C:\输出\新品推广周报_ex3106_ex3114.xlsx"
```

正式开售日期模式按 SKU 指定日期：

```powershell
.venv\Scripts\python -m amazon_weekly_report generate `
  --template "C:\新品推广周报.xlsx" --data "C:\领星导出数据" `
  --mapping "C:\配置\mapping.local.json" --skus "ex3106,ex3114" `
  --d01-mode launch_date --launch-date ex3106=2026-08-01 `
  --launch-date ex3114=2026-08-08 --output "C:\输出\本周报.xlsx"
```

增量更新时另选已有周报，输出仍必须是新文件，不覆盖来源：

```powershell
.venv\Scripts\python -m amazon_weekly_report generate `
  --template "C:\新品推广周报.xlsx" --existing "C:\上周报.xlsx" `
  --data "C:\新增领星导出" --mapping "C:\配置\mapping.local.json" `
  --skus "ex3106" --output "C:\输出\新品推广周报_更新.xlsx"
```

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

`tests/` 使用明确标注的临时测试夹具验证 D61 扩行、手动字段保留、空白花费、关闭日期后花费告警和 MCP Streamable HTTP 协议；这些夹具只用于自动测试，不代表真实业务记录。ex3106 的真实销量、销售额和 8 月 5 日前后广告实际花费仍须由已授权 MCP 查询或领星原始导出确认。若官方 MCP 没有返回按日期、SKU 唯一归属的广告实际花费，不能用广告归因销售数据替代；应改用具备该粒度的领星 OpenAPI 或报表导出。
