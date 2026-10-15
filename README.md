# 亚马逊新品推广周报工具

面向 Windows 10/11 的本地程序，读取用户导出的领星 CSV/XLSX，按原周报模板生成多 SKU 工作表和数据校验日志；另提供 Windows 本地领星 Streamable HTTP MCP 只读连接客户端。

## 当前实现

- 延续原模板七天一周的样式、周次、边框和条件格式，按实际来源覆盖日期扩展 D61、D62……；自然日连续排列，来源缺失的日期指标保持空白。
- 支持多个 SKU、多个来源文件和多个店铺；字段必须在本机 JSON 映射中明确指定，不猜报表表头。
- 新 SKU 页签会清除模板样例指标。增量更新时，实际来源有值的指标会更新；来源没有记录的旧指标保留。`A2` 推广阶段、`G2` 发货时间和 `M` 列日目标销量由用户维护，程序不会覆盖。
- `first_sale` 模式以真实销售来源中首个正销量日期作为 D01；`launch_date` 模式要求按 SKU 填正式开售日期。
- 销量、销售金额、广告实际花费、广告表现和流量可按真实字段映射。广告花费必须使用实际花费字段；只有在用户将 `sku_scope` 明确设为 `SKU_LEVEL_CONFIRMED` 后才接受该映射。重复、跨 SKU 共享的同一活动花费或店铺覆盖不完整的记录会从相关日汇总中排除并写入校验日志。
- 校验日志为 JSONL，记录来源文件名、SKU 的 D01、最新数据覆盖日、销量/销售额/广告花费汇总和异常提示。关闭日期后的实际花费只有在来源提供并映射 `campaign_disabled_at` 时才提示。
- Windows 本地 MCP 客户端支持 MCP 初始化连接、工具列表分页、只读工具查询。请求带 `X-Mcp-Key`、MCP session 和协议版本；仅调用服务端明确标记 `readOnlyHint=true` 的工具。

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

之后可依次使用“测试连接”“获取工具列表”和“运行只读查询测试”。查询工具必须被服务端标为只读；未标注只读的工具会被拒绝调用。查询结果仅在本机窗口/终端显示，不保存到日志或工作簿。命令行连接测试也只从 Windows Credential Manager 读取凭证：

```powershell
.venv\Scripts\python -m amazon_weekly_report mcp-test
.venv\Scripts\python -m amazon_weekly_report mcp-tools
.venv\Scripts\python -m amazon_weekly_report mcp-query 工具名称 --arguments "{}"
```

当前没有已授权的官方地址、密钥或真实工具结果，因此 Codex 云端无法替你完成在线登录或 ex3106 业务验证。MCP 页面用于本机握手、列工具和人工发起只读查询；将领星工具结果稳定映射到日报字段，需要先用你的授权服务确认工具名称、参数和返回字段。完成这项实测前，周报生成器使用下述 Excel/CSV 导入流程，不声称 MCP 已自动提供销量或活动实际花费。

## 领星 OpenAPI 本机测试

OpenAPI 使用固定官方域名 `https://openapi.lingxing.com`。目前已实现 Token 请求：`POST /api/auth-server/oauth/access-token`，以 multipart/form-data 发送 `appId` 和 `appSecret`。在“领星 OpenAPI 连接”页分别填写两项凭证并点击“安全保存凭证”；凭证保存在 Windows Credential Manager 的 `my-amazon-project.lingxing-openapi` 服务项中，不写入配置文件、命令行或日志。之后点击“测试 API 身份验证”，界面只显示成功状态和响应中可识别的有效期，不显示 Access Token。

店铺列表接口 `GET /erp/sc/data/seller/lists` 需要官方定义的 `access_token`、`app_key`、`timestamp`、`sign` 参数。签名规则、店铺响应字段和销售查询接口尚未完成官方文档核验，因此“获取店铺列表”和“测试 SKU 数据”当前会提示未发送请求，不会猜签名、读取模拟数据或把结果写入正式周报。广告接口也未接入。完整身份验证和数据查询仍需在 Windows 本机用你的授权凭证测试；不要将 AppID、AppSecret 或 Token 发到聊天或 GitHub。

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
