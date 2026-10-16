# 新品推广周报 V1.0 实施清单

本计划基于仓库 `main` 的现有实现，限定为新品推广周报。广告活动管理、库存和利润等其他业务不属于 V1.0。

## 基线审查

当前 `main` 已包含：MCP Streamable HTTP 初始化/会话、工具目录与只读许可、Windows Credential Manager、OpenAPI Token 获取、Excel/CSV 映射导入、D01 计算、D61+ 扩行、JSONL 校验日志和模板样式复制。

当前仍缺：MCP 查询到周报生成的连接、真实店铺/SKU/ASIN/MSKU 数据映射、统一的带来源与覆盖状态的日数据、实际广告花费的接口口径核验，以及面向普通用户的 GUI 流程。OpenAPI 的店铺和销售查询保持 fail-closed。现有自动化测试均为本地夹具，不代表真实领星或 Windows 验收。

## 分阶段任务

| 阶段 | 模块 | 输入 → 输出 | 依赖 | 测试 | 完成条件 |
|---|---|---|---|---|---|
| 1 | MCP 参数封装 | action 外层 Schema + search 返回的业务 Schema + 参数 → action 请求体 | 会话中真实 search 结果、已核验只读许可 | MCP 模拟测试覆盖有效 offset/length、缺字段、额外字段、类型错误、外层 ID/params 错误和无效许可 | params 只按业务 Schema 验证；拒绝项不发请求；Windows 上对已核验的只读工具真实执行成功 |
| 2 | 店铺与 SKU 查询 | get_my_sids/业务查询结果 + 选择店铺及 SKU → SID、local SKU、MSKU、ASIN 与日指标原始记录 | Windows 授权 MCP；真实工具名、Schema、分页和响应样本 | 模拟分页、超时、限流/错误、Token 或 session 失效；Windows 对照领星原报表 | EX-shop-US 唯一映射 SID；ex3106 每日销量/销售额与原报表一致，首单日期可追溯 |
| 3 | 统一日数据与校验 | 销售/广告/流量响应 → SKU 日记录、数据来源、统计口径、日期/时区、完整性、实体归属和异常 | 阶段 2 的字段定义和官方口径 | Decimal 精度、缺失与 0、时区、重复、跨 SKU 活动费、关闭后仍有花费测试 | 未知或无法唯一归属的值为空；共享活动费用不分摊/不重复累计；金额计算不先转 float |
| 4 | Excel 周报适配 | 统一日记录 + 模板/已有周报 → 新工作簿 + 脱敏校验日志 | 阶段 2、3；现有 `report_generation.py` 和模板扩展 | 原模板结构回归、D61+、人工字段保护、增量更新；用 ex3106 真实历史周报核对 | D01 与自然日正确；仅写入覆盖完整且已核验指标；其余明确留空/部分完成；公式、样式、合并单元格和人工字段保持 |
| 5 | Windows 日常 GUI | 店铺、SKU → 进度、输出位置、成功/部分完成/失败和诊断 | 阶段 4 稳定接口；凭证/模板/输出本地设置 | UI 模拟测试、Windows 10/11 本地 MCP 与 Excel 测试 | 普通页面无需 JSON、Schema 或 Tool ID；高级设置集中技术连接选项；查询失败不生成貌似完整的报告 |
| 6 | 打包与交付 | Python 项目、依赖、目标机 → EXE、安装说明和验收报告 | 阶段 5；Windows Tk/Tcl、Credential Manager 和 PyInstaller | Windows 干净环境安装、启动、真实数据端到端与 EXE 启动 | ex3106 单 SKU 先验收，再验收 ex3106+ex3114；对照领星原始报表一致；仓库不含密钥、业务报表或未脱敏日志 |

## 第一阶段参数校验约束

`action` 的外层 Schema 负责验证工具 ID、必需 envelope 字段、params 对象类型及外层额外字段。`params` 内的业务字段只使用同一 MCP 会话中 `search` 返回的业务 inputSchema 校验。通用 action params 的空 `properties`、`additionalProperties: false` 或字段数占位限制不能覆盖业务 Schema。只读许可仍绑定当次业务工具 ID、说明和 Schema 指纹；未核验、发生变化或不可确认只读的调用必须在发请求前被拒绝。

## 真实接口所需的安全诊断材料

真实接入前需从 Windows 本地已授权会话确认并以脱敏方式记录：

- `get_my_sids` 工具元数据、inputSchema 和店铺返回中店铺名/SID/站点字段。
- 负责 local SKU 与 ASIN/MSKU 对应、历史首单/每日销量/销售额、广告活动实际花费及可用广告表现指标的只读业务工具名、完整 inputSchema、分页方式和字段定义。
- 每项查询一小段去除账号/订单等敏感值的响应结构，以及同日期领星原始报表上的汇总核对结果。

不要收集或提交 MCP 密钥、Token、AppSecret、买家信息、完整订单明细或未经脱敏的业务报表。当前云环境无领星 MCP 凭证/会话；云端 mock 测试不构成真实数据验收。若接口未提供唯一 SKU 粒度的广告活动实际花费，相关单元必须留空并报告原因。
