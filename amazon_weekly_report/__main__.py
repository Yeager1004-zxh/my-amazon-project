from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any

from .inspection import InspectionError, inspect_folder
from .mapping import load_mapping, validate_mapping
from .template_inspection import TemplateInspectionError, inspect_template
from .template_preparation import TemplatePreparationError, prepare_sku_workbook
from .credential_store import CredentialStoreError, WindowsCredentialStore
from .local_settings import LocalSettingsError, load_settings
from .mcp_client import LingxingMcpClient, McpClientError, redact_sensitive
from .report_generation import ReportGenerationError, generate_reports


def _print_inspection(folder: Path) -> int:
    try:
        result = inspect_folder(folder)
    except InspectionError as exc:
        print(f"检查失败：{exc}", file=sys.stderr)
        return 2

    print(f"目录：{result.directory}")
    if not result.files:
        print("没有找到受支持的 CSV、XLSX 或 XLSM 文件。")
    for item in result.files:
        location = f"，工作表：{item.sheet_name}" if item.sheet_name else ""
        print(f"\n文件：{item.path}{location}")
        if item.encoding:
            print(f"编码：{item.encoding}；分隔符：{item.delimiter!r}")
        if item.header_row is None:
            print("未能识别候选表头行。")
        else:
            print(f"候选表头行：{item.header_row}")
            print("字段：" + " | ".join(item.headers))
        for warning in item.warnings:
            print(f"提示：{warning}")
    for path in result.unsupported_files:
        print(f"\n暂不支持：{path}（旧版 XLS 请先另存为 XLSX）")
    return 0


def _print_mapping(path: Path) -> int:
    try:
        mapping = load_mapping(path)
    except (OSError, ValueError) as exc:
        print(f"读取映射配置失败：{exc}", file=sys.stderr)
        return 2

    issues = validate_mapping(mapping)
    if issues:
        print("映射尚未就绪：")
        for issue in issues:
            print(f"- {issue}")
        return 1
    print("映射配置结构完整。模板结构和数据汇总流程仍需根据真实样本配置。")
    return 0


def _print_template(path: Path) -> int:
    try:
        result = inspect_template(path)
    except TemplateInspectionError as exc:
        print(f"模板检查失败：{exc}", file=sys.stderr)
        return 2
    print(f"工作表：{', '.join(result.sheet_names)}")
    print(f"识别模板：{result.sheet_name}；表头第 {result.header_row} 行")
    print(
        f"日期行：第 {result.first_day_row}–{result.last_day_row} 行，"
        f"共 D01–D{result.day_count:02d}，天数列 {result.day_column}"
    )
    print("周起始行：" + ", ".join(map(str, result.week_start_rows)))
    print("列：" + " | ".join(f"{name}={column}" for name, column in result.columns.items()))
    print("人工输入字段：" + "; ".join(f"{name} {cell}" for name, cell in result.manual_input_cells.items()))
    print(f"合并单元格 {len(result.merged_ranges)} 个；公式 {result.formula_count} 个")
    for warning in result.warnings:
        print(f"提示：{warning}")
    return 0


def _prepare_template(template: Path, output: Path, sku_list: str) -> int:
    skus = [sku for sku in sku_list.replace("，", ",").split(",") if sku.strip()]
    try:
        output_path = prepare_sku_workbook(template, output, skus)
    except (OSError, TemplatePreparationError) as exc:
        print(f"准备工作表失败：{exc}", file=sys.stderr)
        return 2
    print(f"已创建 SKU 工作表框架：{output_path}")
    print("注意：文件仅包含清空样例数据后的模板工作表，尚未导入业务数据或生成周报指标。")
    return 0


def _sku_list(value: str) -> list[str]:
    return [sku.strip() for sku in value.replace("，", ",").split(",") if sku.strip()]


def _parse_d01_dates(values: list[str]) -> dict[str, date]:
    result: dict[str, date] = {}
    for value in values:
        if "=" not in value:
            raise ValueError("D01 日期格式应为 SKU=YYYY-MM-DD")
        sku, day = (part.strip() for part in value.split("=", 1))
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day):
            raise ValueError(f"D01 日期格式应为 YYYY-MM-DD：{sku}")
        try:
            result[sku] = date.fromisoformat(day)
        except ValueError as exc:
            raise ValueError(f"D01 日期不是有效的 YYYY-MM-DD 日期：{sku}") from exc
    return result


def _mcp_client() -> LingxingMcpClient:
    settings = load_settings()
    endpoint = settings.get("endpoint")
    if not endpoint:
        raise LocalSettingsError("尚未设置领星官方 MCP 地址；请先在 Windows 图形界面配置。")
    key = WindowsCredentialStore().get("X-Mcp-Key")
    if not key:
        raise CredentialStoreError("Windows Credential Manager 中没有 X-Mcp-Key；请先在图形界面安全保存。")
    return LingxingMcpClient(endpoint, key)


def _mcp_command(command: str, tool_name: str | None = None, arguments: str = "{}") -> int:
    try:
        client = _mcp_client()
        with client:
            result = client.connect()
            if command == "mcp-test":
                print(f"连接成功；MCP 协议版本：{result.get('protocolVersion')}")
                return 0
            tools = client.list_tools()
            if command == "mcp-tools":
                for tool in tools:
                    marker = "readOnlyHint=true" if tool.read_only else f"readOnlyHint={tool.annotations.get('readOnlyHint', '未提供')}"
                    print(f"\n{tool.name} [{marker}] — {tool.description}")
                    print("annotations:")
                    print(json.dumps(tool.annotations, ensure_ascii=False, indent=2))
                    print("inputSchema:")
                    print(json.dumps(tool.input_schema, ensure_ascii=False, indent=2))
                print(f"\n共 {len(tools)} 个 MCP 工具；基础只读标记 {sum(tool.read_only for tool in tools)} 个。")
                return 0
            selected = next((tool for tool in tools if tool.name == tool_name), None)
            if selected is None:
                raise McpClientError("工具列表中找不到指定工具。")
            try:
                parsed = json.loads(arguments)
            except json.JSONDecodeError as exc:
                raise ValueError(f"查询参数 JSON 格式错误，第 {exc.lineno} 行：{exc.msg}") from exc
            if not isinstance(parsed, dict):
                raise ValueError("查询参数必须是 JSON 对象。")
            output = client.call_read_only_tool(selected, parsed)
            print(json.dumps(redact_sensitive(output), ensure_ascii=False, indent=2, default=str))
            return 0
    except (CredentialStoreError, LocalSettingsError, McpClientError, ValueError, OSError) as exc:
        print(f"MCP 操作失败：{exc}", file=sys.stderr)
        return 2


def _generate(args: argparse.Namespace) -> int:
    try:
        d01_dates = _parse_d01_dates(args.d01)
        result = generate_reports(
            template_path=args.template,
            data_folder=args.data,
            mapping_path=args.mapping,
            skus=_sku_list(args.skus),
            output_path=args.output,
            d01_dates=d01_dates,
            d01_confirmed=args.confirm_d01,
            confirmed_d01_changes=args.confirm_d01_change,
            existing_report_path=args.existing,
        )
    except (ReportGenerationError, OSError, ValueError) as exc:
        print(f"生成失败：{exc}", file=sys.stderr)
        return 2
    print(f"周报：{result.output_path}")
    print(f"数据校验日志：{result.audit_path}")
    for item in result.sku_results:
        dates = f"D01={item.d01}，覆盖至={item.latest_covered_date}，{item.days} 天" if item.d01 else "未生成"
        print(f"{item.sku}: {item.status}，{dates}")
        for warning in item.warnings:
            print(f"  提示：{warning}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="amazon-weekly-report",
        description="从真实领星导出文件生成周报，并提供 Windows 本地 MCP 只读连接测试。",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="检查导出文件和候选表头")
    inspect_parser.add_argument("folder", type=Path, help="领星导出文件所在文件夹")

    mapping_parser = subparsers.add_parser("validate-mapping", help="检查字段映射是否完整")
    mapping_parser.add_argument("file", type=Path, help="JSON 映射配置文件")

    template_parser = subparsers.add_parser("inspect-template", help="检查原周报模板结构")
    template_parser.add_argument("file", type=Path, help="原始周报 XLSX 文件")

    prepare_parser = subparsers.add_parser("prepare-template", help="从格式样例创建空白 SKU 工作表框架")
    prepare_parser.add_argument("template", type=Path, help="原始周报 XLSX 文件")
    prepare_parser.add_argument("--output", type=Path, required=True, help="新建输出 XLSX 路径，不覆盖已有文件")
    prepare_parser.add_argument("--skus", required=True, help="SKU 列表，使用逗号或中文逗号分隔")

    generate_parser = subparsers.add_parser("generate", help="按明确字段映射生成周报和校验日志")
    generate_parser.add_argument("--template", type=Path, required=True, help="原周报格式模板")
    generate_parser.add_argument("--data", type=Path, required=True, help="领星导出报表文件夹")
    generate_parser.add_argument("--mapping", type=Path, required=True, help="已按真实源表头配置的 JSON")
    generate_parser.add_argument("--skus", required=True, help="SKU 列表，使用逗号分隔")
    generate_parser.add_argument("--output", type=Path, required=True, help="新建输出 XLSX，不覆盖已有文件")
    generate_parser.add_argument("--d01", action="append", default=[], help="手动确认的 D01，格式 SKU=YYYY-MM-DD；可重复")
    generate_parser.add_argument("--confirm-d01", action="store_true", required=True, help="确认所填 D01 按所选店铺站点日期解释")
    generate_parser.add_argument("--confirm-d01-change", action="store_true", help="确认已有周报的 D01 变化会使全部 D 编号日期错位")
    generate_parser.add_argument("--existing", type=Path, help="可选的已有周报；更新时手动字段会保留")

    subparsers.add_parser("gui", help="启动 Windows 中文图形界面")
    subparsers.add_parser("mcp-test", help="使用本机凭据进行 MCP 初始化连接测试")
    subparsers.add_parser("mcp-tools", help="获取并列出官方 MCP 工具及只读标记")
    query_parser = subparsers.add_parser("mcp-query", help="调用一个明确标记只读的 MCP 工具")
    query_parser.add_argument("tool_name")
    query_parser.add_argument("--arguments", default="{}", help="JSON 对象；不要在参数中放密钥")

    args = parser.parse_args()
    if args.command == "inspect":
        return _print_inspection(args.folder)
    if args.command == "validate-mapping":
        return _print_mapping(args.file)
    if args.command == "inspect-template":
        return _print_template(args.file)
    if args.command == "prepare-template":
        return _prepare_template(args.template, args.output, args.skus)
    if args.command == "generate":
        return _generate(args)
    if args.command == "gui":
        from .gui import main as gui_main
        return gui_main()
    if args.command in {"mcp-test", "mcp-tools"}:
        return _mcp_command(args.command)
    if args.command == "mcp-query":
        return _mcp_command(args.command, args.tool_name, args.arguments)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
