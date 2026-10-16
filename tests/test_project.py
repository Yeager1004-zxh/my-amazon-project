from __future__ import annotations

import csv
import json
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

import httpx
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Border, Side

from amazon_weekly_report.mcp_client import (
    BusinessTool,
    LingxingMcpClient,
    McpClientError,
    McpTool,
    VerifiedReadOnlyGrant,
    extract_business_tools,
    is_directory_search_tool,
    is_read_only_query_candidate,
    redact_sensitive,
)
from amazon_weekly_report.providers.lingxing_api import (
    LingxingApiError,
    LingxingApiNotReady,
    LingxingOpenApiClient,
)
from amazon_weekly_report.report_generation import generate_reports
from amazon_weekly_report.template_expansion import ensure_day_rows
from amazon_weekly_report.template_inspection import inspect_template
from amazon_weekly_report.timeline import build_timeline


HEADERS = [
    "周次", "天数", "周开始日期", "售价", "日销", "自然单", "广告单", "广告订单\n占比",
    "类目", "小类排名", "评分", "评论数", "日目标销量", "目标达成率", "Sessions",
    "自然点击", "广告曝光", "广告点击", "广告CTR", "广告CVR", "广告花费\n($)",
    "广告销售额\n($)", "整体CVR", "ACOS",
]


def create_template(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "新品推广周报"
    worksheet["A1"] = "推广阶段"
    worksheet["G1"] = "发货时间"
    worksheet.merge_cells("A2:A3")
    worksheet.merge_cells("G2:G3")
    worksheet["B2"] = "SAMPLE_ONLY"
    for column, value in enumerate(HEADERS, start=1):
        worksheet.cell(5, column, value)
    for day_number in range(1, 61):
        row = day_number + 5
        worksheet.cell(row, 1, f"第{(day_number - 1) // 7 + 1}周" if (day_number - 1) % 7 == 0 else None)
        worksheet.cell(row, 2, f"D{day_number:02d}")
        if (day_number - 1) % 7 == 0:
            worksheet.cell(row, 3, date(2026, 1, 1) + timedelta(days=day_number - 1))
        worksheet.cell(row, 5, 9876)  # recognizable sample value that must be cleared
        worksheet.cell(row, 13, None)
        worksheet.cell(row, 2).border = Border(bottom=Side(style="thin")) if day_number == 60 else Border()
    red_fill = "FFFF0000"
    worksheet.conditional_formatting.add("N6:N65", CellIsRule(operator="greaterThan", formula=["1"], fill=None))
    worksheet.conditional_formatting.add("X6:X65", CellIsRule(operator="greaterThan", formula=["1"], fill=None))
    worksheet.column_dimensions["B"].width = 12
    worksheet.row_dimensions[6].height = 23
    worksheet["C6"].number_format = "yyyymmdd"
    workbook.save(path)
    workbook.close()


class TemplateAndReportTests(unittest.TestCase):
    def test_gap_free_timeline_and_extension_past_sixty(self) -> None:
        first = date(2026, 1, 1)
        timeline = build_timeline(first, first + timedelta(days=65))
        self.assertEqual(len(timeline), 66)
        self.assertEqual([item.day_number for item in timeline], list(range(1, 67)))
        self.assertEqual(timeline[60].day_label, "D61")
        self.assertTrue(timeline[63].is_week_start)
        self.assertEqual(timeline[63].week_label, "第10周")

        with tempfile.TemporaryDirectory() as temp_dir:
            template = Path(temp_dir) / "template.xlsx"
            create_template(template)
            summary = inspect_template(template)
            workbook = load_workbook(template)
            worksheet = workbook[summary.sheet_name]
            rows = ensure_day_rows(worksheet, summary, 66)
            self.assertEqual(len(rows), 66)
            self.assertEqual(worksheet.cell(66, 2).style_id, worksheet.cell(10, 2).style_id)
            rows = ensure_day_rows(worksheet, summary, 70)
            self.assertEqual(len(rows), 70)
            ranges = {str(item.sqref) for item in worksheet.conditional_formatting}
            self.assertIn("N66:N71", ranges)
            self.assertIn("N72:N75", ranges)
            self.assertIn("X66:X71", ranges)
            workbook.close()

    def test_generation_clears_sample_and_keeps_missing_spend_blank(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            data = root / "exports"
            data.mkdir()
            create_template(template)
            sales_file = data / "sales.csv"
            with sales_file.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["店铺", "SKU", "日期", "销量", "销售额"])
                for index in range(10):
                    day = date(2026, 1, 1) + timedelta(days=index)
                    writer.writerow(["TEST_STORE", "TEST_SKU", day.isoformat(), 1, "10.00"])
            mapping = {
                "version": 1,
                "status": "SALES_ONLY_READY",
                "reports": {
                    "sales": {
                        "header_row": 1,
                        "sheet_name": "",
                        "file_patterns": ["sales.csv"],
                        "date_format": "%Y-%m-%d",
                        "source_headers": {
                            "store": "店铺", "sku": "SKU", "date": "日期",
                            "units_sold": "销量", "sales_amount": "销售额",
                        },
                    },
                    "actual_ad_spend": {
                        "header_row": None, "sheet_name": "", "sku_scope": "UNVERIFIED",
                        "file_patterns": [], "source_headers": {},
                    },
                    "ad_performance": {"enabled": False, "source_headers": {}},
                    "traffic": {"enabled": False, "source_headers": {}},
                    "product_snapshot": {"enabled": False, "source_headers": {}},
                },
            }
            mapping_path = root / "mapping.json"
            mapping_path.write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
            output = root / "report.xlsx"
            result = generate_reports(
                template_path=template,
                data_folder=data,
                mapping_path=mapping_path,
                skus=["TEST_SKU"],
                output_path=output,
            )
            self.assertTrue(output.is_file())
            self.assertTrue(result.audit_path.is_file())
            self.assertEqual(result.sku_results[0].days, 10)
            workbook = load_workbook(output, data_only=False)
            worksheet = workbook["TEST_SKU"]
            self.assertEqual(worksheet["B2"].value, "TEST_SKU")
            self.assertIsNone(worksheet["A2"].value)
            self.assertIsNone(worksheet["G2"].value)
            self.assertIsNone(worksheet["B16"].value)
            self.assertIsNone(worksheet["E16"].value)
            self.assertNotEqual(worksheet["E6"].value, 9876)
            workbook.close()
            log_text = result.audit_path.read_text(encoding="utf-8")
            self.assertIn("2026-01-10", log_text)
            self.assertIn("广告花费单元格保持空白", log_text)
            daily_records = [json.loads(line) for line in log_text.splitlines()]
            first_day = next(item for item in daily_records if item.get("record_type") == "daily_metrics")
            self.assertEqual(first_day["sales_amount"], 10)
            self.assertTrue(first_day["sales_source_row_present"])
            self.assertIsNone(first_day["actual_ad_spend"])

            workbook = load_workbook(output)
            worksheet = workbook["TEST_SKU"]
            worksheet["A2"] = "D1-D14冷启动"
            worksheet["G2"] = "2026-01-05"
            worksheet["M6"] = 3
            worksheet["M16"] = 7
            workbook.save(output)
            workbook.close()

            with sales_file.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["店铺", "SKU", "日期", "销量", "销售额"])
                for index in range(61):
                    day = date(2026, 1, 1) + timedelta(days=index)
                    writer.writerow(["TEST_STORE", "TEST_SKU", day.isoformat(), 1, "10.00"])
            updated = root / "updated.xlsx"
            update_result = generate_reports(
                template_path=template,
                data_folder=data,
                mapping_path=mapping_path,
                skus=["TEST_SKU"],
                output_path=updated,
                existing_report_path=output,
            )
            workbook = load_workbook(updated, data_only=False)
            worksheet = workbook["TEST_SKU"]
            self.assertEqual(worksheet["A2"].value, "D1-D14冷启动")
            self.assertEqual(worksheet["G2"].value, "2026-01-05")
            self.assertEqual(worksheet["M6"].value, 3)
            self.assertEqual(worksheet["M16"].value, 7)
            self.assertEqual(worksheet["N6"].value, 1 / 3)
            self.assertEqual(worksheet["B66"].value, "D61")
            self.assertEqual(worksheet["E66"].value, 1)
            self.assertIsNone(worksheet["U66"].value)
            self.assertIsNone(worksheet["M66"].value)
            self.assertIsNone(worksheet["E67"].value)
            workbook.close()
            self.assertIn("2026-03-02", update_result.audit_path.read_text(encoding="utf-8"))

    def test_actual_spend_sums_distinct_campaigns_and_flags_post_close_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            data = root / "exports"
            data.mkdir()
            create_template(template)
            with (data / "sales.csv").open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["店铺", "SKU", "日期", "销量", "销售额"])
                writer.writerow(["TEST_STORE", "TEST_SKU", "2026-01-01", 1, 10])
            with (data / "ads.csv").open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["店铺", "SKU", "日期", "活动ID", "实际花费", "关闭日期"])
                writer.writerow(["TEST_STORE", "TEST_SKU", "2026-01-02", "CAMPAIGN_A", "2.00", "2026-01-01"])
                writer.writerow(["TEST_STORE", "TEST_SKU", "2026-01-02", "CAMPAIGN_B", "3.00", "2026-01-01"])
            mapping = {
                "version": 1,
                "status": "READY_FOR_REVIEW",
                "reports": {
                    "sales": {
                        "header_row": 1, "sheet_name": "", "file_patterns": ["sales.csv"],
                        "date_format": "%Y-%m-%d",
                        "source_headers": {"store": "店铺", "sku": "SKU", "date": "日期", "units_sold": "销量", "sales_amount": "销售额"},
                    },
                    "actual_ad_spend": {
                        "header_row": 1, "sheet_name": "", "sku_scope": "SKU_LEVEL_CONFIRMED",
                        "file_patterns": ["ads.csv"], "date_format": "%Y-%m-%d",
                        "source_headers": {"store": "店铺", "sku": "SKU", "date": "日期", "actual_spend": "实际花费", "campaign_id": "活动ID", "campaign_disabled_at": "关闭日期"},
                    },
                    "ad_performance": {"enabled": False, "source_headers": {}},
                    "traffic": {"enabled": False, "source_headers": {}},
                    "product_snapshot": {"enabled": False, "source_headers": {}},
                },
            }
            mapping_path = root / "mapping.json"
            mapping_path.write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
            result = generate_reports(
                template_path=template,
                data_folder=data,
                mapping_path=mapping_path,
                skus=["TEST_SKU"],
                output_path=root / "report.xlsx",
            )
            workbook = load_workbook(result.output_path, data_only=False)
            worksheet = workbook["TEST_SKU"]
            self.assertEqual(worksheet["U7"].value, 5)
            self.assertIsNone(worksheet["U6"].value)
            workbook.close()
            log_text = result.audit_path.read_text(encoding="utf-8")
            self.assertIn("关闭日期 2026-01-01 后仍有实际花费记录 2026-01-02", log_text)
            self.assertIn('"requested_sku_source_total_in_report_windows": 5.0', log_text)

    def test_repeated_campaign_cost_across_skus_is_left_unattributed(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            data = root / "exports"
            data.mkdir()
            create_template(template)
            with (data / "sales.csv").open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["店铺", "SKU", "日期", "销量", "销售额"])
                for sku in ("SKU_A", "SKU_B"):
                    writer.writerow(["TEST_STORE", sku, "2026-01-01", 1, 10])
            with (data / "ads.csv").open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.writer(stream)
                writer.writerow(["店铺", "SKU", "日期", "活动ID", "实际花费"])
                for sku in ("SKU_A", "SKU_B"):
                    writer.writerow(["TEST_STORE", sku, "2026-01-02", "SHARED_CAMPAIGN", "5.00"])
            mapping = {
                "version": 1,
                "status": "READY_FOR_REVIEW",
                "reports": {
                    "sales": {
                        "header_row": 1, "sheet_name": "", "file_patterns": ["sales.csv"],
                        "source_headers": {"store": "店铺", "sku": "SKU", "date": "日期", "units_sold": "销量", "sales_amount": "销售额"},
                    },
                    "actual_ad_spend": {
                        "header_row": 1, "sheet_name": "", "sku_scope": "SKU_LEVEL_CONFIRMED",
                        "file_patterns": ["ads.csv"],
                        "source_headers": {"store": "店铺", "sku": "SKU", "date": "日期", "actual_spend": "实际花费", "campaign_id": "活动ID"},
                    },
                    "ad_performance": {"enabled": False, "source_headers": {}},
                    "traffic": {"enabled": False, "source_headers": {}},
                    "product_snapshot": {"enabled": False, "source_headers": {}},
                },
            }
            mapping_path = root / "mapping.json"
            mapping_path.write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
            result = generate_reports(
                template_path=template,
                data_folder=data,
                mapping_path=mapping_path,
                skus=["SKU_A", "SKU_B"],
                output_path=root / "report.xlsx",
            )
            workbook = load_workbook(result.output_path, data_only=False)
            for sku in ("SKU_A", "SKU_B"):
                self.assertIsNone(workbook[sku]["U7"].value)
            workbook.close()
            log_text = result.audit_path.read_text(encoding="utf-8")
            self.assertIn("相同活动/广告组花费重复关联多个 SKU", log_text)


class McpClientTests(unittest.TestCase):
    def test_streamable_http_init_tools_and_read_only_call(self) -> None:
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            if request.method == "DELETE":
                return httpx.Response(200)
            payload = json.loads(request.content) if request.content else {}
            method = payload.get("method")
            request_id = payload.get("id")
            if method == "initialize":
                return httpx.Response(
                    200,
                    headers={"Content-Type": "application/json", "Mcp-Session-Id": "test-session"},
                    json={"jsonrpc": "2.0", "id": request_id, "result": {"protocolVersion": "2025-03-26", "serverInfo": {"name": "fixture"}}},
                )
            if method == "notifications/initialized":
                return httpx.Response(202)
            if method == "tools/list":
                cursor = payload.get("params", {}).get("cursor")
                tools = ([{"name": "read_report", "description": "read", "annotations": {"readOnlyHint": True}, "inputSchema": {"type": "object"}}]
                         if not cursor else [{"name": "write_report", "annotations": {"readOnlyHint": False}}])
                result = {"tools": tools}
                if not cursor:
                    result["nextCursor"] = "page-2"
                return httpx.Response(200, headers={"Content-Type": "application/json"}, json={"jsonrpc": "2.0", "id": request_id, "result": result})
            if method == "tools/call":
                return httpx.Response(200, headers={"Content-Type": "application/json"}, json={"jsonrpc": "2.0", "id": request_id, "result": {"content": [{"type": "text", "text": "query ok"}]}})
            return httpx.Response(405)

        http = httpx.Client(transport=httpx.MockTransport(handler))
        client = LingxingMcpClient("https://mcp.example.test/mcp", "test-key-never-log", http_client=http)
        self.assertEqual(client.connect()["protocolVersion"], "2025-03-26")
        tools = client.list_tools()
        self.assertEqual([tool.name for tool in tools], ["read_report", "write_report"])
        self.assertTrue(tools[0].read_only)
        self.assertEqual(tools[0].input_schema, {"type": "object"})
        self.assertEqual(tools[0].annotations, {"readOnlyHint": True})
        self.assertEqual(client.call_read_only_tool(tools[0], {})["content"][0]["text"], "query ok")
        with self.assertRaises(McpClientError):
            client.call_read_only_tool(tools[1], {})
        self.assertEqual(requests[0].headers["X-Mcp-Key"], "test-key-never-log")
        self.assertEqual(requests[2].headers["Mcp-Session-Id"], "test-session")
        self.assertEqual(requests[2].headers["MCP-Protocol-Version"], "2025-03-26")
        client.close()
        http.close()

    def test_http_error_does_not_echo_server_body_or_secret(self) -> None:
        def handler(request: httpx.Request) -> httpx.Response:
            if request.method == "DELETE":
                return httpx.Response(200)
            return httpx.Response(403, text="test-key-never-log")

        http = httpx.Client(transport=httpx.MockTransport(handler))
        client = LingxingMcpClient("https://mcp.example.test/mcp", "test-key-never-log", http_client=http)
        with self.assertRaises(McpClientError) as raised:
            client.connect()
        self.assertNotIn("test-key-never-log", str(raised.exception))
        client.close()
        http.close()

    def test_secret_fields_are_redacted_for_display(self) -> None:
        self.assertEqual(redact_sensitive({"X-Mcp-Key": "private", "data": "ok"}), {"X-Mcp-Key": "[已隐藏]", "data": "ok"})
        self.assertEqual(redact_sensitive("echoed key", secret_values=("key",)), "echoed [已隐藏]")

    def test_tool_directory_search_returns_business_schema_even_without_read_only_hint(self) -> None:
        business = {
            "toolId": "get_sku_daily_sales",
            "name": "SKU每日销量与销售额",
            "description": "查询 SKU 每日销量和销售额",
            "inputSchema": {
                "type": "object",
                "properties": {"sku": {"type": "string"}, "start_date": {"type": "string"}},
                "required": ["sku"],
                "additionalProperties": False,
            },
        }
        requests: list[dict] = []

        def handler(request: httpx.Request) -> httpx.Response:
            if request.method == "DELETE":
                return httpx.Response(200)
            payload = json.loads(request.content) if request.content else {}
            method = payload.get("method")
            request_id = payload.get("id")
            if method == "initialize":
                return httpx.Response(200, headers={"Mcp-Session-Id": "s"}, json={"jsonrpc": "2.0", "id": request_id, "result": {"protocolVersion": "2025-03-26"}})
            if method == "notifications/initialized":
                return httpx.Response(202)
            if method == "tools/list":
                return httpx.Response(200, json={"jsonrpc": "2.0", "id": request_id, "result": {"tools": [
                    {"name": "search", "description": "搜索领星业务工具目录并返回对应 Schema", "annotations": {"readOnlyHint": False}, "inputSchema": {"type": "object", "properties": {"keyword": {"type": "string"}}, "additionalProperties": False}},
                    {"name": "action", "description": "执行指定业务工具；调用前必须通过 search 获取 Schema", "annotations": {}, "inputSchema": {"type": "object"}},
                    {"name": "status", "description": "状态", "annotations": {}, "inputSchema": {"type": "object"}},
                ]}})
            if method == "tools/call":
                requests.append(payload)
                return httpx.Response(200, json={"jsonrpc": "2.0", "id": request_id, "result": {"structuredContent": {"tools": [business]}}})
            return httpx.Response(405)

        http = httpx.Client(transport=httpx.MockTransport(handler))
        client = LingxingMcpClient("https://mcp.example.test/mcp", "fixture-key", http_client=http)
        client.connect()
        tools = client.list_tools()
        self.assertEqual(len(tools), 3)
        self.assertTrue(is_directory_search_tool(tools[0]))
        self.assertFalse(tools[1].read_only)
        result = client.call_directory_search(tools[0], {"keyword": "销量"})
        specs = extract_business_tools(result)
        self.assertEqual(len(specs), 1)
        self.assertEqual(specs[0].tool_id, "get_sku_daily_sales")
        self.assertEqual(specs[0].input_schema, business["inputSchema"])
        self.assertEqual(requests[0]["params"]["name"], "search")
        self.assertEqual(requests[0]["params"]["arguments"], {"keyword": "销量"})
        client.close()
        http.close()

    def test_search_tool_without_catalog_contract_is_never_called(self) -> None:
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            if request.method == "DELETE":
                return httpx.Response(200)
            payload = json.loads(request.content) if request.content else {}
            request_id = payload.get("id")
            if payload.get("method") == "initialize":
                return httpx.Response(200, headers={"Mcp-Session-Id": "s"}, json={"jsonrpc": "2.0", "id": request_id, "result": {"protocolVersion": "2025-03-26"}})
            if payload.get("method") == "notifications/initialized":
                return httpx.Response(202)
            if payload.get("method") == "tools/call":
                requests.append(payload)
            return httpx.Response(200, json={"jsonrpc": "2.0", "id": request_id, "result": {"tools": []}})

        http = httpx.Client(transport=httpx.MockTransport(handler))
        client = LingxingMcpClient("https://mcp.example.test/mcp", "fixture-key", http_client=http)
        client.connect()
        unsafe_search = McpTool("search", "查询销售数据", {"type": "object", "properties": {}}, {})
        with self.assertRaises(McpClientError):
            client.call_directory_search(unsafe_search, {})
        self.assertEqual(requests, [])
        client.close()
        http.close()

    def test_action_requires_search_schema_user_grant_and_query_only_candidate(self) -> None:
        business = BusinessTool(
            "query_product_performance_asin_lists",
            "商品表现 ASIN 列表查询",
            "查询商品表现 ASIN 列表",
            {
                "type": "object",
                "properties": {"offset": {"type": "integer"}, "length": {"type": "integer"}},
                "required": ["offset", "length"],
                "additionalProperties": False,
            },
        )
        search_tool = McpTool(
            "search", "搜索业务工具目录和 Schema", {"type": "object", "properties": {}, "additionalProperties": False}, {}
        )
        action_tool = McpTool(
            "action",
            "执行指定业务工具；调用前必须先通过 search 获取 Schema",
            {
                "type": "object",
                "properties": {
                    "toolId": {"type": "string"},
                    "params": {
                        "type": "object",
                        "properties": {},
                        "additionalProperties": False,
                        "maxProperties": 0,
                    },
                },
                "required": ["toolId", "params"],
                "additionalProperties": False,
            },
            {},
        )
        calls: list[dict] = []

        def handler(request: httpx.Request) -> httpx.Response:
            if request.method == "DELETE":
                return httpx.Response(200)
            payload = json.loads(request.content) if request.content else {}
            method = payload.get("method")
            request_id = payload.get("id")
            if method == "initialize":
                return httpx.Response(200, headers={"Mcp-Session-Id": "s"}, json={"jsonrpc": "2.0", "id": request_id, "result": {"protocolVersion": "2025-03-26"}})
            if method == "notifications/initialized":
                return httpx.Response(202)
            if method == "tools/call":
                params = payload["params"]
                calls.append(params)
                if params["name"] == "search":
                    return httpx.Response(200, json={"jsonrpc": "2.0", "id": request_id, "result": {"structuredContent": {"tools": [{"id": business.tool_id, "name": business.name, "description": business.description, "inputSchema": business.input_schema}]}}})
                return httpx.Response(200, json={"jsonrpc": "2.0", "id": request_id, "result": {"content": [{"type": "text", "text": "query complete"}]}})
            return httpx.Response(200, json={"jsonrpc": "2.0", "id": request_id, "result": {"tools": []}})

        http = httpx.Client(transport=httpx.MockTransport(handler))
        client = LingxingMcpClient("https://mcp.example.test/mcp", "fixture-key", http_client=http)
        client.connect()
        client.call_directory_search(search_tool, {})
        grant = VerifiedReadOnlyGrant(business.tool_id, business.permission_fingerprint)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {"offset": 0, "length": 5}, VerifiedReadOnlyGrant("other_id", business.permission_fingerprint))
        self.assertEqual(len(calls), 1, "unverified action must not be sent")
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {"offset": 0}, grant)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {"offset": 0, "length": 5, "unexpected": "value"}, grant)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {"offset": "0", "length": 5}, grant)
        changed_schema = BusinessTool(
            business.tool_id,
            business.name,
            business.description,
            {"type": "object", "properties": {"offset": {"type": "string"}, "length": {"type": "integer"}}, "required": ["offset", "length"], "additionalProperties": False},
        )
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, changed_schema, {"offset": 0, "length": 5}, grant)
        bad_outer_schema = McpTool(
            action_tool.name,
            action_tool.description,
            {
                "type": "object",
                "properties": {
                    "toolId": {"type": "integer"},
                    "params": {"type": "object", "properties": {}, "additionalProperties": False},
                },
                "required": ["toolId", "params"],
                "additionalProperties": False,
            },
            {},
        )
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(bad_outer_schema, business, {"offset": 0, "length": 5}, grant)
        self.assertEqual(len(calls), 1, "invalid params and a changed ID/Schema must not send action")
        result = client.call_verified_business_query(action_tool, business, {"offset": 0, "length": 5}, grant)
        self.assertIn("content", result)
        self.assertEqual(calls[-1], {"name": "action", "arguments": {"toolId": business.tool_id, "params": {"offset": 0, "length": 5}}})
        unsafe = BusinessTool("update_listing", "修改Listing", "更新商品信息", business.input_schema)
        self.assertFalse(is_read_only_query_candidate(unsafe))
        camel_case_query = BusinessTool("getSkuDailySales", "SKU销量", "查询 SKU 每日销量", business.input_schema)
        self.assertTrue(is_read_only_query_candidate(camel_case_query))
        operation_schema = BusinessTool(
            "get_report",
            "查询报表",
            "查询报表数据",
            {"type": "object", "properties": {"update_listing": {"type": "boolean"}}},
        )
        self.assertFalse(is_read_only_query_candidate(operation_schema))
        client.close()
        http.close()


class LingxingOpenApiTests(unittest.TestCase):
    def test_token_exchange_uses_multipart_and_hides_token(self) -> None:
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            return httpx.Response(
                200,
                json={"code": 0, "data": {"access_token": "fixture-token-do-not-display", "expires_in": 3600}},
            )

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            result = client.request_access_token()
        request = requests[0]
        self.assertEqual(request.url.path, "/api/auth-server/oauth/access-token")
        self.assertIn("multipart/form-data; boundary=", request.headers["Content-Type"])
        body = request.content.decode("utf-8")
        self.assertIn('name="appId"', body)
        self.assertIn("fixture-app-id", body)
        self.assertIn('name="appSecret"', body)
        self.assertIn("fixture-app-secret", body)
        self.assertEqual(result.expires_in_seconds, 3600)
        self.assertNotIn("fixture-token-do-not-display", repr(result))
        self.assertNotIn("fixture-token-do-not-display", result.safe_status())
        http.close()

    def test_signed_queries_fail_closed_without_sending_request(self) -> None:
        requests: list[httpx.Request] = []
        http = httpx.Client(transport=httpx.MockTransport(lambda request: requests.append(request) or httpx.Response(200)))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            with self.assertRaises(LingxingApiNotReady):
                client.list_sellers("fixture-token")
            with self.assertRaises(LingxingApiNotReady):
                client.query_sku_sales(
                    access_token="fixture-token",
                    seller_id="fixture-seller",
                    sku="ex3106",
                    start_date="2026-08-01",
                    end_date="2026-08-10",
                )
        self.assertEqual(requests, [])
        http.close()

    def test_token_http_error_does_not_echo_credentials(self) -> None:
        http = httpx.Client(transport=httpx.MockTransport(lambda _request: httpx.Response(403, text="fixture-app-secret")))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            with self.assertRaises(LingxingApiError) as raised:
                client.request_access_token()
        self.assertNotIn("fixture-app-secret", str(raised.exception))
        http.close()


if __name__ == "__main__":
    unittest.main()
