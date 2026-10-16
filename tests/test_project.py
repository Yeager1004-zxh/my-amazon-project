from __future__ import annotations

import csv
import json
import tempfile
import unittest
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Any
from unittest.mock import patch
from urllib.parse import parse_qs

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
    validate_json_schema,
)
from amazon_weekly_report.d01_settings import D01SettingsError, get_d01_dates, load_d01_settings, save_d01_dates
from amazon_weekly_report.providers.lingxing_api import (
    LingxingApiError,
    LingxingApiNotReady,
    LingxingOpenApiClient,
    SkuDailySales,
    SkuDailySalesWindow,
    SkuProductMapping,
    parse_asin_daily_page,
    parse_product_performance_page,
    parse_sp_campaign_report_page,
    parse_sp_product_ad_report_page,
    parse_seller_list_response,
    split_date_windows,
)
from amazon_weekly_report.providers.lingxing_signing import (
    LingxingSignatureError,
    canonical_signing_string,
    encode_query_parameters,
    generate_signature,
    signed_post_query_parameters,
)
from amazon_weekly_report.report_generation import ReportGenerationError, generate_reports
from amazon_weekly_report.template_expansion import ensure_day_rows
from amazon_weekly_report.template_inspection import inspect_template
from amazon_weekly_report.timeline import build_timeline
from amazon_weekly_report.gui import WeeklyReportApp


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
    def test_report_generation_requires_a_manually_confirmed_d01_for_every_sku(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            create_template(template)
            data = root / "exports"
            data.mkdir()
            mapping = root / "mapping.json"
            mapping.write_text("{}", encoding="utf-8")
            with self.assertRaisesRegex(ReportGenerationError, "尚未设置并确认 D01"):
                generate_reports(
                    template_path=template,
                    data_folder=data,
                    mapping_path=mapping,
                    skus=["SKU_A", "SKU_B"],
                    output_path=root / "blocked.xlsx",
                    d01_confirmed=True,
                )
            self.assertFalse((root / "blocked.xlsx").exists())

    def test_existing_report_d01_change_requires_confirmation_then_rebases_dates_and_keeps_manual_fields(self) -> None:
        start = date(2026, 1, 1)
        mapping = SkuProductMapping("API-TEST", ("MSKU-A",), ())
        rows = tuple(
            SkuDailySales(987654, "API-TEST", "MSKU-A", "", start + timedelta(days=index), "USD", index + 1, index + 1, Decimal(f"{index + 1}.00"))
            for index in range(3)
        )
        window = SkuDailySalesWindow(start, start + timedelta(days=2), (mapping,), rows)
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            create_template(template)
            original = root / "original.xlsx"
            generate_reports(
                template_path=template,
                data_folder=None,
                mapping_path=None,
                skus=["API-TEST"],
                output_path=original,
                d01_dates={"API-TEST": start},
                d01_confirmed=True,
                api_sales_windows=[window],
                api_store_sid=987654,
                api_store_name="Synthetic Shop",
            )
            workbook = load_workbook(original)
            sheet = workbook["API-TEST"]
            sheet["A2"] = "保留推广阶段"
            sheet["G2"] = "保留发货时间"
            sheet["M6"] = 5
            sheet["M7"] = 6
            sheet["M8"] = 7
            workbook.save(original)
            workbook.close()

            new_d01 = start + timedelta(days=1)
            with self.assertRaisesRegex(ReportGenerationError, "日期错位风险"):
                generate_reports(
                    template_path=template,
                    data_folder=None,
                    mapping_path=None,
                    skus=["API-TEST"],
                    output_path=root / "unconfirmed.xlsx",
                    d01_dates={"API-TEST": new_d01},
                    d01_confirmed=True,
                    existing_report_path=original,
                    api_sales_windows=[window],
                    api_store_sid=987654,
                    api_store_name="Synthetic Shop",
                )

            updated = root / "confirmed.xlsx"
            generate_reports(
                template_path=template,
                data_folder=None,
                mapping_path=None,
                skus=["API-TEST"],
                output_path=updated,
                d01_dates={"API-TEST": new_d01},
                d01_confirmed=True,
                confirmed_d01_changes=True,
                existing_report_path=original,
                api_sales_windows=[window],
                api_store_sid=987654,
                api_store_name="Synthetic Shop",
            )
            workbook = load_workbook(updated)
            sheet = workbook["API-TEST"]
            self.assertEqual(sheet["C6"].value.date(), new_d01)
            self.assertEqual(sheet["B6"].value, "D01")
            self.assertEqual(sheet["B7"].value, "D02")
            self.assertEqual(sheet["B8"].value, "D03")
            self.assertEqual(sheet["E6"].value, 2)
            self.assertEqual(sheet["E7"].value, 3)
            self.assertIsNone(sheet["E8"].value)
            self.assertEqual(sheet["A2"].value, "保留推广阶段")
            self.assertEqual(sheet["G2"].value, "保留发货时间")
            self.assertEqual([sheet[f"M{row}"].value for row in (6, 7, 8)], [5, 6, 7])
            workbook.close()

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
                d01_dates={"TEST_SKU": date(2026, 1, 1)},
                d01_confirmed=True,
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
                d01_dates={"TEST_SKU": date(2026, 1, 1)},
                d01_confirmed=True,
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
                d01_dates={"TEST_SKU": date(2026, 1, 1)},
                d01_confirmed=True,
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
                d01_dates={"SKU_A": date(2026, 1, 1), "SKU_B": date(2026, 1, 1)},
                d01_confirmed=True,
            )
            workbook = load_workbook(result.output_path, data_only=False)
            for sku in ("SKU_A", "SKU_B"):
                self.assertIsNone(workbook[sku]["U7"].value)
            workbook.close()
            log_text = result.audit_path.read_text(encoding="utf-8")
            self.assertIn("相同活动/广告组花费重复关联多个 SKU", log_text)


class ApiReportGenerationTests(unittest.TestCase):
    def test_openapi_sales_generate_template_d61_and_keep_missing_distinct_from_zero(self) -> None:
        start = date(2026, 1, 1)
        end = start + timedelta(days=61)
        mapping = SkuProductMapping("API-TEST", ("MSKU-A", "MSKU-B"), ("ASINTEST0001",))
        rows: list[SkuDailySales] = []
        for offset in range((end - start).days + 1):
            event_date = start + timedelta(days=offset)
            for seller_sku in mapping.seller_skus:
                if offset == 0:
                    units, orders, amount = 0, 0, Decimal("0.00")
                elif offset == 1:
                    units, orders, amount = None, None, None
                else:
                    units, orders, amount = 1, 1, Decimal("10.00") if seller_sku == "MSKU-A" else Decimal("11.00")
                rows.append(SkuDailySales(
                    sid=987654,
                    local_sku="API-TEST",
                    seller_sku=seller_sku,
                    asin="ASINTEST0001",
                    event_date=event_date,
                    currency_code="USD",
                    units_sold=units,
                    order_count=orders,
                    sales_amount=amount,
                ))
        window = SkuDailySalesWindow(start, end, (mapping,), tuple(rows))

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            output = root / "api-report.xlsx"
            create_template(template)
            result = generate_reports(
                template_path=template,
                data_folder=None,
                mapping_path=None,
                skus=["API-TEST"],
                output_path=output,
                d01_dates={"API-TEST": start},
                d01_confirmed=True,
                api_sales_windows=[window],
                api_store_sid=987654,
                api_store_name="Synthetic Shop",
            )
            self.assertEqual(result.sku_results[0].status, "partial")
            self.assertEqual(result.sku_results[0].days, 62)
            workbook = load_workbook(output, data_only=False)
            sheet = workbook["API-TEST"]
            self.assertEqual(sheet["E6"].value, 0)
            self.assertIsNone(sheet["D6"].value)  # 0/0 is undefined, not zero price.
            self.assertIsNone(sheet["E7"].value)  # Missing API records stay blank.
            self.assertIsNone(sheet["D7"].value)
            self.assertEqual(sheet["E8"].value, 2)
            self.assertEqual(sheet["D8"].value, 10.5)
            self.assertEqual(sheet["B66"].value, "D61")
            self.assertEqual(sheet["E66"].value, 2)
            self.assertIsNone(sheet["U6"].value)  # Actual ad spend is not yet connected.
            self.assertIsNone(sheet["A2"].value)
            self.assertIsNone(sheet["G2"].value)
            self.assertIsNone(sheet["M6"].value)
            workbook.close()

            audit = [json.loads(line) for line in result.audit_path.read_text(encoding="utf-8").splitlines()]
            run = next(item for item in audit if item.get("record_type") == "run")
            self.assertEqual(run["source_details"]["provider"], "lingxing_openapi")
            self.assertEqual(run["source_details"]["store_sid"], 987654)
            first_day = next(item for item in audit if item.get("record_type") == "daily_metrics" and item["day_number"] == 1)
            missing_day = next(item for item in audit if item.get("record_type") == "daily_metrics" and item["day_number"] == 2)
            self.assertEqual(first_day["total_order_count"], "0")
            self.assertEqual(first_day["sales_amount"], "0.00")
            self.assertIsNone(missing_day["total_order_count"])
            self.assertIsNone(missing_day["sales_amount"])
            self.assertEqual(first_day["sales_source"], "lingxing_openapi")

            workbook = load_workbook(output)
            sheet = workbook["API-TEST"]
            sheet["A2"] = "人工推广阶段"
            sheet["G2"] = "人工发货时间"
            sheet["M6"] = 4
            sheet["M66"] = 8
            workbook.save(output)
            workbook.close()

            added_date = end + timedelta(days=1)
            added_rows = tuple(
                SkuDailySales(987654, "API-TEST", seller_sku, "ASINTEST0001", added_date, "USD", 1, 1, Decimal("10.00"))
                for seller_sku in mapping.seller_skus
            )
            updated_window = SkuDailySalesWindow(start, added_date, (mapping,), (*rows, *added_rows))
            updated = root / "api-report-updated.xlsx"
            generate_reports(
                template_path=template,
                data_folder=None,
                mapping_path=None,
                skus=["API-TEST"],
                output_path=updated,
                d01_dates={"API-TEST": start},
                d01_confirmed=True,
                existing_report_path=output,
                api_sales_windows=[updated_window],
                api_store_sid=987654,
                api_store_name="Synthetic Shop",
            )
            workbook = load_workbook(updated)
            sheet = workbook["API-TEST"]
            self.assertEqual(sheet["A2"].value, "人工推广阶段")
            self.assertEqual(sheet["G2"].value, "人工发货时间")
            self.assertEqual(sheet["M6"].value, 4)
            self.assertEqual(sheet["M66"].value, 8)
            self.assertEqual(sheet["B68"].value, "D63")
            self.assertEqual(sheet["E68"].value, 2)
            workbook.close()

    def test_d01_requires_explicit_user_confirmation(self) -> None:
        start = date(2026, 1, 1)
        mapping = SkuProductMapping("API-TEST", ("MSKU-A",), ())
        row = SkuDailySales(987654, "API-TEST", "MSKU-A", "", start, "USD", 1, 1, Decimal("9.99"))
        window = SkuDailySalesWindow(start, start, (mapping,), (row,))
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.xlsx"
            create_template(template)
            with self.assertRaisesRegex(ReportGenerationError, "尚未由用户明确确认"):
                generate_reports(
                    template_path=template,
                    data_folder=None,
                    mapping_path=None,
                    skus=["API-TEST"],
                    output_path=root / "blocked.xlsx",
                    d01_dates={"API-TEST": start},
                    api_sales_windows=[window],
                    api_store_sid=987654,
                    api_store_name="Synthetic Shop",
                )


class D01SettingsTests(unittest.TestCase):
    def test_confirmed_d01_dates_persist_per_store_and_sku(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "d01_dates.json"
            first = date(2026, 8, 1)
            second = date(2026, 9, 3)
            save_d01_dates(510226, {"EX3106": first}, path=path)
            save_d01_dates(510226, {"ex3114": second}, path=path)
            save_d01_dates(510227, {"EX3106": second}, path=path)
            self.assertEqual(get_d01_dates(510226, ["ex3106", "EX3114"], path=path), {
                "ex3106": first, "ex3114": second,
            })
            self.assertEqual(get_d01_dates(510227, ["ex3106"], path=path), {"ex3106": second})
            self.assertEqual(set(load_d01_settings(path)), {"510226", "510227"})
            self.assertNotIn("secret", path.read_text(encoding="utf-8").casefold())

    def test_corrupt_d01_settings_fail_closed_without_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "d01_dates.json"
            path.write_text("{broken", encoding="utf-8")
            with self.assertRaises(D01SettingsError):
                save_d01_dates(510226, {"SKU_A": date(2026, 1, 1)}, path=path)
            self.assertEqual(path.read_text(encoding="utf-8"), "{broken")

    def test_d01_settings_reject_invalid_dates_and_sid(self) -> None:
        with self.assertRaises(D01SettingsError):
            save_d01_dates(True, {"SKU_A": date(2026, 1, 1)}, path=Path("unused.json"))
        with self.assertRaises(D01SettingsError):
            save_d01_dates(510226, {"SKU_A": "2026-01-01"}, path=Path("unused.json"))  # type: ignore[arg-type]


class SpAdApiTests(unittest.TestCase):
    @staticmethod
    def _envelope(records: list[dict[str, Any]], total: int | None = None) -> dict[str, Any]:
        return {
            "code": 0,
            "message": "success",
            "error_details": [],
            "request_id": "synthetic-request",
            "response_time": "2026-08-05 10:00:00",
            "total": len(records) if total is None else total,
            "data": records,
        }

    def test_sp_report_parsers_validate_costs_ids_dates_and_business_code(self) -> None:
        report_day = date(2026, 8, 5)
        campaign = self._envelope([{
            "campaign_id": 123,
            "profile_id": 987,
            "cost": 3.25,
            "report_date": report_day.isoformat(),
        }])
        product = self._envelope([{
            "ad_id": 456,
            "asin": "ASINTEST0001",
            "sku": "MSKU-A",
            "ad_group_id": 789,
            "profile_id": 987,
            "campaign_id": 123,
            "cost": 3.25,
            "report_date": report_day.isoformat(),
        }])
        total, campaign_rows = parse_sp_campaign_report_page(campaign, report_date=report_day)
        self.assertEqual(total, 1)
        self.assertEqual(campaign_rows[0].cost, Decimal("3.25"))
        total, product_rows = parse_sp_product_ad_report_page(product, report_date=report_day)
        self.assertEqual(total, 1)
        self.assertEqual(product_rows[0].sku, "MSKU-A")
        self.assertEqual(product_rows[0].campaign_id, "123")

        wrong_code = dict(campaign, code=1)
        with self.assertRaisesRegex(LingxingApiError, "业务 code"):
            parse_sp_campaign_report_page(wrong_code, report_date=report_day)
        wrong_date = self._envelope([dict(campaign["data"][0], report_date="2026-08-06")])
        with self.assertRaisesRegex(LingxingApiError, "请求日期不一致"):
            parse_sp_campaign_report_page(wrong_date, report_date=report_day)
        wrong_cost = self._envelope([dict(product["data"][0], cost="3.25")])
        with self.assertRaisesRegex(LingxingApiError, "字段 cost 类型无效"):
            parse_sp_product_ad_report_page(wrong_cost, report_date=report_day)

    def test_sp_product_cost_is_unique_per_sku_and_cross_checked_to_campaign_total(self) -> None:
        report_day = date(2026, 8, 5)
        campaigns = [{"campaign_id": 100, "profile_id": 900, "cost": 3, "report_date": report_day.isoformat()}]
        products = [
            {"ad_id": 1, "asin": "ASINTEST0001", "sku": "MSKU-A", "ad_group_id": 10, "profile_id": 900, "campaign_id": 100, "cost": 1.25, "report_date": report_day.isoformat()},
            {"ad_id": 2, "asin": "ASINTEST0002", "sku": "MSKU-B", "ad_group_id": 10, "profile_id": 900, "campaign_id": 100, "cost": 1.75, "report_date": report_day.isoformat()},
        ]
        requests: list[dict[str, Any]] = []
        headers_seen: list[str | None] = []

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content)
            requests.append({"path": request.url.path, **body})
            headers_seen.append(request.headers.get("X-API-VERSION"))
            records = campaigns if request.url.path.endswith("spCampaignReports") else products
            return httpx.Response(200, json=self._envelope(records))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient("1234567890123456", "fixture-app-secret", http_client=http) as client:
            result = client.query_sp_ad_spend_history(
                access_token="fixture-token",
                store_sid=987654,
                mappings=[
                    SkuProductMapping("LOCAL-A", ("MSKU-A",), ()),
                    SkuProductMapping("LOCAL-B", ("MSKU-B",), ()),
                ],
                start_date=report_day,
                end_date=report_day,
            )
        self.assertEqual(len(requests), 2)
        self.assertTrue(all(item["sid"] == 987654 and item["show_detail"] == 0 and item["length"] == 15 for item in requests))
        self.assertEqual(headers_seen, ["2", "2"])
        self.assertEqual(len(result[0].discrepancies), 0)
        self.assertEqual({item.local_sku: item.actual_spend for item in result[0].sku_daily_spends}, {
            "LOCAL-A": Decimal("1.25"), "LOCAL-B": Decimal("1.75"),
        })
        self.assertEqual(sum(item.actual_spend for item in result[0].sku_daily_spends), Decimal("3.00"))
        http.close()

    def test_sp_cost_mismatch_blocks_sku_attribution_and_pagination_is_complete(self) -> None:
        report_day = date(2026, 8, 5)
        campaign_records = [{"campaign_id": 100, "profile_id": 900, "cost": 4, "report_date": report_day.isoformat()}]
        product_records = [
            {"ad_id": index + 1, "asin": "ASINTEST0001", "sku": "MSKU-A", "ad_group_id": 10, "profile_id": 900, "campaign_id": 100, "cost": 0.25, "report_date": report_day.isoformat()}
            for index in range(16)
        ]
        request_offsets: list[tuple[str, int]] = []

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content)
            request_offsets.append((request.url.path, body["offset"]))
            if request.url.path.endswith("spCampaignReports"):
                return httpx.Response(200, json=self._envelope(campaign_records))
            offset = body["offset"]
            return httpx.Response(200, json=self._envelope(product_records[offset:offset + 15], len(product_records)))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient("1234567890123456", "fixture-app-secret", http_client=http) as client:
            result = client.query_sp_ad_spend_history(
                access_token="fixture-token",
                store_sid=987654,
                mappings=[SkuProductMapping("LOCAL-A", ("MSKU-A",), ())],
                start_date=report_day,
                end_date=report_day,
            )
        self.assertEqual(request_offsets, [
            ("/pb/openapi/newad/spCampaignReports", 0),
            ("/pb/openapi/newad/spProductAdReports", 0),
            ("/pb/openapi/newad/spProductAdReports", 15),
        ])
        self.assertEqual(len(result[0].product_rows), 16)
        self.assertEqual(len(result[0].discrepancies), 0)
        self.assertEqual(result[0].sku_daily_spends[0].actual_spend, Decimal("4.00"))
        http.close()

        def mismatch_handler(request: httpx.Request) -> httpx.Response:
            rows = campaign_records if request.url.path.endswith("spCampaignReports") else product_records[:1]
            return httpx.Response(200, json=self._envelope(rows))

        http = httpx.Client(transport=httpx.MockTransport(mismatch_handler))
        with LingxingOpenApiClient("1234567890123456", "fixture-app-secret", http_client=http) as client:
            mismatch = client.query_sp_ad_spend_history(
                access_token="fixture-token",
                store_sid=987654,
                mappings=[SkuProductMapping("LOCAL-A", ("MSKU-A",), ())],
                start_date=report_day,
                end_date=report_day,
            )
        self.assertEqual(len(mismatch[0].discrepancies), 1)
        self.assertEqual(mismatch[0].sku_daily_spends, ())
        http.close()


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
                "properties": {
                    "offset": {"type": "integer"},
                    "length": {"type": "integer"},
                    "search_value": {"type": "array"},
                },
                "required": ["offset", "length", "search_value"],
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
        valid_arguments = {"offset": 0, "length": 5, "search_value": ["local-test-a"]}
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, valid_arguments, VerifiedReadOnlyGrant("other_id", business.permission_fingerprint))
        self.assertEqual(len(calls), 1, "unverified action must not be sent")
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {"length": 5, "search_value": ["local-test-a"]}, grant)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {"offset": 0, "search_value": ["local-test-a"]}, grant)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {**valid_arguments, "unexpected": "value"}, grant)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {**valid_arguments, "offset": "0"}, grant)
        with self.assertRaises(McpClientError):
            client.call_verified_business_query(action_tool, business, {**valid_arguments, "search_value": [{"sku": "local-test-a"}]}, grant)
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
            client.call_verified_business_query(bad_outer_schema, business, valid_arguments, grant)
        self.assertEqual(len(calls), 1, "invalid params and a changed ID/Schema must not send action")
        result = client.call_verified_business_query(action_tool, business, valid_arguments, grant)
        self.assertIn("content", result)
        self.assertEqual(calls[-1], {"name": "action", "arguments": {"toolId": business.tool_id, "params": valid_arguments}})
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

    def test_array_without_items_uses_scalar_allowlist_and_keeps_array_constraints(self) -> None:
        schema = {
            "type": "object",
            "properties": {
                "search_value": {
                    "type": "array",
                    "minItems": 1,
                    "maxItems": 3,
                    "uniqueItems": True,
                }
            },
            "required": ["search_value"],
            "additionalProperties": False,
        }
        validate_json_schema(schema, {"search_value": ["local-test-a"]})
        validate_json_schema(schema, {"search_value": ["local-test-a", 0, None]})
        with self.assertRaises(McpClientError):
            validate_json_schema(schema, {"search_value": []})
        with self.assertRaises(McpClientError):
            validate_json_schema(schema, {"search_value": ["local-test-a", "local-test-a"]})
        with self.assertRaises(McpClientError):
            validate_json_schema(schema, {"search_value": ["local-test-a", "local-test-b", "local-test-c", "local-test-d"]})
        with self.assertRaises(McpClientError):
            validate_json_schema(schema, {"search_value": [{"sku": "local-test-a"}]})
        with self.assertRaises(McpClientError):
            validate_json_schema(schema, {"search_value": [["local-test-a"]]})


class LingxingOpenApiTests(unittest.TestCase):
    @staticmethod
    def _seller_record() -> dict[str, object]:
        return {
            "sid": 987654,
            "mid": 3001,
            "name": "Synthetic Store",
            "seller_id": "synthetic-seller-id",
            "account_name": "synthetic-account",
            "seller_account_id": 40001,
            "region": "R_TEST",
            "country": "ZZ",
            "marketplace_id": "MARKETPLACE_TEST",
            "status": 17,
            "has_ads_setting": 23,
        }

    @classmethod
    def _seller_response(cls, data: object | None = None) -> dict[str, object]:
        return {
            "code": 0,
            "message": "success",
            "error_details": [],
            "request_id": "synthetic-request-id",
            "response_time": "2026-01-02T03:04:05Z",
            "data": [cls._seller_record()] if data is None else data,
        }

    def test_seller_response_requires_business_success_code_and_maps_fields(self) -> None:
        payload = self._seller_response()
        sellers = parse_seller_list_response(payload)
        self.assertEqual(len(sellers), 1)
        seller = sellers[0]
        self.assertEqual(seller.sid, 987654)
        self.assertEqual(seller.mid, 3001)
        self.assertEqual(seller.name, "Synthetic Store")
        self.assertEqual(seller.seller_id, "synthetic-seller-id")
        self.assertEqual(seller.account_name, "synthetic-account")
        self.assertEqual(seller.seller_account_id, 40001)
        self.assertEqual(seller.site_label, "ZZ / R_TEST / MARKETPLACE_TEST")
        # These values remain raw integers; no business meaning is inferred.
        self.assertEqual(seller.has_ads_setting, 23)
        self.assertEqual(seller.status, 17)
        self.assertEqual(seller.display_label, "Synthetic Store — ZZ / R_TEST / MARKETPLACE_TEST — SID 987654")
        self.assertNotIn("synthetic-seller-id", repr(seller))
        self.assertNotIn("987654", repr(seller))

    def test_seller_response_accepts_empty_direct_data_array(self) -> None:
        self.assertEqual(parse_seller_list_response(self._seller_response([])), [])

    def test_seller_response_rejects_business_errors_and_wrong_envelope_types(self) -> None:
        with self.assertRaisesRegex(LingxingApiError, "业务 code 9"):
            parse_seller_list_response({**self._seller_response(), "code": 9})
        with self.assertRaisesRegex(LingxingApiError, "integer 类型的业务 code"):
            parse_seller_list_response({**self._seller_response(), "code": "0"})
        with self.assertRaisesRegex(LingxingApiError, "integer 类型的业务 code"):
            parse_seller_list_response({**self._seller_response(), "code": True})
        with self.assertRaisesRegex(LingxingApiError, "data 数组"):
            parse_seller_list_response({key: value for key, value in self._seller_response().items() if key != "data"})
        with self.assertRaisesRegex(LingxingApiError, "data 必须是数组"):
            parse_seller_list_response(self._seller_response({"list": [self._seller_record()]}))
        for field, wrong_value in (
            ("message", 1),
            ("error_details", "none"),
            ("request_id", 12),
            ("response_time", None),
        ):
            with self.subTest(field=field), self.assertRaisesRegex(LingxingApiError, f"{field} 的类型无效"):
                parse_seller_list_response({**self._seller_response(), field: wrong_value})

    def test_seller_response_rejects_missing_or_wrong_record_field_types(self) -> None:
        for field in ("sid", "mid", "seller_account_id", "status", "has_ads_setting"):
            record = self._seller_record()
            record[field] = "not-an-integer"
            with self.subTest(field=field), self.assertRaisesRegex(LingxingApiError, f"字段 {field} 的类型无效"):
                parse_seller_list_response(self._seller_response([record]))
        for field in ("name", "seller_id", "account_name", "region", "country", "marketplace_id"):
            record = self._seller_record()
            record[field] = 123
            with self.subTest(field=field), self.assertRaisesRegex(LingxingApiError, f"字段 {field} 的类型无效"):
                parse_seller_list_response(self._seller_response([record]))
        record = self._seller_record()
        del record["seller_id"]
        with self.assertRaisesRegex(LingxingApiError, "字段 seller_id 的类型无效"):
            parse_seller_list_response(self._seller_response([record]))
        record = self._seller_record()
        record["sid"] = True
        with self.assertRaisesRegex(LingxingApiError, "字段 sid 的类型无效"):
            parse_seller_list_response(self._seller_response([record]))

    @staticmethod
    def _product_record(local_sku: str = "local-test-a", *, offset: int = 0) -> dict[str, object]:
        return {
            "sku": local_sku,
            "price_list": [
                {"local_sku": local_sku, "seller_sku": f"market-test-{offset + 1}", "sid": "987654"}
            ],
            "asins": [{"asin": f"ASINTEST{offset + 1:04d}", "sid": "987654"}],
        }

    def test_long_sales_history_is_split_into_nonoverlapping_92_day_windows(self) -> None:
        start = date(2026, 1, 1)
        end = start + timedelta(days=93)
        windows = split_date_windows(start, end)
        self.assertEqual([(left, right) for left, right in windows], [
            (start, start + timedelta(days=91)),
            (start + timedelta(days=92), end),
        ])
        self.assertTrue(all((right - left).days + 1 <= 92 for left, right in windows))

        mapping_calls: list[tuple[date, date]] = []
        daily_calls: list[tuple[date, date]] = []
        with LingxingOpenApiClient("synthetic-app-id", "synthetic-secret") as client:
            def query_mappings(**kwargs):
                mapping_calls.append((kwargs["start_date"], kwargs["end_date"]))
                return [SkuProductMapping("local-test-a", ("MSKU-A",), ())]

            def query_sales(**kwargs):
                daily_calls.append((kwargs["start_date"], kwargs["end_date"]))
                return []

            with patch.object(client, "query_product_mappings", side_effect=query_mappings), \
                    patch.object(client, "query_daily_sales", side_effect=query_sales):
                result = client.query_daily_sales_history(
                    access_token="synthetic-token",
                    store_sid=987654,
                    local_skus=["local-test-a"],
                    start_date=start,
                    end_date=end,
                )
        self.assertEqual(mapping_calls, windows)
        self.assertEqual(daily_calls, windows)
        self.assertEqual([(window.start_date, window.end_date) for window in result], windows)

    @staticmethod
    def _product_response(records: list[dict[str, object]], *, total: int | None = None) -> dict[str, object]:
        return {
            "code": 0,
            "msg": "success",
            "trace_id": "synthetic-trace-id",
            "data": {"total": len(records) if total is None else total, "list": records},
        }

    @staticmethod
    def _daily_record(
        day: str,
        metric_value: int | float,
        *,
        seller_sku: str = "market-test-1",
        asin: str = "ASINTEST0001",
    ) -> dict[str, object]:
        return {
            "sid": 987654,
            "r_date": day,
            "currency_code": "USD",
            "seller_sku": seller_sku,
            "asin": asin,
            "product_name": "Synthetic Product",
            "map_value": metric_value,
        }

    @staticmethod
    def _daily_response(records: list[dict[str, object]]) -> dict[str, object]:
        return {
            "code": 0,
            "message": "success",
            "error_details": [],
            "request_id": "synthetic-request-id",
            "response_time": "2026-01-02T03:04:05Z",
            "data": records,
        }

    def test_product_performance_parser_resolves_local_sku_to_store_msku_and_asin_candidates(self) -> None:
        payload = self._product_response([self._product_record()])
        total, mappings = parse_product_performance_page(
            payload,
            store_sid=987654,
            requested_skus={"local-test-a"},
        )
        self.assertEqual(total, 1)
        self.assertEqual(
            mappings,
            [SkuProductMapping("local-test-a", ("market-test-1",), ("ASINTEST0001",))],
        )

    def test_product_performance_parser_rejects_bad_envelope_and_untyped_mapping(self) -> None:
        good = self._product_response([self._product_record()])
        with self.assertRaisesRegex(LingxingApiError, "业务 code 7"):
            parse_product_performance_page(
                {**good, "code": 7}, store_sid=987654, requested_skus={"local-test-a"}
            )
        bad_row = self._product_record()
        bad_row["price_list"] = "not-an-array"
        with self.assertRaisesRegex(LingxingApiError, "price_list/asins 必须是数组"):
            parse_product_performance_page(
                self._product_response([bad_row]),
                store_sid=987654,
                requested_skus={"local-test-a"},
            )
        with self.assertRaisesRegex(LingxingApiError, "price_list.sid 的类型无效"):
            parse_product_performance_page(
                self._product_response([{
                    **self._product_record(),
                    "price_list": [{"local_sku": "local-test-a", "seller_sku": "market-test-1", "sid": True}],
                }]),
                store_sid=987654,
                requested_skus={"local-test-a"},
            )

    def test_daily_parser_preserves_explicit_zero_and_rejects_bad_site_date_or_types(self) -> None:
        day = date(2026, 1, 2)
        parsed = parse_asin_daily_page(
            self._daily_response([self._daily_record(day.isoformat(), 0)]),
            store_sid=987654,
            event_date=day,
            metric_type=2,
        )
        self.assertEqual(parsed[0].metric_value, Decimal("0"))
        self.assertEqual(parsed[0].event_date, day)
        with self.assertRaisesRegex(LingxingApiError, "站点日期不一致"):
            parse_asin_daily_page(
                self._daily_response([self._daily_record("2026-01-03", 1)]),
                store_sid=987654,
                event_date=day,
                metric_type=2,
            )
        with self.assertRaisesRegex(LingxingApiError, "不是整数"):
            parse_asin_daily_page(
                self._daily_response([self._daily_record(day.isoformat(), 1.5)]),
                store_sid=987654,
                event_date=day,
                metric_type=3,
            )
        with self.assertRaisesRegex(LingxingApiError, "SID 不匹配"):
            parse_asin_daily_page(
                self._daily_response([{**self._daily_record(day.isoformat(), 1), "sid": 1}]),
                store_sid=987654,
                event_date=day,
                metric_type=2,
            )

    def test_product_and_daily_queries_send_documented_parameters_and_keep_missing_distinct_from_zero(self) -> None:
        app_id = "ak_0123456789abc"
        token = "fixture-token"
        first_day = date(2026, 1, 2)
        second_day = date(2026, 1, 3)
        requests: list[tuple[httpx.Request, dict[str, object]]] = []

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content.decode("utf-8"))
            requests.append((request, body))
            if request.url.path == "/bd/productPerformance/openApi/asinList":
                return httpx.Response(200, json=self._product_response([self._product_record()]))
            self.assertEqual(request.url.path, "/erp/sc/data/sales_report/asinDailyLists")
            event_date = body["event_date"]
            metric_type = body["type"]
            if event_date == first_day.isoformat():
                values = {1: 12.34, 2: 2, 3: 1}
                records = [self._daily_record(event_date, values[metric_type])]
            elif metric_type == 2:
                # A returned zero is present data; omitted type 1/type 3 records
                # stay None in the normalized result.
                records = [self._daily_record(event_date, 0)]
            else:
                records = []
            return httpx.Response(200, json=self._daily_response(records))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with patch("amazon_weekly_report.providers.lingxing_api._PRODUCT_PERFORMANCE_LAST_RESPONSE_AT", None), \
                patch("amazon_weekly_report.providers.lingxing_api.time.time", return_value=1639734344), \
                patch("amazon_weekly_report.providers.lingxing_api.time.monotonic", return_value=100.0), \
                patch("amazon_weekly_report.providers.lingxing_api.time.sleep"):
            with LingxingOpenApiClient(app_id, "fixture-app-secret", http_client=http) as client:
                mappings = client.query_product_mappings(
                    access_token=token,
                    store_sid=987654,
                    local_skus=["local-test-a"],
                    start_date=first_day,
                    end_date=second_day,
                )
                rows = client.query_daily_sales(
                    access_token=token,
                    store_sid=987654,
                    mappings=mappings,
                    start_date=first_day,
                    end_date=second_day,
                )
        http.close()

        self.assertEqual(len(requests), 7)
        product_request, product_body = requests[0]
        self.assertEqual(product_request.method, "POST")
        self.assertEqual(product_request.url.path, "/bd/productPerformance/openApi/asinList")
        self.assertEqual(product_body["search_field"], "local_sku")
        self.assertEqual(product_body["search_value"], ["local-test-a"])
        self.assertEqual(product_body["length"], 10_000)
        self.assertEqual(product_body["sid"], "987654")
        self.assertEqual(product_body["summary_field"], "sku")
        self.assertIs(product_body["is_recently_enum"], False)
        self.assertEqual(product_body["start_date"], first_day.isoformat())
        self.assertEqual(product_body["end_date"], second_day.isoformat())
        self.assertEqual([request.url.path for request, _ in requests[1:]], [
            "/erp/sc/data/sales_report/asinDailyLists"
        ] * 6)
        for request, body in requests[1:]:
            self.assertEqual(body["sid"], 987654)
            self.assertEqual(body["asin_type"], 2)
            self.assertIn(body["event_date"], (first_day.isoformat(), second_day.isoformat()))
            self.assertIn(body["type"], (1, 2, 3))
            self.assertEqual(body["length"], 1000)
            query = parse_qs(request.url.query.decode("ascii"))
            self.assertEqual(query["access_token"], [token])
            self.assertEqual(query["timestamp"], ["1639734344"])
            self.assertEqual(query["sign"], [signed_post_query_parameters({
                "access_token": token,
                "app_key": app_id,
                "timestamp": 1639734344,
            }, body, app_id)["sign"]])

        self.assertEqual(len(rows), 2)
        first, second = rows
        self.assertEqual((first.units_sold, first.order_count, first.sales_amount), (2, 1, Decimal("12.34")))
        self.assertEqual(first.event_date, first_day)
        self.assertEqual(first.asin, "ASINTEST0001")
        self.assertEqual((second.units_sold, second.order_count, second.sales_amount), (0, None, None))
        self.assertEqual(second.event_date, second_day)
        self.assertEqual(second.asin, "ASINTEST0001")

    def test_product_mapping_paginates_using_offset_and_rejects_uncovered_pages(self) -> None:
        requested = ["local-test-a", "local-test-b"]
        offsets: list[int] = []

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content.decode("utf-8"))
            offsets.append(body["offset"])
            record_index = body["offset"]
            records = [
                self._product_record(requested[record_index], offset=record_index)
            ] if record_index < len(requested) else []
            return httpx.Response(200, json=self._product_response(records, total=len(requested)))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with patch("amazon_weekly_report.providers.lingxing_api.PRODUCT_PERFORMANCE_PAGE_LENGTH", 1), \
                patch("amazon_weekly_report.providers.lingxing_api._PRODUCT_PERFORMANCE_LAST_RESPONSE_AT", None), \
                patch("amazon_weekly_report.providers.lingxing_api.time.monotonic", return_value=100.0), \
                patch("amazon_weekly_report.providers.lingxing_api.time.sleep"):
            with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
                mappings = client.query_product_mappings(
                    access_token="fixture-token",
                    store_sid=987654,
                    local_skus=requested,
                    start_date=date(2026, 1, 1),
                    end_date=date(2026, 1, 2),
                )
        http.close()
        self.assertEqual(offsets, [0, 1])
        self.assertEqual([mapping.local_sku for mapping in mappings], requested)

    def test_product_mapping_rejects_duplicate_local_sku_across_pages(self) -> None:
        record = self._product_record()

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content.decode("utf-8"))
            return httpx.Response(200, json=self._product_response([record], total=2))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with patch("amazon_weekly_report.providers.lingxing_api.PRODUCT_PERFORMANCE_PAGE_LENGTH", 1), \
                patch("amazon_weekly_report.providers.lingxing_api._PRODUCT_PERFORMANCE_LAST_RESPONSE_AT", None), \
                patch("amazon_weekly_report.providers.lingxing_api.time.monotonic", return_value=100.0), \
                patch("amazon_weekly_report.providers.lingxing_api.time.sleep"):
            with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
                with self.assertRaisesRegex(LingxingApiError, "重复 local SKU"):
                    client.query_product_mappings(
                        access_token="fixture-token",
                        store_sid=987654,
                        local_skus=["local-test-a", "local-test-b"],
                        start_date=date(2026, 1, 1),
                        end_date=date(2026, 1, 2),
                    )
        http.close()

    def test_daily_statistics_paginates_serially_and_stops_on_rate_limit(self) -> None:
        day = date(2026, 1, 2)
        offsets: list[tuple[int, int]] = []

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content.decode("utf-8"))
            metric_type = body["type"]
            offset = body["offset"]
            offsets.append((metric_type, offset))
            page_rows = [
                self._daily_record(day.isoformat(), 3, seller_sku="market-other"),
                self._daily_record(day.isoformat(), {1: 12.34, 2: 2, 3: 1}[metric_type]),
            ]
            records = page_rows[offset:offset + 1]
            return httpx.Response(200, json=self._daily_response(records))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with patch("amazon_weekly_report.providers.lingxing_api.ASIN_DAILY_PAGE_LENGTH", 1):
            with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
                rows = client.query_daily_sales(
                    access_token="fixture-token",
                    store_sid=987654,
                    mappings=[SkuProductMapping("local-test-a", ("market-test-1",), ("ASINTEST0001",))],
                    start_date=day,
                    end_date=day,
                )
        http.close()
        self.assertEqual(offsets, [(metric, offset) for metric in (1, 2, 3) for offset in (0, 1, 2)])
        self.assertEqual(len(rows), 1)
        self.assertEqual((rows[0].units_sold, rows[0].order_count, rows[0].sales_amount), (2, 1, Decimal("12.34")))

        limited_http = httpx.Client(
            transport=httpx.MockTransport(lambda _request: httpx.Response(429, headers={"Retry-After": "9"}))
        )
        with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=limited_http) as client:
            with self.assertRaisesRegex(LingxingApiError, "等待 9 秒"):
                client.query_daily_sales(
                    access_token="fixture-token",
                    store_sid=987654,
                    mappings=[SkuProductMapping("local-test-a", ("market-test-1",), ())],
                    start_date=day,
                    end_date=day,
                )
        limited_http.close()

    def test_daily_statistics_rejects_duplicate_msku_across_pages(self) -> None:
        day = date(2026, 1, 2)

        def handler(request: httpx.Request) -> httpx.Response:
            body = json.loads(request.content.decode("utf-8"))
            return httpx.Response(200, json=self._daily_response([self._daily_record(day.isoformat(), 1)]))

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with patch("amazon_weekly_report.providers.lingxing_api.ASIN_DAILY_PAGE_LENGTH", 1):
            with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
                with self.assertRaisesRegex(LingxingApiError, "重复 MSKU"):
                    client.query_daily_sales(
                        access_token="fixture-token",
                        store_sid=987654,
                        mappings=[SkuProductMapping("local-test-a", ("market-test-1",), ())],
                        start_date=day,
                        end_date=day,
                    )
        http.close()

    def test_signature_matches_reference_algorithm_vector(self) -> None:
        app_id = "ak_0123456789abc"
        parameters = {
            "timestamp": 1639734344,
            "access_token": "fixture-token",
            "app_key": app_id,
        }
        self.assertEqual(
            canonical_signing_string(parameters),
            "access_token=fixture-token&app_key=ak_0123456789abc&timestamp=1639734344",
        )
        self.assertEqual(
            generate_signature(parameters, app_id),
            "DGIcRRQN/VQlrOtSjtCmntt5i7TsON7XPBbSVCNFij5IwVhSPzVcd0sPeLkAl3rO",
        )

    def test_signature_serialization_boundaries_fail_closed(self) -> None:
        self.assertEqual(
            canonical_signing_string({"z": "  value  ", "a": "", "enabled": True, "count": 5}),
            "a=&count=5&enabled=true&z=value",
        )
        for value in (None, 1.5, ["x"], {"x": "y"}):
            with self.subTest(value_type=type(value).__name__):
                with self.assertRaises(LingxingSignatureError):
                    canonical_signing_string({"value": value})
        with self.assertRaises(LingxingSignatureError):
            canonical_signing_string({"签名": "value"})
        with self.assertRaises(LingxingSignatureError):
            canonical_signing_string({"sign": "self-reference"})
        with self.assertRaises(LingxingSignatureError):
            generate_signature({"timestamp": 1}, "short")

    def test_signed_post_serializes_arrays_as_compact_json_for_sdk_signature(self) -> None:
        app_id = "ak_0123456789abc"
        auth = {"access_token": "fixture-token", "app_key": app_id, "timestamp": 1639734344}
        body = {
            "search_value": ["local sku", "本地SKU"],
            "sid": "987654",
            "offset": 0,
        }
        signed = signed_post_query_parameters(auth, body, app_id)
        canonical_parameters = {
            **auth,
            "search_value": '["local sku","本地SKU"]',
            "sid": "987654",
            "offset": 0,
        }
        self.assertEqual(signed["sign"], generate_signature(canonical_parameters, app_id))
        self.assertEqual(set(signed), {"access_token", "app_key", "timestamp", "sign"})
        with self.assertRaises(LingxingSignatureError):
            signed_post_query_parameters(auth, {"timestamp": 1}, app_id)

    def test_query_encoding_matches_java_urlencoder_for_utf8_and_base64(self) -> None:
        encoded = encode_query_parameters({"value": "a b+/%~=中", "sign": "A+/="})
        self.assertEqual(encoded, "value=a+b%2B%2F%25%7E%3D%E4%B8%AD&sign=A%2B%2F%3D")

    def test_list_sellers_sends_sdk_signed_get_and_parses_business_response(self) -> None:
        app_id = "ak_0123456789abc"
        token = "fixture-token"
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            return httpx.Response(
                200,
                json={
                    "code": 0,
                    "message": "success",
                    "error_details": [],
                    "request_id": "synthetic-request-id",
                    "response_time": "2026-01-02T03:04:05Z",
                    "data": [
                        self._seller_record(),
                    ],
                },
            )

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient(app_id, "fixture-app-secret", http_client=http) as client:
            with patch("amazon_weekly_report.providers.lingxing_api.time.time", return_value=1639734344):
                response = client.list_sellers(token)

        request = requests[0]
        self.assertEqual(request.method, "GET")
        self.assertEqual(request.url.path, "/erp/sc/data/seller/lists")
        self.assertEqual(request.content, b"")
        self.assertNotIn("fixture-app-secret", str(request.url))
        wire_query = request.url.query.decode("ascii")
        self.assertIn("%2F", wire_query, "Base64 slash must be URL encoded")
        query = parse_qs(wire_query, keep_blank_values=True)
        self.assertEqual(query["access_token"], [token])
        self.assertEqual(query["app_key"], [app_id])
        self.assertEqual(query["timestamp"], ["1639734344"])
        self.assertEqual(
            query["sign"],
            [generate_signature({"access_token": token, "app_key": app_id, "timestamp": 1639734344}, app_id)],
        )
        self.assertEqual(len(response), 1)
        self.assertEqual(response[0].name, "Synthetic Store")
        self.assertEqual(response[0].sid, 987654)
        self.assertEqual(response[0].has_ads_setting, 23)
        http.close()

    def test_list_sellers_rejects_http_200_business_error(self) -> None:
        payload = self._seller_response()
        payload["code"] = 41
        http = httpx.Client(
            transport=httpx.MockTransport(lambda _request: httpx.Response(200, json=payload))
        )
        with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
            with self.assertRaisesRegex(LingxingApiError, "业务 code 41"):
                client.list_sellers("fixture-token")
        http.close()

    def test_list_sellers_http_error_does_not_echo_token_or_signature(self) -> None:
        token = "fixture-access-token"
        http = httpx.Client(
            transport=httpx.MockTransport(
                lambda _request: httpx.Response(403, text=f"denied {token} fake-signature")
            )
        )
        with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
            with self.assertRaises(LingxingApiError) as raised:
                client.list_sellers(token)
        self.assertNotIn(token, str(raised.exception))
        self.assertNotIn("fake-signature", str(raised.exception))
        http.close()

    def test_list_sellers_rejects_non_success_redirect_response(self) -> None:
        http = httpx.Client(
            transport=httpx.MockTransport(lambda _request: httpx.Response(302, json={"code": 0, "data": []}))
        )
        with LingxingOpenApiClient("ak_0123456789abc", "fixture-app-secret", http_client=http) as client:
            with self.assertRaisesRegex(LingxingApiError, "HTTP 302"):
                client.list_sellers("fixture-token")
        http.close()

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

    def test_access_token_cache_reuses_only_tokens_with_expiry(self) -> None:
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            token_number = len(requests)
            return httpx.Response(
                200,
                json={"code": 0, "data": {"access_token": f"fixture-token-{token_number}", "expires_in": 3600}},
            )

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            first = client.get_access_token()
            cached = client.get_access_token()
            self.assertIs(first, cached)
            self.assertEqual(len(requests), 1)

            refreshed = client.get_access_token(force_refresh=True)
            self.assertEqual(refreshed.token, "fixture-token-2")
            self.assertEqual(len(requests), 2)

            client.clear_token_cache()
            after_clear = client.get_access_token()
            self.assertEqual(after_clear.token, "fixture-token-3")
            self.assertEqual(len(requests), 3)
        http.close()

    def test_access_token_without_expiry_is_not_cached(self) -> None:
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            return httpx.Response(200, json={"code": 0, "data": {"access_token": f"fixture-token-{len(requests)}"}})

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            first = client.get_access_token()
            second = client.get_access_token()
        self.assertEqual((first.token, second.token), ("fixture-token-1", "fixture-token-2"))
        self.assertEqual(len(requests), 2)
        http.close()

    def test_access_token_near_expiry_is_refreshed(self) -> None:
        requests: list[httpx.Request] = []

        def handler(request: httpx.Request) -> httpx.Response:
            requests.append(request)
            expires_in = 1 if len(requests) == 1 else 3600
            return httpx.Response(
                200,
                json={"code": 0, "data": {"access_token": f"fixture-token-{len(requests)}", "expires_in": expires_in}},
            )

        http = httpx.Client(transport=httpx.MockTransport(handler))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            first = client.get_access_token()
            second = client.get_access_token()
        self.assertEqual((first.token, second.token), ("fixture-token-1", "fixture-token-2"))
        self.assertEqual(len(requests), 2)
        http.close()

    def test_unverified_sku_sales_query_fails_closed_without_sending_request(self) -> None:
        requests: list[httpx.Request] = []
        http = httpx.Client(transport=httpx.MockTransport(lambda request: requests.append(request) or httpx.Response(200)))
        with LingxingOpenApiClient("fixture-app-id", "fixture-app-secret", http_client=http) as client:
            with self.assertRaises(LingxingApiNotReady):
                client.query_sku_sales(
                    access_token="fixture-token",
                    seller_id="fixture-seller",
                    sku="local-test-a",
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


class OpenApiPreviewTests(unittest.TestCase):
    class _Value:
        def __init__(self) -> None:
            self.value = ""

        def set(self, value: str) -> None:
            self.value = value

    class _Tree:
        def __init__(self) -> None:
            self.rows: list[tuple[object, ...]] = []

        def insert(self, _parent: str, _index: str, *, values: tuple[object, ...]) -> None:
            self.rows.append(values)

    def test_gui_preview_formats_missing_and_explicit_zero_without_claiming_historical_first_sale(self) -> None:
        app = object.__new__(WeeklyReportApp)
        app.report_mapping_status = self._Value()
        app.report_query_status = self._Value()
        app.report_sales_tree = self._Tree()
        day = date(2026, 1, 2)
        result = {
            "_mappings": [SkuProductMapping("local-test-a", ("market-test-1",), ("ASINTEST0001",))],
            "_daily_rows": [
                SkuDailySales(
                    sid=987654,
                    local_sku="local-test-a",
                    seller_sku="market-test-1",
                    asin="ASINTEST0001",
                    event_date=day,
                    currency_code="USD",
                    units_sold=0,
                    order_count=None,
                    sales_amount=None,
                )
            ],
            "_requested_skus": ["local-test-a"],
            "_unresolved_skus": [],
            "_start_date": day,
            "_end_date": day,
            "_store_name": "Synthetic Store",
        }
        app._apply_daily_sales_preview(result)
        self.assertEqual(app.report_sales_tree.rows[0][4:8], ("0", "", "", "USD"))
        self.assertIn("ASIN 候选未假定与某个 MSKU 一一对应", app.report_mapping_status.value)
        self.assertIn("所选范围未发现明确正销量记录", app.report_query_status.value)
        self.assertIn("不是已核验的历史首单日期", app.report_query_status.value)


if __name__ == "__main__":
    unittest.main()
