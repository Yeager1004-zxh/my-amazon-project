from __future__ import annotations

import json
from pathlib import Path
from typing import Any


REPORT_FIELDS: dict[str, tuple[str, ...]] = {
    "sales": ("store", "sku", "date", "units_sold", "sales_amount"),
    "actual_ad_spend": ("store", "sku", "date", "actual_spend"),
}
OPTIONAL_REPORT_FIELDS: dict[str, tuple[str, ...]] = {
    "sales": ("unit_price", "organic_orders"),
    "actual_ad_spend": ("campaign_id", "ad_group_id", "campaign_name", "campaign_disabled_at"),
    "ad_performance": ("store", "sku", "date", "ad_orders", "ad_sales", "impressions", "clicks", "campaign_id", "ad_group_id"),
    "traffic": ("store", "sku", "date", "sessions", "organic_clicks"),
    "product_snapshot": ("store", "sku", "category", "subcategory_rank", "rating", "review_count"),
}


def load_mapping(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON 格式错误，第 {exc.lineno} 行：{exc.msg}") from exc
    if not isinstance(value, dict):
        raise ValueError("映射配置的顶层必须是 JSON 对象")
    return value


def validate_mapping(mapping: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    if mapping.get("version") != 1:
        issues.append("version 必须为 1")
    reports = mapping.get("reports")
    if not isinstance(reports, dict):
        return issues + ["reports 必须是对象"]

    for report_name, required_fields in REPORT_FIELDS.items():
        report = reports.get(report_name)
        if not isinstance(report, dict):
            if report_name == "actual_ad_spend" and mapping.get("status") == "SALES_ONLY_READY":
                continue
            issues.append(f"缺少必需报表配置：{report_name}")
            continue
        headers = report.get("source_headers")
        if not isinstance(headers, dict):
            issues.append(f"{report_name}.source_headers 必须是对象")
            continue
        configured = report_name == "sales" or any(
            isinstance(headers.get(field), str) and headers[field].strip()
            for field in required_fields
        ) or bool(report.get("file_patterns"))
        if report_name == "actual_ad_spend" and not configured and mapping.get("status") == "SALES_ONLY_READY":
            continue
        if report_name == "actual_ad_spend" and not configured:
            issues.append("未配置广告实际花费来源；如需先生成仅含销售数据的周报，请将 status 设为 SALES_ONLY_READY")
            continue
        for field in required_fields:
            value = headers.get(field)
            if not isinstance(value, str) or not value.strip():
                issues.append(f"{report_name}.{field} 尚未映射到真实源表头")
        for field in OPTIONAL_REPORT_FIELDS.get(report_name, ()):
            value = headers.get(field)
            if value not in (None, "") and not isinstance(value, str):
                issues.append(f"{report_name}.{field} 的源表头映射必须是文本")
        header_row = report.get("header_row")
        if not isinstance(header_row, int) or header_row < 1:
            issues.append(f"{report_name}.header_row 必须是从 1 开始的实际表头行号")

        if report_name == "actual_ad_spend":
            if report.get("sku_scope") != "SKU_LEVEL_CONFIRMED":
                issues.append(
                    "actual_ad_spend.sku_scope 未确认为 SKU_LEVEL_CONFIRMED；"
                    "必须先核实原始实际花费记录可唯一对应一个 SKU，禁止按活动或广告组重复分摊"
                )

        patterns = report.get("file_patterns")
        if report_name == "sales" or configured:
            if not isinstance(patterns, list) or not patterns or any(
                not isinstance(pattern, str) or not pattern.strip() for pattern in patterns
            ):
                issues.append(f"{report_name}.file_patterns 必须指定至少一个真实文件名匹配规则")
        date_format = report.get("date_format")
        if date_format not in (None, "") and not isinstance(date_format, str):
            issues.append(f"{report_name}.date_format 必须是日期格式文本")

        mapped_values = [
            headers[field].strip()
            for field in required_fields
            if isinstance(headers.get(field), str) and headers[field].strip()
        ]
        if len(mapped_values) != len(set(mapped_values)):
            issues.append(f"{report_name} 中多个业务字段映射到了同一个源表头")

    for report_name, optional_fields in OPTIONAL_REPORT_FIELDS.items():
        if report_name in REPORT_FIELDS:
            continue
        optional = reports.get(report_name)
        if not isinstance(optional, dict):
            continue
        headers = optional.get("source_headers", {})
        mapped = {
            field: value.strip()
            for field, value in headers.items()
            if field in optional_fields and isinstance(value, str) and value.strip()
        } if isinstance(headers, dict) else {}
        enabled = bool(optional.get("enabled")) or bool(mapped) or bool(optional.get("file_patterns"))
        if not enabled:
            continue
        if not mapped:
            issues.append(f"{report_name} 已启用，但没有映射任何源字段")
        if report_name in {"ad_performance", "traffic"}:
            for join_field in ("store", "sku", "date"):
                if not mapped.get(join_field):
                    issues.append(f"{report_name}.{join_field} 必须配置，才能按店铺/SKU/日期连接数据")
        if not isinstance(optional.get("header_row"), int) or optional["header_row"] < 1:
            issues.append(f"{report_name}.header_row 必须是从 1 开始的实际表头行号")
        patterns = optional.get("file_patterns")
        if not isinstance(patterns, list) or not patterns or any(
            not isinstance(pattern, str) or not pattern.strip() for pattern in patterns
        ):
            issues.append(f"{report_name}.file_patterns 必须指定至少一个真实文件名匹配规则")
        date_format = optional.get("date_format")
        if date_format not in (None, "") and not isinstance(date_format, str):
            issues.append(f"{report_name}.date_format 必须是日期格式文本")
        unknown_fields = set(headers) - set(optional_fields) if isinstance(headers, dict) else set()
        if unknown_fields:
            issues.append(f"{report_name} 包含未知字段：{', '.join(sorted(unknown_fields))}")
        mapped_values = list(mapped.values())
        if len(mapped_values) != len(set(mapped_values)):
            issues.append(f"{report_name} 中多个业务字段映射到了同一个源表头")

    if mapping.get("status") not in {"READY_FOR_REVIEW", "SALES_ONLY_READY"}:
        issues.append("status 需设为 READY_FOR_REVIEW 或 SALES_ONLY_READY；填写前仍须基于真实样本核对")
    if mapping.get("status") == "READY_FOR_REVIEW":
        spend = reports.get("actual_ad_spend")
        if not isinstance(spend, dict) or spend.get("sku_scope") != "SKU_LEVEL_CONFIRMED":
            issues.append("READY_FOR_REVIEW 必须确认实际广告花费源可按 SKU 唯一归属")
    return issues
