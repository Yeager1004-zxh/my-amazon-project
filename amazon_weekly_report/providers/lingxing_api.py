from __future__ import annotations

import math
import json
import re
import threading
import time
from collections import defaultdict
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from typing import Any

import httpx

from ..ad_attribution import PromotionTargetEvidence
from .lingxing_signing import (
    encode_query_parameters,
    signed_post_query_parameters,
    signed_query_parameters,
)


OPENAPI_BASE_URL = "https://openapi.lingxing.com"
ACCESS_TOKEN_PATH = "/api/auth-server/oauth/access-token"
SELLER_LIST_PATH = "/erp/sc/data/seller/lists"
PRODUCT_PERFORMANCE_PATH = "/bd/productPerformance/openApi/asinList"
ASIN_DAILY_LISTS_PATH = "/erp/sc/data/sales_report/asinDailyLists"
SP_CAMPAIGN_REPORTS_PATH = "/pb/openapi/newad/spCampaignReports"
SP_PRODUCT_AD_REPORTS_PATH = "/pb/openapi/newad/spProductAdReports"
SB_CAMPAIGN_REPORTS_PATH = "/pb/openapi/newad/hsaCampaignReports"
SB_AD_GROUP_REPORTS_PATH = "/pb/openapi/newad/hsaAdGroupReports"
SD_CAMPAIGN_REPORTS_PATH = "/pb/openapi/newad/sdCampaignReports"
SD_PRODUCT_AD_REPORTS_PATH = "/pb/openapi/newad/sdProductAdReports"
SB_PRODUCT_ADS_PATH = "/pb/openapi/newad/hsaProductAds"
SB_DIVIDE_ASIN_REPORTS_PATH = "/pb/openapi/newad/sbDivideAsinReports"
SD_PRODUCT_ADS_PATH = "/pb/openapi/newad/sdProductAds"
PRODUCT_PERFORMANCE_PAGE_LENGTH = 10_000
ASIN_DAILY_PAGE_LENGTH = 1_000
AD_REPORT_PAGE_LENGTH = 15
SP_REPORT_PAGE_LENGTH = AD_REPORT_PAGE_LENGTH
MAX_PRODUCT_QUERY_DAYS = 92
_PRODUCT_PERFORMANCE_REQUEST_LOCK = threading.Lock()
_PRODUCT_PERFORMANCE_LAST_RESPONSE_AT: float | None = None
_ASIN_DAILY_REQUEST_LOCK = threading.Lock()
_AD_REPORT_REQUEST_LOCK = threading.Lock()


class LingxingApiError(RuntimeError):
    """A sanitized Lingxing OpenAPI transport or response error."""


class LingxingApiNotReady(LingxingApiError):
    """A query is deliberately blocked until its official contract is verified."""


@dataclass(frozen=True)
class SkuProductMapping:
    """A local SKU's MSKU/ASIN identities for one selected Lingxing store."""

    local_sku: str = field(repr=False)
    seller_skus: tuple[str, ...] = field(repr=False)
    asins: tuple[str, ...] = field(repr=False)


@dataclass(frozen=True)
class ProductPerformanceMetrics:
    """One SKU or MSKU summary returned for an explicitly requested date interval.

    The endpoint response has no per-row date. Callers must retain the request
    interval and must not treat a multi-day value as a daily observation.
    """

    store_sid: int = field(repr=False)
    local_sku: str = field(repr=False)
    seller_skus: tuple[str, ...] = field(repr=False)
    asins: tuple[str, ...] = field(repr=False)
    start_date: date = field(repr=False)
    end_date: date = field(repr=False)
    currency_code: str | None = field(repr=False)
    values: dict[str, Any] = field(repr=False)
    present_fields: frozenset[str] = field(repr=False)
    summary_field: str = field(default="sku", repr=False)
    summary_key: str | None = field(default=None, repr=False)


PRODUCT_PERFORMANCE_PREVIEW_FIELDS: tuple[str, ...] = (
    "sessions_total",
    "ad_order_quantity",
    "ad_sales_amount",
    "impressions",
    "clicks",
    "ctr",
    "ad_cvr",
    "avg_star",
    "reviews_count",
    "small_cate_rank",
    "cate_rank",
    "prev_cate_rank",
    "rank_category",
    "ranking_update_time",
    "spend",
    "ads_sp_cost",
    "shared_ads_sb_cost",
    "shared_ads_sbv_cost",
    "ads_sd_cost",
    "shared_ads_al_cost",
    "shared_ads_cc_cost",
    "shared_ads_sspaot_cost",
    "shared_ads_sar_cost",
    "shared_cost_of_advertising",
    "adv_rate",
    "currency_code",
)

_PRODUCT_PERFORMANCE_INTEGER_FIELDS = frozenset({
    "sessions_total", "ad_order_quantity", "impressions", "clicks", "reviews_count",
    "cate_rank", "prev_cate_rank", "volume", "order_items",
})
_PRODUCT_PERFORMANCE_NUMBER_FIELDS = frozenset({
    "ad_sales_amount", "ctr", "ad_cvr", "avg_star", "spend", "ads_sp_cost",
    "shared_ads_sb_cost", "shared_ads_sbv_cost", "ads_sd_cost", "shared_ads_al_cost",
    "shared_ads_cc_cost", "shared_ads_sspaot_cost", "shared_ads_sar_cost",
    "shared_cost_of_advertising", "adv_rate", "amount", "acos",
})


@dataclass(frozen=True)
class _DailyMetricRecord:
    sid: int = field(repr=False)
    event_date: date = field(repr=False)
    currency_code: str = field(repr=False)
    seller_sku: str = field(repr=False)
    asin: str = field(repr=False)
    product_name: str = field(repr=False)
    metric_value: Decimal = field(repr=False)


@dataclass(frozen=True)
class SkuDailySales:
    """Daily API values; None means no API record, while zero is explicit zero."""

    sid: int = field(repr=False)
    local_sku: str = field(repr=False)
    seller_sku: str = field(repr=False)
    asin: str = field(repr=False)
    event_date: date = field(repr=False)
    currency_code: str = field(repr=False)
    units_sold: int | None = field(repr=False)
    order_count: int | None = field(repr=False)
    sales_amount: Decimal | None = field(repr=False)


@dataclass(frozen=True)
class SkuDailySalesWindow:
    """One successfully queried date window and its exact SKU/MSKU join map."""

    start_date: date = field(repr=False)
    end_date: date = field(repr=False)
    mappings: tuple[SkuProductMapping, ...] = field(repr=False)
    rows: tuple[SkuDailySales, ...] = field(repr=False)


@dataclass(frozen=True)
class SpCampaignSpend:
    report_date: date = field(repr=False)
    campaign_id: str = field(repr=False)
    cost: Decimal = field(repr=False)
    profile_id: str = field(default="", repr=False)


@dataclass(frozen=True)
class SpProductAdSpend:
    report_date: date = field(repr=False)
    sku: str = field(repr=False)
    asin: str = field(repr=False)
    ad_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    cost: Decimal = field(repr=False)
    profile_id: str = field(default="", repr=False)
    source: str = field(default="product_report", repr=False)
    allocation_verified: bool = field(default=False, repr=False)


@dataclass(frozen=True)
class AdProductCatalogRecord:
    """Current read-only ad/product identity record from the ads catalog APIs."""

    ad_type: str = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    ad_id: str = field(default="", repr=False)
    ad_creative_id: str = field(default="", repr=False)
    seller_sku: str = field(default="", repr=False)
    asins: tuple[str, ...] = field(default=(), repr=False)
    state: str = field(default="", repr=False)
    creative_type: str = field(default="", repr=False)
    serving_status: str = field(default="", repr=False)
    creation_date_value: Decimal | None = field(default=None, repr=False)
    last_updated_date_value: Decimal | None = field(default=None, repr=False)
    video_creative_types: tuple[str, ...] = field(default=(), repr=False)


@dataclass(frozen=True)
class SkuDailyAdSpend:
    store_sid: int = field(repr=False)
    local_sku: str = field(repr=False)
    seller_sku: str = field(repr=False)
    report_date: date = field(repr=False)
    actual_spend: Decimal = field(repr=False)
    campaign_count: int = field(repr=False)


@dataclass(frozen=True)
class SpAdSpendDiscrepancy:
    report_date: date = field(repr=False)
    campaign_id: str = field(repr=False)
    campaign_cost: Decimal | None = field(repr=False)
    detail_cost: Decimal | None = field(repr=False)
    detail_type: str = field(repr=False)
    profile_id: str = field(default="", repr=False)

    @property
    def product_cost(self) -> Decimal | None:
        """Backward-compatible name for the former SP product comparison field."""
        return self.detail_cost


@dataclass(frozen=True)
class SbAdGroupSpend:
    report_date: date = field(repr=False)
    campaign_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    cost: Decimal = field(repr=False)
    profile_id: str = field(default="", repr=False)


@dataclass(frozen=True)
class SpAdSpendWindow:
    start_date: date = field(repr=False)
    end_date: date = field(repr=False)
    campaign_rows: tuple[SpCampaignSpend, ...] = field(repr=False)
    product_rows: tuple[SpProductAdSpend, ...] = field(repr=False)
    sku_daily_spends: tuple[SkuDailyAdSpend, ...] = field(repr=False)
    discrepancies: tuple[SpAdSpendDiscrepancy, ...] = field(repr=False)
    ad_type: str = field(default="SP", repr=False)
    ad_group_rows: tuple[SbAdGroupSpend, ...] = field(default=(), repr=False)
    campaign_cost_total: Decimal | None = field(default=None, repr=False)
    detail_cost_total: Decimal | None = field(default=None, repr=False)
    unattributed_cost: Decimal | None = field(default=None, repr=False)
    report_complete: bool = field(default=False, repr=False)
    promotion_targets: tuple[PromotionTargetEvidence, ...] = field(default=(), repr=False)
    promotion_catalog_complete: bool = field(default=False, repr=False)
    product_catalog_rows: tuple[AdProductCatalogRecord, ...] = field(default=(), repr=False)
    product_catalog_observed_at: datetime | None = field(default=None, repr=False)


def split_date_windows(start_date: date, end_date: date, *, max_days: int = MAX_PRODUCT_QUERY_DAYS) -> list[tuple[date, date]]:
    """Split an inclusive range into contiguous, non-overlapping API windows."""
    if type(start_date) is not date or type(end_date) is not date or end_date < start_date:
        raise LingxingApiError("查询日期范围无效。")
    if isinstance(max_days, bool) or not isinstance(max_days, int) or max_days < 1:
        raise LingxingApiError("查询窗口长度无效。")
    windows: list[tuple[date, date]] = []
    current = start_date
    while current <= end_date:
        window_end = min(end_date, current + timedelta(days=max_days - 1))
        windows.append((current, window_end))
        current = window_end + timedelta(days=1)
    return windows


def _business_code(payload: Any, endpoint_label: str, *, message_field: str) -> dict[str, Any]:
    if not isinstance(payload, dict):
        raise LingxingApiError(f"{endpoint_label}业务响应必须是 JSON 对象。")
    code = payload.get("code")
    if isinstance(code, bool) or not isinstance(code, int):
        raise LingxingApiError(f"{endpoint_label}业务响应 code 类型无效。")
    if code != 0:
        raise LingxingApiError(f"{endpoint_label}业务查询失败（业务 code {code}）。")
    if not isinstance(payload.get(message_field), str):
        raise LingxingApiError(f"{endpoint_label}响应字段 {message_field} 类型无效。")
    return payload


def _documented_sid(value: Any, field_name: str) -> int:
    # The Product Performance field table says integer, but its published JSON
    # success example encodes nested SID values as decimal strings.
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if isinstance(value, str) and value.isdecimal():
        return int(value)
    raise LingxingApiError(f"商品映射响应字段 {field_name} 的类型无效。")


def _required_text(value: Any, field_name: str, endpoint_label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 类型或内容无效。")
    return value.strip()


def _decimal_number(value: Any, field_name: str, endpoint_label: str) -> Decimal:
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 类型无效；要求 number。")
    try:
        result = Decimal(str(value))
    except InvalidOperation as exc:
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 数值无效。") from exc
    if not result.is_finite():
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 数值无效。")
    return result


def _report_identifier(value: Any, field_name: str, endpoint_label: str) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 类型无效；要求 number。")
    try:
        numeric = Decimal(str(value))
    except InvalidOperation as exc:
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 无效。") from exc
    if not numeric.is_finite() or numeric <= 0 or numeric != numeric.to_integral_value():
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 无效。")
    return format(numeric.quantize(Decimal("1")), "f")


def _report_page_records(payload: Any, endpoint_label: str) -> tuple[int, list[dict[str, Any]]]:
    response = _business_code(payload, endpoint_label, message_field="message")
    if not isinstance(response.get("error_details"), list):
        raise LingxingApiError(f"{endpoint_label}响应字段 error_details 类型无效。")
    if not isinstance(response.get("request_id"), str) or not isinstance(response.get("response_time"), str):
        raise LingxingApiError(f"{endpoint_label}响应缺少 request_id/response_time。")
    total = response.get("total")
    if isinstance(total, bool) or not isinstance(total, int) or total < 0:
        raise LingxingApiError(f"{endpoint_label}响应字段 total 类型无效。")
    records = response.get("data")
    if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
        raise LingxingApiError(f"{endpoint_label}响应字段 data 必须是对象数组。")
    return total, records


def _report_date(value: Any, expected_date: date, field_name: str, endpoint_label: str) -> date:
    if not isinstance(value, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 必须是 YYYY-MM-DD。")
    try:
        parsed = date.fromisoformat(value)
    except ValueError as exc:
        raise LingxingApiError(f"{endpoint_label}响应字段 {field_name} 不是有效日期。") from exc
    if parsed != expected_date:
        raise LingxingApiError(f"{endpoint_label}返回日期与请求日期不一致。")
    return parsed


def parse_sp_campaign_report_page(
    payload: Any,
    *,
    report_date: date,
) -> tuple[int, list[SpCampaignSpend]]:
    return _parse_campaign_report_page(payload, report_date=report_date, endpoint="SP 广告活动报表")


def parse_sb_campaign_report_page(
    payload: Any,
    *,
    report_date: date,
) -> tuple[int, list[SpCampaignSpend]]:
    return _parse_campaign_report_page(payload, report_date=report_date, endpoint="HSA 品牌广告活动报表")


def parse_sd_campaign_report_page(
    payload: Any,
    *,
    report_date: date,
) -> tuple[int, list[SpCampaignSpend]]:
    return _parse_campaign_report_page(payload, report_date=report_date, endpoint="SD 广告活动报表")


def _parse_campaign_report_page(
    payload: Any,
    *,
    report_date: date,
    endpoint: str,
) -> tuple[int, list[SpCampaignSpend]]:
    total, records = _report_page_records(payload, endpoint)
    parsed: list[SpCampaignSpend] = []
    for item in records:
        campaign_id = _report_identifier(item.get("campaign_id"), "campaign_id", endpoint)
        profile_id = _report_identifier(item.get("profile_id"), "profile_id", endpoint)
        row_date = _report_date(item.get("report_date"), report_date, "report_date", endpoint)
        cost = _decimal_number(item.get("cost"), "cost", endpoint)
        parsed.append(SpCampaignSpend(row_date, campaign_id, cost, profile_id))
    return total, parsed


def parse_sp_product_ad_report_page(
    payload: Any,
    *,
    report_date: date,
) -> tuple[int, list[SpProductAdSpend]]:
    return _parse_product_ad_report_page(payload, report_date=report_date, endpoint="SP 广告商品报表")


def parse_sd_product_ad_report_page(
    payload: Any,
    *,
    report_date: date,
) -> tuple[int, list[SpProductAdSpend]]:
    return _parse_product_ad_report_page(payload, report_date=report_date, endpoint="SD 广告商品报表")


def _catalog_page_envelope(payload: Any, endpoint: str) -> tuple[int, list[dict[str, Any]], str]:
    total, records = _report_page_records(payload, endpoint)
    next_token = payload.get("next_token")
    if not isinstance(next_token, str):
        raise LingxingApiError(f"{endpoint}响应字段 next_token 类型无效。")
    return total, records, next_token.strip()


def _catalog_timestamp(value: Any, field_name: str, endpoint: str) -> Decimal:
    # The docs call these numbers but give inconsistent epoch-like examples.
    # Keep the raw numeric value; do not infer seconds, milliseconds or dates.
    return _decimal_number(value, field_name, endpoint)


def parse_sb_product_ads_page(payload: Any) -> tuple[int, list[AdProductCatalogRecord], str]:
    """Parse the current HSA creative catalog without treating it as history."""
    endpoint = "HSA 广告创意目录"
    total, records, next_token = _catalog_page_envelope(payload, endpoint)
    parsed: list[AdProductCatalogRecord] = []
    for item in records:
        profile_id = _report_identifier(item.get("profile_id"), "profile_id", endpoint)
        campaign_id = _report_identifier(item.get("campaign_id"), "campaign_id", endpoint)
        ad_group_id = _report_identifier(item.get("ad_group_id"), "ad_group_id", endpoint)
        ad_creative_id = _report_identifier(item.get("ad_creative_id"), "ad_creative_id", endpoint)
        raw_asins = item.get("asin")
        if not isinstance(raw_asins, list) or any(not isinstance(asin, str) or not asin.strip() for asin in raw_asins):
            raise LingxingApiError(f"{endpoint}响应字段 asin 必须是非空字符串组成的数组。")
        # video_creatives is an explicit documented field, so it can provide a
        # format candidate without relying on a campaign name. It is still a
        # current catalog snapshot and does not prove historical HSA ownership.
        raw_video_creatives = item.get("video_creatives", [])
        if not isinstance(raw_video_creatives, list) or any(not isinstance(entry, dict) for entry in raw_video_creatives):
            raise LingxingApiError(f"{endpoint}响应字段 video_creatives 必须是对象数组。")
        video_creative_types: list[str] = []
        for entry in raw_video_creatives:
            creative_type = entry.get("creative_type")
            if creative_type is not None and not isinstance(creative_type, str):
                raise LingxingApiError(f"{endpoint}响应字段 video_creatives[].creative_type 类型无效。")
            if isinstance(creative_type, str) and creative_type.strip():
                video_creative_types.append(creative_type.strip())
        parsed.append(AdProductCatalogRecord(
            ad_type="HSA",
            profile_id=profile_id,
            campaign_id=campaign_id,
            ad_group_id=ad_group_id,
            ad_creative_id=ad_creative_id,
            asins=tuple(asin.strip() for asin in raw_asins),
            state=_required_text(item.get("state"), "state", endpoint),
            creative_type=_required_text(item.get("creative_type"), "creative_type", endpoint),
            serving_status=_required_text(item.get("serving_status"), "serving_status", endpoint),
            creation_date_value=_catalog_timestamp(item.get("creation_date"), "creation_date", endpoint),
            last_updated_date_value=_catalog_timestamp(item.get("last_updated_date"), "last_updated_date", endpoint),
            video_creative_types=tuple(video_creative_types),
        ))
    return total, parsed, next_token


def parse_sd_product_ads_page(payload: Any) -> tuple[int, list[AdProductCatalogRecord], str]:
    """Parse the current SD product-ad catalog as a cross-check only."""
    endpoint = "SD 广告商品目录"
    total, records, next_token = _catalog_page_envelope(payload, endpoint)
    parsed: list[AdProductCatalogRecord] = []
    for item in records:
        parsed.append(AdProductCatalogRecord(
            ad_type="SD",
            profile_id=_report_identifier(item.get("profile_id"), "profile_id", endpoint),
            campaign_id=_report_identifier(item.get("campaign_id"), "campaign_id", endpoint),
            ad_group_id=_report_identifier(item.get("ad_group_id"), "ad_group_id", endpoint),
            ad_id=_report_identifier(item.get("ad_id"), "ad_id", endpoint),
            seller_sku=_required_text(item.get("sku"), "sku", endpoint),
            asins=(_required_text(item.get("asin"), "asin", endpoint),),
            state=_required_text(item.get("state"), "state", endpoint),
            serving_status=_required_text(item.get("serving_status"), "serving_status", endpoint),
            creation_date_value=_catalog_timestamp(item.get("creation_date"), "creation_date", endpoint),
            last_updated_date_value=_catalog_timestamp(item.get("last_updated_date"), "last_updated_date", endpoint),
        ))
    return total, parsed, next_token


def _parse_product_ad_report_page(
    payload: Any,
    *,
    report_date: date,
    endpoint: str,
) -> tuple[int, list[SpProductAdSpend]]:
    total, records = _report_page_records(payload, endpoint)
    parsed: list[SpProductAdSpend] = []
    for item in records:
        sku = _required_text(item.get("sku"), "sku", endpoint)
        asin = _required_text(item.get("asin"), "asin", endpoint)
        ad_id = _report_identifier(item.get("ad_id"), "ad_id", endpoint)
        ad_group_id = _report_identifier(item.get("ad_group_id"), "ad_group_id", endpoint)
        profile_id = _report_identifier(item.get("profile_id"), "profile_id", endpoint)
        campaign_id = _report_identifier(item.get("campaign_id"), "campaign_id", endpoint)
        row_date = _report_date(item.get("report_date"), report_date, "report_date", endpoint)
        cost = _decimal_number(item.get("cost"), "cost", endpoint)
        parsed.append(SpProductAdSpend(row_date, sku, asin, ad_id, ad_group_id, campaign_id, cost, profile_id))
    return total, parsed


def parse_sb_ad_group_report_page(
    payload: Any,
    *,
    report_date: date,
) -> tuple[int, list[SbAdGroupSpend]]:
    endpoint = "HSA 品牌广告组报表"
    total, records = _report_page_records(payload, endpoint)
    parsed: list[SbAdGroupSpend] = []
    for item in records:
        campaign_id = _report_identifier(item.get("campaign_id"), "campaign_id", endpoint)
        ad_group_id = _report_identifier(item.get("ad_group_id"), "ad_group_id", endpoint)
        profile_id = _report_identifier(item.get("profile_id"), "profile_id", endpoint)
        row_date = _report_date(item.get("report_date"), report_date, "report_date", endpoint)
        cost = _decimal_number(item.get("cost"), "cost", endpoint)
        parsed.append(SbAdGroupSpend(row_date, campaign_id, ad_group_id, cost, profile_id))
    return total, parsed


def parse_product_performance_page(
    payload: Any,
    *,
    store_sid: int,
    requested_skus: set[str],
) -> tuple[int, list[SkuProductMapping]]:
    """Parse one documented ``asinList`` page without leaking raw records."""
    response = _business_code(payload, "产品表现", message_field="msg")
    if not isinstance(response.get("trace_id"), str):
        raise LingxingApiError("产品表现响应字段 trace_id 类型无效。")
    data = response.get("data")
    if not isinstance(data, dict):
        raise LingxingApiError("产品表现响应字段 data 必须是对象。")
    total = data.get("total")
    if isinstance(total, bool) or not isinstance(total, int) or total < 0:
        raise LingxingApiError("产品表现响应字段 total 类型无效。")
    records = data.get("list")
    if not isinstance(records, list):
        raise LingxingApiError("产品表现响应字段 data.list 必须是数组。")

    mappings: list[SkuProductMapping] = []
    for item in records:
        if not isinstance(item, dict):
            raise LingxingApiError("产品表现 data.list 中的记录必须是对象。")
        local_sku = _required_text(item.get("sku"), "sku", "产品表现")
        if local_sku not in requested_skus:
            continue
        price_list = item.get("price_list")
        asin_records = item.get("asins")
        if not isinstance(price_list, list) or not isinstance(asin_records, list):
            raise LingxingApiError("产品表现记录的 price_list/asins 必须是数组。")

        seller_skus: set[str] = set()
        for price in price_list:
            if not isinstance(price, dict):
                raise LingxingApiError("产品表现 price_list 记录必须是对象。")
            row_local_sku = _required_text(price.get("local_sku"), "price_list.local_sku", "产品表现")
            seller_sku = _required_text(price.get("seller_sku"), "price_list.seller_sku", "产品表现")
            row_sid = _documented_sid(price.get("sid"), "price_list.sid")
            if row_sid == store_sid and row_local_sku == local_sku:
                seller_skus.add(seller_sku)

        asins: set[str] = set()
        for asin_record in asin_records:
            if not isinstance(asin_record, dict):
                raise LingxingApiError("产品表现 asins 记录必须是对象。")
            asin = _required_text(asin_record.get("asin"), "asins.asin", "产品表现")
            row_sid = _documented_sid(asin_record.get("sid"), "asins.sid")
            if row_sid == store_sid:
                asins.add(asin)

        mappings.append(
            SkuProductMapping(
                local_sku=local_sku,
                seller_skus=tuple(sorted(seller_skus)),
                asins=tuple(sorted(asins)),
            )
        )
    return total, mappings


def _performance_number(value: Any, field_name: str, *, integer: bool = False) -> Decimal | int | None:
    """Accept JSON numbers and numeric strings documented by SDK examples.

    The published schema declares many Product Performance fields as numbers,
    while its success example serializes them as strings. This parser accepts
    only finite decimal representations and keeps absent/null values distinct
    from explicit zero.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        raise LingxingApiError(f"产品表现响应字段 {field_name} 数值类型无效。")
    if isinstance(value, str):
        value = value.strip()
        if not value or not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", value):
            raise LingxingApiError(f"产品表现响应字段 {field_name} 数值格式无效。")
    elif not isinstance(value, (int, float, Decimal)):
        raise LingxingApiError(f"产品表现响应字段 {field_name} 数值类型无效。")
    try:
        parsed = Decimal(str(value))
    except InvalidOperation as exc:
        raise LingxingApiError(f"产品表现响应字段 {field_name} 数值无效。") from exc
    if not parsed.is_finite():
        raise LingxingApiError(f"产品表现响应字段 {field_name} 数值无效。")
    if integer:
        if parsed != parsed.to_integral_value():
            raise LingxingApiError(f"产品表现响应字段 {field_name} 应为整数。")
        return int(parsed)
    return parsed


def parse_product_performance_metrics_page(
    payload: Any,
    *,
    store_sid: int,
    requested_skus: set[str],
    start_date: date,
    end_date: date,
    summary_field: str = "sku",
) -> tuple[int, list[ProductPerformanceMetrics]]:
    """Parse one official AsinListNew summary dimension without asserting daily scope."""
    if summary_field not in {"sku", "msku"}:
        raise LingxingApiError("产品表现预览仅支持已记录的 sku 或 msku 汇总维度。")
    endpoint = "产品表现新版"
    response = _business_code(payload, endpoint, message_field="msg")
    if not isinstance(response.get("trace_id"), str):
        raise LingxingApiError(f"{endpoint}响应字段 trace_id 类型无效。")
    data = response.get("data")
    if not isinstance(data, dict):
        raise LingxingApiError(f"{endpoint}响应字段 data 必须是对象。")
    total = data.get("total")
    if isinstance(total, bool) or not isinstance(total, int) or total < 0:
        raise LingxingApiError(f"{endpoint}响应字段 total 类型无效。")
    records = data.get("list")
    if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
        raise LingxingApiError(f"{endpoint}响应字段 data.list 必须是对象数组。")

    parsed_rows: list[ProductPerformanceMetrics] = []
    seen_summary_keys: set[tuple[str, str]] = set()
    for item in records:
        raw_sids = item.get("sids")
        if not isinstance(raw_sids, list) or not raw_sids:
            raise LingxingApiError(f"{endpoint}响应字段 sids 必须是非空数组。")
        row_sids = {_documented_sid(value, "sids") for value in raw_sids}
        if row_sids != {store_sid}:
            raise LingxingApiError(f"{endpoint}汇总行的店铺范围与所选 SID 不一致或不唯一。")

        raw_price_list = item.get("price_list")
        if not isinstance(raw_price_list, list):
            raise LingxingApiError(f"{endpoint}响应字段 price_list 必须是数组。")
        seller_skus: set[str] = set()
        matched_local_skus: set[str] = set()
        for price in raw_price_list:
            if not isinstance(price, dict):
                raise LingxingApiError(f"{endpoint}响应字段 price_list 必须包含对象。")
            row_local_sku = _required_text(price.get("local_sku"), "price_list.local_sku", endpoint)
            row_sid = _documented_sid(price.get("sid"), "price_list.sid")
            if row_sid == store_sid and row_local_sku in requested_skus:
                matched_local_skus.add(row_local_sku)
                seller_skus.add(_required_text(price.get("seller_sku"), "price_list.seller_sku", endpoint))
        if not seller_skus:
            raise LingxingApiError(f"{endpoint}没有可核验的本店 MSKU 与 local SKU 对应关系。")

        if summary_field == "sku":
            local_sku = _required_text(item.get("sku"), "sku", endpoint)
            if local_sku not in requested_skus:
                raise LingxingApiError(f"{endpoint}返回了本次未请求的 local SKU；已停止归属。")
            if matched_local_skus != {local_sku}:
                raise LingxingApiError(f"{endpoint} SKU 汇总行与 price_list 的本店 SKU 映射不一致。")
            if local_sku in {key[0] for key in seen_summary_keys}:
                raise LingxingApiError(f"{endpoint}在 SKU 汇总维度返回重复 SKU {local_sku}；已停止归属。")
            summary_key = local_sku
        else:
            if len(matched_local_skus) != 1:
                raise LingxingApiError(
                    f"{endpoint} MSKU 汇总行无法唯一匹配一个本店、已请求的 local SKU；已停止归属。"
                )
            if len(seller_skus) != 1:
                raise LingxingApiError(f"{endpoint} MSKU 汇总行未唯一标识 seller_sku；已停止归属。")
            local_sku = next(iter(matched_local_skus))
            summary_key = next(iter(seller_skus))
        identity = (local_sku, summary_key)
        if identity in seen_summary_keys:
            raise LingxingApiError(f"{endpoint}返回重复的 {summary_field} 汇总行；已停止归属。")
        seen_summary_keys.add(identity)

        raw_asins = item.get("asins")
        if not isinstance(raw_asins, list):
            raise LingxingApiError(f"{endpoint}响应字段 asins 必须是数组。")
        asins: set[str] = set()
        for asin_record in raw_asins:
            if not isinstance(asin_record, dict):
                raise LingxingApiError(f"{endpoint}响应字段 asins 必须包含对象。")
            row_sid = _documented_sid(asin_record.get("sid"), "asins.sid")
            if row_sid == store_sid:
                asins.add(_required_text(asin_record.get("asin"), "asins.asin", endpoint))

        currency_code = item.get("currency_code")
        if currency_code is not None and not isinstance(currency_code, str):
            raise LingxingApiError(f"{endpoint}响应字段 currency_code 类型无效。")

        values: dict[str, Any] = {}
        present_fields = frozenset(item)
        for field_name in _PRODUCT_PERFORMANCE_INTEGER_FIELDS:
            values[field_name] = _performance_number(
                item.get(field_name), field_name, integer=True
            )
        for field_name in _PRODUCT_PERFORMANCE_NUMBER_FIELDS:
            values[field_name] = _performance_number(item.get(field_name), field_name)
        for field_name in ("rank_category", "ranking_update_time"):
            value = item.get(field_name)
            if value is not None and not isinstance(value, str):
                raise LingxingApiError(f"{endpoint}响应字段 {field_name} 类型无效。")
            values[field_name] = value

        raw_ranks = item.get("small_cate_rank")
        if raw_ranks is None:
            values["small_cate_rank"] = None
        elif not isinstance(raw_ranks, list):
            raise LingxingApiError(f"{endpoint}响应字段 small_cate_rank 应为数组或 null。")
        else:
            ranks: list[dict[str, Any]] = []
            for rank_row in raw_ranks:
                if not isinstance(rank_row, dict):
                    raise LingxingApiError(f"{endpoint}响应字段 small_cate_rank 必须包含对象。")
                category = rank_row.get("category")
                if category is not None and not isinstance(category, str):
                    raise LingxingApiError(f"{endpoint}响应字段 small_cate_rank.category 类型无效。")
                ranks.append({
                    "category": category,
                    "rank": _performance_number(rank_row.get("rank"), "small_cate_rank.rank", integer=True),
                    "prev_rank": _performance_number(rank_row.get("prev_rank"), "small_cate_rank.prev_rank", integer=True),
                })
            values["small_cate_rank"] = tuple(ranks)
        values["currency_code"] = currency_code
        parsed_rows.append(ProductPerformanceMetrics(
            store_sid=store_sid,
            local_sku=local_sku,
            seller_skus=tuple(sorted(seller_skus)),
            asins=tuple(sorted(asins)),
            start_date=start_date,
            end_date=end_date,
            currency_code=currency_code,
            values=values,
            present_fields=present_fields,
            summary_field=summary_field,
            summary_key=summary_key,
        ))
    return total, parsed_rows


def parse_asin_daily_page(
    payload: Any,
    *,
    store_sid: int,
    event_date: date,
    metric_type: int,
) -> list[_DailyMetricRecord]:
    """Parse one daily statistics page for the documented metric selector."""
    if isinstance(metric_type, bool) or not isinstance(metric_type, int) or metric_type not in (1, 2, 3):
        raise LingxingApiError("每日销量统计 type 必须是文档规定的 1、2 或 3。")
    response = _business_code(payload, "每日销量统计", message_field="message")
    if not isinstance(response.get("error_details"), list):
        raise LingxingApiError("每日销量统计响应字段 error_details 类型无效。")
    if not isinstance(response.get("request_id"), str):
        raise LingxingApiError("每日销量统计响应字段 request_id 类型无效。")
    if not isinstance(response.get("response_time"), str):
        raise LingxingApiError("每日销量统计响应字段 response_time 类型无效。")
    records = response.get("data")
    if not isinstance(records, list):
        raise LingxingApiError("每日销量统计响应字段 data 必须是数组。")

    result: list[_DailyMetricRecord] = []
    for item in records:
        if not isinstance(item, dict):
            raise LingxingApiError("每日销量统计 data 中的记录必须是对象。")
        sid = item.get("sid")
        if isinstance(sid, bool) or not isinstance(sid, int):
            raise LingxingApiError("每日销量统计响应字段 sid 类型无效。")
        raw_date = _required_text(item.get("r_date"), "r_date", "每日销量统计")
        try:
            parsed_date = date.fromisoformat(raw_date)
        except ValueError as exc:
            raise LingxingApiError("每日销量统计响应字段 r_date 不是 YYYY-MM-DD。") from exc
        if parsed_date != event_date:
            raise LingxingApiError("每日销量统计响应日期与请求的站点日期不一致。")
        currency_code = _required_text(item.get("currency_code"), "currency_code", "每日销量统计")
        seller_sku = _required_text(item.get("seller_sku"), "seller_sku", "每日销量统计")
        asin = _required_text(item.get("asin"), "asin", "每日销量统计")
        product_name = item.get("product_name")
        if not isinstance(product_name, str):
            raise LingxingApiError("每日销量统计响应字段 product_name 类型无效。")
        metric_value = _decimal_number(item.get("map_value"), "map_value", "每日销量统计")
        if metric_type in (2, 3) and metric_value != metric_value.to_integral_value():
            raise LingxingApiError("每日销量统计销量/订单量不是整数，已停止映射。")
        if sid != store_sid:
            raise LingxingApiError("每日销量统计返回了与所选店铺 SID 不匹配的记录。")
        result.append(
            _DailyMetricRecord(
                sid=sid,
                event_date=parsed_date,
                currency_code=currency_code,
                seller_sku=seller_sku,
                asin=asin,
                product_name=product_name,
                metric_value=metric_value,
            )
        )
    return result


@dataclass(frozen=True)
class LingxingSeller:
    """A store record parsed from the documented seller-list response."""

    sid: int = field(repr=False)
    mid: int = field(repr=False)
    name: str = field(repr=False)
    seller_id: str = field(repr=False)
    account_name: str = field(repr=False)
    seller_account_id: int = field(repr=False)
    region: str = field(repr=False)
    country: str = field(repr=False)
    marketplace_id: str = field(repr=False)
    status: int = field(repr=False)
    has_ads_setting: int = field(repr=False)

    @property
    def site_label(self) -> str:
        return " / ".join(value for value in (self.country, self.region, self.marketplace_id) if value) or "未知"

    @property
    def display_label(self) -> str:
        return f"{self.name} — {self.site_label} — SID {self.sid}"


_SELLER_INTEGER_FIELDS = ("sid", "mid", "seller_account_id", "status", "has_ads_setting")
_SELLER_STRING_FIELDS = ("name", "seller_id", "account_name", "region", "country", "marketplace_id")


def _seller_integer(record: dict[str, Any], key: str) -> int:
    value = record.get(key)
    if isinstance(value, bool) or not isinstance(value, int):
        raise LingxingApiError(f"店铺响应字段 {key} 的类型无效；要求 integer。")
    return value


def _seller_string(record: dict[str, Any], key: str) -> str:
    value = record.get(key)
    if not isinstance(value, str):
        raise LingxingApiError(f"店铺响应字段 {key} 的类型无效；要求 string。")
    if key == "name" and not value.strip():
        raise LingxingApiError("店铺响应字段 name 不能为空。")
    return value.strip()


def parse_seller_records(records: Any) -> list[LingxingSeller]:
    """Map store records without retaining unrecognized or sensitive response fields."""
    if not isinstance(records, list):
        raise LingxingApiError("店铺列表业务响应中的 data 必须是数组。")
    sellers: list[LingxingSeller] = []
    seen_sids: set[int] = set()
    for index, record in enumerate(records):
        if not isinstance(record, dict):
            raise LingxingApiError(f"店铺列表第 {index + 1} 条记录不是对象。")
        # Validate the complete, confirmed SellerLists record contract. Do not
        # coerce strings/numbers: silently changing IDs can bind a query wrongly.
        for key in _SELLER_INTEGER_FIELDS:
            _seller_integer(record, key)
        for key in _SELLER_STRING_FIELDS:
            _seller_string(record, key)
        sid = record["sid"]
        if sid in seen_sids:
            raise LingxingApiError("店铺列表包含重复 SID，已停止映射。")
        seen_sids.add(sid)
        sellers.append(
            LingxingSeller(
                sid=sid,
                mid=record["mid"],
                name=record["name"].strip(),
                seller_id=record["seller_id"].strip(),
                account_name=record["account_name"].strip(),
                seller_account_id=record["seller_account_id"],
                region=record["region"].strip(),
                country=record["country"].strip(),
                marketplace_id=record["marketplace_id"].strip(),
                status=record["status"],
                has_ads_setting=record["has_ads_setting"],
            )
        )
    return sellers


def parse_seller_list_response(payload: Any) -> list[LingxingSeller]:
    """Parse the confirmed SellerLists envelope and its direct ``data`` array."""
    if not isinstance(payload, dict):
        raise LingxingApiError("店铺列表业务响应必须是 JSON 对象。")
    code = payload.get("code")
    if isinstance(code, bool) or not isinstance(code, int):
        raise LingxingApiError("店铺列表响应缺少 integer 类型的业务 code，已拒绝接受响应。")
    if code != 0:
        raise LingxingApiError(f"店铺列表业务查询失败（业务 code {code}）。")

    for key, expected_type in (
        ("message", str),
        ("error_details", list),
        ("request_id", str),
        ("response_time", str),
    ):
        if not isinstance(payload.get(key), expected_type):
            raise LingxingApiError(f"店铺列表响应字段 {key} 的类型无效。")
    if "data" not in payload:
        raise LingxingApiError("店铺列表响应缺少 data 数组，已拒绝接受响应。")
    return parse_seller_records(payload["data"])


@dataclass(frozen=True)
class AccessTokenResult:
    # Keep the token out of accidental repr/log output.
    token: str = field(repr=False)
    expires_at: datetime | None = None
    expires_in_seconds: int | None = None

    def safe_status(self) -> str:
        if self.expires_at is not None:
            expiry = self.expires_at.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
            return f"API 身份验证成功；Access Token 有效期至 {expiry}。Token 已隐藏。"
        if self.expires_in_seconds is not None:
            return f"API 身份验证成功；Token 有效期 {self.expires_in_seconds} 秒。Token 已隐藏。"
        return "API 身份验证成功；响应未提供可识别的有效期。Token 已隐藏。"


def _object_layers(payload: dict[str, Any]) -> list[dict[str, Any]]:
    layers = [payload]
    for field_name in ("data", "result"):
        nested = payload.get(field_name)
        if isinstance(nested, dict):
            layers.append(nested)
    return layers


def _read_expiry(layers: list[dict[str, Any]]) -> tuple[datetime | None, int | None]:
    seconds: int | None = None
    for layer in layers:
        for name in ("expires_in", "expiresIn"):
            value = layer.get(name)
            if isinstance(value, (int, float)) and not isinstance(value, bool) and value >= 0:
                seconds = int(value)
                break
            if isinstance(value, str) and value.strip().isdigit():
                seconds = int(value.strip())
                break
        if seconds is not None:
            break
    if seconds is not None:
        return datetime.now(timezone.utc) + timedelta(seconds=seconds), seconds

    for layer in layers:
        for name in ("expires_at", "expiresAt", "expire_time", "expireTime"):
            value = layer.get(name)
            if isinstance(value, datetime):
                return (value.replace(tzinfo=timezone.utc) if value.tzinfo is None else value), None
            if isinstance(value, str) and value.strip():
                try:
                    parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
                except ValueError:
                    continue
                return (parsed.replace(tzinfo=timezone.utc) if parsed.tzinfo is None else parsed), None
    return None, None


def _safe_error_message(payload: dict[str, Any], secrets: tuple[str, ...]) -> str:
    for layer in _object_layers(payload):
        for name in ("message", "msg", "error_description", "error"):
            value = layer.get(name)
            if isinstance(value, str) and value.strip():
                message = value.strip()
                for secret in secrets:
                    if secret:
                        message = message.replace(secret, "[已隐藏]")
                message = re.sub(r"(?i)(access[_-]?token\s*[:=]\s*)[^,\s]+", r"\1[已隐藏]", message)
                return message[:500]
    return "领星未返回 Access Token；请检查 AppID、AppSecret、账号权限及 IP 白名单。"


def _redact_api_payload(value: Any, secrets: tuple[str, ...]) -> Any:
    if isinstance(value, dict):
        safe: dict[str, Any] = {}
        for key, item in value.items():
            normalized_key = re.sub(r"[^a-z0-9]", "", str(key).casefold())
            # A next_token is a pagination cursor, not an authentication token.
            # The client must retain it to fetch the next read-only page.
            if normalized_key == "nexttoken":
                safe[key] = item
            elif any(marker in normalized_key for marker in ("token", "secret", "password", "sign", "authorization", "appkey", "appid")):
                safe[key] = "[已隐藏]"
            else:
                safe[key] = _redact_api_payload(item, secrets)
        return safe
    if isinstance(value, list):
        return [_redact_api_payload(item, secrets) for item in value]
    if isinstance(value, str):
        safe_value = value
        for secret in secrets:
            if secret:
                safe_value = safe_value.replace(secret, "[已隐藏]")
        safe_value = re.sub(r"(?i)(access[_-]?token\s*[:=]\s*)[^,\s]+", r"\1[已隐藏]", safe_value)
        return safe_value
    return value


class LingxingOpenApiClient:
    """Official-host client for signed read-only seller, sales, and ad queries."""

    def __init__(
        self,
        app_id: str,
        app_secret: str,
        *,
        timeout_seconds: float = 30,
        token_refresh_skew_seconds: float = 60,
        http_client: httpx.Client | None = None,
    ) -> None:
        if not app_id.strip() or not app_secret:
            raise LingxingApiError("请填写 AppID 和 AppSecret。")
        if "\r" in app_id or "\n" in app_id or "\r" in app_secret or "\n" in app_secret:
            raise LingxingApiError("AppID 或 AppSecret 格式无效。")
        self._app_id = app_id.strip()
        self._app_secret = app_secret
        if (
            isinstance(token_refresh_skew_seconds, bool)
            or not isinstance(token_refresh_skew_seconds, (int, float))
            or not math.isfinite(token_refresh_skew_seconds)
            or token_refresh_skew_seconds < 0
        ):
            raise LingxingApiError("Token 提前刷新时间必须是有限的非负数字。")
        self._token_refresh_skew = timedelta(seconds=token_refresh_skew_seconds)
        self._token_lock = threading.Lock()
        self._cached_token: AccessTokenResult | None = None
        self._http = http_client or httpx.Client(timeout=timeout_seconds, follow_redirects=False)
        self._owns_http = http_client is None

    def request_access_token(self) -> AccessTokenResult:
        try:
            response = self._http.post(
                OPENAPI_BASE_URL + ACCESS_TOKEN_PATH,
                files={
                    "appId": (None, self._app_id),
                    "appSecret": (None, self._app_secret),
                },
            )
        except httpx.TimeoutException as exc:
            raise LingxingApiError("连接领星 OpenAPI 超时。") from exc
        except httpx.HTTPError as exc:
            raise LingxingApiError("无法连接领星 OpenAPI；请检查网络、域名及 IP 白名单。") from exc

        if response.status_code >= 400:
            # Never include a response body: the service could echo credentials.
            raise LingxingApiError(f"领星 Token 接口返回 HTTP {response.status_code}。")
        try:
            payload = response.json()
        except (ValueError, UnicodeDecodeError) as exc:
            raise LingxingApiError("领星 Token 接口返回了无法解析的 JSON。") from exc
        if not isinstance(payload, dict):
            raise LingxingApiError("领星 Token 接口响应结构无效。")

        layers = _object_layers(payload)
        token: str | None = None
        for layer in layers:
            candidate = layer.get("access_token")
            if candidate is None:
                candidate = layer.get("accessToken")
            if isinstance(candidate, str) and candidate.strip():
                token = candidate.strip()
                break
        if token is None:
            message = _safe_error_message(payload, (self._app_id, self._app_secret))
            raise LingxingApiError(message)

        expires_at, expires_in = _read_expiry(layers)
        return AccessTokenResult(token, expires_at, expires_in)

    def get_access_token(self, *, force_refresh: bool = False) -> AccessTokenResult:
        """Return a fresh token, reusing it only when the response supplied expiry.

        Tokens without an explicit, parseable expiry are never cached because
        their lifetime cannot safely be inferred from the local client.
        """
        with self._token_lock:
            now = datetime.now(timezone.utc)
            cached = self._cached_token
            if (
                not force_refresh
                and cached is not None
                and cached.expires_at is not None
                and cached.expires_at - now > self._token_refresh_skew
            ):
                return cached

            self._cached_token = None
            refreshed = self.request_access_token()
            self._cached_token = (
                refreshed
                if refreshed.expires_at is not None and refreshed.expires_at > now
                else None
            )
            return refreshed

    def clear_token_cache(self) -> None:
        """Forget the in-memory token after authentication changes or failures."""
        with self._token_lock:
            self._cached_token = None

    def _request_signed_get(
        self,
        path: str,
        parameters: dict[str, str],
        *,
        sensitive_values: tuple[str, ...],
    ) -> Any:
        query = encode_query_parameters(parameters)
        # The query was already serialized with java.net.URLEncoder-compatible
        # escaping. httpx.URL preserves existing percent escapes when parsed.
        try:
            url = httpx.URL(f"{OPENAPI_BASE_URL}{path}?{query}")
            response = self._http.get(url)
        except httpx.InvalidURL as exc:
            raise LingxingApiError("领星 OpenAPI 查询请求格式无效。") from exc
        except httpx.TimeoutException as exc:
            raise LingxingApiError("领星 OpenAPI 查询超时。") from exc
        except httpx.HTTPError as exc:
            raise LingxingApiError("无法连接领星 OpenAPI；请检查网络、域名及 IP 白名单。") from exc

        if not 200 <= response.status_code < 300:
            # Do not include the URL or response body; either can echo credentials/signature.
            raise LingxingApiError(f"领星 OpenAPI 查询返回 HTTP {response.status_code}。")
        try:
            payload = response.json()
        except (ValueError, UnicodeDecodeError) as exc:
            raise LingxingApiError("领星 OpenAPI 查询返回了无法解析的 JSON。") from exc
        if not isinstance(payload, (dict, list)):
            raise LingxingApiError("领星 OpenAPI 查询响应结构无效。")
        return _redact_api_payload(
            payload,
            (self._app_id, self._app_secret, *sensitive_values),
        )

    def _request_signed_post(
        self,
        path: str,
        *,
        access_token: str,
        body: dict[str, Any],
        headers: dict[str, str] | None = None,
    ) -> Any:
        """Send an SDK-style signed JSON POST; never expose URL/body on errors."""
        auth_parameters = {
            "access_token": access_token,
            "app_key": self._app_id,
            "timestamp": str(int(time.time())),
        }
        signed_query = signed_post_query_parameters(auth_parameters, body, self._app_id)
        query = encode_query_parameters(signed_query)
        try:
            url = httpx.URL(f"{OPENAPI_BASE_URL}{path}?{query}")
            response = self._http.post(url, json=body, headers=headers)
        except httpx.InvalidURL as exc:
            raise LingxingApiError("领星 OpenAPI 查询请求格式无效。") from exc
        except httpx.TimeoutException as exc:
            raise LingxingApiError("领星 OpenAPI 查询超时。") from exc
        except httpx.HTTPError as exc:
            raise LingxingApiError("无法连接领星 OpenAPI；请检查网络、域名及 IP 白名单。") from exc

        if not 200 <= response.status_code < 300:
            if response.status_code == 429:
                retry_after = response.headers.get("Retry-After", "")
                if retry_after.isdecimal():
                    raise LingxingApiError(
                        f"领星 OpenAPI 限流（HTTP 429）；服务端建议等待 {retry_after} 秒后再试。"
                    )
                raise LingxingApiError("领星 OpenAPI 限流（HTTP 429）；未自动重发，请稍后重试。")
            raise LingxingApiError(f"领星 OpenAPI 查询返回 HTTP {response.status_code}。")
        try:
            payload = json.loads(response.content, parse_float=Decimal)
        except (ValueError, UnicodeDecodeError) as exc:
            raise LingxingApiError("领星 OpenAPI 查询返回了无法解析的 JSON。") from exc
        if not isinstance(payload, (dict, list)):
            raise LingxingApiError("领星 OpenAPI 查询响应结构无效。")
        return _redact_api_payload(
            payload,
            (self._app_id, self._app_secret, access_token, signed_query["sign"]),
        )

    def _request_product_performance_page(
        self,
        *,
        access_token: str,
        body: dict[str, Any],
    ) -> Any:
        """Honor Product Performance's documented single-store 1 s interval."""
        global _PRODUCT_PERFORMANCE_LAST_RESPONSE_AT
        with _PRODUCT_PERFORMANCE_REQUEST_LOCK:
            if _PRODUCT_PERFORMANCE_LAST_RESPONSE_AT is not None:
                remaining = 1.0 - (time.monotonic() - _PRODUCT_PERFORMANCE_LAST_RESPONSE_AT)
                if remaining > 0:
                    time.sleep(remaining)
            try:
                return self._request_signed_post(
                    PRODUCT_PERFORMANCE_PATH,
                    access_token=access_token,
                    body=body,
                )
            finally:
                # Count failed responses too so a subsequent retry cannot burst.
                _PRODUCT_PERFORMANCE_LAST_RESPONSE_AT = time.monotonic()

    @staticmethod
    def _validate_query_range(start_date: date, end_date: date) -> None:
        if type(start_date) is not date or type(end_date) is not date:
            raise LingxingApiError("查询日期必须使用 YYYY-MM-DD 站点日期。")
        if end_date < start_date:
            raise LingxingApiError("结束日期不能早于开始日期。")
        if (end_date - start_date).days + 1 > MAX_PRODUCT_QUERY_DAYS:
            raise LingxingApiError("本次查询范围不能超过 92 个自然日（含首尾日期）。")

    def query_product_performance_single_day(
        self,
        *,
        access_token: str,
        store_sid: int,
        local_skus: list[str],
        event_date: date,
        summary_field: str = "sku",
    ) -> list[ProductPerformanceMetrics]:
        """Read one supported summary dimension for one site date; never writes report data.

        This deliberately accepts a single date only. The endpoint returns no
        row-level date, so a multi-day aggregate must not be mistaken for daily
        values. Historical rating and ranking scope still requires external
        verification even for a one-day request.
        """
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError("Access Token 为空，未发送产品表现指标查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
            raise LingxingApiError("产品表现指标查询必须使用动态店铺 SID。")
        if type(event_date) is not date:
            raise LingxingApiError("产品表现指标查询日期必须是有效的站点日期。")
        if summary_field not in {"sku", "msku"}:
            raise LingxingApiError("产品表现指标查询汇总维度仅支持 sku 或 msku。")
        if not isinstance(local_skus, list) or not 1 <= len(local_skus) <= 50:
            raise LingxingApiError("产品表现指标查询需要 1–50 个 local SKU。")
        normalized_skus: list[str] = []
        for sku in local_skus:
            if not isinstance(sku, str) or not sku.strip():
                raise LingxingApiError("产品表现指标查询包含空白或无效 SKU。")
            value = sku.strip()
            if value not in normalized_skus:
                normalized_skus.append(value)
        requested_skus = set(normalized_skus)

        offset = 0
        expected_total: int | None = None
        seen_summary_keys: set[tuple[str, str]] = set()
        rows_by_key: dict[tuple[str, str], ProductPerformanceMetrics] = {}
        while True:
            body: dict[str, Any] = {
                "offset": offset,
                "length": PRODUCT_PERFORMANCE_PAGE_LENGTH,
                "sort_field": "volume",
                "sort_type": "desc",
                "search_field": "local_sku",
                "search_value": normalized_skus,
                # Official docs specify a string SID for a single-store query.
                "sid": str(store_sid),
                "start_date": event_date.isoformat(),
                "end_date": event_date.isoformat(),
                "summary_field": summary_field,
                "is_recently_enum": False,
                # Explicitly retain the documented default refund-date rule.
                "purchase_status": 0,
            }
            payload = self._request_product_performance_page(
                access_token=access_token.strip(),
                body=body,
            )
            total, page_rows = parse_product_performance_metrics_page(
                payload,
                store_sid=store_sid,
                requested_skus=requested_skus,
                start_date=event_date,
                end_date=event_date,
                summary_field=summary_field,
            )
            raw_records = payload["data"]["list"]
            if expected_total is None:
                expected_total = total
            elif total != expected_total:
                raise LingxingApiError("产品表现指标分页期间 total 发生变化，已停止查询。")
            if len(raw_records) > PRODUCT_PERFORMANCE_PAGE_LENGTH or offset + len(raw_records) > total:
                raise LingxingApiError("产品表现指标分页返回数量无效。")
            if not raw_records and offset < total:
                raise LingxingApiError("产品表现指标分页提前返回空页，数据可能不完整。")
            for row in page_rows:
                identity = (row.local_sku, row.summary_key or "")
                if identity in seen_summary_keys:
                    raise LingxingApiError(f"产品表现指标分页返回重复 {summary_field} 汇总行；已停止归属。")
                seen_summary_keys.add(identity)
                rows_by_key[identity] = row
            if offset + len(raw_records) >= total:
                break
            offset += len(raw_records)

        return sorted(
            rows_by_key.values(),
            key=lambda row: (normalized_skus.index(row.local_sku), row.summary_key or ""),
        )

    def query_product_performance_cost_diagnostic(
        self,
        *,
        access_token: str,
        store_sid: int,
        local_sku: str,
        start_date: date,
        end_date: date,
    ) -> list[ProductPerformanceMetrics]:
        """Compare three per-day SKU summaries with the middle-day MSKU summary.

        This is a deliberately small, read-only diagnostic: it makes exactly
        four paginated query sets, does not aggregate unlike dimensions, and
        does not write or log returned business metrics.
        """
        if not isinstance(local_sku, str) or not local_sku.strip():
            raise LingxingApiError("产品表现诊断需要一个 local SKU。")
        if type(start_date) is not date or type(end_date) is not date:
            raise LingxingApiError("产品表现诊断日期必须使用 YYYY-MM-DD 站点日期。")
        if (end_date - start_date).days != 2:
            raise LingxingApiError("产品表现日期诊断必须覆盖连续三天（开始日、次日、第三日）。")

        requested_sku = local_sku.strip()
        days = (start_date, start_date + timedelta(days=1), end_date)
        rows: list[ProductPerformanceMetrics] = []
        for event_date in days:
            rows.extend(self.query_product_performance_single_day(
                access_token=access_token,
                store_sid=store_sid,
                local_skus=[requested_sku],
                event_date=event_date,
                summary_field="sku",
            ))
        rows.extend(self.query_product_performance_single_day(
            access_token=access_token,
            store_sid=store_sid,
            local_skus=[requested_sku],
            event_date=days[1],
            summary_field="msku",
        ))
        return rows

    def query_product_mappings(
        self,
        *,
        access_token: str,
        store_sid: int,
        local_skus: list[str],
        start_date: date,
        end_date: date,
    ) -> list[SkuProductMapping]:
        """Find local SKU -> MSKU/ASIN mappings using the documented API."""
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError("Access Token 为空，未发送产品表现查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int):
            raise LingxingApiError("产品表现查询必须使用动态店铺 SID。")
        self._validate_query_range(start_date, end_date)
        if not isinstance(local_skus, list) or not 1 <= len(local_skus) <= 50:
            raise LingxingApiError("产品表现查询需要 1–50 个 local SKU。")
        normalized_skus: list[str] = []
        for sku in local_skus:
            if not isinstance(sku, str) or not sku.strip():
                raise LingxingApiError("产品表现查询包含空白或无效 SKU。")
            value = sku.strip()
            if value not in normalized_skus:
                normalized_skus.append(value)
        requested_skus = set(normalized_skus)

        offset = 0
        expected_total: int | None = None
        seen_local_skus: set[str] = set()
        mappings: dict[str, SkuProductMapping] = {}
        while True:
            body: dict[str, Any] = {
                "offset": offset,
                "length": PRODUCT_PERFORMANCE_PAGE_LENGTH,
                "sort_field": "volume",
                "sort_type": "desc",
                "search_field": "local_sku",
                "search_value": normalized_skus,
                # SellerLists returns SID as integer; the Product Performance
                # docs specify a string for a single store.
                "sid": str(store_sid),
                "start_date": start_date.isoformat(),
                "end_date": end_date.isoformat(),
                "summary_field": "sku",
                "is_recently_enum": False,
            }
            payload = self._request_product_performance_page(
                access_token=access_token.strip(),
                body=body,
            )
            total, page_mappings = parse_product_performance_page(
                payload,
                store_sid=store_sid,
                requested_skus=requested_skus,
            )
            raw_records = payload["data"]["list"]
            if expected_total is None:
                expected_total = total
            elif total != expected_total:
                raise LingxingApiError("产品表现分页期间 total 发生变化，已停止映射。")
            if len(raw_records) > PRODUCT_PERFORMANCE_PAGE_LENGTH:
                raise LingxingApiError("产品表现分页返回数量超过请求长度。")
            if not raw_records and offset < total:
                raise LingxingApiError("产品表现分页提前返回空页，数据可能不完整。")
            for mapping in page_mappings:
                if mapping.local_sku in seen_local_skus:
                    raise LingxingApiError("产品表现分页包含重复 local SKU，已停止映射。")
                seen_local_skus.add(mapping.local_sku)
                mappings[mapping.local_sku] = mapping
            if offset + len(raw_records) >= total or not raw_records:
                break
            offset += len(raw_records)
        return [mappings[sku] for sku in normalized_skus if sku in mappings]

    def _request_daily_page(
        self,
        *,
        access_token: str,
        body: dict[str, Any],
    ) -> Any:
        # The official page publishes a token-bucket capacity of 5 but no refill
        # rate. Serialize calls, and fail on rate-limit responses without guessing
        # a retry delay or automatically replaying requests.
        with _ASIN_DAILY_REQUEST_LOCK:
            return self._request_signed_post(
                ASIN_DAILY_LISTS_PATH,
                access_token=access_token,
                body=body,
            )

    def query_daily_sales(
        self,
        *,
        access_token: str,
        store_sid: int,
        mappings: list[SkuProductMapping],
        start_date: date,
        end_date: date,
    ) -> list[SkuDailySales]:
        """Query daily sales/units/orders by site date, then bind rows by MSKU."""
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError("Access Token 为空，未发送每日销量查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int):
            raise LingxingApiError("每日销量查询必须使用动态店铺 SID。")
        self._validate_query_range(start_date, end_date)
        if not mappings or any(not isinstance(item, SkuProductMapping) for item in mappings):
            raise LingxingApiError("没有已验证的 SKU→MSKU 映射，未发送每日销量查询。")

        msku_to_local: dict[str, str] = {}
        seen_local_skus: set[str] = set()
        for mapping in mappings:
            if not isinstance(mapping.local_sku, str) or not mapping.local_sku.strip():
                raise LingxingApiError("SKU→MSKU 映射包含无效 local SKU，已停止查询。")
            if mapping.local_sku in seen_local_skus:
                raise LingxingApiError("SKU→MSKU 映射包含重复 local SKU，已停止查询。")
            seen_local_skus.add(mapping.local_sku)
            if not mapping.seller_skus:
                raise LingxingApiError(f"SKU {mapping.local_sku} 没有可用于查询的 MSKU。")
            if any(not isinstance(value, str) or not value.strip() for value in mapping.seller_skus):
                raise LingxingApiError("SKU→MSKU 映射包含无效 MSKU，已停止查询。")
            if len(set(mapping.seller_skus)) != len(mapping.seller_skus):
                raise LingxingApiError("SKU→MSKU 映射包含重复 MSKU，已停止查询。")
            for seller_sku in mapping.seller_skus:
                prior_sku = msku_to_local.get(seller_sku)
                if prior_sku is not None and prior_sku != mapping.local_sku:
                    raise LingxingApiError("同一店铺 MSKU 映射到多个 local SKU，已停止归属。")
                msku_to_local[seller_sku] = mapping.local_sku

        values: dict[tuple[date, str], dict[int, _DailyMetricRecord]] = {}
        day_count = (end_date - start_date).days + 1
        for day_offset in range(day_count):
            day = start_date + timedelta(days=day_offset)
            for metric_type in (1, 2, 3):
                offset = 0
                seen_seller_skus: set[str] = set()
                while True:
                    body = {
                        "sid": store_sid,
                        "event_date": day.isoformat(),
                        "asin_type": 2,
                        "type": metric_type,
                        "offset": offset,
                        "length": ASIN_DAILY_PAGE_LENGTH,
                    }
                    payload = self._request_daily_page(
                        access_token=access_token.strip(),
                        body=body,
                    )
                    page = parse_asin_daily_page(
                        payload,
                        store_sid=store_sid,
                        event_date=day,
                        metric_type=metric_type,
                    )
                    if len(page) > ASIN_DAILY_PAGE_LENGTH:
                        raise LingxingApiError("每日销量统计分页返回数量超过请求长度。")
                    for record in page:
                        if record.seller_sku not in msku_to_local:
                            continue
                        if record.seller_sku in seen_seller_skus:
                            raise LingxingApiError("每日销量统计分页包含重复 MSKU，已停止归属。")
                        seen_seller_skus.add(record.seller_sku)
                        key = (day, record.seller_sku)
                        existing = values.setdefault(key, {})
                        for other in existing.values():
                            if (
                                other.asin != record.asin
                                or other.currency_code != record.currency_code
                            ):
                                raise LingxingApiError("同一 MSKU 的每日 ASIN/币种记录不一致，已停止归属。")
                        existing[metric_type] = record
                    if len(page) < ASIN_DAILY_PAGE_LENGTH:
                        break
                    offset += len(page)

        result: list[SkuDailySales] = []
        for day_offset in range(day_count):
            day = start_date + timedelta(days=day_offset)
            for mapping in mappings:
                for seller_sku in mapping.seller_skus:
                    metric_records = values.get((day, seller_sku), {})
                    identity = next(iter(metric_records.values()), None)
                    # Product Performance exposes ASIN and MSKU lists, but does
                    # not define a one-to-one relationship between those lists.
                    # Only an ASIN returned on this exact daily MSKU row is
                    # safe to display here.
                    asin = identity.asin if identity is not None else ""
                    currency_code = identity.currency_code if identity is not None else ""
                    units_value = metric_records.get(2)
                    orders_value = metric_records.get(3)
                    sales_value = metric_records.get(1)
                    result.append(
                        SkuDailySales(
                            sid=store_sid,
                            local_sku=mapping.local_sku,
                            seller_sku=seller_sku,
                            asin=asin,
                            event_date=day,
                            currency_code=currency_code,
                            units_sold=(int(units_value.metric_value) if units_value is not None else None),
                            order_count=(int(orders_value.metric_value) if orders_value is not None else None),
                            sales_amount=(sales_value.metric_value if sales_value is not None else None),
                        )
                    )
        return result

    def _query_ad_report_pages(
        self,
        *,
        path: str,
        endpoint_label: str,
        access_token: str,
        store_sid: int,
        report_date: date,
        parser: Any,
        identity: Any,
        show_detail: int | None = 0,
    ) -> list[Any]:
        offset = 0
        expected_total: int | None = None
        seen: set[Any] = set()
        output: list[Any] = []
        while True:
            body = {
                "sid": store_sid,
                "report_date": report_date.isoformat(),
                "offset": offset,
                "length": AD_REPORT_PAGE_LENGTH,
            }
            if show_detail is not None:
                body["show_detail"] = show_detail
            with _AD_REPORT_REQUEST_LOCK:
                payload = self._request_signed_post(
                    path,
                    access_token=access_token,
                    body=body,
                    headers={"X-API-VERSION": "2"},
                )
            total, page = parser(payload, report_date=report_date)
            if expected_total is None:
                expected_total = total
            elif total != expected_total:
                raise LingxingApiError(f"{endpoint_label}分页期间 total 发生变化，已停止查询。")
            if len(page) > AD_REPORT_PAGE_LENGTH or offset + len(page) > total:
                raise LingxingApiError(f"{endpoint_label}分页返回数量无效。")
            if not page and offset < total:
                raise LingxingApiError(f"{endpoint_label}分页提前返回空页，数据可能不完整。")
            for record in page:
                record_id = identity(record)
                if record_id in seen:
                    raise LingxingApiError(f"{endpoint_label}分页返回重复记录，已停止累计。")
                seen.add(record_id)
                output.append(record)
            if offset + len(page) >= total:
                break
            offset += len(page)
        return output

    def _query_ad_product_catalog_pages(
        self,
        *,
        path: str,
        endpoint_label: str,
        access_token: str,
        store_sid: int,
        parser: Any,
        identity: Any,
        api_version: str | None = None,
    ) -> tuple[AdProductCatalogRecord, ...]:
        """Read a complete current ad catalog using documented cursor/offset pagination."""
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError(f"Access Token 为空，未发送 {endpoint_label}查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
            raise LingxingApiError(f"{endpoint_label}查询必须使用动态店铺 SID。")
        offset = 0
        cursor = ""
        expected_total: int | None = None
        seen_tokens: set[str] = set()
        seen_records: set[Any] = set()
        output: list[AdProductCatalogRecord] = []
        while True:
            body: dict[str, Any] = {"sid": store_sid, "offset": offset, "length": AD_REPORT_PAGE_LENGTH}
            if cursor:
                # The documented cursor takes precedence over offset. Omit offset
                # on cursor requests to avoid ambiguity between API versions.
                body.pop("offset")
                body["next_token"] = cursor
            headers = {"X-API-VERSION": api_version} if api_version else None
            with _AD_REPORT_REQUEST_LOCK:
                payload = self._request_signed_post(
                    path,
                    access_token=access_token.strip(),
                    body=body,
                    headers=headers,
                )
            total, page, next_token = parser(payload)
            if expected_total is None:
                expected_total = total
            elif total != expected_total:
                raise LingxingApiError(f"{endpoint_label}分页期间 total 发生变化，已停止查询。")
            if len(page) > AD_REPORT_PAGE_LENGTH or len(output) + len(page) > total:
                raise LingxingApiError(f"{endpoint_label}分页返回数量无效。")
            if not page and len(output) < total:
                raise LingxingApiError(f"{endpoint_label}分页提前返回空页，数据可能不完整。")
            for record in page:
                record_id = identity(record)
                if record_id in seen_records:
                    raise LingxingApiError(f"{endpoint_label}分页返回重复广告记录，已停止查询。")
                seen_records.add(record_id)
                output.append(record)
            offset += len(page)
            if len(output) >= total:
                break
            if next_token:
                if next_token in seen_tokens:
                    raise LingxingApiError(f"{endpoint_label}分页游标重复，已停止查询。")
                seen_tokens.add(next_token)
                cursor = next_token
            else:
                # Both catalog schemas document offset. Use it only when the
                # service returned no cursor, and still validate total/duplicates.
                cursor = ""
                if offset >= total:
                    break
        return tuple(output)

    def query_sb_product_ads(
        self,
        *,
        access_token: str,
        store_sid: int,
    ) -> tuple[AdProductCatalogRecord, ...]:
        """Read current SB creative ASINs; the endpoint does not provide history."""
        return self._query_ad_product_catalog_pages(
            path=SB_PRODUCT_ADS_PATH,
            endpoint_label="SB 广告创意目录",
            access_token=access_token,
            store_sid=store_sid,
            parser=parse_sb_product_ads_page,
            identity=lambda row: (row.profile_id, row.campaign_id, row.ad_group_id, row.ad_creative_id),
        )

    def query_sd_product_ads(
        self,
        *,
        access_token: str,
        store_sid: int,
    ) -> tuple[AdProductCatalogRecord, ...]:
        """Read current SD ad/product identities for cross-checking historical rows."""
        return self._query_ad_product_catalog_pages(
            path=SD_PRODUCT_ADS_PATH,
            endpoint_label="SD 广告商品目录",
            access_token=access_token,
            store_sid=store_sid,
            parser=parse_sd_product_ads_page,
            identity=lambda row: (row.profile_id, row.ad_id),
            api_version="2",
        )

    def query_sb_divide_asin_reports(self, **_kwargs: Any) -> None:
        """Block the SB allocation endpoint until its malformed contract is corrected."""
        raise LingxingApiNotReady(
            f"{SB_DIVIDE_ASIN_REPORTS_PATH} 的 Schema 与示例字段错位，spends/percent 的结构和金额含义无法核验；未发送请求。"
        )

    def query_sp_ad_spend_history(
        self,
        *,
        access_token: str,
        store_sid: int,
        mappings: list[SkuProductMapping],
        start_date: date,
        end_date: date,
    ) -> list[SpAdSpendWindow]:
        """Read SP product costs and reconcile by campaign before SKU attribution."""
        return self._query_product_ad_spend_history(
            ad_type="SP",
            campaign_path=SP_CAMPAIGN_REPORTS_PATH,
            product_path=SP_PRODUCT_AD_REPORTS_PATH,
            campaign_parser=parse_sp_campaign_report_page,
            product_parser=parse_sp_product_ad_report_page,
            access_token=access_token,
            store_sid=store_sid,
            mappings=mappings,
            start_date=start_date,
            end_date=end_date,
        )

    def query_sd_ad_spend_history(
        self,
        *,
        access_token: str,
        store_sid: int,
        mappings: list[SkuProductMapping],
        start_date: date,
        end_date: date,
        include_catalog: bool = False,
    ) -> list[SpAdSpendWindow]:
        """Read historical SD product costs; current catalog is optional diagnostic context."""
        windows = self._query_product_ad_spend_history(
            ad_type="SD",
            campaign_path=SD_CAMPAIGN_REPORTS_PATH,
            product_path=SD_PRODUCT_AD_REPORTS_PATH,
            campaign_parser=parse_sd_campaign_report_page,
            product_parser=parse_sd_product_ad_report_page,
            access_token=access_token,
            store_sid=store_sid,
            mappings=mappings,
            start_date=start_date,
            end_date=end_date,
        )
        # The historical product report itself supplies report_date, MSKU, ASIN,
        # campaign/ad identifiers and cost. A present-day catalog is not required
        # to attribute that date-scoped row. Keep the catalog as best-effort
        # preview context only, after the historical query has succeeded.
        if not include_catalog:
            return windows
        try:
            catalog_rows = self.query_sd_product_ads(access_token=access_token, store_sid=store_sid)
        except LingxingApiError:
            catalog_rows = []
        return [replace(window, product_catalog_rows=catalog_rows) for window in windows]

    def _query_product_ad_spend_history(
        self,
        *,
        ad_type: str,
        campaign_path: str,
        product_path: str,
        campaign_parser: Any,
        product_parser: Any,
        access_token: str,
        store_sid: int,
        mappings: list[SkuProductMapping],
        start_date: date,
        end_date: date,
    ) -> list[SpAdSpendWindow]:
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError(f"Access Token 为空，未发送 {ad_type} 广告报表查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int):
            raise LingxingApiError(f"{ad_type} 广告报表查询必须使用动态店铺 SID。")
        self._validate_query_range(start_date, end_date)
        if not mappings or any(not isinstance(item, SkuProductMapping) for item in mappings):
            raise LingxingApiError(f"{ad_type} 广告查询需要已验证的 local SKU→MSKU 映射。")
        local_by_msku: dict[str, set[str]] = defaultdict(set)
        local_by_asin: dict[str, set[str]] = defaultdict(set)
        seen_local: set[str] = set()
        local_names: dict[str, str] = {}
        for mapping in mappings:
            local_key = mapping.local_sku.casefold()
            if local_key in seen_local or not (mapping.seller_skus or mapping.asins):
                raise LingxingApiError(f"{ad_type} 广告查询的 SKU 映射重复或没有可识别的 MSKU/ASIN。")
            seen_local.add(local_key)
            local_names[local_key] = mapping.local_sku
            for seller_sku in mapping.seller_skus:
                local_by_msku[seller_sku.casefold()].add(local_key)
            for asin in mapping.asins:
                local_by_asin[asin.casefold()].add(local_key)

        windows: list[SpAdSpendWindow] = []
        day_count = (end_date - start_date).days + 1
        for day_offset in range(day_count):
            report_day = start_date + timedelta(days=day_offset)
            campaigns = self._query_ad_report_pages(
                path=campaign_path,
                endpoint_label=f"{ad_type} 广告活动报表",
                access_token=access_token.strip(),
                store_sid=store_sid,
                report_date=report_day,
                parser=campaign_parser,
                identity=lambda row: (row.profile_id, row.campaign_id),
            )
            products = self._query_ad_report_pages(
                path=product_path,
                endpoint_label=f"{ad_type} 广告商品报表",
                access_token=access_token.strip(),
                store_sid=store_sid,
                report_date=report_day,
                parser=product_parser,
                identity=lambda row: (row.profile_id, row.ad_id),
            )

            campaign_costs = {(row.profile_id, row.campaign_id): row.cost for row in campaigns}
            product_costs: dict[tuple[str, str], Decimal] = defaultdict(lambda: Decimal("0"))
            for row in products:
                product_costs[(row.profile_id, row.campaign_id)] += row.cost
            discrepancies: list[SpAdSpendDiscrepancy] = []
            for profile_id, campaign_id in sorted(set(campaign_costs) | set(product_costs)):
                campaign_key = (profile_id, campaign_id)
                campaign_cost = campaign_costs.get(campaign_key)
                product_cost = product_costs.get(campaign_key)
                if campaign_cost is None or product_cost is None or campaign_cost != product_cost:
                    discrepancies.append(
                        SpAdSpendDiscrepancy(
                            report_day, campaign_id, campaign_cost, product_cost, "商品", profile_id
                        )
                    )

            sku_spends: list[SkuDailyAdSpend] = []
            report_complete = bool(campaigns and products and not discrepancies)
            if report_complete:
                blocked_campaigns = {(item.profile_id, item.campaign_id) for item in discrepancies}
                grouped: dict[tuple[str, str], list[SpProductAdSpend]] = defaultdict(list)
                attributed_product_indexes: set[int] = set()
                for index, row in enumerate(products):
                    if (row.profile_id, row.campaign_id) in blocked_campaigns:
                        continue
                    # Historical report identifiers are authoritative for their
                    # own report_date. Resolve exact MSKU and ASIN independently;
                    # if both are known they must identify the same local SKU.
                    owners = set(local_by_msku.get(row.sku.casefold(), set()))
                    if row.asin:
                        owners.update(local_by_asin.get(row.asin.casefold(), set()))
                    if len(owners) == 1:
                        local_key = next(iter(owners))
                        grouped[(local_key, row.sku.casefold())].append(row)
                        attributed_product_indexes.add(index)
                for (local_key, _msku_key), rows in grouped.items():
                    sku_spends.append(
                        SkuDailyAdSpend(
                            store_sid=store_sid,
                            local_sku=local_names[local_key],
                            seller_sku=rows[0].sku,
                            report_date=report_day,
                            actual_spend=sum((row.cost for row in rows), Decimal("0")),
                            campaign_count=len({(row.profile_id, row.campaign_id) for row in rows}),
                        )
                    )
            else:
                attributed_product_indexes = set()
            unmapped_cost = (
                sum(
                    (row.cost for index, row in enumerate(products) if index not in attributed_product_indexes),
                    Decimal("0"),
                )
                if report_complete else None
            )
            windows.append(SpAdSpendWindow(
                start_date=report_day,
                end_date=report_day,
                campaign_rows=tuple(campaigns),
                product_rows=tuple(products),
                sku_daily_spends=tuple(sku_spends),
                discrepancies=tuple(discrepancies),
                ad_type=ad_type,
                campaign_cost_total=(sum((row.cost for row in campaigns), Decimal("0")) if campaigns else None),
                detail_cost_total=(sum((row.cost for row in products), Decimal("0")) if products else None),
                unattributed_cost=unmapped_cost,
                report_complete=report_complete,
            ))
        return windows

    def query_sb_ad_spend_history(
        self,
        *,
        access_token: str,
        store_sid: int,
        start_date: date,
        end_date: date,
    ) -> list[SpAdSpendWindow]:
        """Read combined HSA campaign/ad-group cost; campaign totals only reconcile."""
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError("Access Token 为空，未发送 HSA 品牌广告报表查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int):
            raise LingxingApiError("HSA 品牌广告报表查询必须使用动态店铺 SID。")
        self._validate_query_range(start_date, end_date)
        # The creative catalog is a current snapshot, not historical ownership
        # proof. It is diagnostic only and must not block historical cost reads.
        try:
            catalog_rows = self.query_sb_product_ads(access_token=access_token, store_sid=store_sid)
            catalog_complete = True
            catalog_observed_at = datetime.now(timezone.utc)
        except LingxingApiError:
            catalog_rows = ()
            catalog_complete = False
            catalog_observed_at = None
        windows: list[SpAdSpendWindow] = []
        day_count = (end_date - start_date).days + 1
        for day_offset in range(day_count):
            report_day = start_date + timedelta(days=day_offset)
            campaigns = self._query_ad_report_pages(
                path=SB_CAMPAIGN_REPORTS_PATH,
                endpoint_label="HSA 品牌广告活动报表",
                access_token=access_token.strip(),
                store_sid=store_sid,
                report_date=report_day,
                parser=parse_sb_campaign_report_page,
                identity=lambda row: (row.profile_id, row.campaign_id),
                show_detail=None,
            )
            ad_groups = self._query_ad_report_pages(
                path=SB_AD_GROUP_REPORTS_PATH,
                endpoint_label="HSA 品牌广告组报表",
                access_token=access_token.strip(),
                store_sid=store_sid,
                report_date=report_day,
                parser=parse_sb_ad_group_report_page,
                identity=lambda row: (row.profile_id, row.campaign_id, row.ad_group_id),
                show_detail=None,
            )
            campaign_costs = {(row.profile_id, row.campaign_id): row.cost for row in campaigns}
            group_costs: dict[tuple[str, str], Decimal] = defaultdict(lambda: Decimal("0"))
            for row in ad_groups:
                group_costs[(row.profile_id, row.campaign_id)] += row.cost
            discrepancies: list[SpAdSpendDiscrepancy] = []
            for profile_id, campaign_id in sorted(set(campaign_costs) | set(group_costs)):
                campaign_key = (profile_id, campaign_id)
                campaign_cost = campaign_costs.get(campaign_key)
                group_cost = group_costs.get(campaign_key)
                if campaign_cost is None or group_cost is None or campaign_cost != group_cost:
                    discrepancies.append(
                        SpAdSpendDiscrepancy(
                            report_day, campaign_id, campaign_cost, group_cost, "广告组", profile_id
                        )
                    )
            report_complete = bool(campaigns and ad_groups and not discrepancies)
            campaign_total = sum((row.cost for row in campaigns), Decimal("0")) if campaigns else None
            detail_total = sum((row.cost for row in ad_groups), Decimal("0")) if ad_groups else None
            windows.append(SpAdSpendWindow(
                start_date=report_day,
                end_date=report_day,
                campaign_rows=tuple(campaigns),
                product_rows=(),
                sku_daily_spends=(),
                discrepancies=tuple(discrepancies),
                # HSA groups brand advertising together; subtype does not gate
                # SKU ownership or cost attribution.
                ad_type="HSA",
                ad_group_rows=tuple(ad_groups),
                campaign_cost_total=campaign_total,
                detail_cost_total=detail_total,
                # The preview applies date-scoped operator confirmations to
                # ad-group rows. Never allocate the activity amount itself.
                unattributed_cost=campaign_total if report_complete else None,
                report_complete=report_complete,
                promotion_targets=tuple(
                    PromotionTargetEvidence(
                        ad_type="HSA",
                        store_sid=store_sid,
                        report_date=report_day,
                        profile_id=row.profile_id,
                        campaign_id=row.campaign_id,
                        ad_group_id=row.ad_group_id,
                        asin=asin,
                        ad_creative_id=row.ad_creative_id,
                        # This catalog has no report-date history or effective
                        # timezone, so it is candidate identity only.
                        identity_valid_for_report=False,
                    )
                    for row in catalog_rows
                    for asin in row.asins
                ),
                promotion_catalog_complete=catalog_complete,
                product_catalog_rows=catalog_rows,
                product_catalog_observed_at=catalog_observed_at,
            ))
        return windows

    def query_daily_sales_history(
        self,
        *,
        access_token: str,
        store_sid: int,
        local_skus: list[str],
        start_date: date,
        end_date: date,
    ) -> list[SkuDailySalesWindow]:
        """Query a long date span in documented <=92-day product windows.

        Each window keeps its own mapping because the documented API does not
        promise that a local SKU's MSKU list is timeless. A missing mapping or
        a rate-limit/error in any window aborts the whole result; callers must
        not generate a partial-looking workbook from incomplete history.
        """
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError("Access Token 为空，未发送历史销量查询。")
        if isinstance(store_sid, bool) or not isinstance(store_sid, int):
            raise LingxingApiError("历史销量查询必须使用动态店铺 SID。")
        if not isinstance(local_skus, list) or not local_skus:
            raise LingxingApiError("历史销量查询至少需要一个 local SKU。")
        normalized_skus = list(dict.fromkeys(
            sku.strip() for sku in local_skus if isinstance(sku, str) and sku.strip()
        ))
        if len(normalized_skus) != len(local_skus):
            raise LingxingApiError("历史销量查询包含空白或重复 SKU，已停止。")

        windows: list[SkuDailySalesWindow] = []
        for window_start, window_end in split_date_windows(start_date, end_date):
            mappings = self.query_product_mappings(
                access_token=access_token,
                store_sid=store_sid,
                local_skus=normalized_skus,
                start_date=window_start,
                end_date=window_end,
            )
            mapping_by_sku = {item.local_sku.casefold(): item for item in mappings}
            unresolved = [
                sku for sku in normalized_skus
                if not mapping_by_sku.get(sku.casefold())
                or not mapping_by_sku[sku.casefold()].seller_skus
            ]
            if unresolved:
                raise LingxingApiError(
                    f"{window_start.isoformat()} 至 {window_end.isoformat()} 的 SKU→MSKU 映射不完整："
                    + ", ".join(unresolved)
                    + "。未生成部分历史结果。"
                )
            selected_mappings = [mapping_by_sku[sku.casefold()] for sku in normalized_skus]
            rows = self.query_daily_sales(
                access_token=access_token,
                store_sid=store_sid,
                mappings=selected_mappings,
                start_date=window_start,
                end_date=window_end,
            )
            windows.append(
                SkuDailySalesWindow(
                    start_date=window_start,
                    end_date=window_end,
                    mappings=tuple(selected_mappings),
                    rows=tuple(rows),
                )
            )
        return windows

    def list_sellers(self, access_token: str) -> list[LingxingSeller]:
        """Call SellerLists and return records only after business/schema validation."""
        if not isinstance(access_token, str) or not access_token.strip():
            raise LingxingApiError("Access Token 为空，未发送店铺列表请求。")
        parameters = {
            "access_token": access_token.strip(),
            "app_key": self._app_id,
            "timestamp": str(int(time.time())),
        }
        signed = signed_query_parameters(parameters, self._app_id)
        payload = self._request_signed_get(
            SELLER_LIST_PATH,
            signed,
            sensitive_values=(access_token, signed["sign"]),
        )
        return parse_seller_list_response(payload)

    def query_sku_sales(
        self,
        *,
        access_token: str,
        seller_id: str,
        sku: str,
        start_date: str,
        end_date: str,
    ) -> list[dict[str, Any]]:
        del access_token, seller_id, sku, start_date, end_date
        raise LingxingApiNotReady(
            "销售查询接口路径、参数和字段尚未依据领星官方文档核验，未发送请求。"
        )

    def close(self) -> None:
        self.clear_token_cache()
        if self._owns_http:
            self._http.close()

    def __enter__(self) -> "LingxingOpenApiClient":
        return self

    def __exit__(self, *_exc: object) -> None:
        self.close()
