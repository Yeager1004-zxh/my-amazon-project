from __future__ import annotations

import math
import json
import re
import threading
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from typing import Any

import httpx

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
PRODUCT_PERFORMANCE_PAGE_LENGTH = 10_000
ASIN_DAILY_PAGE_LENGTH = 1_000
MAX_PRODUCT_QUERY_DAYS = 92
_PRODUCT_PERFORMANCE_REQUEST_LOCK = threading.Lock()
_PRODUCT_PERFORMANCE_LAST_RESPONSE_AT: float | None = None
_ASIN_DAILY_REQUEST_LOCK = threading.Lock()


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
            if any(marker in normalized_key for marker in ("token", "secret", "password", "sign", "authorization", "appkey", "appid")):
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
    """Official-host client for signed seller, product mapping, and daily sales reads."""

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
            response = self._http.post(url, json=body)
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
