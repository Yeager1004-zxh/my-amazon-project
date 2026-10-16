from __future__ import annotations

import math
import re
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx

from .lingxing_signing import encode_query_parameters, signed_query_parameters


OPENAPI_BASE_URL = "https://openapi.lingxing.com"
ACCESS_TOKEN_PATH = "/api/auth-server/oauth/access-token"
SELLER_LIST_PATH = "/erp/sc/data/seller/lists"


class LingxingApiError(RuntimeError):
    """A sanitized Lingxing OpenAPI transport or response error."""


class LingxingApiNotReady(LingxingApiError):
    """A query is deliberately blocked until its official contract is verified."""


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
    """Official-host client for token exchange and the signed seller-list query."""

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
