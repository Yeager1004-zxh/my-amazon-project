from __future__ import annotations

import math
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any

import httpx


OPENAPI_BASE_URL = "https://openapi.lingxing.com"
ACCESS_TOKEN_PATH = "/api/auth-server/oauth/access-token"
SELLER_LIST_PATH = "/erp/sc/data/seller/lists"


class LingxingApiError(RuntimeError):
    """A sanitized Lingxing OpenAPI transport or response error."""


class LingxingApiNotReady(LingxingApiError):
    """A query is deliberately blocked until its official contract is verified."""


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


class LingxingOpenApiClient:
    """Minimal official-host client for the documented multipart token exchange.

    Calls that require request signing remain fail-closed until the precise
    official signing contract is available and covered by tests.
    """

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

    def list_sellers(self, access_token: str) -> list[dict[str, Any]]:
        del access_token
        raise LingxingApiNotReady(
            "店铺列表接口需要 app_key、timestamp 和 sign；领星官方签名规则尚未核验，未发送请求。"
        )

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
