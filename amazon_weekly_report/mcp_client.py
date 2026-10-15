from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any

import httpx

from .local_settings import validate_mcp_endpoint


MCP_PROTOCOL_VERSION = "2025-03-26"
CLIENT_NAME = "my-amazon-project"
CLIENT_VERSION = "0.2.0"
_SENSITIVE_KEY = re.compile(
    r"(?:secret|token|api.?key|mcp.?key|authorization|password|credential)", re.IGNORECASE
)


class McpClientError(RuntimeError):
    """A sanitized MCP transport or protocol error."""


@dataclass(frozen=True)
class McpTool:
    name: str
    description: str
    input_schema: dict[str, Any]
    annotations: dict[str, Any]

    @property
    def read_only(self) -> bool:
        return self.annotations.get("readOnlyHint") is True


def redact_sensitive(value: Any, *, secret_values: tuple[str, ...] = ()) -> Any:
    """Return a display-safe copy without common credential fields."""
    if isinstance(value, dict):
        return {
            str(key): "[已隐藏]" if _SENSITIVE_KEY.search(str(key)) else redact_sensitive(item, secret_values=secret_values)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [redact_sensitive(item, secret_values=secret_values) for item in value]
    if isinstance(value, str):
        safe = value
        for secret in secret_values:
            if len(secret) >= 6:
                safe = safe.replace(secret, "[已隐藏]")
        return safe
    return value


class LingxingMcpClient:
    """Minimal MCP Streamable HTTP client using the vendor-provided endpoint."""

    def __init__(
        self,
        endpoint: str,
        mcp_key: str,
        *,
        timeout_seconds: float = 45,
        http_client: httpx.Client | None = None,
    ) -> None:
        self.endpoint = validate_mcp_endpoint(endpoint)
        if not mcp_key or "\r" in mcp_key or "\n" in mcp_key:
            raise McpClientError("X-Mcp-Key 缺失或格式无效。")
        self._mcp_key = mcp_key
        self._http = http_client or httpx.Client(timeout=timeout_seconds, follow_redirects=False)
        self._owns_http = http_client is None
        self._session_id: str | None = None
        self._protocol_version: str | None = None
        self._initialized = False
        self._next_id = 1

    @property
    def connected(self) -> bool:
        return self._initialized

    def _headers(self, *, initialized: bool) -> dict[str, str]:
        headers = {
            "Accept": "application/json, text/event-stream",
            "Content-Type": "application/json",
            "X-Mcp-Key": self._mcp_key,
        }
        if self._session_id:
            headers["Mcp-Session-Id"] = self._session_id
        if initialized and self._protocol_version:
            headers["MCP-Protocol-Version"] = self._protocol_version
        return headers

    @staticmethod
    def _parse_sse(body: str) -> list[dict[str, Any]]:
        messages: list[dict[str, Any]] = []
        data_lines: list[str] = []
        for line in body.splitlines() + [""]:
            if not line:
                if data_lines:
                    try:
                        message = json.loads("\n".join(data_lines))
                    except json.JSONDecodeError as exc:
                        raise McpClientError("MCP 服务返回了无法解析的 SSE 消息。") from exc
                    if isinstance(message, dict):
                        messages.append(message)
                    data_lines.clear()
            elif line.startswith("data:"):
                data_lines.append(line[5:].lstrip())
        return messages

    def _decode_response(self, response: httpx.Response, *, request_id: int | None) -> dict[str, Any] | None:
        if response.status_code >= 400:
            # Do not include the response body or request headers: a server may echo credentials.
            raise McpClientError(f"MCP 请求失败，HTTP 状态码 {response.status_code}。")
        session_id = response.headers.get("Mcp-Session-Id")
        if session_id:
            self._session_id = session_id
        body = response.content
        if not body:
            return None
        content_type = response.headers.get("Content-Type", "").lower()
        try:
            if "text/event-stream" in content_type:
                messages = self._parse_sse(body.decode("utf-8", "replace"))
            else:
                message = json.loads(body.decode("utf-8"))
                messages = [message] if isinstance(message, dict) else []
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise McpClientError("MCP 服务返回了无法解析的 JSON-RPC 响应。") from exc
        if request_id is None:
            return next((message for message in messages if "id" in message), None)
        for message in messages:
            if message.get("id") == request_id:
                if "error" in message:
                    error = message.get("error")
                    code = error.get("code") if isinstance(error, dict) else None
                    raise McpClientError(f"MCP JSON-RPC 返回错误（代码 {code!s}）。")
                return message
        raise McpClientError("MCP 响应中没有对应的 JSON-RPC 请求编号。")

    def _request(self, method: str, params: dict[str, Any] | None = None, *, initialized: bool = True) -> Any:
        if not self._initialized and method != "initialize":
            raise McpClientError("请先连接并完成 MCP 初始化。")
        request_id = self._next_id
        self._next_id += 1
        payload: dict[str, Any] = {
            "jsonrpc": "2.0",
            "id": request_id,
            "method": method,
            "params": params or {},
        }
        try:
            response = self._http.post(
                self.endpoint,
                headers=self._headers(initialized=initialized),
                json=payload,
            )
        except httpx.TimeoutException as exc:
            raise McpClientError("连接领星 MCP 超时。") from exc
        except httpx.HTTPError as exc:
            raise McpClientError("无法连接领星 MCP，请检查官方地址、网络和 X-Mcp-Key。") from exc
        message = self._decode_response(response, request_id=request_id)
        if message is None or "result" not in message:
            raise McpClientError("MCP 响应缺少 result。")
        return message["result"]

    def _notification(self, method: str, params: dict[str, Any] | None = None) -> None:
        payload = {"jsonrpc": "2.0", "method": method, "params": params or {}}
        try:
            response = self._http.post(
                self.endpoint,
                headers=self._headers(initialized=True),
                json=payload,
            )
        except httpx.HTTPError as exc:
            raise McpClientError("发送 MCP 初始化通知失败。") from exc
        if response.status_code >= 400:
            raise McpClientError(f"MCP 初始化通知失败，HTTP 状态码 {response.status_code}。")

    def connect(self) -> dict[str, Any]:
        if self._initialized:
            return {"serverInfo": {}, "protocolVersion": self._protocol_version}
        result = self._request(
            "initialize",
            {
                "protocolVersion": MCP_PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": CLIENT_NAME, "version": CLIENT_VERSION},
            },
            initialized=False,
        )
        if not isinstance(result, dict) or not isinstance(result.get("protocolVersion"), str):
            raise McpClientError("领星 MCP 初始化响应缺少协议版本。")
        self._protocol_version = result["protocolVersion"]
        self._initialized = True
        self._notification("notifications/initialized")
        return redact_sensitive(result, secret_values=(self._mcp_key,))

    def list_tools(self, *, max_pages: int = 20) -> list[McpTool]:
        cursor: str | None = None
        tools: list[McpTool] = []
        for _ in range(max_pages):
            params = {"cursor": cursor} if cursor else {}
            result = self._request("tools/list", params)
            if not isinstance(result, dict) or not isinstance(result.get("tools"), list):
                raise McpClientError("MCP tools/list 响应格式无效。")
            for item in result["tools"]:
                if not isinstance(item, dict) or not isinstance(item.get("name"), str):
                    continue
                tools.append(
                    McpTool(
                        name=item["name"],
                        description=redact_sensitive(str(item.get("description") or ""), secret_values=(self._mcp_key,)),
                        input_schema=redact_sensitive(item.get("inputSchema") if isinstance(item.get("inputSchema"), dict) else {}, secret_values=(self._mcp_key,)),
                        annotations=redact_sensitive(item.get("annotations") if isinstance(item.get("annotations"), dict) else {}, secret_values=(self._mcp_key,)),
                    )
                )
            cursor = result.get("nextCursor")
            if not cursor:
                return tools
        raise McpClientError("MCP 工具列表分页超过安全上限，已停止读取。")

    def call_read_only_tool(self, tool: McpTool, arguments: dict[str, Any]) -> dict[str, Any]:
        if not tool.read_only:
            raise McpClientError(
                "为避免写入业务数据，只能调用 tools/list 明确标记 readOnlyHint=true 的工具。"
            )
        result = self._request("tools/call", {"name": tool.name, "arguments": arguments})
        if not isinstance(result, dict):
            raise McpClientError("MCP 工具结果格式无效。")
        if result.get("isError") is True:
            raise McpClientError(f"只读工具 {tool.name} 返回错误；请检查参数和该账号的数据权限。")
        return redact_sensitive(result, secret_values=(self._mcp_key,))

    def close(self) -> None:
        if self._session_id:
            try:
                response = self._http.delete(
                    self.endpoint,
                    headers=self._headers(initialized=True),
                )
                response.close()
            except Exception:
                pass
            self._session_id = None
        if self._owns_http:
            self._http.close()

    def __enter__(self) -> "LingxingMcpClient":
        return self

    def __exit__(self, *_exc: object) -> None:
        self.close()
