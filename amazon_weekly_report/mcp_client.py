from __future__ import annotations

import json
import hashlib
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


@dataclass(frozen=True)
class BusinessTool:
    """A business operation and its schema returned by the MCP search tool."""

    tool_id: str
    name: str
    description: str
    input_schema: dict[str, Any]

    @property
    def permission_fingerprint(self) -> str:
        body = json.dumps(
            {
                "id": self.tool_id,
                "name": self.name,
                "description": self.description,
                "input_schema": self.input_schema,
            },
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
        return hashlib.sha256(body).hexdigest()


@dataclass(frozen=True)
class VerifiedReadOnlyGrant:
    """Session-only approval bound to one business tool's exact returned schema."""

    tool_id: str
    permission_fingerprint: str


_WRITE_WORDS = (
    "create", "update", "delete", "remove", "modify", "edit", "write", "submit", "save",
    "insert", "add", "set", "change", "enable", "disable", "pause", "resume", "cancel",
    "创建", "新增", "添加", "修改", "更新", "删除", "移除", "编辑", "写入", "提交", "保存",
    "启用", "停用", "暂停", "恢复", "取消", "变更", "设置",
)
_READ_WORDS = ("query", "get", "list", "fetch", "read", "search", "查询", "获取", "列表", "读取", "搜索")


def _contains_policy_word(value: str, words: tuple[str, ...]) -> bool:
    value = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", value).replace("_", " ").replace("-", " ").casefold()
    for word in words:
        if word.isascii() and word.replace("_", "").isalnum():
            if re.search(rf"\b{re.escape(word)}\b", value, flags=re.IGNORECASE):
                return True
        elif word in value:
            return True
    return False


def is_directory_search_tool(tool: McpTool) -> bool:
    """Allow only an explicitly described tool-catalog/schema discovery call."""
    if tool.name.casefold() != "search":
        return False
    description = tool.description.casefold()
    if _contains_policy_word(description, _WRITE_WORDS):
        return False
    describes_tools = any(word in description for word in ("tool", "工具", "业务功能"))
    describes_schema = any(word in description for word in ("schema", "参数", "目录", "工具列表"))
    return describes_tools and describes_schema and isinstance(tool.input_schema, dict)


def is_read_only_query_candidate(tool: BusinessTool) -> bool:
    """A conservative candidate check; a user must still verify and grant it."""
    identity = f"{tool.tool_id} {tool.name} {tool.description}"
    if _contains_policy_word(identity, _WRITE_WORDS):
        return False
    return (
        _contains_policy_word(identity, _READ_WORDS)
        and tool.input_schema.get("type") == "object"
        and tool.input_schema.get("additionalProperties") is not True
        and not _schema_has_write_fields(tool.input_schema)
        and not _schema_has_ambiguous_operation_fields(tool.input_schema)
    )


def _schema_has_write_fields(schema: Any) -> bool:
    if not isinstance(schema, dict):
        return True
    properties = schema.get("properties", {})
    if not isinstance(properties, dict):
        return True
    for name, child in properties.items():
        normalized = str(name)
        if _contains_policy_word(normalized, _WRITE_WORDS) or _schema_has_write_fields(child):
            return True
    items = schema.get("items")
    return _schema_has_write_fields(items) if isinstance(items, dict) else False


def _schema_has_ambiguous_operation_fields(schema: Any) -> bool:
    if not isinstance(schema, dict):
        return True
    properties = schema.get("properties", {})
    if not isinstance(properties, dict):
        return True
    routing_fields = {"action", "operation", "command", "method", "endpoint", "mode", "tool", "tool_id", "toolid"}
    for name, child in properties.items():
        normalized = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(name)).replace("_", " ").replace("-", " ").casefold()
        if normalized in routing_fields or _schema_has_ambiguous_operation_fields(child):
            return True
    items = schema.get("items")
    return _schema_has_ambiguous_operation_fields(items) if isinstance(items, dict) else False


def is_action_wrapper_tool(tool: McpTool) -> bool:
    """Identify the vendor wrapper that executes a previously searched business tool."""
    name_match = tool.name.casefold() == "action" or re.search(r"(?:^|[._-])action$", tool.name.casefold()) is not None
    description = tool.description.casefold()
    return (
        name_match
        and any(word in description for word in ("执行指定业务工具", "execute the specified business tool", "execute specified business tool"))
        and "search" in description
        and "schema" in description
    )


def _decode_schema(value: Any) -> dict[str, Any] | None:
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            return None
    return value if isinstance(value, dict) and value.get("type") == "object" else None


def _candidate_business_tool(value: Any) -> BusinessTool | None:
    if not isinstance(value, dict):
        return None
    tool_id = next((value.get(key) for key in ("tool_id", "toolId", "business_tool_id", "businessToolId", "function_id", "functionId", "id", "tool_name", "toolName", "name") if isinstance(value.get(key), (str, int))), None)
    schema_value = next((value.get(key) for key in ("inputSchema", "input_schema", "schema", "parametersSchema", "parameter_schema", "parameters") if value.get(key) is not None), None)
    schema = _decode_schema(schema_value)
    if tool_id is None or schema is None:
        return None
    name = next((value.get(key) for key in ("name", "tool_name", "toolName", "title") if isinstance(value.get(key), str) and value.get(key).strip()), str(tool_id))
    description = next((value.get(key) for key in ("description", "desc", "summary") if isinstance(value.get(key), str)), "")
    return BusinessTool(str(tool_id), name.strip(), description.strip(), redact_sensitive(schema))


def extract_business_tools(value: Any) -> list[BusinessTool]:
    """Extract catalog entries from MCP structuredContent or JSON text content."""
    roots: list[Any] = [value]
    if isinstance(value, dict):
        for key in ("structuredContent", "data", "result", "tools", "items", "results", "functions", "business_tools", "content"):
            nested = value.get(key)
            if nested is not None:
                roots.append(nested)
    decoded_roots: list[Any] = []
    for root in roots:
        if isinstance(root, list):
            for item in root:
                if isinstance(item, dict) and item.get("type") == "text" and isinstance(item.get("text"), str):
                    try:
                        decoded_roots.append(json.loads(item["text"]))
                    except json.JSONDecodeError:
                        continue
                else:
                    decoded_roots.append(item)
        elif isinstance(root, dict):
            decoded_roots.append(root)
        elif isinstance(root, str):
            try:
                decoded_roots.append(json.loads(root))
            except json.JSONDecodeError:
                continue

    found: dict[tuple[str, str], BusinessTool] = {}
    queue = list(decoded_roots)
    while queue:
        current = queue.pop(0)
        candidate = _candidate_business_tool(current)
        if candidate is not None:
            found[(candidate.tool_id, candidate.permission_fingerprint)] = candidate
            continue
        if isinstance(current, dict):
            for item in current.values():
                if isinstance(item, (dict, list)):
                    queue.append(item)
        elif isinstance(current, list):
            queue.extend(current)
    return list(found.values())


def validate_json_schema(schema: dict[str, Any], value: Any, *, path: str = "参数") -> None:
    """Validate the strict JSON Schema subset used for MCP object parameters."""
    if not isinstance(schema, dict) or schema.get("type") != "object":
        raise McpClientError("Schema 不支持安全校验，已阻止调用。")
    if any(key in schema for key in ("$ref", "oneOf", "anyOf", "allOf", "not", "if", "then", "else", "patternProperties")):
        raise McpClientError("Schema 含有尚未支持的组合规则，已阻止调用。")
    properties = schema.get("properties", {})
    required = schema.get("required", [])
    if not isinstance(properties, dict) or not isinstance(required, list) or any(not isinstance(key, str) for key in required):
        raise McpClientError("Schema 的 properties/required 格式无效，已阻止调用。")
    if not isinstance(value, dict):
        raise McpClientError(f"{path}必须是 JSON 对象。")
    if len(value) < schema.get("minProperties", 0):
        raise McpClientError(f"{path}字段数量少于 Schema 要求。")
    if "maxProperties" in schema and len(value) > schema["maxProperties"]:
        raise McpClientError(f"{path}字段数量超过 Schema 限制。")
    missing = [key for key in required if key not in value]
    if missing:
        raise McpClientError(f"{path}缺少必填字段：{', '.join(missing)}。")
    for key, item in value.items():
        item_schema = properties.get(key)
        if item_schema is not None:
            _validate_schema_value(item_schema, item, f"{path}.{key}")
            continue
        additional = schema.get("additionalProperties", False)
        if additional is True:
            continue
        if isinstance(additional, dict):
            _validate_schema_value(additional, item, f"{path}.{key}")
            continue
        raise McpClientError(f"{path}字段“{key}”未在 Schema 中定义，已阻止调用。")


def _validate_schema_value(schema: Any, value: Any, path: str) -> None:
    if not isinstance(schema, dict):
        raise McpClientError(f"{path}的 Schema 无效，已阻止调用。")
    if any(key in schema for key in ("$ref", "oneOf", "anyOf", "allOf", "not", "if", "then", "else", "patternProperties")):
        raise McpClientError(f"{path}包含尚未支持的 Schema 组合规则，已阻止调用。")
    expected = schema.get("type")
    matches = {
        "string": lambda v: isinstance(v, str),
        "integer": lambda v: isinstance(v, int) and not isinstance(v, bool),
        "number": lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
        "boolean": lambda v: isinstance(v, bool),
        "object": lambda v: isinstance(v, dict),
        "array": lambda v: isinstance(v, list),
        "null": lambda v: v is None,
    }
    if expected not in matches:
        raise McpClientError(f"{path}包含不支持的 Schema 类型，已阻止调用。")
    if not matches[expected](value):
        raise McpClientError(f"{path}类型不符合 Schema（应为 {expected}）。")
    if "enum" in schema and value not in schema["enum"]:
        raise McpClientError(f"{path}不在 Schema 允许的选项中。")
    if "const" in schema and value != schema["const"]:
        raise McpClientError(f"{path}不符合 Schema 固定值。")
    if expected == "string":
        if len(value) < schema.get("minLength", 0):
            raise McpClientError(f"{path}短于 Schema 要求。")
        if "maxLength" in schema and len(value) > schema["maxLength"]:
            raise McpClientError(f"{path}超过 Schema 长度限制。")
        if "pattern" in schema:
            try:
                if re.search(str(schema["pattern"]), value) is None:
                    raise McpClientError(f"{path}格式不符合 Schema。")
            except re.error as exc:
                raise McpClientError(f"{path}的 Schema 正则表达式无效，已阻止调用。") from exc
        if "format" in schema:
            format_name = schema["format"]
            if format_name == "date":
                from datetime import date
                try:
                    date.fromisoformat(value)
                except ValueError as exc:
                    raise McpClientError(f"{path}应为 YYYY-MM-DD 日期。") from exc
            elif format_name == "date-time":
                from datetime import datetime
                try:
                    datetime.fromisoformat(value.replace("Z", "+00:00"))
                except ValueError as exc:
                    raise McpClientError(f"{path}应为有效日期时间。") from exc
            else:
                raise McpClientError(f"{path}包含尚未支持的 Schema format，已阻止调用。")
    if expected in {"integer", "number"}:
        if "minimum" in schema and value < schema["minimum"]:
            raise McpClientError(f"{path}小于 Schema 最小值。")
        if "maximum" in schema and value > schema["maximum"]:
            raise McpClientError(f"{path}大于 Schema 最大值。")
        if "exclusiveMinimum" in schema and value <= schema["exclusiveMinimum"]:
            raise McpClientError(f"{path}未超过 Schema 最小边界。")
        if "exclusiveMaximum" in schema and value >= schema["exclusiveMaximum"]:
            raise McpClientError(f"{path}超过 Schema 最大边界。")
        if "multipleOf" in schema and value % schema["multipleOf"] != 0:
            raise McpClientError(f"{path}不是 Schema 要求的倍数。")
    if expected == "object":
        validate_json_schema(schema, value, path=path)
    elif expected == "array":
        if len(value) < schema.get("minItems", 0):
            raise McpClientError(f"{path}数组元素数量少于 Schema 要求。")
        if "maxItems" in schema and len(value) > schema["maxItems"]:
            raise McpClientError(f"{path}数组元素数量超过 Schema 限制。")
        if schema.get("uniqueItems") is True and len(value) != len({json.dumps(item, sort_keys=True, ensure_ascii=False) for item in value}):
            raise McpClientError(f"{path}数组含有重复项，违反 Schema 要求。")
        item_schema = schema.get("items")
        if not isinstance(item_schema, dict):
            raise McpClientError(f"{path}数组缺少可验证的 items Schema，已阻止调用。")
        for index, item in enumerate(value):
            _validate_schema_value(item_schema, item, f"{path}[{index}]")


def build_action_arguments(action_tool: McpTool, business_tool: BusinessTool, arguments: dict[str, Any]) -> dict[str, Any]:
    """Map a verified business ID and args into the wrapper's declared schema."""
    schema = action_tool.input_schema
    properties = schema.get("properties") if isinstance(schema, dict) else None
    if not isinstance(properties, dict):
        raise McpClientError("action 的 inputSchema 不包含可识别的 properties，已阻止调用。")
    normalized = {str(key).casefold(): key for key in properties}
    id_names = ("tool_id", "toolid", "business_tool_id", "businesstoolid", "business_tool_name", "businesstoolname", "tool_name", "toolname", "tool", "name", "id")
    args_names = ("arguments", "params", "parameters", "input", "inputs", "tool_input", "toolinput", "tool_arguments", "toolarguments")
    id_key = next((normalized[name] for name in id_names if name in normalized), None)
    args_key = next((normalized[name] for name in args_names if name in normalized and normalized[name] != id_key), None)
    if id_key is None or args_key is None:
        raise McpClientError("action Schema 未声明业务工具 ID 和参数对象字段，已阻止调用。")
    envelope = {id_key: business_tool.tool_id, args_key: arguments}
    validate_json_schema(schema, envelope, path="action 参数")
    return envelope


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
            if secret:
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
        self._discovered_business_tools: dict[tuple[str, str], BusinessTool] = {}

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
        validate_json_schema(tool.input_schema, arguments)
        result = self._request("tools/call", {"name": tool.name, "arguments": arguments})
        if not isinstance(result, dict):
            raise McpClientError("MCP 工具结果格式无效。")
        if result.get("isError") is True:
            raise McpClientError(f"只读工具 {tool.name} 返回错误；请检查参数和该账号的数据权限。")
        return redact_sensitive(result, secret_values=(self._mcp_key,))

    def call_directory_search(self, tool: McpTool, arguments: dict[str, Any]) -> dict[str, Any]:
        """Call only the tool explicitly described as a business-tool Schema catalog search."""
        if not is_directory_search_tool(tool):
            raise McpClientError("search 未能确认为只查询业务工具目录的发现工具，已阻止调用。")
        validate_json_schema(tool.input_schema, arguments)
        result = self._call_tool(tool.name, arguments, error_label="search 目录发现工具")
        specs = extract_business_tools(result)
        self._discovered_business_tools = {(spec.tool_id, spec.permission_fingerprint): spec for spec in specs}
        return result

    def call_verified_business_query(
        self,
        action_tool: McpTool,
        business_tool: BusinessTool,
        arguments: dict[str, Any],
        grant: VerifiedReadOnlyGrant,
    ) -> dict[str, Any]:
        """Call action only for a query candidate explicitly approved in this session."""
        if not is_action_wrapper_tool(action_tool):
            raise McpClientError("未找到领星 action 包装工具，已阻止调用。")
        if (business_tool.tool_id, business_tool.permission_fingerprint) not in self._discovered_business_tools:
            raise McpClientError("action 的业务工具 ID/Schema 未在本次 MCP 会话的 search 结果中出现，已阻止调用。")
        if not is_read_only_query_candidate(business_tool):
            raise McpClientError("该业务工具无法确认是只读查询，已阻止 action 调用。")
        if not isinstance(grant, VerifiedReadOnlyGrant):
            raise McpClientError("缺少本次会话核验的只读许可，已阻止 action 调用。")
        if grant.tool_id != business_tool.tool_id or grant.permission_fingerprint != business_tool.permission_fingerprint:
            raise McpClientError("只读许可与当前业务工具 ID/Schema 不匹配，已阻止 action 调用。")
        validate_json_schema(business_tool.input_schema, arguments, path="业务查询参数")
        envelope = build_action_arguments(action_tool, business_tool, arguments)
        return self._call_tool(action_tool.name, envelope, error_label=f"业务查询 {business_tool.tool_id}")

    def _call_tool(self, name: str, arguments: dict[str, Any], *, error_label: str) -> dict[str, Any]:
        result = self._request("tools/call", {"name": name, "arguments": arguments})
        if not isinstance(result, dict):
            raise McpClientError("MCP 工具结果格式无效。")
        if result.get("isError") is True:
            raise McpClientError(f"{error_label}返回错误；请检查 Schema 和账号权限。")
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
