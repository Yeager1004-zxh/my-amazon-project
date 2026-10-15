from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from urllib.parse import urlsplit


class LocalSettingsError(ValueError):
    """Invalid or inaccessible local connection settings."""


def validate_mcp_endpoint(endpoint: str) -> str:
    value = endpoint.strip()
    if not value:
        raise LocalSettingsError("请输入领星官方 MCP 文档提供的 Streamable HTTP 地址。")
    parsed = urlsplit(value)
    if parsed.scheme not in {"https", "http"} or not parsed.hostname:
        raise LocalSettingsError("MCP 地址必须是有效的 HTTPS URL。")
    if parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise LocalSettingsError("MCP 地址不能包含用户名、密码、查询参数或片段。")
    if parsed.scheme == "http":
        host = parsed.hostname.casefold()
        if host not in {"127.0.0.1", "localhost", "::1"}:
            raise LocalSettingsError("仅允许 HTTPS；HTTP 只可用于本机协议测试。")
    return value.rstrip("/")


def settings_path() -> Path:
    if os.name == "nt":
        root = os.environ.get("APPDATA")
        if not root:
            raise LocalSettingsError("Windows 未提供 APPDATA 目录。")
        return Path(root) / "MyAmazonProject" / "settings.json"
    return Path.home() / ".config" / "my-amazon-project" / "settings.json"


def load_settings() -> dict[str, str]:
    path = settings_path()
    if not path.is_file():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise LocalSettingsError("本地设置文件无法读取。") from exc
    if not isinstance(value, dict):
        raise LocalSettingsError("本地设置文件格式错误。")
    if "x_mcp_key" in value or "secret" in value or "api_key" in value:
        raise LocalSettingsError(
            "本地设置文件含有不应落盘的凭证字段。请删除该字段，并在 Windows Credential Manager 中重新保存凭证。"
        )
    endpoint = value.get("endpoint", "")
    if endpoint:
        value["endpoint"] = validate_mcp_endpoint(str(endpoint))
    return {key: str(item) for key, item in value.items() if key in {"endpoint"}}


def save_settings(*, endpoint: str) -> Path:
    normalized = validate_mcp_endpoint(endpoint)
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps({"endpoint": normalized}, ensure_ascii=False, indent=2) + "\n"
    fd, temp_name = tempfile.mkstemp(prefix="settings-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_name, path)
    finally:
        try:
            os.unlink(temp_name)
        except FileNotFoundError:
            pass
    return path
