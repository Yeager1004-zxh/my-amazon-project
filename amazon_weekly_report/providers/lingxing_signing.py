"""Lingxing OpenAPI request signing based on its bundled Java SDK.

Reference behavior: ``ApiSign.sign`` and ``AesUtil.encryptEcb`` from the
user-provided ``openapi-sdk-java-20240730`` archive. The archive's logging
statements are deliberately not copied because they print signing inputs.
"""

from __future__ import annotations

import base64
import hashlib
import json
from collections.abc import Mapping
from urllib.parse import quote_plus
from typing import Any

from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes


class LingxingSignatureError(ValueError):
    """An input cannot be serialized according to the reference SDK safely."""


def _serialize_sdk_value(value: Any, *, name: str) -> str:
    # ApiSign uses String.valueOf(value).trim(). Restrict accepted values to
    # deterministic scalar forms; Java Map/List toString ordering is not stable.
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, int):
        return str(value)
    raise LingxingSignatureError(
        f"签名参数 {name} 必须是字符串、布尔值或整数；复杂/空值序列化不确定，已阻止请求。"
    )


def normalize_signing_parameters(parameters: Mapping[str, Any]) -> dict[str, str]:
    """Normalize simple query values using the Java SDK's String.valueOf/trim behavior."""
    if not isinstance(parameters, Mapping):
        raise LingxingSignatureError("签名参数必须是键值映射。")
    names = list(parameters)
    if any(not isinstance(name, str) or not name or not name.isascii() for name in names):
        raise LingxingSignatureError("签名参数名必须是非空 ASCII 字符串。")
    normalized: dict[str, str] = {}
    for name in sorted(names):
        if name.casefold() == "sign":
            raise LingxingSignatureError("sign 字段不能参与自身签名。")
        normalized[name] = _serialize_sdk_value(parameters[name], name=name)
    return normalized


def canonical_signing_string(parameters: Mapping[str, Any]) -> str:
    """Return the SDK's sorted ``key=value&...`` string before MD5."""
    normalized = normalize_signing_parameters(parameters)
    return "&".join(f"{name}={value}" for name, value in normalized.items())


def generate_signature(parameters: Mapping[str, Any], app_id: str) -> str:
    """Create the Base64 AES-ECB/PKCS5 signature used by the Java SDK sample.

    The SDK sample passes AppID to ``ApiSign.sign`` as its AES key, despite the
    method parameter being named ``appSecret``. It hashes the UTF-8 canonical
    string with MD5, uppercases the hex digest, then encrypts those ASCII bytes.
    """
    if not isinstance(app_id, str) or not app_id or not app_id.isascii():
        raise LingxingSignatureError("AppID 必须是非空 ASCII 字符串。")
    key = app_id.encode("utf-8")
    if len(key) not in {16, 24, 32}:
        raise LingxingSignatureError("AppID 字节长度不符合 AES 密钥长度要求；未发送请求。")

    canonical = canonical_signing_string(parameters)
    digest = hashlib.md5(canonical.encode("utf-8"), usedforsecurity=False).hexdigest().upper()
    padder = padding.PKCS7(algorithms.AES.block_size).padder()
    padded = padder.update(digest.encode("ascii")) + padder.finalize()
    encryptor = Cipher(algorithms.AES(key), modes.ECB()).encryptor()
    ciphertext = encryptor.update(padded) + encryptor.finalize()
    return base64.b64encode(ciphertext).decode("ascii")


def signed_query_parameters(parameters: Mapping[str, Any], app_id: str) -> dict[str, str]:
    """Normalize the transmitted values and append their SDK-compatible sign."""
    normalized = normalize_signing_parameters(parameters)
    signed = dict(normalized)
    signed["sign"] = generate_signature(normalized, app_id)
    return signed


def signed_post_query_parameters(
    query_parameters: Mapping[str, Any],
    body_parameters: Mapping[str, Any],
    app_id: str,
) -> dict[str, str]:
    """Sign auth query parameters together with JSON POST body fields.

    The supplied Java SDK POST examples merge query/body maps before signing.
    For collection-valued body fields, its examples first JSON-serialize the
    collection for signing, then send the collection itself in the JSON body.
    """
    normalized_query = normalize_signing_parameters(query_parameters)
    if not isinstance(body_parameters, Mapping):
        raise LingxingSignatureError("POST JSON body 必须是键值映射。")
    signed_values: dict[str, Any] = dict(normalized_query)
    for name, value in body_parameters.items():
        if not isinstance(name, str) or not name or not name.isascii():
            raise LingxingSignatureError("POST body 参数名必须是非空 ASCII 字符串。")
        if name in signed_values:
            raise LingxingSignatureError(f"POST body 与签名查询参数重名：{name}。")
        if isinstance(value, (list, dict)):
            try:
                signed_values[name] = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
            except (TypeError, ValueError) as exc:
                raise LingxingSignatureError(f"POST body 参数 {name} 不能按 SDK 规则序列化。") from exc
        else:
            # Validate scalar serialization using the same behavior as ApiSign.
            signed_values[name] = _serialize_sdk_value(value, name=name)

    return {
        **normalized_query,
        "sign": generate_signature(signed_values, app_id),
    }


def _java_urlencode(value: str) -> str:
    # The SDK uses java.net.URLEncoder with UTF-8: spaces become '+', and the
    # form-safe set is alphanumeric plus '.', '-', '*', '_'. Python always
    # leaves '~' unescaped, so encode it after quote_plus to match Java.
    return quote_plus(value, safe="*-._").replace("~", "%7E")


def encode_query_parameters(parameters: Mapping[str, str]) -> str:
    """Encode query parameters using the Java SDK's UTF-8 form encoding."""
    if not isinstance(parameters, Mapping):
        raise LingxingSignatureError("请求参数必须是键值映射。")
    pairs: list[str] = []
    for name, value in parameters.items():
        if not isinstance(name, str) or not name or not name.isascii() or not isinstance(value, str):
            raise LingxingSignatureError("请求参数名/值格式无效，已阻止请求。")
        pairs.append(f"{_java_urlencode(name)}={_java_urlencode(value)}")
    return "&".join(pairs)
