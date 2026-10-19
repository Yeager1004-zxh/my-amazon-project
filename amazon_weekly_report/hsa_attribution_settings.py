"""Local, date-scoped operator confirmations for HSA ad-group ownership.

These records contain identifiers and local SKU labels only. They never store
API credentials or report amounts and are kept outside the repository.
"""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any


class HsaAttributionSettingsError(ValueError):
    """Invalid or inaccessible local HSA attribution settings."""


@dataclass(frozen=True)
class HsaAdGroupSkuMapping:
    store_sid: int = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    local_sku: str = field(repr=False)
    valid_from: date = field(repr=False)
    valid_through: date = field(repr=False)


def hsa_attribution_settings_path() -> Path:
    if os.name == "nt":
        root = os.environ.get("APPDATA")
        if not root:
            raise HsaAttributionSettingsError("Windows 未提供 APPDATA 目录，无法保存 HSA 归属确认。")
        return Path(root) / "MyAmazonProject" / "hsa_ad_group_mappings.json"
    return Path.home() / ".config" / "my-amazon-project" / "hsa_ad_group_mappings.json"


def _sid(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise HsaAttributionSettingsError("HSA 归属确认必须绑定有效的店铺 SID。")
    return value


def _text(value: Any, name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise HsaAttributionSettingsError(f"HSA 归属确认中的 {name} 不能为空。")
    return value.strip()


def _day(value: Any, name: str) -> date:
    if type(value) is date:
        return value
    if not isinstance(value, str) or len(value) != 10:
        raise HsaAttributionSettingsError(f"HSA 归属确认中的 {name} 必须是 YYYY-MM-DD。")
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise HsaAttributionSettingsError(f"HSA 归属确认中的 {name} 必须是有效日期。") from exc


def _validate_mapping(mapping: HsaAdGroupSkuMapping) -> HsaAdGroupSkuMapping:
    if not isinstance(mapping, HsaAdGroupSkuMapping):
        raise HsaAttributionSettingsError("HSA 归属确认格式无效。")
    normalized = HsaAdGroupSkuMapping(
        store_sid=_sid(mapping.store_sid),
        profile_id=_text(mapping.profile_id, "profile_id"),
        campaign_id=_text(mapping.campaign_id, "campaign_id"),
        ad_group_id=_text(mapping.ad_group_id, "ad_group_id"),
        local_sku=_text(mapping.local_sku, "local_sku"),
        valid_from=_day(mapping.valid_from, "valid_from"),
        valid_through=_day(mapping.valid_through, "valid_through"),
    )
    if normalized.valid_through < normalized.valid_from:
        raise HsaAttributionSettingsError("HSA 归属确认的有效结束日期早于开始日期。")
    return normalized


def _overlaps(left: HsaAdGroupSkuMapping, right: HsaAdGroupSkuMapping) -> bool:
    return left.valid_from <= right.valid_through and right.valid_from <= left.valid_through


def _same_group(left: HsaAdGroupSkuMapping, right: HsaAdGroupSkuMapping) -> bool:
    return (
        left.store_sid == right.store_sid
        and left.profile_id.casefold() == right.profile_id.casefold()
        and left.campaign_id.casefold() == right.campaign_id.casefold()
        and left.ad_group_id.casefold() == right.ad_group_id.casefold()
    )


def _check_conflicts(mappings: list[HsaAdGroupSkuMapping]) -> None:
    for index, left in enumerate(mappings):
        for right in mappings[index + 1:]:
            if (
                _same_group(left, right)
                and left.local_sku.casefold() != right.local_sku.casefold()
                and _overlaps(left, right)
            ):
                raise HsaAttributionSettingsError(
                    "同一店铺、Profile、Campaign 和 Ad Group 的有效期重叠且指向不同 SKU；已阻止保存/使用。"
                )


def load_hsa_ad_group_mappings(
    path: Path | None = None,
) -> tuple[HsaAdGroupSkuMapping, ...]:
    settings_path = path or hsa_attribution_settings_path()
    if not settings_path.is_file():
        return ()
    try:
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise HsaAttributionSettingsError("本地 HSA 归属设置无法读取；未覆盖现有设置。") from exc
    if not isinstance(payload, dict) or payload.get("version") != 1 or not isinstance(payload.get("mappings"), list):
        raise HsaAttributionSettingsError("本地 HSA 归属设置结构无效。")
    mappings: list[HsaAdGroupSkuMapping] = []
    for row in payload["mappings"]:
        if not isinstance(row, dict):
            raise HsaAttributionSettingsError("本地 HSA 归属设置包含无效记录。")
        mappings.append(_validate_mapping(HsaAdGroupSkuMapping(
            store_sid=row.get("store_sid"),
            profile_id=row.get("profile_id"),
            campaign_id=row.get("campaign_id"),
            ad_group_id=row.get("ad_group_id"),
            local_sku=row.get("local_sku"),
            valid_from=_day(row.get("valid_from"), "valid_from"),
            valid_through=_day(row.get("valid_through"), "valid_through"),
        )))
    _check_conflicts(mappings)
    return tuple(mappings)


def save_hsa_ad_group_mapping(
    mapping: HsaAdGroupSkuMapping,
    *,
    path: Path | None = None,
) -> Path:
    """Atomically save an explicit operator-confirmed mapping, rejecting conflicts."""
    normalized = _validate_mapping(mapping)
    settings_file = path or hsa_attribution_settings_path()
    existing = list(load_hsa_ad_group_mappings(settings_file))
    # Re-saving the exact same record is idempotent.
    if normalized not in existing:
        candidate = [*existing, normalized]
        _check_conflicts(candidate)
        existing = candidate
    payload = {
        "version": 1,
        "mappings": [
            {
                "store_sid": item.store_sid,
                "profile_id": item.profile_id,
                "campaign_id": item.campaign_id,
                "ad_group_id": item.ad_group_id,
                "local_sku": item.local_sku,
                "valid_from": item.valid_from.isoformat(),
                "valid_through": item.valid_through.isoformat(),
            }
            for item in sorted(
                existing,
                key=lambda item: (
                    item.store_sid, item.profile_id.casefold(), item.campaign_id.casefold(),
                    item.ad_group_id.casefold(), item.valid_from, item.valid_through,
                ),
            )
        ],
    }
    settings_file.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix="hsa-mappings-", suffix=".tmp", dir=settings_file.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_name, settings_file)
    finally:
        try:
            os.unlink(temp_name)
        except FileNotFoundError:
            pass
    return settings_file
