"""Local, date-scoped operator confirmations for HSA ad-group ownership.

These records contain identifiers and local SKU labels only. They never store
API credentials or report amounts and are kept outside the repository.
"""

from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
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


@dataclass(frozen=True)
class HsaCreativeRelationSnapshot:
    """A timestamped current-directory observation, never a historical interval."""

    store_sid: int = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    ad_creative_id: str = field(repr=False)
    asins: tuple[str, ...] = field(repr=False)
    observed_at: datetime = field(repr=False)
    state: str = field(default="", repr=False)
    serving_status: str = field(default="", repr=False)
    creative_type: str = field(default="", repr=False)


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


def _observed_at(value: Any) -> datetime:
    if not isinstance(value, datetime) or value.tzinfo is None or value.utcoffset() is None:
        raise HsaAttributionSettingsError("HSA 关系快照必须包含带时区的 observed_at。")
    return value.astimezone(timezone.utc)


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


def _empty_settings() -> dict[str, Any]:
    return {"version": 2, "mappings": [], "snapshots": []}


def _read_settings_payload(settings_path: Path) -> dict[str, Any]:
    if not settings_path.is_file():
        return _empty_settings()
    try:
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise HsaAttributionSettingsError("本地 HSA 归属设置无法读取；未覆盖现有设置。") from exc
    if not isinstance(payload, dict) or payload.get("version") not in (1, 2):
        raise HsaAttributionSettingsError("本地 HSA 归属设置结构无效。")
    if not isinstance(payload.get("mappings"), list):
        raise HsaAttributionSettingsError("本地 HSA 归属设置中的 mappings 必须是数组。")
    snapshots = payload.get("snapshots", [])
    if payload.get("version") == 2 and not isinstance(snapshots, list):
        raise HsaAttributionSettingsError("本地 HSA 关系快照必须是数组。")
    return {"version": 2, "mappings": payload["mappings"], "snapshots": snapshots}


def _atomic_write_settings(settings_path: Path, payload: dict[str, Any]) -> None:
    settings_path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix="hsa-settings-", suffix=".tmp", dir=settings_path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp_name, settings_path)
    finally:
        try:
            os.unlink(temp_name)
        except FileNotFoundError:
            pass


def load_hsa_ad_group_mappings(
    path: Path | None = None,
) -> tuple[HsaAdGroupSkuMapping, ...]:
    settings_path = path or hsa_attribution_settings_path()
    payload = _read_settings_payload(settings_path)
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


def save_hsa_ad_group_mappings(
    mappings: list[HsaAdGroupSkuMapping] | tuple[HsaAdGroupSkuMapping, ...],
    *,
    path: Path | None = None,
) -> Path:
    """Atomically save a reviewed batch of operator-confirmed mappings."""
    if not isinstance(mappings, (list, tuple)) or not mappings:
        raise HsaAttributionSettingsError("至少需要一条经过核验的 HSA 广告组归属。")
    normalized_mappings = [_validate_mapping(mapping) for mapping in mappings]
    settings_file = path or hsa_attribution_settings_path()
    payload = _read_settings_payload(settings_file)
    existing = list(load_hsa_ad_group_mappings(settings_file))
    load_hsa_creative_relation_snapshots(settings_file)
    candidate = list(existing)
    for normalized in normalized_mappings:
        if normalized not in candidate:
            candidate.append(normalized)
    _check_conflicts(candidate)
    payload = {
        "version": 2,
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
                candidate,
                key=lambda item: (
                    item.store_sid, item.profile_id.casefold(), item.campaign_id.casefold(),
                    item.ad_group_id.casefold(), item.valid_from, item.valid_through,
                ),
            )
        ],
        "snapshots": payload["snapshots"],
    }
    _atomic_write_settings(settings_file, payload)
    return settings_file


def save_hsa_ad_group_mapping(
    mapping: HsaAdGroupSkuMapping,
    *,
    path: Path | None = None,
) -> Path:
    """Backward-compatible single-row wrapper around the batch saver."""
    return save_hsa_ad_group_mappings([mapping], path=path)


def record_hsa_creative_catalog_snapshot(
    store_sid: int,
    catalog_rows: list[Any] | tuple[Any, ...],
    *,
    observed_at: datetime | None = None,
    path: Path | None = None,
) -> tuple[Path, int]:
    """Persist one observation of the HSA catalog without inferring date coverage.

    Only exact identifiers, top-level ASINs and state metadata are kept. Creative
    names, request details, credentials and report amounts are omitted.
    """
    sid = _sid(store_sid)
    if not isinstance(catalog_rows, (list, tuple)) or not catalog_rows:
        return (path or hsa_attribution_settings_path(), 0)
    timestamp = _observed_at(observed_at or datetime.now(timezone.utc))
    snapshots: list[dict[str, Any]] = []
    for row in catalog_rows:
        try:
            profile_id = _text(row.profile_id, "profile_id")
            campaign_id = _text(row.campaign_id, "campaign_id")
            ad_group_id = _text(row.ad_group_id, "ad_group_id")
            ad_creative_id = _text(row.ad_creative_id, "ad_creative_id")
            asins = row.asins
            state = getattr(row, "state", "")
            serving_status = getattr(row, "serving_status", "")
            creative_type = getattr(row, "creative_type", "")
        except AttributeError as exc:
            raise HsaAttributionSettingsError("HSA 创意目录记录缺少快照所需的 ID 字段。") from exc
        if not isinstance(asins, (tuple, list)) or any(not isinstance(asin, str) or not asin.strip() for asin in asins):
            raise HsaAttributionSettingsError("HSA 关系快照 ASIN 必须是非空文本数组。")
        if any(not isinstance(value, str) for value in (state, serving_status, creative_type)):
            raise HsaAttributionSettingsError("HSA 关系快照状态字段类型无效。")
        snapshots.append({
            "store_sid": sid,
            "profile_id": profile_id,
            "campaign_id": campaign_id,
            "ad_group_id": ad_group_id,
            "ad_creative_id": ad_creative_id,
            "asins": sorted({asin.strip() for asin in asins}, key=str.casefold),
            "state": state,
            "serving_status": serving_status,
            "creative_type": creative_type,
            "observed_at": timestamp.isoformat(),
        })
    settings_file = path or hsa_attribution_settings_path()
    payload = _read_settings_payload(settings_file)
    load_hsa_ad_group_mappings(settings_file)
    load_hsa_creative_relation_snapshots(settings_file)
    known = {
        (
            item.get("store_sid"), item.get("profile_id"), item.get("campaign_id"),
            item.get("ad_group_id"), item.get("ad_creative_id"), item.get("observed_at"),
        )
        for item in payload["snapshots"] if isinstance(item, dict)
    }
    new_rows: list[dict[str, Any]] = []
    for item in snapshots:
        key = (
            item["store_sid"], item["profile_id"], item["campaign_id"], item["ad_group_id"],
            item["ad_creative_id"], item["observed_at"],
        )
        if key not in known:
            new_rows.append(item)
            known.add(key)
    payload["version"] = 2
    payload["snapshots"].extend(new_rows)
    if len(payload["snapshots"]) > 20_000:
        payload["snapshots"] = payload["snapshots"][-20_000:]
    _atomic_write_settings(settings_file, payload)
    return settings_file, len(new_rows)


def load_hsa_creative_relation_snapshots(
    path: Path | None = None,
) -> tuple[HsaCreativeRelationSnapshot, ...]:
    """Read local observations for diagnostics; never converts them to mappings."""
    payload = _read_settings_payload(path or hsa_attribution_settings_path())
    result: list[HsaCreativeRelationSnapshot] = []
    for row in payload["snapshots"]:
        if not isinstance(row, dict):
            raise HsaAttributionSettingsError("本地 HSA 快照设置包含无效记录。")
        try:
            observed = datetime.fromisoformat(row.get("observed_at"))
            asins = row.get("asins")
            if not isinstance(asins, list) or any(not isinstance(asin, str) or not asin.strip() for asin in asins):
                raise ValueError("asins")
            result.append(HsaCreativeRelationSnapshot(
                store_sid=_sid(row.get("store_sid")),
                profile_id=_text(row.get("profile_id"), "profile_id"),
                campaign_id=_text(row.get("campaign_id"), "campaign_id"),
                ad_group_id=_text(row.get("ad_group_id"), "ad_group_id"),
                ad_creative_id=_text(row.get("ad_creative_id"), "ad_creative_id"),
                asins=tuple(asins),
                observed_at=_observed_at(observed),
                state=_text(row.get("state", ""), "state") if row.get("state") else "",
                serving_status=_text(row.get("serving_status", ""), "serving_status") if row.get("serving_status") else "",
                creative_type=_text(row.get("creative_type", ""), "creative_type") if row.get("creative_type") else "",
            ))
        except (TypeError, ValueError, OverflowError) as exc:
            raise HsaAttributionSettingsError("本地 HSA 关系快照结构或时间戳无效。") from exc
    return tuple(result)
