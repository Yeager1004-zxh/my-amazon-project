from __future__ import annotations

import json
import os
import tempfile
from datetime import date
from pathlib import Path
from typing import Any


class D01SettingsError(ValueError):
    """Invalid or inaccessible locally confirmed D01 settings."""


def d01_settings_path() -> Path:
    if os.name == "nt":
        root = os.environ.get("APPDATA")
        if not root:
            raise D01SettingsError("Windows 未提供 APPDATA 目录，无法保存 D01 设置。")
        return Path(root) / "MyAmazonProject" / "d01_dates.json"
    return Path.home() / ".config" / "my-amazon-project" / "d01_dates.json"


def _store_key(store_sid: int) -> str:
    if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
        raise D01SettingsError("D01 设置必须绑定有效的领星店铺 SID。")
    return str(store_sid)


def _sku_key(value: Any) -> str:
    if not isinstance(value, str) or not value.strip():
        raise D01SettingsError("D01 设置中包含无效 SKU。")
    return value.strip().casefold()


def load_d01_settings(path: Path | None = None) -> dict[str, dict[str, date]]:
    """Load confirmed dates as ``{sid: {casefolded_sku: date}}``."""
    settings_path = path or d01_settings_path()
    if not settings_path.is_file():
        return {}
    try:
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise D01SettingsError("本地 D01 设置无法读取；未覆盖现有配置。") from exc
    if not isinstance(payload, dict) or payload.get("version") != 1 or not isinstance(payload.get("stores"), dict):
        raise D01SettingsError("本地 D01 设置格式无效；未覆盖现有配置。")
    result: dict[str, dict[str, date]] = {}
    for raw_sid, store_data in payload["stores"].items():
        if not isinstance(raw_sid, str) or not raw_sid.isdecimal() or int(raw_sid) <= 0:
            raise D01SettingsError("本地 D01 设置包含无效店铺标识。")
        if not isinstance(store_data, dict) or not isinstance(store_data.get("sku_dates"), dict):
            raise D01SettingsError("本地 D01 店铺设置结构无效。")
        sku_dates: dict[str, date] = {}
        for raw_sku, raw_date in store_data["sku_dates"].items():
            sku_key = _sku_key(raw_sku)
            if not isinstance(raw_date, str) or len(raw_date) != 10:
                raise D01SettingsError("本地 D01 日期格式无效；要求 YYYY-MM-DD。")
            try:
                parsed_date = date.fromisoformat(raw_date)
            except ValueError as exc:
                raise D01SettingsError("本地 D01 日期格式无效；要求 YYYY-MM-DD。") from exc
            if sku_key in sku_dates:
                raise D01SettingsError("本地 D01 设置中存在重复 SKU。")
            sku_dates[sku_key] = parsed_date
        result[raw_sid] = sku_dates
    return result


def get_d01_dates(
    store_sid: int,
    skus: list[str] | tuple[str, ...] | None = None,
    *,
    path: Path | None = None,
) -> dict[str, date]:
    store_dates = load_d01_settings(path).get(_store_key(store_sid), {})
    if skus is None:
        return dict(store_dates)
    return {key: store_dates[key] for sku in skus if (key := _sku_key(sku)) in store_dates}


def save_d01_dates(store_sid: int, values: dict[str, date], *, path: Path | None = None) -> Path:
    """Atomically merge explicitly confirmed dates for one store."""
    sid = _store_key(store_sid)
    if not isinstance(values, dict) or not values:
        raise D01SettingsError("至少需要一个已输入并确认的 SKU D01 日期。")
    normalized: dict[str, date] = {}
    for sku, day in values.items():
        key = _sku_key(sku)
        if type(day) is not date:
            raise D01SettingsError("D01 日期必须是有效的站点日期。")
        if key in normalized:
            raise D01SettingsError("同一店铺中 SKU 不区分大小写且不能重复。")
        normalized[key] = day

    settings_file = path or d01_settings_path()
    all_dates = load_d01_settings(settings_file)
    all_dates.setdefault(sid, {}).update(normalized)
    payload = {
        "version": 1,
        "stores": {
            stored_sid: {
                "sku_dates": {sku: day.isoformat() for sku, day in sorted(sku_dates.items())}
            }
            for stored_sid, sku_dates in sorted(all_dates.items(), key=lambda item: int(item[0]))
        },
    }
    settings_file.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix="d01-dates-", suffix=".tmp", dir=settings_file.parent)
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
