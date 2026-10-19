"""Exact-identifier advertising cost attribution for preview and validation.

This module deliberately does not infer product ownership from campaign names,
attributed sales, or proportional splits. Provider adapters must normalize only
documented product-ad or explicitly verified allocation rows into these types.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Iterable, Mapping, Sequence

from .hsa_attribution_settings import HsaAdGroupSkuMapping


@dataclass(frozen=True)
class AttributionMapping:
    """Store-scoped local SKU identities; optional validity dates are inclusive."""

    local_sku: str = field(repr=False)
    seller_skus: tuple[str, ...] = field(default=(), repr=False)
    asins: tuple[str, ...] = field(default=(), repr=False)
    valid_from: date | None = field(default=None, repr=False)
    valid_through: date | None = field(default=None, repr=False)
    validity_verified: bool = field(default=False, repr=False)


@dataclass(frozen=True)
class AdCostEvidence:
    """One product-level cost or explicitly verified SKU allocation.

    ``allocation_verified`` must be set only by an adapter whose official
    contract establishes that ``cost`` is already an allocated amount. A raw
    percentage or campaign spend must not be passed as an allocated cost.
    """

    ad_type: str = field(repr=False)
    store_sid: int = field(repr=False)
    report_date: date = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    cost: Decimal = field(repr=False)
    ad_group_id: str = field(default="", repr=False)
    ad_id: str = field(default="", repr=False)
    seller_sku: str = field(default="", repr=False)
    asin: str = field(default="", repr=False)
    source: str = field(default="product_report", repr=False)
    allocation_verified: bool = field(default=False, repr=False)


@dataclass(frozen=True)
class CampaignCostEvidence:
    """Activity-level cost used only to reconcile product/allocation lines."""

    ad_type: str = field(repr=False)
    store_sid: int = field(repr=False)
    report_date: date = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    cost: Decimal = field(repr=False)


@dataclass(frozen=True)
class AdGroupCostEvidence:
    ad_type: str = field(repr=False)
    store_sid: int = field(repr=False)
    report_date: date = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    cost: Decimal = field(repr=False)


@dataclass(frozen=True)
class PromotionTargetEvidence:
    """Exact product identity promoted by an ad in a specific ad group."""

    ad_type: str = field(repr=False)
    store_sid: int = field(repr=False)
    report_date: date = field(repr=False)
    profile_id: str = field(repr=False)
    campaign_id: str = field(repr=False)
    ad_group_id: str = field(repr=False)
    ad_id: str = field(default="", repr=False)
    seller_sku: str = field(default="", repr=False)
    asin: str = field(default="", repr=False)
    ad_creative_id: str = field(default="", repr=False)
    # Current catalog identities are not historical proof. Set this only when
    # the official source establishes that the identity was valid on report_date.
    identity_valid_for_report: bool = field(default=False, repr=False)


@dataclass(frozen=True)
class ExclusiveAdGroupAttribution:
    evidence: AdGroupCostEvidence = field(repr=False)
    status: str = field(repr=False)
    local_sku: str = field(default="", repr=False)


@dataclass(frozen=True)
class AdGroupAttributionResult:
    rows: tuple[ExclusiveAdGroupAttribution, ...]
    reconciled_campaigns: frozenset[tuple[str, int, date, str, str]]
    blocked_campaigns: frozenset[tuple[str, int, date, str, str]]

    def daily_costs(self) -> Mapping[tuple[str, date, str], Decimal]:
        totals: dict[tuple[str, date, str], Decimal] = defaultdict(lambda: Decimal("0"))
        for row in self.rows:
            if row.local_sku:
                key = (row.evidence.ad_type, row.evidence.report_date, row.local_sku.casefold())
                totals[key] += row.evidence.cost
        return dict(totals)


@dataclass(frozen=True)
class AttributedAdCost:
    evidence: AdCostEvidence = field(repr=False)
    status: str = field(repr=False)
    local_sku: str = field(default="", repr=False)
    matched_by: str = field(default="", repr=False)
    campaign_reconciled: bool = field(default=False, repr=False)
    formal_ready: bool = field(default=False, repr=False)


@dataclass(frozen=True)
class AttributionResult:
    rows: tuple[AttributedAdCost, ...]
    reconciled_campaigns: frozenset[tuple[str, int, date, str, str]]
    blocked_campaigns: frozenset[tuple[str, int, date, str, str]]

    def daily_costs(self) -> Mapping[tuple[str, date, str], Decimal]:
        totals: dict[tuple[str, date, str], Decimal] = defaultdict(lambda: Decimal("0"))
        for row in self.rows:
            if row.local_sku and row.campaign_reconciled:
                key = (row.evidence.ad_type, row.evidence.report_date, row.local_sku.casefold())
                totals[key] += row.evidence.cost
        return dict(totals)


def _norm(value: str) -> str:
    return value.strip().casefold()


def _validate_cost(value: Decimal, label: str) -> None:
    if not isinstance(value, Decimal) or not value.is_finite() or value < 0:
        raise ValueError(f"{label} cost must be a finite, non-negative Decimal.")


def _campaign_key(ad_type: str, sid: int, report_day: date, profile_id: str, campaign_id: str) -> tuple[str, int, date, str, str]:
    return (ad_type.upper(), sid, report_day, _norm(profile_id), _norm(campaign_id))


def _identity_indexes(
    mappings: Sequence[AttributionMapping],
) -> tuple[
    dict[str, str],
    dict[str, list[tuple[str, date | None, date | None]]],
    dict[str, list[tuple[str, date | None, date | None]]],
]:
    by_sku: dict[str, str] = {}
    msku_owners: dict[str, list[tuple[str, date | None, date | None]]] = defaultdict(list)
    asin_owners: dict[str, list[tuple[str, date | None, date | None]]] = defaultdict(list)
    for mapping in mappings:
        if not isinstance(mapping.local_sku, str) or not mapping.local_sku.strip():
            raise ValueError("local SKU mapping must be non-empty text.")
        sku_key = _norm(mapping.local_sku)
        by_sku.setdefault(sku_key, mapping.local_sku.strip())
        if mapping.valid_through and mapping.valid_from and mapping.valid_through < mapping.valid_from:
            raise ValueError("mapping validity end precedes its start.")
        if (mapping.valid_from or mapping.valid_through) and not mapping.validity_verified:
            continue
        local = by_sku[sku_key]
        for seller_sku in mapping.seller_skus:
            if isinstance(seller_sku, str) and seller_sku.strip():
                msku_owners[_norm(seller_sku)].append((local, mapping.valid_from, mapping.valid_through))
        for asin in mapping.asins:
            if isinstance(asin, str) and asin.strip():
                asin_owners[_norm(asin)].append((local, mapping.valid_from, mapping.valid_through))
    return by_sku, msku_owners, asin_owners


def _resolve_identity(
    *,
    report_day: date,
    seller_sku: str,
    asin: str,
    msku_owners: Mapping[str, list[tuple[str, date | None, date | None]]],
    asin_owners: Mapping[str, list[tuple[str, date | None, date | None]]],
) -> tuple[str, str, str]:
    def owners(index: Mapping[str, list[tuple[str, date | None, date | None]]], key: str) -> set[str]:
        return {
            local
            for local, valid_from, valid_through in index.get(key, ())
            if (valid_from is None or report_day >= valid_from)
            and (valid_through is None or report_day <= valid_through)
        }

    msku_key = _norm(seller_sku) if seller_sku.strip() else ""
    asin_key = _norm(asin) if asin.strip() else ""
    msku_matches = owners(msku_owners, msku_key) if msku_key else set()
    asin_matches = owners(asin_owners, asin_key) if asin_key else set()
    if not msku_key and not asin_key:
        return "", "", "缺少 MSKU/ASIN 推广商品标识，待确认"
    if len(msku_matches) > 1 or len(asin_matches) > 1:
        return "", "", "推广商品标识映射到多个本地 SKU，存在冲突"
    matches = msku_matches | asin_matches
    if len(matches) > 1:
        return "", "", "MSKU 与 ASIN 指向不同本地 SKU，存在冲突"
    if not matches:
        return "", "", "没有匹配到已查询的本地 SKU，保留为未归属"
    matched_by = "+".join(
        label for label, values in (("MSKU", msku_matches), ("ASIN", asin_matches)) if values
    )
    return next(iter(matches)), matched_by, ""


def attribute_exclusive_ad_group_costs(
    *,
    store_sid: int,
    mappings: Sequence[AttributionMapping],
    promoted_targets: Iterable[PromotionTargetEvidence],
    ad_group_costs: Iterable[AdGroupCostEvidence],
    campaign_costs: Iterable[CampaignCostEvidence],
    promotion_catalog_complete: bool,
    confirmed_mappings: Sequence[HsaAdGroupSkuMapping] = (),
) -> AdGroupAttributionResult:
    """Attribute HSA group spend from historical product evidence or confirmation.

    Current HSA creative catalogs are not historical proof. A local,
    date-scoped operator confirmation may establish a group-to-SKU mapping;
    otherwise only date-valid, complete product-target evidence can authorize
    attribution. Campaign costs reconcile group costs and are never added.
    """
    if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
        raise ValueError("store_sid must be a positive integer.")
    _by_sku, msku_owners, asin_owners = _identity_indexes(mappings)
    groups = list(ad_group_costs)
    targets = list(promoted_targets)
    campaign_rows = list(campaign_costs)
    campaign_amounts: dict[tuple[str, int, date, str, str], Decimal] = {}
    duplicate_campaigns: set[tuple[str, int, date, str, str]] = set()
    for campaign in campaign_rows:
        if campaign.store_sid != store_sid:
            continue
        _validate_cost(campaign.cost, "campaign")
        key = _campaign_key(campaign.ad_type, campaign.store_sid, campaign.report_date, campaign.profile_id, campaign.campaign_id)
        if key in campaign_amounts:
            duplicate_campaigns.add(key)
        else:
            campaign_amounts[key] = campaign.cost

    group_amounts: dict[tuple[str, int, date, str, str, str], Decimal] = {}
    duplicate_groups: set[tuple[str, int, date, str, str, str]] = set()
    group_campaign_totals: dict[tuple[str, int, date, str, str], Decimal] = defaultdict(lambda: Decimal("0"))
    for group in groups:
        if group.store_sid != store_sid:
            continue
        _validate_cost(group.cost, "ad group")
        campaign_key = _campaign_key(group.ad_type, group.store_sid, group.report_date, group.profile_id, group.campaign_id)
        group_key = (*campaign_key, _norm(group.ad_group_id))
        if group_key in group_amounts:
            duplicate_groups.add(group_key)
            continue
        group_amounts[group_key] = group.cost
        group_campaign_totals[campaign_key] += group.cost

    reconciled: set[tuple[str, int, date, str, str]] = set()
    blocked: set[tuple[str, int, date, str, str]] = set(duplicate_campaigns)
    all_campaign_keys = set(campaign_amounts) | set(group_campaign_totals) | duplicate_campaigns
    for campaign_key in all_campaign_keys:
        keys_for_campaign = [key for key in group_amounts if key[:5] == campaign_key]
        duplicate = any(key[:5] == campaign_key for key in duplicate_groups)
        if (
            campaign_key not in campaign_amounts
            or not keys_for_campaign
            or duplicate
            or campaign_key in duplicate_campaigns
            or group_campaign_totals.get(campaign_key) != campaign_amounts.get(campaign_key)
        ):
            blocked.add(campaign_key)
        else:
            reconciled.add(campaign_key)

    targets_by_group: dict[tuple[str, int, date, str, str, str], list[PromotionTargetEvidence]] = defaultdict(list)
    duplicate_target_groups: set[tuple[str, int, date, str, str, str]] = set()
    seen_targets: set[tuple[object, ...]] = set()
    target_bindings_by_ad_id: dict[tuple[object, ...], set[tuple[tuple[str, int, date, str, str, str], str, str]]] = defaultdict(set)
    for target in targets:
        if target.store_sid != store_sid:
            continue
        key = _campaign_key(target.ad_type, target.store_sid, target.report_date, target.profile_id, target.campaign_id)
        group_key = (*key, _norm(target.ad_group_id))
        target_identity = (
            *group_key, _norm(target.ad_id), _norm(target.ad_creative_id), _norm(target.seller_sku), _norm(target.asin),
        )
        if target_identity in seen_targets:
            duplicate_target_groups.add(group_key)
        seen_targets.add(target_identity)
        targets_by_group[group_key].append(target)
        if target.ad_id.strip():
            ad_identity = (target.ad_type.upper(), target.store_sid, target.report_date, _norm(target.profile_id), "ad", _norm(target.ad_id))
            target_bindings_by_ad_id[ad_identity].add((group_key, _norm(target.seller_sku), _norm(target.asin)))
        if target.ad_creative_id.strip():
            creative_identity = (target.ad_type.upper(), target.store_sid, target.report_date, _norm(target.profile_id), "creative", _norm(target.ad_creative_id))
            target_bindings_by_ad_id[creative_identity].add((group_key, _norm(target.seller_sku), _norm(target.asin)))
    for bindings in target_bindings_by_ad_id.values():
        if len(bindings) > 1:
            duplicate_target_groups.update(binding[0] for binding in bindings)

    confirmed_by_group: dict[tuple[str, str, str, int, date, str], set[str]] = defaultdict(set)
    for mapping in confirmed_mappings:
        if mapping.store_sid != store_sid:
            continue
        if mapping.valid_through < mapping.valid_from:
            raise ValueError("HSA 归属确认的有效结束日期早于开始日期。")
        for group in groups:
            if (
                group.store_sid == store_sid
                and _norm(group.profile_id) == _norm(mapping.profile_id)
                and _norm(group.campaign_id) == _norm(mapping.campaign_id)
                and _norm(group.ad_group_id) == _norm(mapping.ad_group_id)
                and mapping.valid_from <= group.report_date <= mapping.valid_through
            ):
                confirmed_by_group[
                    (group.ad_type.upper(), _norm(group.profile_id), _norm(group.campaign_id), store_sid, group.report_date)
                    + (_norm(group.ad_group_id),)
                ].add(mapping.local_sku)

    result: list[ExclusiveAdGroupAttribution] = []
    for group in groups:
        if group.store_sid != store_sid:
            result.append(ExclusiveAdGroupAttribution(group, status="SID 与所选店铺不一致，阻止归属"))
            continue
        group_key = (*_campaign_key(group.ad_type, group.store_sid, group.report_date, group.profile_id, group.campaign_id), _norm(group.ad_group_id))
        campaign_key = group_key[:5]
        if group_key in duplicate_groups:
            result.append(ExclusiveAdGroupAttribution(group, status="广告组重复，阻止归属"))
            continue
        if campaign_key not in reconciled:
            result.append(ExclusiveAdGroupAttribution(group, status="活动成本与广告组汇总未核平，阻止归属"))
            continue
        confirmation_key = (
            group.ad_type.upper(), _norm(group.profile_id), _norm(group.campaign_id),
            store_sid, group.report_date, _norm(group.ad_group_id),
        )
        confirmed_owners = confirmed_by_group.get(confirmation_key, set())
        if len(confirmed_owners) > 1:
            result.append(ExclusiveAdGroupAttribution(group, status="本地有效期确认记录冲突，阻止归属"))
            continue
        if len(confirmed_owners) == 1:
            confirmed_sku = next(iter(confirmed_owners))
            # A source explicitly marked historically valid can veto a local
            # confirmation if it proves a different product. Current snapshots
            # remain diagnostic and do not override the operator's dated record.
            historical_targets = [
                target for target in targets_by_group.get(group_key, ())
                if target.identity_valid_for_report
            ]
            resolved_historical: set[str] = set()
            historical_error = ""
            for target in historical_targets:
                resolved, _matched_by, error = _resolve_identity(
                    report_day=target.report_date,
                    seller_sku=target.seller_sku,
                    asin=target.asin,
                    msku_owners=msku_owners,
                    asin_owners=asin_owners,
                )
                if error:
                    historical_error = error
                    break
                resolved_historical.add(resolved)
            if historical_error or (resolved_historical and resolved_historical != {confirmed_sku.casefold()}):
                result.append(ExclusiveAdGroupAttribution(
                    group, status=historical_error or "有效期内的历史商品记录与本地确认不一致，阻止归属"
                ))
                continue
            result.append(ExclusiveAdGroupAttribution(
                group,
                local_sku=confirmed_sku,
                status="按运营确认的店铺/Profile/Campaign/Ad Group 与有效期归属；活动总额仅用于核对",
            ))
            continue
        if not promotion_catalog_complete:
            result.append(ExclusiveAdGroupAttribution(group, status="商品广告目录完整性未核验，不能推定广告组只推广单一 SKU"))
            continue
        if group_key in duplicate_target_groups:
            result.append(ExclusiveAdGroupAttribution(group, status="推广商品目录含重复记录，阻止广告组归属"))
            continue
        group_targets = targets_by_group.get(group_key, ())
        if not group_targets:
            result.append(ExclusiveAdGroupAttribution(group, status="该广告组没有可核验的推广商品记录"))
            continue
        if any(not target.identity_valid_for_report for target in group_targets):
            result.append(ExclusiveAdGroupAttribution(group, status="商品目录仅为当前快照，未证明在报表日有效；费用待确认"))
            continue
        resolved: set[str] = set()
        failure = ""
        for target in group_targets:
            if target.ad_type.upper() != group.ad_type.upper() or target.report_date != group.report_date:
                failure = "推广商品记录与广告组日期/类型不一致"
                break
            local, _matched_by, error = _resolve_identity(
                report_day=target.report_date,
                seller_sku=target.seller_sku,
                asin=target.asin,
                msku_owners=msku_owners,
                asin_owners=asin_owners,
            )
            if error:
                failure = error
                break
            resolved.add(local)
        if failure:
            result.append(ExclusiveAdGroupAttribution(group, status=failure))
        elif len(resolved) != 1:
            result.append(ExclusiveAdGroupAttribution(group, status="同一广告组推广多个本地 SKU，整组费用保留未归属"))
        else:
            result.append(ExclusiveAdGroupAttribution(
                group, local_sku=next(iter(resolved)), status="广告组商品目录完整且目标唯一；已按广告组实际花费自动归属"
            ))

    return AdGroupAttributionResult(tuple(result), frozenset(reconciled), frozenset(blocked))


def attribute_ad_costs(
    *,
    store_sid: int,
    mappings: Sequence[AttributionMapping],
    product_costs: Iterable[AdCostEvidence],
    campaign_costs: Iterable[CampaignCostEvidence],
    currency_and_date_basis_verified: bool = False,
    mapping_temporal_scope_verified: bool = False,
) -> AttributionResult:
    """Resolve product spend by exact IDs and reconcile each campaign separately.

    MSKU and ASIN indexes are store scoped by ``store_sid``. If both identifiers
    are known, they must resolve to the same local SKU. A campaign mismatch
    blocks only that campaign, so independent reconciled campaigns remain
    usable in the preview. Missing rows never become zero values.
    """
    if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
        raise ValueError("store_sid must be a positive integer.")

    # Index aliases with their validity windows. Date-scoped aliases are
    # considered only when their validity came from a verified source.
    _by_sku, msku_owners, asin_owners = _identity_indexes(mappings)

    campaigns: dict[tuple[str, int, date, str, str], Decimal] = {}
    duplicate_campaigns: set[tuple[str, int, date, str, str]] = set()
    for campaign in campaign_costs:
        if campaign.store_sid != store_sid:
            continue
        _validate_cost(campaign.cost, "campaign")
        key = _campaign_key(campaign.ad_type, campaign.store_sid, campaign.report_date, campaign.profile_id, campaign.campaign_id)
        if key in campaigns:
            duplicate_campaigns.add(key)
        else:
            campaigns[key] = campaign.cost

    lines = list(product_costs)
    detail_totals: dict[tuple[str, int, date, str, str], Decimal] = defaultdict(lambda: Decimal("0"))
    invalid_lines: set[int] = set()
    seen_lines: set[tuple[object, ...]] = set()
    duplicate_line_campaigns: set[tuple[str, int, date, str, str]] = set()
    seen_ad_ids: dict[tuple[str, int, date, str, str], tuple[tuple[object, ...], int]] = {}
    for index, line in enumerate(lines):
        _validate_cost(line.cost, "product")
        if line.ad_type.upper() not in {"SP", "HSA", "SD"}:
            invalid_lines.add(index)
            continue
        if line.store_sid != store_sid:
            invalid_lines.add(index)
            continue
        if line.source == "verified_asin_allocation" and (
            not line.allocation_verified or line.ad_type.upper() != "HSA"
        ):
            invalid_lines.add(index)
            continue
        if line.source not in {"product_report", "verified_asin_allocation"}:
            invalid_lines.add(index)
            continue
        identity = (
            line.ad_type.upper(), line.store_sid, line.report_date, _norm(line.profile_id),
            _norm(line.campaign_id), _norm(line.ad_group_id), _norm(line.ad_id),
            _norm(line.seller_sku), _norm(line.asin),
        )
        if identity in seen_lines:
            invalid_lines.add(index)
            duplicate_line_campaigns.add(
                _campaign_key(line.ad_type, line.store_sid, line.report_date, line.profile_id, line.campaign_id)
            )
            continue
        if line.ad_id.strip():
            ad_identity = (line.ad_type.upper(), line.store_sid, line.report_date, _norm(line.profile_id), _norm(line.ad_id))
            prior = seen_ad_ids.get(ad_identity)
            if prior is not None and prior[0] != identity:
                invalid_lines.add(index)
                duplicate_line_campaigns.add(
                    _campaign_key(line.ad_type, line.store_sid, line.report_date, line.profile_id, line.campaign_id)
                )
                prior_line = lines[prior[1]]
                duplicate_line_campaigns.add(
                    _campaign_key(prior_line.ad_type, prior_line.store_sid, prior_line.report_date, prior_line.profile_id, prior_line.campaign_id)
                )
                continue
            seen_ad_ids[ad_identity] = (identity, index)
        seen_lines.add(identity)
        key = _campaign_key(line.ad_type, line.store_sid, line.report_date, line.profile_id, line.campaign_id)
        detail_totals[key] += line.cost

    reconciled: set[tuple[str, int, date, str, str]] = set()
    blocked: set[tuple[str, int, date, str, str]] = set()
    for key in set(campaigns) | set(detail_totals) | duplicate_campaigns | duplicate_line_campaigns:
        if key not in campaigns or key not in detail_totals or key in duplicate_campaigns or campaigns.get(key) != detail_totals.get(key):
            blocked.add(key)
        else:
            reconciled.add(key)
    reconciled.difference_update(duplicate_line_campaigns)
    blocked.update(duplicate_line_campaigns)

    output: list[AttributedAdCost] = []
    for index, line in enumerate(lines):
        campaign_key = _campaign_key(line.ad_type, line.store_sid, line.report_date, line.profile_id, line.campaign_id)
        if line.store_sid != store_sid:
            output.append(AttributedAdCost(line, status="SID 与所选店铺不一致，阻止归属"))
            continue
        if index in invalid_lines:
            output.append(AttributedAdCost(line, status="记录身份/来源无效或重复，阻止归属"))
            continue
        if campaign_key not in reconciled:
            reason = "活动总额与商品/分摊明细未核平，阻止该活动归属"
            output.append(AttributedAdCost(line, status=reason))
            continue

        msku_key = _norm(line.seller_sku) if line.seller_sku.strip() else ""
        asin_key = _norm(line.asin) if line.asin.strip() else ""

        def active_owners(index: Mapping[str, list[tuple[str, date | None, date | None]]], key: str) -> set[str]:
            return {
                local
                for local, valid_from, valid_through in index.get(key, ())
                if (valid_from is None or line.report_date >= valid_from)
                and (valid_through is None or line.report_date <= valid_through)
            }

        msku_matches = active_owners(msku_owners, msku_key) if msku_key else set()
        asin_matches = active_owners(asin_owners, asin_key) if asin_key else set()
        if not msku_key and not asin_key:
            output.append(AttributedAdCost(line, status="缺少 MSKU/ASIN 推广商品标识，待确认", campaign_reconciled=True))
            continue
        if len(msku_matches) > 1 or len(asin_matches) > 1:
            output.append(AttributedAdCost(line, status="推广商品标识映射到多个本地 SKU，存在冲突", campaign_reconciled=True))
            continue
        all_matches = msku_matches | asin_matches
        if len(all_matches) > 1:
            output.append(AttributedAdCost(line, status="MSKU 与 ASIN 指向不同本地 SKU，存在冲突", campaign_reconciled=True))
            continue
        if not all_matches:
            output.append(AttributedAdCost(line, status="没有匹配到已查询的本地 SKU，保留为未归属", campaign_reconciled=True))
            continue

        local_sku = next(iter(all_matches))
        matched_by = "+".join(
            label for label, matches in (("MSKU", msku_matches), ("ASIN", asin_matches)) if matches
        )
        formal_ready = bool(currency_and_date_basis_verified and mapping_temporal_scope_verified)
        status = (
            f"活动成本已核平；{matched_by} 精确匹配并自动归属"
            if formal_ready
            else f"活动成本已核平；{matched_by} 精确匹配并自动归属；币种/日期口径或历史映射时效仍待验证"
        )
        output.append(AttributedAdCost(
            line, local_sku=local_sku, status=status, matched_by=matched_by,
            campaign_reconciled=True, formal_ready=formal_ready,
        ))

    return AttributionResult(tuple(output), frozenset(reconciled), frozenset(blocked))
