"""Safe, read-only presentation rows for SP, HSA, and SD ad reports."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from typing import Mapping, Sequence

from .ad_attribution import (
    AdCostEvidence,
    AdGroupCostEvidence,
    AttributionMapping,
    AttributedAdCost,
    CampaignCostEvidence,
    ExclusiveAdGroupAttribution,
    attribute_ad_costs,
    attribute_exclusive_ad_group_costs,
)
from .hsa_attribution_settings import HsaAdGroupSkuMapping
from .providers.lingxing_api import SpAdSpendWindow, SkuProductMapping
from .statistics_scope import SP_SD_AD_TYPES, SP_SD_ONLY_CONFIRMED

LEGACY_AD_DIAGNOSTIC_SCOPE = "LEGACY_SP_HSA_SD_DIAGNOSTIC"


@dataclass(frozen=True)
class AdSpendSummaryPreview:
    report_date: str
    local_sku: str
    sp_cost: str
    hsa_cost: str
    sd_cost: str
    confirmed_total: str
    unattributed_cost: str
    integrity: str


@dataclass(frozen=True)
class AdSpendDetailPreview:
    report_date: str
    ad_type: str
    report_level: str
    local_sku: str
    report_sku: str
    asin: str
    ad_id: str
    profile_id: str
    campaign_id: str
    ad_group_id: str
    cost: str
    integrity: str
    ad_creative_id: str = ""
    candidate_local_skus: tuple[str, ...] = ()
    candidate_resolution_complete: bool = False


@dataclass(frozen=True)
class _WindowAttribution:
    product_rows: tuple[AttributedAdCost, ...] = ()
    ad_group_rows: tuple[ExclusiveAdGroupAttribution, ...] = ()


def _money(value: Decimal | None) -> str:
    return "" if value is None else format(value, "f")


def _type_state(kind: str, window: SpAdSpendWindow | None, *, matched_rows: int = 0) -> str:
    if window is None:
        return f"{kind} 查询结果缺失（未知，不视为零）"
    if matched_rows:
        blocked = len(window.discrepancies)
        suffix = f"；另有 {blocked} 个活动未核平并被隔离" if blocked else ""
        if kind == "HSA":
            return f"HSA 广告组费用按本地有效期确认记录归属{suffix}"
        return f"{kind} 商品费用按 MSKU/ASIN 精确匹配本地 SKU{suffix}"
    if not window.report_complete:
        if window.discrepancies:
            return f"{kind} 活动与明细成本未核平；相关记录不归属"
        if kind == "HSA":
            if window.campaign_rows and not window.ad_group_rows:
                return "HSA 活动报表有记录，广告组报表无记录；SKU 花费未知"
            if window.ad_group_rows and not window.campaign_rows:
                return "HSA 广告组报表有记录，活动报表无记录；金额未核平"
        else:
            if window.campaign_rows and not window.product_rows:
                return f"{kind} 活动报表有记录，商品报表无记录；SKU 花费未知"
            if window.product_rows and not window.campaign_rows:
                return f"{kind} 商品报表有记录，活动报表无记录；金额未核平"
        if not window.campaign_rows and not window.product_rows and not window.ad_group_rows:
            return f"{kind} 报表接口返回空记录；未知，不视为零花费"
        return f"{kind} 活动或明细报表不完整；金额未知"
    if kind == "HSA":
        return "HSA 活动与广告组成本已核平；本 SKU 没有已确认归属记录（不推定为零）"
    return f"{kind} 活动与商品成本已核平；本 SKU 没有匹配记录（不推定为零）"


def _window_state(kind: str, window: SpAdSpendWindow | None) -> str:
    if window is None:
        return f"{kind} 查询缺失"
    if window.report_complete:
        return f"{kind} 活动与{'广告组' if kind == 'HSA' else '商品'}成本已核平"
    if window.discrepancies:
        return f"{kind} 活动与明细成本不一致，金额未知"
    return f"{kind} 活动或明细报表缺失，金额未知"


def _normalized_attribution(
    window: SpAdSpendWindow,
    store_sid: int,
    mappings: Sequence[SkuProductMapping],
    confirmed_hsa_mappings: Sequence[HsaAdGroupSkuMapping],
) -> _WindowAttribution:
    campaigns = [
        CampaignCostEvidence(window.ad_type, store_sid, row.report_date, row.profile_id, row.campaign_id, row.cost)
        for row in window.campaign_rows
    ]
    product_costs = [] if window.ad_type == "HSA" else [
        AdCostEvidence(
            ad_type=window.ad_type,
            store_sid=store_sid,
            report_date=row.report_date,
            profile_id=row.profile_id,
            campaign_id=row.campaign_id,
            ad_group_id=row.ad_group_id,
            ad_id=row.ad_id,
            seller_sku=row.sku,
            asin=row.asin,
            cost=row.cost,
            source=row.source,
            allocation_verified=row.allocation_verified,
        )
        for row in window.product_rows
    ]
    result = attribute_ad_costs(
        store_sid=store_sid,
        mappings=[AttributionMapping(item.local_sku, item.seller_skus, item.asins) for item in mappings],
        product_costs=product_costs,
        campaign_costs=campaigns,
    )
    group_rows: tuple[ExclusiveAdGroupAttribution, ...] = ()
    if window.ad_type == "HSA" and window.ad_group_rows:
        group_result = attribute_exclusive_ad_group_costs(
            store_sid=store_sid,
            mappings=[AttributionMapping(item.local_sku, item.seller_skus, item.asins) for item in mappings],
            promoted_targets=window.promotion_targets,
            ad_group_costs=[
                AdGroupCostEvidence(
                    "HSA", store_sid, row.report_date, row.profile_id, row.campaign_id, row.ad_group_id, row.cost
                )
                for row in window.ad_group_rows
            ],
            campaign_costs=campaigns,
            promotion_catalog_complete=window.promotion_catalog_complete,
            confirmed_mappings=confirmed_hsa_mappings,
        )
        group_rows = group_result.rows
    return _WindowAttribution(tuple(result.rows), group_rows)


def build_ad_spend_preview(
    *,
    windows_by_type: Mapping[str, Sequence[SpAdSpendWindow]],
    mappings: Sequence[SkuProductMapping],
    store_sid: int,
    requested_skus: Sequence[str],
    start_date: date,
    end_date: date,
    confirmed_hsa_mappings: Sequence[HsaAdGroupSkuMapping] = (),
    scope: str = SP_SD_ONLY_CONFIRMED,
    currency_and_date_basis_verified: bool = False,
) -> tuple[list[AdSpendSummaryPreview], list[AdSpendDetailPreview]]:
    """Build a preview without adding activity, group, or product levels twice.

    HSA unites Sponsored Brands and video spend. An HSA group can be assigned
    only from historical target evidence or a local confirmation whose SID,
    profile, campaign, group, and effective dates match the report row.
    Cross-type totals stay blank until currency/date basis is verified.
    """
    if end_date < start_date:
        raise ValueError("广告预览日期范围无效。")
    if scope not in {SP_SD_ONLY_CONFIRMED, LEGACY_AD_DIAGNOSTIC_SCOPE}:
        raise ValueError("广告预览统计范围无效；拒绝调用或汇总未明确配置的广告类型。")
    if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
        raise ValueError("广告预览必须使用所选店铺的有效 SID。")
    kinds = SP_SD_AD_TYPES if scope == SP_SD_ONLY_CONFIRMED else ("SP", "HSA", "SD")
    by_kind_date: dict[str, dict[date, SpAdSpendWindow]] = {}
    outcomes_by_window: dict[tuple[str, date], _WindowAttribution] = {}
    mapped_costs: dict[tuple[str, date, str], Decimal] = {}
    matched_counts: dict[tuple[str, date, str], int] = {}
    unmatched_costs: dict[tuple[str, date], Decimal | None] = {}

    for kind in kinds:
        day_map: dict[date, SpAdSpendWindow] = {}
        for window in windows_by_type.get(kind, ()):
            if window.ad_type != kind or window.start_date != window.end_date:
                raise ValueError(f"{kind} 预览要求单日且类型匹配的报表窗口。")
            if window.start_date in day_map:
                raise ValueError(f"{kind} 报表日期重复，不能构建汇总预览。")
            day_map[window.start_date] = window
            outcomes = _normalized_attribution(window, store_sid, mappings, confirmed_hsa_mappings)
            outcomes_by_window[(kind, window.start_date)] = outcomes
            for outcome in outcomes.product_rows:
                if outcome.local_sku:
                    key = (kind, window.start_date, outcome.local_sku.casefold())
                    mapped_costs[key] = mapped_costs.get(key, Decimal("0")) + outcome.evidence.cost
                    matched_counts[key] = matched_counts.get(key, 0) + 1
            for outcome in outcomes.ad_group_rows:
                if outcome.local_sku:
                    key = (kind, window.start_date, outcome.local_sku.casefold())
                    mapped_costs[key] = mapped_costs.get(key, Decimal("0")) + outcome.evidence.cost
                    matched_counts[key] = matched_counts.get(key, 0) + 1
            if window.report_complete:
                unmatched_costs[(kind, window.start_date)] = sum(
                    (row.evidence.cost for row in outcomes.product_rows if not row.local_sku), Decimal("0")
                ) + sum(
                    (row.evidence.cost for row in outcomes.ad_group_rows if not row.local_sku), Decimal("0")
                )
            else:
                unmatched_costs[(kind, window.start_date)] = None
        by_kind_date[kind] = day_map

    def catalog_sku_candidates(catalog: Any, asin: str) -> tuple[str, ...]:
        owners: set[str] = set()
        for mapping in mappings:
            if catalog.seller_sku and any(
                catalog.seller_sku.casefold() == value.casefold() for value in mapping.seller_skus
            ):
                owners.add(mapping.local_sku)
            if asin and any(asin.casefold() == value.casefold() for value in mapping.asins):
                owners.add(mapping.local_sku)
        return tuple(sorted(owners, key=str.casefold))

    mapping_by_key = {mapping.local_sku.casefold(): mapping for mapping in mappings}

    def sku_scope_is_complete(day: date, sku_key: str) -> bool:
        if scope != SP_SD_ONLY_CONFIRMED or not currency_and_date_basis_verified:
            return False
        mapping = mapping_by_key.get(sku_key)
        if mapping is None or not (mapping.seller_skus or mapping.asins):
            return False
        target_mskus = {value.casefold() for value in mapping.seller_skus}
        target_asins = {value.casefold() for value in mapping.asins}
        for kind in SP_SD_AD_TYPES:
            window = by_kind_date[kind].get(day)
            if (
                window is None
                or not window.report_complete
                or window.campaign_cost_total is None
                or window.detail_cost_total is None
                or window.campaign_cost_total != window.detail_cost_total
            ):
                return False
            outcomes = outcomes_by_window.get((kind, day), _WindowAttribution())
            if len(outcomes.product_rows) != len(window.product_rows) or any(
                not attributed.campaign_reconciled
                or "冲突" in attributed.status
                or "阻止" in attributed.status
                for attributed in outcomes.product_rows
            ):
                return False
            for attributed in outcomes.product_rows:
                evidence = attributed.evidence
                matches_target = (
                    evidence.seller_sku.casefold() in target_mskus
                    or evidence.asin.casefold() in target_asins
                )
                if not matches_target:
                    continue
                if (
                    attributed.local_sku.casefold() != sku_key
                    or not attributed.campaign_reconciled
                    or "冲突" in attributed.status
                    or "阻止" in attributed.status
                ):
                    return False
        return True

    summaries: list[AdSpendSummaryPreview] = []
    details: list[AdSpendDetailPreview] = []
    day_count = (end_date - start_date).days + 1
    for day_offset in range(day_count):
        report_day = start_date + timedelta(days=day_offset)
        day_text = report_day.isoformat()
        day_windows = {kind: by_kind_date[kind].get(report_day) for kind in kinds}
        for local_sku in requested_skus:
            sku_key = local_sku.casefold()
            values: dict[str, str] = {}
            states: list[str] = []
            for kind in kinds:
                key = (kind, report_day, sku_key)
                values[kind] = _money(mapped_costs.get(key))
                states.append(_type_state(kind, day_windows[kind], matched_rows=matched_counts.get(key, 0)))
            current_scope_complete = sku_scope_is_complete(report_day, sku_key)
            confirmed_total = ""
            if current_scope_complete:
                total_value = sum(
                    (mapped_costs.get((kind, report_day, sku_key), Decimal("0")) for kind in SP_SD_AD_TYPES),
                    Decimal("0"),
                )
                confirmed_total = _money(total_value)
                states.append("统计范围 SP+SD 已核平并按商品级费用归属")
            elif scope == SP_SD_ONLY_CONFIRMED:
                states.append("SP+SD 任一类型缺失、未核平或币种/日期未确认；合计留空")
            summaries.append(AdSpendSummaryPreview(
                report_date=day_text,
                local_sku=local_sku,
                sp_cost=values["SP"],
                hsa_cost=values.get("HSA", ""),
                sd_cost=values["SD"],
                confirmed_total=confirmed_total,
                unattributed_cost="",
                integrity="；".join(
                    states + (["本版统计口径：SP+SD；SB/SBV 不在 V1.0 统计范围"] if scope == SP_SD_ONLY_CONFIRMED else ["预览合计留空：跨类型币种/日期口径与完整覆盖尚未共同验收"])
                ),
            ))

        unallocated_parts: list[str] = []
        unallocated_values: dict[str, str] = {}
        for kind in kinds:
            window = day_windows[kind]
            value = unmatched_costs.get((kind, report_day)) if window is not None else None
            unallocated_values[kind] = _money(value)
            if window is None:
                unallocated_parts.append(f"{kind}=查询结果缺失（未知）")
            elif value is None:
                unallocated_parts.append(f"{kind}=未知（报表未完整核对）")
            else:
                unallocated_parts.append(f"{kind}未归属={_money(value)}")
        summaries.append(AdSpendSummaryPreview(
            report_date=day_text,
            local_sku="未归属（店铺级）",
            sp_cost=unallocated_values["SP"],
            hsa_cost=unallocated_values.get("HSA", ""),
            sd_cost=unallocated_values["SD"],
            confirmed_total="",
            unattributed_cost="；".join(unallocated_parts),
            integrity="；".join(
                [_window_state(kind, day_windows[kind]) for kind in kinds]
                + ["活动总额仅用于核对；广告组或商品明细计费一次"]
            ),
        ))

        for kind in kinds:
            window = day_windows[kind]
            if window is None:
                continue
            discrepancy_campaigns = {(item.profile_id, item.campaign_id) for item in window.discrepancies}
            for row in window.campaign_rows:
                state = (
                    "活动级核对基准；不与明细相加；活动存在差异"
                    if (row.profile_id, row.campaign_id) in discrepancy_campaigns
                    else "活动级核对基准；不与明细相加"
                )
                details.append(AdSpendDetailPreview(
                    day_text, kind, "活动总额（仅核对）", "", "", "", "", row.profile_id,
                    row.campaign_id, "", _money(row.cost), state,
                ))
            outcomes = outcomes_by_window.get((kind, report_day), _WindowAttribution())
            for index, row in enumerate(window.product_rows):
                outcome = outcomes.product_rows[index] if index < len(outcomes.product_rows) else None
                details.append(AdSpendDetailPreview(
                    day_text, kind, "商品级实际花费", outcome.local_sku if outcome else "", row.sku, row.asin,
                    row.ad_id, row.profile_id, row.campaign_id, row.ad_group_id, _money(row.cost),
                    outcome.status if outcome else "没有归属结果；未计入 SKU 汇总",
                ))
            for row in window.ad_group_rows:
                outcome = next((item for item in outcomes.ad_group_rows if (
                    item.evidence.report_date == row.report_date
                    and item.evidence.profile_id == row.profile_id
                    and item.evidence.campaign_id == row.campaign_id
                    and item.evidence.ad_group_id == row.ad_group_id
                )), None)
                has_candidate = bool(outcome and outcome.candidate_local_skus)
                report_level = (
                    "广告组实际花费（已归属）" if outcome and outcome.local_sku
                    else "广告组实际花费（SKU候选）" if has_candidate
                    else "广告组实际花费（未归属）"
                )
                details.append(AdSpendDetailPreview(
                    day_text, kind, report_level,
                    outcome.local_sku if outcome and outcome.local_sku else "", "",
                    ", ".join(outcome.candidate_asins) if outcome else "", "", row.profile_id,
                    row.campaign_id, row.ad_group_id, _money(row.cost),
                    outcome.status if outcome else "仅供活动总额核对，未归属",
                    ad_creative_id=", ".join(outcome.candidate_creative_ids) if outcome else "",
                    candidate_local_skus=outcome.candidate_local_skus if outcome else (),
                    candidate_resolution_complete=bool(outcome and outcome.candidate_resolution_complete),
                ))
            # Current creative catalogs help locate IDs, but contain no historical
            # validity interval and therefore cannot authorize historical spend.
            if report_day == start_date and kind == "HSA":
                for catalog in window.product_catalog_rows:
                    asins = catalog.asins or ("",)
                    for asin in asins:
                        details.append(AdSpendDetailPreview(
                            day_text, "HSA", "当前创意目录快照（非历史费用）", "", "", asin,
                            catalog.ad_id or catalog.ad_creative_id, catalog.profile_id, catalog.campaign_id,
                            catalog.ad_group_id, "", "仅用于核对标识；不证明报表日推广关系，不参与费用归属或计费",
                            ad_creative_id=catalog.ad_creative_id,
                            candidate_local_skus=catalog_sku_candidates(catalog, asin),
                        ))
    return summaries, details
