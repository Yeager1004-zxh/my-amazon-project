"""Safe, read-only presentation rows for multi-type advertising reports."""

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
from .providers.lingxing_api import SpAdSpendWindow, SkuProductMapping


@dataclass(frozen=True)
class AdSpendSummaryPreview:
    report_date: str
    local_sku: str
    sp_cost: str
    sb_cost: str
    sbv_cost: str
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


@dataclass(frozen=True)
class _WindowAttribution:
    product_rows: tuple[AttributedAdCost, ...] = ()
    ad_group_rows: tuple[ExclusiveAdGroupAttribution, ...] = ()


def _money(value: Decimal | None) -> str:
    return "" if value is None else format(value, "f")


def _type_state(kind: str, window: SpAdSpendWindow | None, *, matched_rows: int = 0) -> str:
    if window is None:
        if kind == "SBV":
            return "SBV 查询结果缺失（未知，不视为零）"
        return f"{kind} 查询结果缺失（未知）"
    if matched_rows:
        blocked = len(window.discrepancies)
        suffix = f"；另有 {blocked} 个活动未核平并被隔离" if blocked else ""
        return f"{kind} {matched_rows} 条费用按 MSKU/ASIN 精确匹配本地 SKU{suffix}"
    if kind == "HSA":
        if window.report_complete:
            if window.promotion_targets and window.promotion_catalog_complete:
                return "HSA（普通 SB/SBV 未拆分）：当前创意目录已读取；报表日推广身份未核验"
            return "HSA（普通 SB/SBV 未拆分）：活动/广告组成本已核平；类型与商品归属待确认"
        if window.discrepancies:
            return "HSA（普通 SB/SBV 未拆分）：活动与广告组成本未核平"
        return "HSA（普通 SB/SBV 未拆分）无记录（未知，不视为零）"
    if kind == "SBV":
        if window.discrepancies:
            return "SBV 活动与广告组成本未核平；相关费用不归属"
        if window.report_complete and window.ad_group_rows:
            return "SBV 活动与广告组成本已核平；报表日推广商品尚未核验"
        return "SBV 报表/历史商品关联缺失（未知，不视为零）"
    if not window.report_complete:
        if window.discrepancies:
            return f"{kind} 部分活动未核平；相关记录阻止归属"
        if kind == "HSA":
            if window.campaign_rows and not window.ad_group_rows:
                return f"HSA 活动报表有 {len(window.campaign_rows)} 条记录，广告组报表无记录；SB/SBV 类型和 SKU 花费未知"
            if window.ad_group_rows and not window.campaign_rows:
                return f"HSA 广告组报表有 {len(window.ad_group_rows)} 条记录，活动报表无记录；金额未核平"
        else:
            if window.campaign_rows and not window.product_rows:
                return f"{kind} 活动报表有 {len(window.campaign_rows)} 条记录，商品报表无记录；SKU 花费未知"
            if window.product_rows and not window.campaign_rows:
                return f"{kind} 商品报表有 {len(window.product_rows)} 条记录，活动报表无记录；金额未核平"
        if not window.campaign_rows and not window.product_rows and not window.ad_group_rows:
            return f"{kind} 报表接口已返回空记录（不视为零花费）"
        return f"{kind} 活动或商品明细无记录（未知，不视为零）"
    return f"{kind} 商品报表无该 SKU 记录（未知，不视为零）"


def _window_state(kind: str, window: SpAdSpendWindow | None) -> str:
    if window is None:
        return f"{kind} 查询结果缺失"
    if window.report_complete:
        if kind == "HSA":
            detail = "商品/分摊明细" if window.product_rows else "广告组（SB/SBV 类型未拆分）"
        else:
            detail = "商品"
        return f"{kind} 活动与{detail}成本已核平"
    if window.discrepancies:
        return f"{kind} 活动与明细成本不一致，相关活动金额未知"
    if kind == "HSA":
        if window.campaign_rows and not window.ad_group_rows:
            return "HSA 活动接口有记录，但广告组接口返回空记录；SB/SBV 类型未分"
        if window.ad_group_rows and not window.campaign_rows:
            return "HSA 广告组接口有记录，但活动接口返回空记录"
    else:
        if window.campaign_rows and not window.product_rows:
            return f"{kind} 活动接口有记录，但商品接口返回空记录"
        if window.product_rows and not window.campaign_rows:
            return f"{kind} 商品接口有记录，但活动接口返回空记录"
    if not window.campaign_rows and not window.product_rows and not window.ad_group_rows:
        return f"{kind} 活动及明细接口返回空记录；花费未知，不作为零值"
    return f"{kind} 活动或明细报表缺记录，金额未知"


def _normalized_attribution(window: SpAdSpendWindow, store_sid: int, mappings: Sequence[SkuProductMapping]) -> _WindowAttribution:
    campaigns = [
        CampaignCostEvidence(
            window.ad_type, store_sid, row.report_date, row.profile_id, row.campaign_id, row.cost
        )
        for row in window.campaign_rows
    ]
    # The Lingxing HSA report family combines ordinary SB and SBV. Never
    # attribute its rows as ordinary SB merely because the endpoint is HSA.
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
    if window.ad_type == "SBV" and not window.product_rows and window.ad_group_rows:
        group_result = attribute_exclusive_ad_group_costs(
            store_sid=store_sid,
            mappings=[AttributionMapping(item.local_sku, item.seller_skus, item.asins) for item in mappings],
            promoted_targets=window.promotion_targets,
            ad_group_costs=[
                AdGroupCostEvidence(window.ad_type, store_sid, row.report_date, row.profile_id, row.campaign_id, row.ad_group_id, row.cost)
                for row in window.ad_group_rows
            ],
            campaign_costs=campaigns,
            promotion_catalog_complete=window.promotion_catalog_complete,
        )
        group_rows = tuple(group_result.rows)
    else:
        group_rows = ()
    return _WindowAttribution(tuple(result.rows), group_rows)


def build_ad_spend_preview(
    *,
    windows_by_type: Mapping[str, Sequence[SpAdSpendWindow]],
    mappings: Sequence[SkuProductMapping],
    store_sid: int,
    requested_skus: Sequence[str],
    start_date: date,
    end_date: date,
) -> tuple[list[AdSpendSummaryPreview], list[AdSpendDetailPreview]]:
    """Build preview rows using exact product identifiers, without adding levels.

    SP/SD product rows are attributed by unique, exact MSKU and/or ASIN matches.
    SB can use the same path once an adapter supplies verified product/allocation
    rows. Activity totals remain reconciliation references only. Currency,
    report-date timezone and historical identity validity are not yet certified,
    so these values remain previews and the cross-type total remains blank.
    """
    if end_date < start_date:
        raise ValueError("广告预览日期范围无效。")
    if isinstance(store_sid, bool) or not isinstance(store_sid, int) or store_sid <= 0:
        raise ValueError("广告预览必须使用所选店铺的有效 SID。")
    # HSA activity/group reports combine Sponsored Brands and Sponsored
    # Brands Video. Keep that amount separate until the official source proves
    # both the ad subtype and its historical product target.
    kinds = ("SP", "SB", "SBV", "SD", "HSA")
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
            outcomes = _normalized_attribution(window, store_sid, mappings)
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

            if window.report_complete and kind == "HSA":
                # The HSA campaign total is the only safe single-level amount
                # while ordinary SB and SBV are not distinguished.
                unmatched_costs[(kind, window.start_date)] = window.campaign_cost_total
            elif window.report_complete:
                unmatched_costs[(kind, window.start_date)] = sum(
                    (row.evidence.cost for row in outcomes.product_rows if not row.local_sku), Decimal("0")
                ) + sum(
                    (row.evidence.cost for row in outcomes.ad_group_rows if not row.local_sku), Decimal("0")
                )
            else:
                unmatched_costs[(kind, window.start_date)] = None
        by_kind_date[kind] = day_map

    def catalog_sku_candidate(seller_sku: str, asins: Sequence[str]) -> str:
        owners: set[str] = set()
        for mapping in mappings:
            if seller_sku and any(seller_sku.casefold() == value.casefold() for value in mapping.seller_skus):
                owners.add(mapping.local_sku)
            if any(asin.casefold() == value.casefold() for asin in asins for value in mapping.asins):
                owners.add(mapping.local_sku)
        if len(owners) == 1:
            return next(iter(owners))
        if owners:
            return "归属冲突候选：" + ", ".join(sorted(owners))
        return ""

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
            summaries.append(AdSpendSummaryPreview(
                report_date=day_text,
                local_sku=local_sku,
                sp_cost=values["SP"],
                sb_cost=values["SB"],
                sbv_cost=values["SBV"],
                sd_cost=values["SD"],
                confirmed_total="",
                unattributed_cost="",
                integrity="；".join(states + ["精确 ID 自动归属供预览；币种/日期口径/历史映射时效尚未全部验收，合计留空"]),
            ))

        unallocated_parts: list[str] = []
        unallocated_values: dict[str, str] = {}
        for kind in kinds:
            window = day_windows[kind]
            value = unmatched_costs.get((kind, report_day)) if window is not None else None
            unallocated_values[kind] = _money(value)
            if window is None:
                unallocated_parts.append(f"{kind}=查询结果缺失")
            elif value is None:
                unallocated_parts.append(f"{kind}=未知（未完整核对）")
            elif kind == "HSA":
                unallocated_parts.append(f"HSA普通SB/SBV未拆分={_money(value)}")
            else:
                unallocated_parts.append(f"{kind}={_money(value)}")
        summaries.append(AdSpendSummaryPreview(
            report_date=day_text,
            local_sku="未归属（店铺级）",
            sp_cost=unallocated_values["SP"],
            # Do not place the combined HSA report total into the ordinary-SB
            # or SBV column. It remains visible only in unattributed_cost.
            sb_cost="",
            sbv_cost=unallocated_values["SBV"],
            sd_cost=unallocated_values["SD"],
            confirmed_total="",
            unattributed_cost="；".join(unallocated_parts),
            integrity="；".join(
                [_window_state(kind, day_windows[kind]) for kind in kinds]
                + ["HSA 活动费用暂不拆成普通 SB/SBV；活动费用不与明细相加；未归属金额保留在说明中"]
            ),
        ))

        for kind in kinds:
            window = day_windows[kind]
            if window is None:
                continue
            display_kind = kind
            discrepancy_campaigns = {(item.profile_id, item.campaign_id) for item in window.discrepancies}
            for row in window.campaign_rows:
                state = (
                    "活动级核对基准；不与明细相加；活动存在差异"
                    if (row.profile_id, row.campaign_id) in discrepancy_campaigns
                    else "活动级核对基准；不与明细相加"
                )
                details.append(AdSpendDetailPreview(
                    day_text, display_kind, "活动总额（仅核对）", "", "", "", "", row.profile_id,
                    row.campaign_id, "", _money(row.cost), state,
                ))
            outcomes = outcomes_by_window.get((kind, report_day), _WindowAttribution())
            for index, row in enumerate(window.product_rows):
                outcome = outcomes.product_rows[index] if index < len(outcomes.product_rows) else None
                details.append(AdSpendDetailPreview(
                    day_text, display_kind, "商品/分摊明细", outcome.local_sku if outcome else "", row.sku, row.asin, row.ad_id,
                    row.profile_id, row.campaign_id, row.ad_group_id, _money(row.cost),
                    outcome.status if outcome else "HSA 类型未拆分；不将记录归入普通 SB 或 SBV",
                ))
            for row in window.ad_group_rows:
                outcome = next((item for item in outcomes.ad_group_rows if (
                    item.evidence.report_date == row.report_date
                    and item.evidence.profile_id == row.profile_id
                    and item.evidence.campaign_id == row.campaign_id
                    and item.evidence.ad_group_id == row.ad_group_id
                )), None)
                details.append(AdSpendDetailPreview(
                    day_text, display_kind, "广告组费用" if outcome and outcome.local_sku else "广告组记录（不重复计费）",
                    outcome.local_sku if outcome else "", "", "", "", row.profile_id,
                    row.campaign_id, row.ad_group_id, _money(row.cost),
                    outcome.status if outcome else "只用于和活动总额核对，不重复计费",
                ))
            # Catalog endpoints expose current identities, not report-date
            # history. Show their exact IDs for diagnosis, but never attach cost.
            if report_day == start_date:
                for catalog in window.product_catalog_rows:
                    catalog_asins = catalog.asins or ("",)
                    for asin in catalog_asins:
                        has_video_field = any(
                            value.strip().upper() == "VIDEO" for value in catalog.video_creative_types
                        )
                        is_video_candidate = catalog.creative_type.strip().upper() == "VIDEO" or has_video_field
                        catalog_type = "SBV候选" if is_video_candidate else "HSA创意候选"
                        creative_note = (
                            (
                                "creative_type=VIDEO；仅为 SBV 类型候选，非报表日历史证明"
                                if catalog.creative_type.strip().upper() == "VIDEO"
                                else "video_creatives[].creative_type=VIDEO；仅为 SBV 类型候选，非报表日历史证明"
                            )
                            if is_video_candidate
                            else f"creative_type={catalog.creative_type}；HSA 报表未给出 SB/SBV 类型，历史也未证明"
                        )
                        details.append(AdSpendDetailPreview(
                            day_text,
                            catalog_type,
                            "当前商品目录快照（非历史费用）",
                            catalog_sku_candidate(catalog.seller_sku, (asin,)),
                            catalog.seller_sku,
                            asin,
                            catalog.ad_id or catalog.ad_creative_id,
                            catalog.profile_id,
                            catalog.campaign_id,
                            catalog.ad_group_id,
                            "",
                            f"{creative_note}；当前目录 ID 关系仅供定位，不参与费用归属或合计",
                        ))
    return summaries, details
