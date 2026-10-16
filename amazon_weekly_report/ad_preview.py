"""Safe, read-only presentation rows for multi-type advertising reports."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from decimal import Decimal
from typing import Mapping, Sequence

from .providers.lingxing_api import SpAdSpendWindow, SkuProductMapping


@dataclass(frozen=True)
class AdSpendSummaryPreview:
    report_date: str
    local_sku: str
    sp_cost: str
    sb_cost: str
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
    profile_id: str
    campaign_id: str
    ad_group_id: str
    cost: str
    integrity: str


def _money(value: Decimal | None) -> str:
    return "" if value is None else format(value, "f")


def _type_state(kind: str, window: SpAdSpendWindow | None, *, sku_cost_known: bool = False) -> str:
    if window is None:
        return f"{kind} 查询结果缺失（未知）"
    if not window.report_complete:
        if window.discrepancies:
            return f"{kind} 活动与明细成本不一致（归属被阻止）"
        return f"{kind} 活动或明细无记录（未知，不视为零）"
    if kind == "SB":
        return "SB 活动与广告组成本核平；无 SKU/ASIN 字段，保持店铺级未归属"
    if sku_cost_known:
        return f"{kind} 活动与商品成本核平；SKU 候选归属待 Windows 验证"
    return f"{kind} 商品报表无该 SKU 记录（未知，不视为零）"


def _window_state(kind: str, window: SpAdSpendWindow | None) -> str:
    if window is None:
        return f"{kind} 查询结果缺失（未知）"
    if window.report_complete:
        detail = "广告组" if kind == "SB" else "商品"
        return f"{kind} 活动与{detail}成本已核平"
    if window.discrepancies:
        return f"{kind} 活动与明细成本不一致，金额未知"
    return f"{kind} 活动或明细报表缺记录，金额未知"


def build_ad_spend_preview(
    *,
    windows_by_type: Mapping[str, Sequence[SpAdSpendWindow]],
    mappings: Sequence[SkuProductMapping],
    requested_skus: Sequence[str],
    start_date: date,
    end_date: date,
) -> tuple[list[AdSpendSummaryPreview], list[AdSpendDetailPreview]]:
    """Build preview rows without treating absent records as zero or adding types.

    Campaign totals are reconciliation references only. Product/ad-group detail
    costs are used for their own level, and never added again to campaign costs.
    Currency and report-date timezone are undocumented, so cross-type totals are
    deliberately blank.
    """
    if end_date < start_date:
        raise ValueError("广告预览日期范围无效。")
    kinds = ("SP", "SB", "SD")
    by_kind_date: dict[str, dict[date, SpAdSpendWindow]] = {}
    for kind in kinds:
        day_map: dict[date, SpAdSpendWindow] = {}
        for window in windows_by_type.get(kind, ()):
            if window.ad_type != kind or window.start_date != window.end_date:
                raise ValueError(f"{kind} 预览要求单日且类型匹配的报表窗口。")
            if window.start_date in day_map:
                raise ValueError(f"{kind} 报表日期重复，不能构建汇总预览。")
            day_map[window.start_date] = window
        by_kind_date[kind] = day_map

    summaries: list[AdSpendSummaryPreview] = []
    details: list[AdSpendDetailPreview] = []
    day_count = (end_date - start_date).days + 1
    for day_offset in range(day_count):
        report_day = start_date + timedelta(days=day_offset)
        day_text = report_day.isoformat()
        day_windows = {kind: by_kind_date[kind].get(report_day) for kind in kinds}

        mapped_costs: dict[str, dict[str, Decimal]] = {kind: {} for kind in ("SP", "SD")}
        for kind in ("SP", "SD"):
            window = day_windows[kind]
            if window is not None and window.report_complete:
                for spend in window.sku_daily_spends:
                    key = spend.local_sku.casefold()
                    mapped_costs[kind][key] = mapped_costs[kind].get(key, Decimal("0")) + spend.actual_spend

        for local_sku in requested_skus:
            sku_key = local_sku.casefold()
            sku_states: list[str] = []
            values: dict[str, str] = {}
            for kind in kinds:
                window = day_windows[kind]
                value = mapped_costs.get(kind, {}).get(sku_key) if kind in ("SP", "SD") else None
                known = value is not None
                values[kind] = _money(value)
                sku_states.append(_type_state(kind, window, sku_cost_known=known))
            summaries.append(AdSpendSummaryPreview(
                report_date=day_text,
                local_sku=local_sku,
                sp_cost=values["SP"],
                sb_cost=values["SB"],
                sd_cost=values["SD"],
                confirmed_total="",
                unattributed_cost="",
                integrity="；".join(sku_states + ["跨类型币种/日期口径未核验，合计留空"]),
            ))

        unallocated_parts: list[str] = []
        unallocated_costs: dict[str, str] = {}
        for kind in kinds:
            window = day_windows[kind]
            cost = window.unattributed_cost if window is not None and window.report_complete else None
            unallocated_costs[kind] = _money(cost)
            if window is None:
                unallocated_parts.append(f"{kind}=查询结果缺失")
            elif not window.report_complete:
                unallocated_parts.append(f"{kind}=未知（报表未完整核对）")
            else:
                unallocated_parts.append(f"{kind}={_money(cost)}")
        summaries.append(AdSpendSummaryPreview(
            report_date=day_text,
            local_sku="未归属（店铺级）",
            sp_cost=unallocated_costs["SP"],
            sb_cost=unallocated_costs["SB"],
            sd_cost=unallocated_costs["SD"],
            confirmed_total="",
            unattributed_cost="；".join(unallocated_parts),
            integrity="；".join(
                [_window_state(kind, day_windows[kind]) for kind in kinds]
                + ["各类型未归属金额分开列示；币种/日期口径未核验，不相加"]
            ),
        ))

        for kind in kinds:
            window = day_windows[kind]
            if window is None:
                continue
            discrepancy_campaigns = {(item.profile_id, item.campaign_id) for item in window.discrepancies}
            for row in window.campaign_rows:
                state = (
                    "活动级核对基准；不要与明细费用相加；该活动存在差异"
                    if (row.profile_id, row.campaign_id) in discrepancy_campaigns
                    else "活动级核对基准；不要与明细费用相加"
                )
                details.append(AdSpendDetailPreview(
                    day_text, kind, "活动总额（仅核对）", "", "", row.profile_id,
                    row.campaign_id, "", _money(row.cost), state
                ))
            for row in window.product_rows:
                local_sku = next(
                    (mapping.local_sku for mapping in mappings
                     if any(seller_sku.casefold() == row.sku.casefold() for seller_sku in mapping.seller_skus)),
                    "",
                )
                if (row.profile_id, row.campaign_id) in discrepancy_campaigns or not window.report_complete:
                    state = "活动与商品成本未核平；不能归属到 SKU"
                    local_sku = ""
                elif local_sku:
                    state = "MSKU 精确匹配；候选归属待 Windows 真实验证"
                else:
                    state = "未匹配本次 SKU 映射；留在店铺级未归属费用"
                details.append(AdSpendDetailPreview(
                    day_text, kind, "商品记录", local_sku, row.sku, row.profile_id, row.campaign_id,
                    row.ad_group_id, _money(row.cost), state,
                ))
            for row in window.ad_group_rows:
                details.append(AdSpendDetailPreview(
                    day_text, kind, "广告组记录（不重复计费）", "", "", row.profile_id, row.campaign_id,
                    row.ad_group_id, _money(row.cost), "无 SKU/ASIN 字段；仅用于和活动总额核对",
                ))
    return summaries, details
