from __future__ import annotations

import json
import queue
import re
import threading
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Any

from .ad_preview import build_ad_spend_preview
from .credential_store import (
    CredentialStoreError,
    LingxingOpenApiCredentialStore,
    WindowsCredentialStore,
)
from .d01_settings import D01SettingsError, get_d01_dates, save_d01_dates
from .hsa_attribution_settings import (
    HsaAdGroupSkuMapping,
    HsaAttributionSettingsError,
    save_hsa_ad_group_mappings,
)
from .local_settings import LocalSettingsError, load_settings, save_settings, validate_mcp_endpoint
from .mcp_client import (
    BusinessTool,
    LingxingMcpClient,
    McpClientError,
    McpTool,
    VerifiedReadOnlyGrant,
    extract_business_tools,
    is_action_wrapper_tool,
    is_directory_search_tool,
    is_read_only_query_candidate,
    redact_sensitive,
)
from .providers.lingxing_api import (
    LingxingApiError,
    LingxingOpenApiClient,
    LingxingSeller,
    PRODUCT_PERFORMANCE_PREVIEW_FIELDS,
    ProductPerformanceMetrics,
    SkuDailySales,
    SkuDailySalesWindow,
    SkuProductMapping,
)
from .report_generation import ReportGenerationError, generate_reports, read_existing_d01_dates
from .statistics_scope import SP_SD_ONLY_CONFIRMED


_PRODUCT_PERFORMANCE_FIELD_LABELS = {
    "sessions_total": "Sessions-Total",
    "ad_order_quantity": "广告订单量",
    "ad_sales_amount": "广告销售额",
    "impressions": "曝光量",
    "clicks": "点击量",
    "ctr": "CTR",
    "ad_cvr": "广告 CVR",
    "avg_star": "当前评分字段 avg_star",
    "reviews_count": "当前评论数字段 reviews_count",
    "small_cate_rank": "小类排名结构",
    "cate_rank": "大类排名",
    "prev_cate_rank": "上一次大类排名",
    "rank_category": "排名类别",
    "ranking_update_time": "排名更新时间（文档标记已废弃）",
    "spend": "广告花费总计 spend",
    "ads_sp_cost": "SP 花费",
    "shared_ads_sb_cost": "产品表现字段 shared_ads_sb_cost（诊断）",
    "shared_ads_sbv_cost": "产品表现字段 shared_ads_sbv_cost（诊断）",
    "ads_sd_cost": "SD 花费",
    "shared_ads_al_cost": "其他广告花费 shared_ads_al_cost",
    "shared_ads_cc_cost": "其他广告花费 shared_ads_cc_cost",
    "shared_ads_sspaot_cost": "其他广告花费 shared_ads_sspaot_cost",
    "shared_ads_sar_cost": "其他广告花费 shared_ads_sar_cost",
    "shared_cost_of_advertising": "差异分摊字段",
    "adv_rate": "广告订单量占比",
    "currency_code": "币种",
}

_PRODUCT_PERFORMANCE_COST_DIAGNOSTIC_FIELDS = (
    "spend",
    "ads_sp_cost",
    "shared_ads_sb_cost",
    "shared_ads_sbv_cost",
    "ads_sd_cost",
    "shared_cost_of_advertising",
    "currency_code",
)


def _product_performance_value(row: ProductPerformanceMetrics, field_name: str) -> str:
    if field_name not in row.present_fields:
        return "字段未返回"
    value = row.values.get(field_name)
    if value is None:
        return "null（接口返回空值）"
    if isinstance(value, Decimal):
        return format(value, "f")
    if isinstance(value, tuple):
        return json.dumps(value, ensure_ascii=False, default=str)
    return str(value)


def _product_performance_scope_note(field_name: str) -> str:
    if field_name in {"avg_star", "reviews_count", "small_cate_rank", "cate_rank", "prev_cate_rank", "rank_category", "ranking_update_time"}:
        return "历史时点未由 Schema 证明；禁止回填历史日期"
    if field_name in {
        "spend", "ads_sp_cost", "shared_ads_sb_cost", "shared_ads_sbv_cost", "ads_sd_cost",
        "shared_ads_al_cost", "shared_ads_cc_cost", "shared_ads_sspaot_cost", "shared_ads_sar_cost",
        "shared_cost_of_advertising",
    }:
        return "SKU 汇总候选；须核对币种、日期和广告明细后才可入周报"
    return "仅限请求的单日筛选；响应行无日期，须与后台单日数据核对"


class WeeklyReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("亚马逊新品推广周报工具")
        self.geometry("1100x960")
        self.minsize(920, 760)
        self._events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._busy = False
        self._tools: list[McpTool] = []
        self._business_tools: list[BusinessTool] = []
        self._verified_query_grants: dict[str, VerifiedReadOnlyGrant] = {}
        self._viewed_business_fingerprints: set[str] = set()
        self._sellers: list[LingxingSeller] = []
        self._seller_choices: dict[str, LingxingSeller] = {}
        self._d01_refreshing = False
        self._pending_hsa_candidates: tuple[dict[str, Any], ...] = ()
        self._pending_hsa_candidate_context: tuple[int, tuple[str, ...], date, date] | None = None
        self._build_ui()
        self._load_mcp_settings()
        self.after(150, self._poll_events)

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        report_tab = ttk.Frame(notebook, padding=12)
        mcp_tab = ttk.Frame(notebook, padding=12)
        openapi_tab = ttk.Frame(notebook, padding=12)
        notebook.add(report_tab, text="生成周报")
        notebook.add(mcp_tab, text="领星 MCP 连接")
        notebook.add(openapi_tab, text="领星 OpenAPI 连接")
        self._build_scrollable_report_tab(report_tab)
        self._build_mcp_tab(mcp_tab)
        self._build_openapi_tab(openapi_tab)

    def _build_scrollable_report_tab(self, parent: ttk.Frame) -> None:
        """Keep the expanded workflow usable on typical Windows laptop screens."""
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content = ttk.Frame(canvas, padding=12)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))
        self._build_report_tab(content)

    def _path_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, *, directory: bool = False) -> None:
        ttk.Label(parent, text=label, width=20).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=6, pady=5)
        command = (lambda: filedialog.askdirectory()) if directory else (
            lambda: filedialog.askopenfilename(filetypes=[("Excel / JSON", "*.xlsx *.xlsm *.json"), ("All files", "*.*")])
        )
        ttk.Button(parent, text="浏览…", command=lambda: self._set_path(variable, command())).grid(row=row, column=2, pady=5)

    @staticmethod
    def _set_path(variable: tk.StringVar, value: str) -> None:
        if value:
            variable.set(value)

    def _build_report_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        self.template_var = tk.StringVar()
        self.data_var = tk.StringVar()
        self.mapping_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.existing_var = tk.StringVar()
        self.skus_var = tk.StringVar()
        self.report_store_var = tk.StringVar()
        self.query_start_date_var = tk.StringVar()
        self.query_end_date_var = tk.StringVar()
        self._path_row(parent, 0, "Excel 格式模板", self.template_var)
        self._path_row(parent, 1, "领星导出数据文件夹", self.data_var, directory=True)
        self._path_row(parent, 2, "字段映射 JSON", self.mapping_var)
        self._path_row(parent, 3, "输出文件夹", self.output_var, directory=True)
        self._path_row(parent, 4, "已有周报（可选）", self.existing_var)
        ttk.Label(parent, text="领星店铺", width=20).grid(row=5, column=0, sticky="w", pady=5)
        self.report_store_combo = ttk.Combobox(parent, textvariable=self.report_store_var, state="readonly")
        self.report_store_combo.grid(row=5, column=1, sticky="ew", padx=6, pady=5)
        self.report_store_combo.bind("<<ComboboxSelected>>", self._on_report_store_selected)
        ttk.Button(parent, text="刷新店铺列表", command=self._openapi_seller_list).grid(row=5, column=2, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="选择项会自动绑定接口返回的 SID；本地 Excel/CSV 模式仍按字段映射读取。", wraplength=850).grid(
            row=6, column=1, columnspan=2, sticky="w", padx=6, pady=(0, 4)
        )
        ttk.Label(parent, text="SKU（逗号分隔）", width=20).grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.skus_var).grid(row=7, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        self.skus_var.trace_add("write", self._on_report_skus_changed)
        self.query_start_date_var.trace_add("write", self._invalidate_hsa_candidate_review)
        self.query_end_date_var.trace_add("write", self._invalidate_hsa_candidate_review)

        query_box = ttk.LabelFrame(parent, text="领星 SKU 与每日销量预览（只读，不写入周报）", padding=8)
        query_box.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=6)
        ttk.Label(query_box, text="查询开始日（站点时间）").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(query_box, textvariable=self.query_start_date_var, width=13).grid(row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Label(query_box, text="查询结束日").grid(row=0, column=2, sticky="e", padx=4, pady=4)
        ttk.Entry(query_box, textvariable=self.query_end_date_var, width=13).grid(row=0, column=3, sticky="w", padx=4, pady=4)
        self.report_query_button = ttk.Button(query_box, text="查询并预览真实数据", command=self._query_report_sku_data)
        self.report_query_button.grid(row=0, column=4, sticky="w", padx=6, pady=4)
        self.product_performance_query_button = ttk.Button(
            query_box,
            text="AsinListNew\n单日只读查询",
            command=self._query_product_performance_preview,
        )
        self.product_performance_query_button.grid(row=0, column=5, sticky="w", padx=6, pady=4)
        self.product_performance_cost_diagnostic_button = ttk.Button(
            query_box,
            text="费用口径诊断\n4 次只读查询",
            command=self._query_product_performance_cost_diagnostic,
        )
        self.product_performance_cost_diagnostic_button.grid(row=0, column=6, sticky="w", padx=6, pady=4)
        self.report_mapping_status = tk.StringVar(value="商品映射尚未查询。")
        ttk.Label(query_box, textvariable=self.report_mapping_status, wraplength=900).grid(
            row=1, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 5)
        )
        self.report_query_status = tk.StringVar(
            value="日期按所选站点自然日查询，最多 92 天。空白表示接口未返回该指标，0 表示明确返回零。"
        )
        ttk.Label(query_box, textvariable=self.report_query_status, wraplength=900).grid(
            row=2, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 5)
        )
        preview_columns = ("event_date", "local_sku", "seller_sku", "asin", "units_sold", "order_count", "sales_amount", "currency_code")
        self.report_sales_tree = ttk.Treeview(query_box, columns=preview_columns, show="headings", height=5)
        for column, heading, width in (
            ("event_date", "站点日期", 95),
            ("local_sku", "本地 SKU", 120),
            ("seller_sku", "MSKU", 130),
            ("asin", "ASIN", 125),
            ("units_sold", "销量", 65),
            ("order_count", "订单量", 70),
            ("sales_amount", "销售额", 100),
            ("currency_code", "币种", 75),
        ):
            self.report_sales_tree.heading(column, text=heading)
            self.report_sales_tree.column(column, width=width, anchor="w")
        self.report_sales_tree.grid(row=3, column=0, columnspan=5, sticky="nsew", padx=4, pady=4)
        sales_scrollbar = ttk.Scrollbar(query_box, orient="vertical", command=self.report_sales_tree.yview)
        sales_scrollbar.grid(row=3, column=5, sticky="ns", pady=4)
        self.report_sales_tree.configure(yscrollcommand=sales_scrollbar.set)

        performance_box = ttk.LabelFrame(
            query_box,
            text="AsinListNew 产品表现（只读预览；不写入 Excel 或日志）",
            padding=6,
        )
        performance_box.grid(row=4, column=0, columnspan=6, sticky="nsew", padx=4, pady=(7, 3))
        self.product_performance_status = tk.StringVar(
            value="单日预览要求开始日和结束日相同；费用诊断要求连续三天，比较每日 SKU 与中间日 MSKU 汇总。"
        )
        ttk.Label(performance_box, textvariable=self.product_performance_status, wraplength=950).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=4, pady=(0, 4)
        )
        performance_columns = (
            "event_date", "summary_field", "summary_key", "local_sku",
            "field", "value", "currency", "scope",
        )
        self.product_performance_tree = ttk.Treeview(
            performance_box, columns=performance_columns, show="headings", height=8
        )
        for column, heading, width in (
            ("event_date", "请求日期", 92),
            ("summary_field", "汇总维度", 78),
            ("summary_key", "维度键", 125),
            ("local_sku", "本地 SKU", 125),
            ("field", "API 字段", 180),
            ("value", "返回值（空白 / null / 0 分开）", 190),
            ("currency", "币种", 60),
            ("scope", "口径核验状态", 420),
        ):
            self.product_performance_tree.heading(column, text=heading)
            self.product_performance_tree.column(column, width=width, anchor="w")
        self.product_performance_tree.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=4, pady=3)
        performance_scrollbar = ttk.Scrollbar(
            performance_box, orient="vertical", command=self.product_performance_tree.yview
        )
        performance_scrollbar.grid(row=1, column=4, sticky="ns", pady=3)
        self.product_performance_tree.configure(yscrollcommand=performance_scrollbar.set)
        performance_xscrollbar = ttk.Scrollbar(
            performance_box, orient="horizontal", command=self.product_performance_tree.xview
        )
        performance_xscrollbar.grid(row=2, column=0, columnspan=4, sticky="ew", padx=4)
        self.product_performance_tree.configure(xscrollcommand=performance_xscrollbar.set)
        performance_box.columnconfigure(3, weight=1)
        performance_box.rowconfigure(1, weight=1)

        query_box.columnconfigure(4, weight=1)
        query_box.rowconfigure(3, weight=1)
        self.report_ad_query_button = ttk.Button(
            query_box,
            text="查询 SP + SD 广告花费并预览",
            command=self._query_report_ad_data,
        )
        self.report_ad_query_button.grid(row=5, column=0, sticky="w", padx=4, pady=(8, 3))
        ttk.Label(query_box, text="V1.0 统计范围：广告花费（SP+SD）；不含普通 SB / SBV。完整核验后才写入周报。").grid(
            row=5, column=1, columnspan=4, sticky="w", padx=4, pady=(8, 3)
        )
        self.report_ad_status = tk.StringVar(
            value="SP 和 SD 均按历史商品报表归属 SKU；活动级花费仅用于核对。"
        )
        ttk.Label(query_box, textvariable=self.report_ad_status, wraplength=900).grid(
            row=6, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 3)
        )
        ttk.Label(query_box, text="按站点日期及 SKU 汇总；SP+SD 商品级费用各计一次，活动报表仅核对。缺失、失败、未归属和明确零值分别显示。预览不查询每日销售币种，因此合计暂留空；正式生成会用同店铺币种核验。", wraplength=900).grid(
            row=7, column=0, columnspan=5, sticky="w", padx=4, pady=(2, 2)
        )
        ad_summary_columns = (
            "report_date", "local_sku", "sp_cost", "sd_cost",
            "confirmed_total", "unattributed_cost", "integrity",
        )
        self.report_ad_summary_tree = ttk.Treeview(
            query_box, columns=ad_summary_columns, show="headings", height=4
        )
        for column, heading, width in (
            ("report_date", "报表日期", 95),
            ("local_sku", "本地 SKU / 范围", 155),
            ("sp_cost", "SP 实际花费", 100),
            ("sd_cost", "SD 实际花费", 100),
            ("confirmed_total", "已核验花费（SP+SD）", 145),
            ("unattributed_cost", "未归属花费", 300),
            ("integrity", "完整性 / 核对状态", 420),
        ):
            self.report_ad_summary_tree.heading(column, text=heading)
            self.report_ad_summary_tree.column(column, width=width, anchor="w")
        self.report_ad_summary_tree.grid(row=8, column=0, columnspan=5, sticky="nsew", padx=4, pady=4)
        ad_summary_scrollbar = ttk.Scrollbar(query_box, orient="vertical", command=self.report_ad_summary_tree.yview)
        ad_summary_scrollbar.grid(row=8, column=5, sticky="ns", pady=4)
        ad_summary_xscrollbar = ttk.Scrollbar(query_box, orient="horizontal", command=self.report_ad_summary_tree.xview)
        ad_summary_xscrollbar.grid(row=9, column=0, columnspan=5, sticky="ew", padx=4)
        self.report_ad_summary_tree.configure(
            yscrollcommand=ad_summary_scrollbar.set, xscrollcommand=ad_summary_xscrollbar.set
        )
        ttk.Label(query_box, text="原始活动与商品 / 广告组行；活动行只用于核对，不与明细重复累计。", wraplength=900).grid(
            row=10, column=0, columnspan=5, sticky="w", padx=4, pady=(5, 2)
        )
        ad_detail_columns = (
            "report_date", "ad_type", "report_level", "local_sku", "report_sku", "asin", "ad_id",
            "ad_creative_id", "candidate_skus", "profile_id", "campaign_id", "ad_group_id", "cost", "integrity",
        )
        self.report_ad_detail_tree = ttk.Treeview(
            query_box, columns=ad_detail_columns, show="headings", height=4
        )
        for column, heading, width in (
            ("report_date", "报表日期", 95),
            ("ad_type", "类型", 55),
            ("report_level", "报表粒度", 155),
            ("local_sku", "本地 SKU", 120),
            ("report_sku", "报表 SKU/MSKU", 130),
            ("asin", "ASIN", 115),
            ("ad_id", "Ad ID", 110),
            ("ad_creative_id", "Creative ID", 145),
            ("candidate_skus", "SKU 候选", 150),
            ("profile_id", "Profile ID", 135),
            ("campaign_id", "Campaign ID", 145),
            ("ad_group_id", "Ad Group ID", 130),
            ("cost", "Cost", 85),
            ("integrity", "归属 / 核对状态", 350),
        ):
            self.report_ad_detail_tree.heading(column, text=heading)
            self.report_ad_detail_tree.column(column, width=width, anchor="w")
        self.report_ad_detail_tree.grid(row=11, column=0, columnspan=5, sticky="nsew", padx=4, pady=4)
        ad_detail_scrollbar = ttk.Scrollbar(query_box, orient="vertical", command=self.report_ad_detail_tree.yview)
        ad_detail_scrollbar.grid(row=11, column=5, sticky="ns", pady=4)
        ad_detail_xscrollbar = ttk.Scrollbar(query_box, orient="horizontal", command=self.report_ad_detail_tree.xview)
        ad_detail_xscrollbar.grid(row=12, column=0, columnspan=5, sticky="ew", padx=4)
        self.report_ad_detail_tree.configure(
            yscrollcommand=ad_detail_scrollbar.set, xscrollcommand=ad_detail_xscrollbar.set
        )
        query_box.rowconfigure(8, weight=1)
        query_box.rowconfigure(11, weight=1)

        mode_box = ttk.LabelFrame(parent, text="D01 起算日期（手动确认）", padding=8)
        mode_box.grid(row=9, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Label(
            mode_box,
            text="每行填写 SKU=YYYY-MM-DD。日期按当前所选店铺的站点日期理解，并按店铺与 SKU 保存到本机。"
                 "程序不会从商品创建日、订单或销量推断 D01。",
            wraplength=850,
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=4, pady=(2, 5))
        self.d01_dates_text = tk.Text(mode_box, height=3, wrap="none")
        self.d01_dates_text.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4)
        ttk.Button(mode_box, text="确认并保存 D01", command=self._save_confirmed_d01).grid(row=1, column=2, sticky="w", padx=6)
        self.d01_status = tk.StringVar(value="请选择店铺并输入 SKU，加载或确认各自的 D01。")
        ttk.Label(mode_box, textvariable=self.d01_status, wraplength=850).grid(row=2, column=0, columnspan=3, sticky="w", padx=4, pady=(5, 0))
        mode_box.columnconfigure(1, weight=1)

        self.progress = ttk.Progressbar(parent, mode="indeterminate")
        self.progress.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        self.report_status = tk.StringVar(value="待处理。字段映射必须基于真实报表表头配置。")
        ttk.Label(parent, textvariable=self.report_status, wraplength=880).grid(row=11, column=0, columnspan=3, sticky="w", pady=4)
        buttons = ttk.Frame(parent)
        buttons.grid(row=12, column=0, columnspan=3, sticky="ew", pady=10)
        self.generate_button = ttk.Button(buttons, text="开始生成周报", command=self._start_generation)
        self.generate_button.pack(side="left")
        self.generate_api_button = ttk.Button(buttons, text="领星 OpenAPI 查询并生成", command=self._start_api_generation)
        self.generate_api_button.pack(side="left", padx=8)
        ttk.Button(buttons, text="打开输出文件夹", command=self._open_output_folder).pack(side="left", padx=8)
        ttk.Label(
            parent,
            text="程序不会自动推断导出字段；无法唯一匹配或缺少来源的数据留空，并写入校验日志。\n"
                 "新增周次按原模板七天一周延续。A2 推广阶段、G2 发货时间、M 列日目标销量由你手动维护。",
            justify="left",
        ).grid(row=13, column=0, columnspan=3, sticky="w", pady=8)

    def _build_mcp_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(13, weight=1)
        self.endpoint_var = tk.StringVar()
        self.mcp_key_var = tk.StringVar()
        self.tool_var = tk.StringVar()
        ttk.Label(parent, text="官方 Streamable HTTP 地址", width=25).grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.endpoint_var).grid(row=0, column=1, columnspan=3, sticky="ew", padx=6, pady=5)
        ttk.Label(parent, text="X-Mcp-Key", width=25).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.mcp_key_var, show="*").grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="安全保存本机凭证", command=self._save_mcp_settings).grid(row=1, column=3, padx=4, pady=5)
        action_row = ttk.Frame(parent)
        action_row.grid(row=2, column=0, columnspan=4, sticky="ew", pady=8)
        ttk.Button(action_row, text="测试连接", command=self._test_mcp_connection).pack(side="left")
        ttk.Button(action_row, text="获取工具列表", command=self._refresh_tools).pack(side="left", padx=8)
        self.mcp_status = tk.StringVar(value="尚未连接。地址需从领星官方说明取得。")
        ttk.Label(parent, textvariable=self.mcp_status, wraplength=880).grid(row=3, column=0, columnspan=4, sticky="w", pady=4)
        ttk.Label(parent, text="MCP 工具", width=25).grid(row=4, column=0, sticky="w", pady=5)
        self.tool_combo = ttk.Combobox(parent, textvariable=self.tool_var, state="readonly")
        self.tool_combo.grid(row=4, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="查看工具参数", command=self._show_selected_mcp_tool).grid(row=4, column=3, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="基础工具参数 JSON", width=25).grid(row=5, column=0, sticky="nw", pady=4)
        self.arguments_text = tk.Text(parent, height=2, wrap="none")
        self.arguments_text.insert("1.0", "{}")
        self.arguments_text.grid(row=5, column=1, columnspan=3, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="运行 readOnlyHint=true 基础只读工具", command=self._query_tool).grid(row=6, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(parent, text="search 参数 JSON", width=25).grid(row=7, column=0, sticky="nw", pady=4)
        self.search_arguments_text = tk.Text(parent, height=2, wrap="none")
        self.search_arguments_text.insert("1.0", "{}")
        self.search_arguments_text.grid(row=7, column=1, columnspan=3, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="搜索业务工具", command=self._search_business_tools).grid(row=8, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(parent, text="search 返回的业务工具", width=25).grid(row=9, column=0, sticky="w", pady=4)
        self.business_tool_var = tk.StringVar()
        self.business_tool_combo = ttk.Combobox(parent, textvariable=self.business_tool_var, state="readonly")
        self.business_tool_combo.grid(row=9, column=1, sticky="ew", padx=6, pady=4)
        business_actions = ttk.Frame(parent)
        business_actions.grid(row=9, column=2, columnspan=2, sticky="w", padx=6, pady=4)
        ttk.Button(business_actions, text="查看工具参数", command=self._show_selected_business_tool).pack(side="left")
        ttk.Button(business_actions, text="核验本次只读范围", command=self._verify_selected_business_tool).pack(side="left", padx=6)
        ttk.Label(parent, text="业务查询参数 JSON", width=25).grid(row=10, column=0, sticky="nw", pady=4)
        self.business_arguments_text = tk.Text(parent, height=3, wrap="none")
        self.business_arguments_text.insert("1.0", "{}")
        self.business_arguments_text.grid(row=10, column=1, columnspan=3, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="执行已验证的只读查询", command=self._execute_business_query).grid(row=11, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(parent, text="MCP 结果（仅显示在本机，不写入日志或文件）").grid(row=12, column=0, columnspan=4, sticky="w", pady=(6, 2))
        self.mcp_output = tk.Text(parent, height=14, wrap="word", state="disabled")
        self.mcp_output.grid(row=13, column=0, columnspan=4, sticky="nsew", pady=4)
        ttk.Label(
            parent,
            text="search 仅在工具说明明确表示它只发现业务工具/Schema 时用于目录搜索。action 只接受本次会话中已核验、ID 与 Schema 指纹完全匹配的只读查询；不符合规则的工具会被阻止。",
            wraplength=880,
        ).grid(row=14, column=0, columnspan=4, sticky="w", pady=6)

    def _build_openapi_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(7, weight=1)
        self.openapi_app_id_var = tk.StringVar()
        self.openapi_app_secret_var = tk.StringVar()

        ttk.Label(parent, text="AppID", width=22).grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.openapi_app_id_var).grid(row=0, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        ttk.Label(parent, text="AppSecret", width=22).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.openapi_app_secret_var, show="*").grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="安全保存凭证", command=self._save_openapi_credentials).grid(row=1, column=3, padx=4, pady=5)

        actions = ttk.Frame(parent)
        actions.grid(row=2, column=0, columnspan=4, sticky="ew", pady=8)
        ttk.Button(actions, text="测试 API 身份验证", command=self._test_openapi_auth).pack(side="left")
        ttk.Button(actions, text="刷新店铺列表", command=self._openapi_seller_list).pack(side="left", padx=8)

        ttk.Label(parent, text="授权店铺（业务码验证通过后动态映射）").grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 2))
        self.openapi_seller_tree = ttk.Treeview(
            parent,
            columns=("seller_name", "sid", "site", "ad_auth", "status"),
            show="headings",
            height=3,
        )
        for column, heading, width in (
            ("seller_name", "店铺名称", 260),
            ("sid", "SID", 150),
            ("site", "站点", 160),
            ("ad_auth", "has_ads_setting 原值", 180),
            ("status", "status 原值", 120),
        ):
            self.openapi_seller_tree.heading(column, text=heading)
            self.openapi_seller_tree.column(column, width=width, anchor="w")
        self.openapi_seller_tree.grid(row=4, column=0, columnspan=4, sticky="ew", pady=4)
        self.openapi_status = tk.StringVar(value="OpenAPI 地址固定为官方域名；密钥仅保存在 Windows Credential Manager。")
        ttk.Label(parent, textvariable=self.openapi_status, wraplength=900).grid(row=5, column=0, columnspan=4, sticky="w", pady=6)
        ttk.Label(parent, text="结果 / 说明（不显示 Token，不写入日志）").grid(row=6, column=0, columnspan=4, sticky="w", pady=(8, 2))
        self.openapi_output = tk.Text(parent, height=14, wrap="word", state="disabled")
        self.openapi_output.grid(row=7, column=0, columnspan=4, sticky="nsew", pady=4)
        ttk.Label(
            parent,
            text="已启用 Token 身份验证和只读店铺列表请求。SKU 与销量预览位于“生成周报”页；查询结果只显示，不写入工作簿。",
            wraplength=900,
        ).grid(row=8, column=0, columnspan=4, sticky="w", pady=6)

    def _save_openapi_credentials(self) -> None:
        app_id = self.openapi_app_id_var.get().strip()
        app_secret = self.openapi_app_secret_var.get()
        if not app_id or not app_secret:
            messagebox.showwarning("缺少凭证", "请填写 AppID 和 AppSecret。")
            return
        try:
            store = LingxingOpenApiCredentialStore()
            store.set("AppID", app_id)
            store.set("AppSecret", app_secret)
        except CredentialStoreError as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        self.openapi_app_id_var.set("")
        self.openapi_app_secret_var.set("")
        self.openapi_status.set("AppID 与 AppSecret 已保存到 Windows Credential Manager。")

    def _openapi_credentials(self) -> tuple[str, str]:
        store = LingxingOpenApiCredentialStore()
        app_id = store.get("AppID")
        app_secret = store.get("AppSecret")
        if not app_id or not app_secret:
            raise CredentialStoreError("Windows Credential Manager 中没有 OpenAPI 凭证，请先安全保存。")
        return app_id, app_secret

    def _test_openapi_auth(self) -> None:
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return

        def authenticate() -> dict[str, str]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.request_access_token()
            return {"message": token_result.safe_status()}

        self._start_task(authenticate, self.openapi_status, "API 身份验证完成。")

    def _apply_sellers(self, sellers: list[LingxingSeller]) -> None:
        previous = self._seller_choices.get(self.report_store_var.get())
        previous_sid = previous.sid if previous else None
        self._sellers = list(sellers)
        self._seller_choices = {seller.display_label: seller for seller in self._sellers}

        for item_id in self.openapi_seller_tree.get_children():
            self.openapi_seller_tree.delete(item_id)
        for index, seller in enumerate(self._sellers):
            self.openapi_seller_tree.insert(
                "",
                "end",
                iid=f"seller-{index}",
                values=(
                    seller.name,
                    seller.sid,
                    seller.site_label,
                    seller.has_ads_setting,
                    seller.status,
                ),
            )

        labels = tuple(self._seller_choices)
        self.report_store_combo.configure(values=labels)
        selected = next(
            (seller.display_label for seller in self._sellers if seller.sid == previous_sid),
            labels[0] if labels else "",
        )
        self.report_store_var.set(selected)
        if hasattr(self, "d01_status"):
            self._populate_d01_editor(preserve_current=bool(previous_sid and selected and self._seller_choices[selected].sid == previous_sid))
        if selected:
            self.report_status.set(f"已加载 {len(self._sellers)} 家授权店铺；所选店铺的 SID 已自动关联。")
        else:
            self.report_status.set("店铺列表为空；无法进行基于 OpenAPI 的店铺查询。")

    def _selected_report_seller(self) -> LingxingSeller | None:
        return self._seller_choices.get(self.report_store_var.get())

    def _selected_report_sid(self) -> int | None:
        """Resolve the selected display label to the API SID returned by Lingxing."""
        seller = self._selected_report_seller()
        return seller.sid if seller is not None else None

    def _on_report_store_selected(self, _event: Any = None) -> None:
        self._invalidate_hsa_candidate_review()
        self._populate_d01_editor(preserve_current=False)

    def _on_report_skus_changed(self, *_args: Any) -> None:
        self._invalidate_hsa_candidate_review()
        if not self._d01_refreshing:
            self._populate_d01_editor(preserve_current=True)

    def _invalidate_hsa_candidate_review(self, *_args: Any) -> None:
        self._pending_hsa_candidates = ()
        self._pending_hsa_candidate_context = None

    def _openapi_seller_list(self) -> None:
        if self._busy:
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return
        self._apply_sellers([])

        def query_sellers() -> dict[str, str]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.get_access_token()
                sellers = client.list_sellers(token_result.token)
            if not isinstance(sellers, list) or any(not isinstance(item, LingxingSeller) for item in sellers):
                raise LingxingApiError("店铺响应业务码或记录结构尚未完成官方契约核验，未更新店铺列表。")
            return {
                "message": f"领星店铺列表业务查询成功，共 {len(sellers)} 家店铺。",
                "output": f"已解析并显示 {len(sellers)} 家授权店铺；敏感账号字段不会显示。",
                "_sellers": sellers,
            }

        self._start_task(query_sellers, self.openapi_status, "店铺列表请求完成。")

    def _query_product_performance_preview(self) -> None:
        if self._busy:
            return
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先刷新并选择领星授权店铺。")
            return

        local_skus = list(dict.fromkeys(
            value.strip()
            for value in self.skus_var.get().replace("，", ",").split(",")
            if value.strip()
        ))
        if not 1 <= len(local_skus) <= 50:
            messagebox.showwarning("SKU 数量无效", "请输入 1–50 个不重复的本地 SKU。")
            return

        start_text = self.query_start_date_var.get().strip()
        end_text = self.query_end_date_var.get().strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_text) or start_text != end_text:
            messagebox.showwarning(
                "单日范围无效",
                "AsinListNew 预览测试只接受单日范围。请将查询开始日和结束日都设为同一个 YYYY-MM-DD 站点日期。",
            )
            return
        try:
            event_date = date.fromisoformat(start_text)
        except ValueError:
            messagebox.showwarning("日期格式错误", "请输入有效的 YYYY-MM-DD 站点日期。")
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return

        for item_id in self.product_performance_tree.get_children():
            self.product_performance_tree.delete(item_id)
        self.product_performance_status.set("正在按 local_sku + sku 汇总维度执行单日只读查询…")

        def query() -> dict[str, Any]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.get_access_token()
                rows = client.query_product_performance_single_day(
                    access_token=token_result.token,
                    store_sid=seller.sid,
                    local_skus=local_skus,
                    event_date=event_date,
                )
            return {
                "_product_performance_rows": rows,
                "_product_performance_skus": local_skus,
                "_product_performance_date": event_date,
                "_product_performance_store": seller.name,
                "_product_performance_sid": seller.sid,
            }

        self._start_task(
            query,
            self.product_performance_status,
            "AsinListNew 单日只读查询完成。",
        )

    def _query_product_performance_cost_diagnostic(self) -> None:
        if self._busy:
            return
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先刷新并选择领星授权店铺。")
            return

        local_skus = list(dict.fromkeys(
            value.strip()
            for value in self.skus_var.get().replace("，", ",").split(",")
            if value.strip()
        ))
        if len(local_skus) != 1:
            messagebox.showwarning("仅支持单 SKU 诊断", "请输入一个本地 SKU，以避免结果归属混淆。")
            return

        start_text = self.query_start_date_var.get().strip()
        end_text = self.query_end_date_var.get().strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_text) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_text):
            messagebox.showwarning("日期格式错误", "请输入 YYYY-MM-DD 站点日期。")
            return
        try:
            start_date = date.fromisoformat(start_text)
            end_date = date.fromisoformat(end_text)
        except ValueError:
            messagebox.showwarning("日期格式错误", "请输入有效的 YYYY-MM-DD 站点日期。")
            return
        if (end_date - start_date).days != 2:
            messagebox.showwarning(
                "诊断日期范围无效",
                "请设置连续三天，例如 2026-08-28 至 2026-08-30。程序将额外以 MSKU 维度查询中间日。",
            )
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return

        for item_id in self.product_performance_tree.get_children():
            self.product_performance_tree.delete(item_id)
        self.product_performance_status.set(
            "正在执行 4 组串行只读请求：三天 sku 汇总，以及中间日 msku 汇总…"
        )

        def query() -> dict[str, Any]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.get_access_token()
                rows = client.query_product_performance_cost_diagnostic(
                    access_token=token_result.token,
                    store_sid=seller.sid,
                    local_sku=local_skus[0],
                    start_date=start_date,
                    end_date=end_date,
                )
            return {
                "_product_performance_cost_diagnostic": True,
                "_product_performance_rows": rows,
                "_product_performance_skus": local_skus,
                "_product_performance_start_date": start_date,
                "_product_performance_end_date": end_date,
                "_product_performance_store": seller.name,
                "_product_performance_sid": seller.sid,
            }

        self._start_task(
            query,
            self.product_performance_status,
            "AsinListNew 日期与汇总维度诊断查询完成。",
        )

    def _apply_product_performance_preview(self, result: dict[str, Any]) -> None:
        rows: list[ProductPerformanceMetrics] = result["_product_performance_rows"]
        requested_skus: list[str] = result["_product_performance_skus"]
        event_date: date = result["_product_performance_date"]
        row_by_sku = {row.local_sku: row for row in rows}
        self.product_performance_status.set(
            f"OpenAPI 请求已成功（HTTP 2xx，业务 code=0）；店铺 {result['_product_performance_store']} "
            f"（本次 SID {result['_product_performance_sid']}），站点日期筛选 {event_date.isoformat()}，"
            f"返回 {len(rows)} 条 SKU 汇总记录。结果只在本机预览，不写入工作簿或日志。"
            "响应行没有日期字段；需与领星同日数据核对后才能判断为日值。"
        )
        for sku in requested_skus:
            row = row_by_sku.get(sku)
            if row is None:
                self.product_performance_tree.insert(
                    "", "end", values=(
                        event_date.isoformat(), "sku", sku, sku,
                        "记录状态", "未返回商品记录（不是零值）", "", "查询范围内没有对应 data.list 记录",
                    )
                )
                continue
            for field_name in PRODUCT_PERFORMANCE_PREVIEW_FIELDS:
                self.product_performance_tree.insert(
                    "",
                    "end",
                    values=(
                        event_date.isoformat(),
                        row.summary_field,
                        row.summary_key or "",
                        row.local_sku,
                        _PRODUCT_PERFORMANCE_FIELD_LABELS.get(field_name, field_name),
                        _product_performance_value(row, field_name),
                        row.currency_code or "",
                        _product_performance_scope_note(field_name),
                    ),
                )

    def _apply_product_performance_cost_diagnostic(self, result: dict[str, Any]) -> None:
        rows: list[ProductPerformanceMetrics] = result["_product_performance_rows"]
        start_date: date = result["_product_performance_start_date"]
        end_date: date = result["_product_performance_end_date"]
        sku = result["_product_performance_skus"][0]
        middle_date = start_date + (end_date - start_date) / 2
        self.product_performance_status.set(
            f"4 组只读查询成功（每页均通过业务 code=0）：店铺 {result['_product_performance_store']} "
            f"（SID {result['_product_performance_sid']}），SKU {sku}；sku 汇总日期 "
            f"{start_date.isoformat()}、{middle_date.isoformat()}、{end_date.isoformat()}，另查询 "
            f"{middle_date.isoformat()} 的 msku 汇总。响应行不含报表日期；日期仅代表请求筛选条件，"
            "数值变化可帮助排查，但不能单独证明费用字段的统计日期或分摊口径。结果只在本机预览。"
        )
        rows_by_query: dict[tuple[date, str], list[ProductPerformanceMetrics]] = {}
        for row in rows:
            rows_by_query.setdefault((row.start_date, row.summary_field), []).append(row)

        for event_date, summary_field in (
            (start_date, "sku"),
            (middle_date, "sku"),
            (end_date, "sku"),
            (middle_date, "msku"),
        ):
            query_rows = rows_by_query.get((event_date, summary_field), [])
            if not query_rows:
                self.product_performance_tree.insert(
                    "", "end", values=(
                        event_date.isoformat(), summary_field, sku, sku,
                        "记录状态", "未返回商品记录（不是零值）", "", "本次请求没有对应 data.list 记录",
                    )
                )
                continue
            for row in query_rows:
                for field_name in _PRODUCT_PERFORMANCE_COST_DIAGNOSTIC_FIELDS:
                    scope = (
                        "请求日期筛选；Schema 未定义该费用字段的日期/分摊口径，需与后台同日核对"
                        if field_name != "currency_code"
                        else "接口币种字段；须与领星后台金额币种一致"
                    )
                    self.product_performance_tree.insert(
                        "", "end", values=(
                            event_date.isoformat(),
                            summary_field,
                            row.summary_key or "",
                            row.local_sku,
                            _PRODUCT_PERFORMANCE_FIELD_LABELS.get(field_name, field_name),
                            _product_performance_value(row, field_name),
                            row.currency_code or "",
                            scope,
                        ),
                    )

    def _query_report_sku_data(self) -> None:
        if self._busy:
            return
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先刷新并选择领星授权店铺。")
            return

        local_skus = [
            value.strip()
            for value in self.skus_var.get().replace("，", ",").split(",")
            if value.strip()
        ]
        local_skus = list(dict.fromkeys(local_skus))
        if not 1 <= len(local_skus) <= 50:
            messagebox.showwarning("SKU 数量无效", "请输入 1–50 个不重复的本地 SKU，以逗号分隔。")
            return

        start_text = self.query_start_date_var.get().strip()
        end_text = self.query_end_date_var.get().strip()
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_text) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_text):
            messagebox.showwarning("日期格式错误", "请输入 YYYY-MM-DD 格式的站点日期。")
            return
        try:
            start_date = date.fromisoformat(start_text)
            end_date = date.fromisoformat(end_text)
        except ValueError:
            messagebox.showwarning("日期格式错误", "开始日和结束日必须是有效的 YYYY-MM-DD 日期。")
            return
        if end_date < start_date:
            messagebox.showwarning("日期范围错误", "结束日期不能早于开始日期。")
            return
        if (end_date - start_date).days + 1 > 92:
            messagebox.showwarning("日期范围过长", "产品表现接口的日期范围上限为 92 个自然日（含首尾）。")
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return

        for item_id in self.report_sales_tree.get_children():
            self.report_sales_tree.delete(item_id)
        self.report_mapping_status.set("正在查询所选店铺的 SKU 商品映射…")

        def query() -> dict[str, Any]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.get_access_token()
                mappings = client.query_product_mappings(
                    access_token=token_result.token,
                    store_sid=seller.sid,
                    local_skus=local_skus,
                    start_date=start_date,
                    end_date=end_date,
                )
                mapping_by_sku = {mapping.local_sku: mapping for mapping in mappings}
                daily_mappings = [mapping for mapping in mappings if mapping.seller_skus]
                unresolved_skus = [
                    sku for sku in local_skus
                    if sku not in mapping_by_sku or not mapping_by_sku[sku].seller_skus
                ]
                rows = (
                    client.query_daily_sales(
                        access_token=token_result.token,
                        store_sid=seller.sid,
                        mappings=daily_mappings,
                        start_date=start_date,
                        end_date=end_date,
                    )
                    if daily_mappings
                    else []
                )
            return {
                "_daily_rows": rows,
                "_mappings": mappings,
                "_requested_skus": local_skus,
                "_unresolved_skus": unresolved_skus,
                "_start_date": start_date,
                "_end_date": end_date,
                "_store_name": seller.name,
            }

        self._start_task(query, self.report_query_status, "SKU 每日数据查询完成。")

    def _query_report_ad_data(self) -> None:
        if self._busy:
            return
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先刷新并选择领星授权店铺。")
            return
        local_skus = self._report_skus()
        if not 1 <= len(local_skus) <= 50:
            messagebox.showwarning("SKU 数量无效", "请输入 1–50 个本地 SKU。")
            return
        start_text = self.query_start_date_var.get().strip()
        end_text = self.query_end_date_var.get().strip()
        try:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_text) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_text):
                raise ValueError
            start_date = date.fromisoformat(start_text)
            end_date = date.fromisoformat(end_text)
        except ValueError:
            messagebox.showwarning("日期格式错误", "广告报表日期请输入有效的 YYYY-MM-DD。")
            return
        if end_date < start_date:
            messagebox.showwarning("日期范围错误", "结束日期不能早于开始日期。")
            return
        if (end_date - start_date).days + 1 > 92:
            messagebox.showwarning("日期范围过长", "为控制请求量，广告预览一次最多查询 92 个日期；遇到 429 会停止且不自动重试。")
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return
        for tree in (self.report_ad_summary_tree, self.report_ad_detail_tree):
            for item_id in tree.get_children():
                tree.delete(item_id)
        self.report_ad_status.set("正在读取 SP 和 SD 商品报表并逐日核对；SB/SBV 不在 V1.0 统计范围……")

        def query() -> dict[str, Any]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.get_access_token()
                mappings = client.query_product_mappings(
                    access_token=token_result.token,
                    store_sid=seller.sid,
                    local_skus=local_skus,
                    start_date=start_date,
                    end_date=end_date,
                )
                by_sku = {item.local_sku.casefold(): item for item in mappings}
                missing = [
                    sku for sku in local_skus
                    if sku.casefold() not in by_sku
                    or not (by_sku[sku.casefold()].seller_skus or by_sku[sku.casefold()].asins)
                ]
                if missing:
                    raise LingxingApiError("以下 SKU 没有可用于广告识别的 MSKU/ASIN 映射：" + ", ".join(missing))
                selected_mappings = [by_sku[sku.casefold()] for sku in local_skus]
                windows_by_type: dict[str, list[Any]] = {"SP": [], "SD": []}
                errors: dict[str, str] = {}
                for kind, query_method in (
                    ("SP", client.query_sp_ad_spend_history),
                    ("SD", client.query_sd_ad_spend_history),
                ):
                    try:
                        windows_by_type[kind] = query_method(
                            access_token=token_result.token,
                            store_sid=seller.sid,
                            mappings=selected_mappings,
                            start_date=start_date,
                            end_date=end_date,
                            **({"include_catalog": False} if kind == "SD" else {}),
                        )
                    except LingxingApiError as exc:
                        errors[kind] = f"查询失败：{exc}"
            return {
                "_ad_windows_by_type": windows_by_type,
                "_ad_mappings": mappings,
                "_ad_query_errors": errors,
                "_ad_store_sid": seller.sid,
                "_ad_requested_skus": local_skus,
                "_ad_store_name": seller.name,
                "_ad_start": start_date,
                "_ad_end": end_date,
                # This preview does not query the daily sales endpoint for a
                # currency check. Keep component costs visible, but do not
                # present a combined amount without currency/date evidence.
                "_ad_currency_and_date_basis_verified": False,
                "message": "SP+SD 广告费用预览查询完成；缺少同店铺币种核验时合计保持空白。",
            }

        self._start_task(query, self.report_ad_status, "SP+SD 广告费用查询完成。")

    def _review_hsa_candidates(self) -> None:
        """Review all unique current-catalog candidates in one date-scoped batch."""
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先选择要核验的领星店铺。")
            return
        candidates = self._pending_hsa_candidates
        if not candidates:
            messagebox.showinfo(
                "没有可批量核验的候选",
                "当前预览没有同时满足目录完整、ASIN 唯一映射和活动成本核平条件的 HSA 候选组。\n"
                "未匹配或共享候选仍留在未归属状态，不会猜测归属。",
            )
            return
        try:
            current_context = (
                seller.sid,
                tuple(sorted(sku.casefold() for sku in self._report_skus())),
                date.fromisoformat(self.query_start_date_var.get().strip()),
                date.fromisoformat(self.query_end_date_var.get().strip()),
            )
        except ValueError:
            current_context = None
        if current_context != self._pending_hsa_candidate_context:
            self._invalidate_hsa_candidate_review()
            messagebox.showwarning(
                "预览范围已变化",
                "店铺、SKU 或日期与产生这些候选时不同。请重新查询广告数据后再核验，避免将候选保存到错误范围。",
            )
            return

        dialog = tk.Toplevel(self)
        dialog.title("批量核验 HSA 广告组 SKU 候选")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("1050x520")
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)
        ttk.Label(
            dialog,
            text=(
                f"找到 {len(candidates)} 个唯一 SKU 候选广告组。候选来自当前 HSA 创意目录的 Creative ID / ASIN 与本地 SKU 映射；"
                "目录没有历史有效期，快照之间也不代表关系连续。请逐项核对后，可一次确认多个广告组。"
            ),
            wraplength=1000,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", padx=10, pady=8)
        columns = ("profile_id", "campaign_id", "ad_group_id", "creative_ids", "asins", "candidate_sku", "observed_days")
        tree = ttk.Treeview(dialog, columns=columns, show="headings", selectmode="extended", height=12)
        labels = (
            ("profile_id", "Profile ID", 120),
            ("campaign_id", "Campaign ID", 150),
            ("ad_group_id", "Ad Group ID", 145),
            ("creative_ids", "Creative ID", 170),
            ("asins", "目录 ASIN", 200),
            ("candidate_sku", "候选本地 SKU", 150),
            ("observed_days", "查询中出现日期", 150),
        )
        for column, heading, width in labels:
            tree.heading(column, text=heading)
            tree.column(column, width=width, anchor="w")
        tree.grid(row=1, column=0, sticky="nsew", padx=10, pady=4)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=4)
        tree.configure(yscrollcommand=scrollbar.set)
        candidate_by_iid: dict[str, dict[str, Any]] = {}
        for index, candidate in enumerate(candidates):
            iid = f"hsa_candidate_{index}"
            tree.insert("", "end", iid=iid, values=(
                candidate["profile_id"], candidate["campaign_id"], candidate["ad_group_id"],
                ", ".join(candidate["creative_ids"]), ", ".join(candidate["asins"]),
                candidate["local_sku"], ", ".join(candidate["query_dates"]),
            ))
            candidate_by_iid[iid] = candidate
            tree.selection_add(iid)

        range_box = ttk.Frame(dialog)
        range_box.grid(row=2, column=0, sticky="ew", padx=10, pady=8)
        ttk.Label(range_box, text="人工核验的有效开始日期（YYYY-MM-DD）").grid(row=0, column=0, sticky="w")
        start_var = tk.StringVar()
        ttk.Entry(range_box, textvariable=start_var, width=14).grid(row=0, column=1, padx=(6, 18), sticky="w")
        ttk.Label(range_box, text="有效结束日期（YYYY-MM-DD）").grid(row=0, column=2, sticky="w")
        end_var = tk.StringVar()
        ttk.Entry(range_box, textvariable=end_var, width=14).grid(row=0, column=3, padx=6, sticky="w")
        ttk.Label(
            dialog,
            text="日期范围不会从目录快照或查询范围自动推断。只有你确认所选各组在整个填写范围内均只推广对应候选 SKU 时，才保存确认。",
            wraplength=1000,
            justify="left",
        ).grid(row=3, column=0, sticky="w", padx=10, pady=(0, 8))

        def save_selected() -> None:
            selected_ids = tree.selection()
            if not selected_ids:
                messagebox.showwarning("未选择候选", "请选择要一起核验的 HSA 广告组。", parent=dialog)
                return
            try:
                valid_from = date.fromisoformat(start_var.get().strip())
                valid_through = date.fromisoformat(end_var.get().strip())
            except ValueError:
                messagebox.showwarning("日期格式错误", "请填写有效的 YYYY-MM-DD 起止日期。", parent=dialog)
                return
            if valid_through < valid_from:
                messagebox.showwarning("日期范围错误", "有效结束日期不能早于开始日期。", parent=dialog)
                return
            selected_candidates = [candidate_by_iid[item] for item in selected_ids]
            mapping_preview = "\n".join(
                f"Campaign {candidate['campaign_id']} / Ad Group {candidate['ad_group_id']} → {candidate['local_sku']}"
                for candidate in selected_candidates
            )
            if not messagebox.askyesno(
                "确认批量 HSA 归属",
                f"确认以下 {len(selected_candidates)} 个广告组在 {valid_from.isoformat()} 至 "
                f"{valid_through.isoformat()}（含首尾）均只推广对应 SKU？\n\n{mapping_preview}\n\n"
                "当前目录本身不证明历史归属。若任何组推广多个 SKU 或有效期不确定，请取消或缩小选择/日期范围。",
                parent=dialog,
            ):
                return
            mappings = [HsaAdGroupSkuMapping(
                store_sid=seller.sid,
                profile_id=candidate["profile_id"],
                campaign_id=candidate["campaign_id"],
                ad_group_id=candidate["ad_group_id"],
                local_sku=candidate["local_sku"],
                valid_from=valid_from,
                valid_through=valid_through,
            ) for candidate in selected_candidates]
            try:
                path = save_hsa_ad_group_mappings(mappings)
            except HsaAttributionSettingsError as exc:
                messagebox.showerror("保存失败", str(exc), parent=dialog)
                return
            self._pending_hsa_candidates = ()
            self._pending_hsa_candidate_context = None
            self.report_ad_status.set(
                f"已在本机批量保存 {len(mappings)} 条 HSA 广告组归属确认；重新查询后按有效期应用。文件：{path.name}。"
            )
            dialog.destroy()
            messagebox.showinfo(
                "归属确认已保存",
                f"已在本机保存 {len(mappings)} 条 HSA 广告组确认。请重新查询，再检查活动与广告组核平及归属结果。",
            )

        buttons = ttk.Frame(dialog)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", padx=10, pady=8)
        ttk.Button(buttons, text="确认所选候选组", command=save_selected).pack(side="left", padx=4)
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=4)
        self.wait_window(dialog)

    def _apply_ad_spend_preview(self, result: dict[str, Any]) -> None:
        summaries, details = build_ad_spend_preview(
            windows_by_type=result["_ad_windows_by_type"],
            mappings=result["_ad_mappings"],
            store_sid=result["_ad_store_sid"],
            requested_skus=result["_ad_requested_skus"],
            start_date=result["_ad_start"],
            end_date=result["_ad_end"],
            scope=SP_SD_ONLY_CONFIRMED,
            currency_and_date_basis_verified=result["_ad_currency_and_date_basis_verified"],
        )
        self._pending_hsa_candidates = ()
        self._pending_hsa_candidate_context = None
        for row in summaries:
            self.report_ad_summary_tree.insert(
                "", "end", values=(
                    row.report_date, row.local_sku, row.sp_cost, row.sd_cost,
                    row.confirmed_total, row.unattributed_cost, row.integrity,
                )
            )
        for row in details:
            self.report_ad_detail_tree.insert(
                "", "end", values=(
                    row.report_date, row.ad_type, row.report_level, row.local_sku,
                    row.report_sku, row.asin, row.ad_id, row.ad_creative_id,
                    ", ".join(row.candidate_local_skus), row.profile_id, row.campaign_id,
                    row.ad_group_id, row.cost, row.integrity,
                )
            )
        windows_by_type = result["_ad_windows_by_type"]
        complete_by_type = {
            kind: sum(1 for window in windows_by_type[kind] if window.report_complete)
            for kind in ("SP", "SD")
        }
        discrepancies = {
            kind: sum(len(window.discrepancies) for window in windows_by_type[kind])
            for kind in ("SP", "SD")
        }
        report_row_counts = {
            "SP": (
                sum(len(window.campaign_rows) for window in windows_by_type["SP"]),
                sum(len(window.product_rows) for window in windows_by_type["SP"]),
            ),
            "SD": (
                sum(len(window.campaign_rows) for window in windows_by_type["SD"]),
                sum(len(window.product_rows) for window in windows_by_type["SD"]),
            ),
        }
        errors = result.get("_ad_query_errors", {})
        error_text = "；".join(f"{kind}：{message}" for kind, message in sorted(errors.items()))
        self.report_ad_status.set(
            f"店铺 {result['_ad_store_name']}：{result['_ad_start'].isoformat()} 至 "
            f"{result['_ad_end'].isoformat()}（报表日期）。SP/SD 已核平日期数="
            f"{complete_by_type['SP']}/{complete_by_type['SD']}；"
            f"差异 campaign 数 SP/SD={discrepancies['SP']}/{discrepancies['SD']}。"
            f"接口成功返回的活动/明细行数 SP={report_row_counts['SP'][0]}/{report_row_counts['SP'][1]}、"
            f"SD={report_row_counts['SD'][0]}/{report_row_counts['SD'][1]}。"
            "商品报表中的 MSKU/ASIN 精确匹配 SKU；活动级金额仅用于核对，不重复相加。普通 SB/SBV 不属于 V1.0 范围。"
            f"{error_text}"
        )

    def _apply_daily_sales_preview(self, result: dict[str, Any]) -> None:
        mappings: list[SkuProductMapping] = result["_mappings"]
        rows: list[SkuDailySales] = result["_daily_rows"]
        requested_skus: list[str] = result["_requested_skus"]
        unresolved_skus: list[str] = result["_unresolved_skus"]
        start_date: date = result["_start_date"]
        end_date: date = result["_end_date"]

        summary_parts: list[str] = []
        mapping_by_sku = {mapping.local_sku: mapping for mapping in mappings}
        for sku in requested_skus:
            mapping = mapping_by_sku.get(sku)
            if mapping is None:
                summary_parts.append(f"{sku}：未找到本店商品映射")
                continue
            msku_text = ", ".join(mapping.seller_skus) if mapping.seller_skus else "无已确认 MSKU"
            asin_text = ", ".join(mapping.asins) if mapping.asins else "无 ASIN 候选"
            summary_parts.append(f"{sku}：MSKU [{msku_text}]；本店 ASIN 候选 [{asin_text}]")
        self.report_mapping_status.set(
            "商品映射（ASIN 候选未假定与某个 MSKU 一一对应）：" + "；".join(summary_parts)
            if summary_parts
            else "产品表现接口没有返回所请求 SKU 的映射。"
        )

        for row in rows:
            self.report_sales_tree.insert(
                "",
                "end",
                values=(
                    row.event_date.isoformat(),
                    row.local_sku,
                    row.seller_sku,
                    row.asin,
                    "" if row.units_sold is None else str(row.units_sold),
                    "" if row.order_count is None else str(row.order_count),
                    "" if row.sales_amount is None else format(row.sales_amount, "f"),
                    row.currency_code,
                ),
            )

        first_positive_day: dict[str, date] = {}
        for row in rows:
            if row.units_sold is not None and row.units_sold > 0:
                prior = first_positive_day.get(row.local_sku)
                if prior is None or row.event_date < prior:
                    first_positive_day[row.local_sku] = row.event_date
        first_positive_text = "；".join(
            f"{sku}={first_positive_day[sku].isoformat()}"
            for sku in requested_skus
            if sku in first_positive_day
        ) or "所选范围未发现明确正销量记录"
        populated_metric_count = sum(
            value is not None
            for row in rows
            for value in (row.sales_amount, row.units_sold, row.order_count)
        )
        status_parts = [
            f"店铺 {result['_store_name']}：{start_date.isoformat()} 至 {end_date.isoformat()}（站点日期）",
            f"预览列出 {len(rows)} 个 SKU×MSKU×日期位置，接口明确返回 {populated_metric_count} 项指标值",
            f"范围内首个明确正销量日：{first_positive_text}。这不是已核验的历史首单日期。",
            "本次结果仅供核对，不写入周报；无记录留空，明确零值显示为 0。",
        ]
        if unresolved_skus:
            status_parts.append("未能唯一映射的 SKU：" + ", ".join(unresolved_skus))
        self.report_query_status.set(" ".join(status_parts))

    def _show_openapi_result(self, content: str) -> None:
        self.openapi_output.configure(state="normal")
        self.openapi_output.delete("1.0", "end")
        self.openapi_output.insert("1.0", content)
        self.openapi_output.configure(state="disabled")

    def _load_mcp_settings(self) -> None:
        try:
            settings = load_settings()
            self.endpoint_var.set(settings.get("endpoint", ""))
        except LocalSettingsError as exc:
            self.mcp_status.set(str(exc))

    def _save_mcp_settings(self) -> None:
        try:
            endpoint = save_settings(endpoint=self.endpoint_var.get())
            key = self.mcp_key_var.get()
            if key:
                WindowsCredentialStore().set("X-Mcp-Key", key)
                self.mcp_key_var.set("")
            self.mcp_status.set(f"本机设置已保存：{endpoint}；密钥未写入设置文件。")
        except (LocalSettingsError, CredentialStoreError) as exc:
            messagebox.showerror("保存失败", str(exc))

    def _credentials(self) -> tuple[str, str]:
        endpoint = validate_mcp_endpoint(self.endpoint_var.get())
        store = WindowsCredentialStore()
        key = self.mcp_key_var.get() or store.get("X-Mcp-Key")
        if not key:
            raise CredentialStoreError("Windows Credential Manager 中没有 X-Mcp-Key，请先安全保存。")
        return endpoint, key

    def _start_task(self, function, status_var: tk.StringVar, success_title: str) -> None:
        if self._busy:
            return
        self._busy = True
        self.generate_button.configure(state="disabled")
        self.generate_api_button.configure(state="disabled")
        self.report_query_button.configure(state="disabled")
        self.product_performance_query_button.configure(state="disabled")
        self.product_performance_cost_diagnostic_button.configure(state="disabled")
        self.report_ad_query_button.configure(state="disabled")
        self.progress.start(10)
        status_var.set("处理中…")

        def worker() -> None:
            try:
                result = function()
                self._events.put(("success", (status_var, success_title, result)))
            except Exception as exc:
                # Display only controlled exception messages. Never log args or request headers.
                self._events.put(("error", (status_var, str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_events(self) -> None:
        finished = False
        try:
            while True:
                kind, payload = self._events.get_nowait()
                if kind == "success":
                    status_var, title, result = payload
                    status_var.set(title)
                    if status_var is self.mcp_status:
                        if isinstance(result, dict) and result.get("_tools") is not None:
                            self._tools = result["_tools"]
                        if isinstance(result, dict) and result.get("_business_tools") is not None:
                            self._business_tools = result["_business_tools"]
                            self.business_tool_combo.configure(
                                values=tuple(
                                    f"{tool.tool_id} — {tool.name} — {tool.description[:100]}"
                                    for tool in self._business_tools
                                )
                            )
                            if self._business_tools:
                                self.business_tool_combo.current(0)
                        if isinstance(result, dict) and result.get("_tool_names") is not None:
                            self.tool_combo.configure(values=result["_tool_names"])
                            if result["_tool_names"]:
                                self.tool_combo.current(0)
                        self._show_mcp_result(result)
                    elif status_var is self.openapi_status and isinstance(result, dict):
                        status_var.set(result.get("message", title))
                        if result.get("_sellers") is not None:
                            self._apply_sellers(result["_sellers"])
                        self._show_openapi_result(result.get("output", status_var.get()))
                    elif status_var is self.report_query_status and isinstance(result, dict):
                        self._apply_daily_sales_preview(result)
                    elif status_var is self.product_performance_status and isinstance(result, dict):
                        if result.get("_product_performance_cost_diagnostic"):
                            self._apply_product_performance_cost_diagnostic(result)
                        else:
                            self._apply_product_performance_preview(result)
                    elif status_var is self.report_ad_status and isinstance(result, dict) and result.get("_ad_windows_by_type") is not None:
                        self._apply_ad_spend_preview(result)
                    elif isinstance(result, dict):
                        status_var.set(result.get("message", title))
                        if result.get("output"):
                            messagebox.showinfo("周报生成完成", result["message"])
                else:
                    status_var, error = payload
                    if status_var is self.report_ad_status:
                        status_var.set("广告查询失败；SP+SD 花费保持未知，不作为零值或正式周报数据。")
                    elif status_var is self.product_performance_status:
                        status_var.set("AsinListNew 查询失败；没有数据写入周报或日志。")
                    else:
                        status_var.set("处理失败。")
                    messagebox.showerror("处理失败", error)
                finished = True
        except queue.Empty:
            pass
        if finished and self._busy:
            self._busy = False
            self.progress.stop()
            self.generate_button.configure(state="normal")
            self.generate_api_button.configure(state="normal")
            self.report_query_button.configure(state="normal")
            self.product_performance_query_button.configure(state="normal")
            self.product_performance_cost_diagnostic_button.configure(state="normal")
            self.report_ad_query_button.configure(state="normal")
        self.after(150, self._poll_events)

    def _show_mcp_result(self, value: Any) -> None:
        if isinstance(value, dict) and "_output_text" in value:
            content = value["_output_text"]
            self.mcp_status.set(value.get("_status", "MCP 查询完成。"))
        else:
            content = json.dumps(redact_sensitive(value), ensure_ascii=False, indent=2, default=str)
        self.mcp_output.configure(state="normal")
        self.mcp_output.delete("1.0", "end")
        max_length = 80_000
        self.mcp_output.insert("1.0", content[:max_length])
        if len(content) > max_length:
            self.mcp_output.insert("end", "\n…结果过长，界面仅显示前 80,000 个字符；完整结果不会保存。")
        self.mcp_output.configure(state="disabled")

    def _run_mcp(
        self,
        endpoint: str,
        key: str,
        action: str,
        *,
        tool_name: str | None = None,
        arguments: dict[str, Any] | None = None,
        search_arguments: dict[str, Any] | None = None,
        business_tool: BusinessTool | None = None,
        grant: VerifiedReadOnlyGrant | None = None,
    ):
        client = LingxingMcpClient(endpoint, key)
        try:
            init = client.connect()
            if action == "connect":
                return {"_status": f"连接成功；协议版本 {init.get('protocolVersion')}。", "_output_text": "MCP 初始化握手成功。"}
            tools = client.list_tools()
            if action == "tools":
                display = [
                    {"name": tool.name, "description": tool.description, "annotations": tool.annotations, "inputSchema": tool.input_schema}
                    for tool in tools
                ]
                tool_names = tuple(
                    f"{tool.name} [{self._tool_annotation_label(tool)}] — {tool.description[:80]}"
                    for tool in tools
                )
                safe_search_count = sum(is_directory_search_tool(tool) for tool in tools)
                read_only_count = sum(tool.read_only for tool in tools)
                return {
                    "_status": f"已获取 {len(tools)} 个工具；readOnlyHint=true {read_only_count} 个；安全目录发现工具 {safe_search_count} 个。",
                    "_output_text": json.dumps(display, ensure_ascii=False, indent=2),
                    "_tool_names": tool_names,
                    "_tools": tools,
                }
            if action == "search_business_tools":
                search_tool = next((tool for tool in tools if tool.name.casefold() == "search"), None)
                if search_tool is None:
                    raise McpClientError("工具列表中没有名为 search 的发现工具。")
                result = client.call_directory_search(search_tool, search_arguments or {})
                business_tools = extract_business_tools(result)
                return {
                    "_status": f"search 已返回结果；解析出 {len(business_tools)} 个带 Schema 的业务工具。未解析的内容可在结果区查看。",
                    "_output_text": json.dumps(result, ensure_ascii=False, indent=2, default=str),
                    "_business_tools": business_tools,
                }
            if action == "business_query":
                if business_tool is None or grant is None:
                    raise McpClientError("缺少本次会话核验的业务工具许可，已阻止 action 调用。")
                search_tool = next((tool for tool in tools if tool.name.casefold() == "search"), None)
                action_tool = next((tool for tool in tools if is_action_wrapper_tool(tool)), None)
                if search_tool is None or action_tool is None:
                    raise McpClientError("工具列表缺少 search 或 action，已阻止查询。")
                fresh_result = client.call_directory_search(search_tool, search_arguments or {})
                fresh_tools = extract_business_tools(fresh_result)
                fresh_spec = next((item for item in fresh_tools if item.tool_id == business_tool.tool_id), None)
                if fresh_spec is None or fresh_spec.permission_fingerprint != business_tool.permission_fingerprint:
                    raise McpClientError("search 返回的业务工具 ID/Schema 已变化，旧许可失效；请重新搜索并核验。")
                result = client.call_verified_business_query(action_tool, fresh_spec, arguments or {}, grant)
                return {
                    "_status": f"已执行已核验的只读查询：{fresh_spec.tool_id}。结果仅显示在本机。",
                    "_output_text": json.dumps(result, ensure_ascii=False, indent=2, default=str),
                }
            selected = next((tool for tool in tools if tool.name == tool_name), None)
            if selected is None:
                raise McpClientError("工具列表中找不到所选工具，请先刷新列表。")
            result = client.call_read_only_tool(selected, arguments or {})
            return {"_status": f"只读查询完成：{selected.name}。结果仅显示在本机界面。", "_output_text": json.dumps(result, ensure_ascii=False, indent=2, default=str)}
        finally:
            client.close()

    @staticmethod
    def _tool_annotation_label(tool: McpTool) -> str:
        value = tool.annotations.get("readOnlyHint")
        if value is True:
            return "readOnlyHint=true"
        if value is False:
            return "readOnlyHint=false"
        return "未提供 readOnlyHint"

    def _start_mcp_action(
        self,
        action: str,
        *,
        tool_name: str | None = None,
        arguments: dict[str, Any] | None = None,
        search_arguments: dict[str, Any] | None = None,
        business_tool: BusinessTool | None = None,
        grant: VerifiedReadOnlyGrant | None = None,
    ) -> None:
        try:
            endpoint, key = self._credentials()
        except (CredentialStoreError, LocalSettingsError) as exc:
            messagebox.showerror("MCP 配置错误", str(exc))
            return
        self._start_task(
            lambda: self._run_mcp(
                endpoint,
                key,
                action,
                tool_name=tool_name,
                arguments=arguments,
                search_arguments=search_arguments,
                business_tool=business_tool,
                grant=grant,
            ),
            self.mcp_status,
            "MCP 操作完成。",
        )

    def _test_mcp_connection(self) -> None:
        self._start_mcp_action("connect")

    def _refresh_tools(self) -> None:
        self._start_mcp_action("tools")

    def _query_tool(self) -> None:
        index = self.tool_combo.current()
        if index < 0 or index >= len(self._tools):
            messagebox.showwarning("请选择工具", "请先获取工具列表并选择工具。")
            return
        tool = self._tools[index]
        if not tool.read_only:
            messagebox.showwarning("工具未标记只读", "基础工具调用仍要求服务端明确标记 readOnlyHint=true。业务 action 请先使用搜索并核验流程。")
            return
        try:
            args = json.loads(self.arguments_text.get("1.0", "end").strip() or "{}")
        except json.JSONDecodeError as exc:
            messagebox.showerror("参数格式错误", f"JSON 第 {exc.lineno} 行：{exc.msg}")
            return
        if not isinstance(args, dict):
            messagebox.showerror("参数格式错误", "查询参数必须是 JSON 对象。")
            return
        self._start_mcp_action("query", tool_name=tool.name, arguments=args)

    def _show_selected_mcp_tool(self) -> None:
        index = self.tool_combo.current()
        if index < 0 or index >= len(self._tools):
            messagebox.showwarning("请选择工具", "请先获取工具列表。")
            return
        tool = self._tools[index]
        self._show_json_window(
            f"MCP 工具：{tool.name}",
            {"name": tool.name, "description": tool.description, "annotations": tool.annotations, "inputSchema": tool.input_schema},
        )

    def _selected_business_tool(self) -> BusinessTool | None:
        index = self.business_tool_combo.current()
        if index < 0 or index >= len(self._business_tools):
            return None
        return self._business_tools[index]

    def _show_selected_business_tool(self) -> None:
        tool = self._selected_business_tool()
        if tool is None:
            messagebox.showwarning("请选择业务工具", "请先点击“搜索业务工具”并选择一项结果。")
            return
        self._viewed_business_fingerprints.add(tool.permission_fingerprint)
        self._show_json_window(
            f"业务工具参数：{tool.tool_id}",
            {"id": tool.tool_id, "name": tool.name, "description": tool.description, "inputSchema": tool.input_schema},
        )

    def _verify_selected_business_tool(self) -> None:
        tool = self._selected_business_tool()
        if tool is None:
            messagebox.showwarning("请选择业务工具", "请先搜索并选择一项业务工具。")
            return
        if tool.permission_fingerprint not in self._viewed_business_fingerprints:
            messagebox.showwarning("请先查看 Schema", "请先打开并检查当前业务工具的名称、说明和完整参数 Schema，再登记只读范围。")
            return
        if not is_read_only_query_candidate(tool):
            messagebox.showwarning("不能核验为只读", "该工具缺少明确查询描述，或 ID/描述包含写入操作，系统不会为它创建只读许可。")
            return
        approved = messagebox.askyesno(
            "核验只读权限",
            f"请先根据领星官方说明确认这项业务工具只读取数据，不创建、修改或删除数据。\n\n"
            f"ID：{tool.tool_id}\n名称：{tool.name}\n描述：{tool.description}\nSchema 指纹：{tool.permission_fingerprint[:16]}…\n\n"
            "许可仅在本次程序会话有效，并绑定当前 ID、说明和 Schema。确认后才允许通过 action 执行。",
        )
        if approved:
            self._verified_query_grants[tool.tool_id] = VerifiedReadOnlyGrant(tool.tool_id, tool.permission_fingerprint)
            self.mcp_status.set(f"已在本次会话核验只读范围：{tool.tool_id}。工具 Schema 变化后许可自动失效。")

    def _search_business_tools(self) -> None:
        try:
            args = json.loads(self.search_arguments_text.get("1.0", "end").strip() or "{}")
        except json.JSONDecodeError as exc:
            messagebox.showerror("search 参数格式错误", f"JSON 第 {exc.lineno} 行：{exc.msg}")
            return
        if not isinstance(args, dict):
            messagebox.showerror("search 参数格式错误", "search 参数必须是 JSON 对象。")
            return
        self._start_mcp_action("search_business_tools", search_arguments=args)

    def _execute_business_query(self) -> None:
        tool = self._selected_business_tool()
        if tool is None:
            messagebox.showwarning("请选择业务工具", "请先搜索并选择一项业务工具。")
            return
        grant = self._verified_query_grants.get(tool.tool_id)
        if grant is None or grant.permission_fingerprint != tool.permission_fingerprint:
            messagebox.showwarning("只读范围未核验", "请先查看参数并点击“核验本次只读范围”；没有许可时 action 调用会被阻止。")
            return
        try:
            search_args = json.loads(self.search_arguments_text.get("1.0", "end").strip() or "{}")
            args = json.loads(self.business_arguments_text.get("1.0", "end").strip() or "{}")
        except json.JSONDecodeError as exc:
            messagebox.showerror("参数格式错误", f"JSON 第 {exc.lineno} 行：{exc.msg}")
            return
        if not isinstance(search_args, dict) or not isinstance(args, dict):
            messagebox.showerror("参数格式错误", "search 和业务查询参数都必须是 JSON 对象。")
            return
        self._start_mcp_action(
            "business_query",
            search_arguments=search_args,
            business_tool=tool,
            arguments=args,
            grant=grant,
        )

    def _show_json_window(self, title: str, value: Any) -> None:
        window = tk.Toplevel(self)
        window.title(title)
        window.geometry("760x600")
        content = tk.Text(window, wrap="none")
        content.pack(fill="both", expand=True, padx=8, pady=8)
        content.insert("1.0", json.dumps(redact_sensitive(value), ensure_ascii=False, indent=2, default=str))
        content.configure(state="disabled")

    def _report_skus(self) -> list[str]:
        return list(dict.fromkeys(
            value.strip()
            for value in self.skus_var.get().replace("，", ",").split(",")
            if value.strip()
        ))

    def _parse_d01_dates(self) -> dict[str, date]:
        result: dict[str, date] = {}
        content = self.d01_dates_text.get("1.0", "end").strip()
        for line_number, line in enumerate(content.splitlines(), start=1):
            if not line.strip():
                continue
            if "=" not in line:
                raise ReportGenerationError(f"D01 第 {line_number} 行格式应为 SKU=YYYY-MM-DD。")
            sku, day_text = (part.strip() for part in line.split("=", 1))
            if not sku:
                raise ReportGenerationError(f"D01 第 {line_number} 行缺少 SKU。")
            if not day_text:
                continue
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day_text):
                raise ReportGenerationError(f"D01 第 {line_number} 行日期格式应为 YYYY-MM-DD。")
            try:
                day = date.fromisoformat(day_text)
            except ValueError as exc:
                raise ReportGenerationError(f"D01 第 {line_number} 行不是有效日期。") from exc
            key = sku.casefold()
            if key in result:
                raise ReportGenerationError(f"D01 中 SKU {sku} 重复；SKU 不区分大小写。")
            result[key] = day
        return result

    def _populate_d01_editor(self, *, preserve_current: bool) -> None:
        seller = self._selected_report_seller()
        skus = self._report_skus()
        if not hasattr(self, "d01_dates_text"):
            return
        current: dict[str, date] = {}
        if preserve_current:
            try:
                current = self._parse_d01_dates()
            except ReportGenerationError:
                current = {}
        saved: dict[str, date] = {}
        if seller is not None:
            try:
                saved = get_d01_dates(seller.sid, skus)
            except D01SettingsError as exc:
                self.d01_status.set(str(exc))
        rendered: list[str] = []
        for sku in skus:
            key = sku.casefold()
            value = current.get(key) if preserve_current and key in current else saved.get(key)
            rendered.append(f"{sku}={value.isoformat() if value else ''}")
        self._d01_refreshing = True
        try:
            self.d01_dates_text.delete("1.0", "end")
            self.d01_dates_text.insert("1.0", "\n".join(rendered))
        finally:
            self._d01_refreshing = False
        if seller is None:
            self.d01_status.set("先刷新并选择店铺；D01 会按店铺 SID 与 SKU 保存在本机。")
        elif not skus:
            self.d01_status.set("输入 SKU 后会自动载入该店铺已保存的 D01。")
        else:
            missing = [sku for sku in skus if sku.casefold() not in saved]
            self.d01_status.set(
                f"已从本机载入 {len(skus) - len(missing)}/{len(skus)} 个 D01。"
                + ("待设置：" + ", ".join(missing) if missing else "修改后请再次确认并保存。")
            )

    def _save_confirmed_d01(self) -> None:
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先刷新并选择领星店铺，D01 才能按店铺保存。")
            return
        skus = self._report_skus()
        if not skus:
            messagebox.showwarning("缺少 SKU", "请先输入至少一个 SKU。")
            return
        try:
            dates = self._parse_d01_dates()
        except ReportGenerationError as exc:
            messagebox.showerror("D01 日期错误", str(exc))
            return
        requested = {sku.casefold() for sku in skus}
        extra = sorted(set(dates) - requested)
        if extra:
            messagebox.showerror("D01 SKU 不匹配", "D01 输入中包含本次 SKU 列表以外的项目：" + ", ".join(extra))
            return
        if not dates:
            messagebox.showwarning("缺少 D01", "至少为一个 SKU 输入 YYYY-MM-DD 日期后再确认。")
            return
        missing = [sku for sku in skus if sku.casefold() not in dates]
        existing = Path(self.existing_var.get().strip()) if self.existing_var.get().strip() else None
        changes: dict[str, tuple[date, date]] = {}
        if existing is not None and existing.is_file() and self.template_var.get().strip():
            try:
                prior_dates = read_existing_d01_dates(existing, Path(self.template_var.get()), skus)
            except ReportGenerationError as exc:
                messagebox.showerror("无法核对已有周报", str(exc))
                return
            changes = {
                sku: (prior_dates[sku.casefold()], dates[sku.casefold()])
                for sku in skus
                if sku.casefold() in prior_dates
                and sku.casefold() in dates
                and prior_dates[sku.casefold()] != dates[sku.casefold()]
            }
        details = "\n".join(f"{sku}: {dates[sku.casefold()].isoformat()}" for sku in skus if sku.casefold() in dates)
        confirmation = (
            f"请确认这些日期均为用户指定的 D01，并按 {seller.site_label} 店铺站点日期解释：\n\n{details}"
        )
        if missing:
            confirmation += "\n\n尚未设置：" + ", ".join(missing) + "。这些 SKU 仍不能生成周报。"
        if changes:
            changed_details = "\n".join(
                f"{sku}: {before.isoformat()} → {after.isoformat()}" for sku, (before, after) in changes.items()
            )
            confirmation += (
                "\n\n日期错位风险：更新已有周报会改变该 SKU 的 D01 至后续日期对应关系；"
                "程序将按新日期重新映射数据，并保留手工目标和页眉字段。请检查：\n" + changed_details
            )
        if not messagebox.askyesno("确认并保存 D01", confirmation):
            return
        try:
            saved_path = save_d01_dates(seller.sid, dates)
        except D01SettingsError as exc:
            messagebox.showerror("保存 D01 失败", str(exc))
            return
        self._populate_d01_editor(preserve_current=False)
        self.d01_status.set(f"已确认并保存 {len(dates)} 个 D01（店铺 SID {seller.sid}）；本地设置：{saved_path.name}。")

    def _confirmed_d01_for_generation(
        self,
        seller: LingxingSeller,
        skus: list[str],
        existing: Path | None,
    ) -> tuple[dict[str, date], bool] | None:
        try:
            entered = self._parse_d01_dates()
            saved = get_d01_dates(seller.sid, skus)
        except (ReportGenerationError, D01SettingsError) as exc:
            messagebox.showerror("D01 设置错误", str(exc))
            return None
        missing = [sku for sku in skus if sku.casefold() not in entered or sku.casefold() not in saved]
        if missing:
            messagebox.showwarning("缺少已确认 D01", "以下 SKU 没有已保存并确认的 D01，不能生成正式周报：" + ", ".join(missing))
            return None
        if any(entered[sku.casefold()] != saved[sku.casefold()] for sku in skus):
            messagebox.showwarning("D01 尚未保存", "当前输入与本机已确认设置不同。请点击“确认并保存 D01”，再生成周报。")
            return None
        dates = {sku: saved[sku.casefold()] for sku in skus}
        if existing is None:
            return dates, False
        try:
            prior_dates = read_existing_d01_dates(existing, Path(self.template_var.get()), skus)
        except ReportGenerationError as exc:
            messagebox.showerror("无法读取已有周报", str(exc))
            return None
        changes = {
            sku: (prior_dates[sku.casefold()], dates[sku])
            for sku in skus
            if sku.casefold() in prior_dates and prior_dates[sku.casefold()] != dates[sku]
        }
        if not changes:
            return dates, False
        changed_details = "\n".join(
            f"{sku}: {before.isoformat()} → {after.isoformat()}" for sku, (before, after) in changes.items()
        )
        accepted = messagebox.askyesno(
            "确认已有周报 D01 变更",
            "风险提示：D01 改变会使该 SKU 的 D01、D02……对应到不同站点日期。生成时会按新日期重新映射数据，"
            "并保留日目标销量、发货时间、推广阶段等人工内容。请再次确认：\n\n" + changed_details,
        )
        return (dates, True) if accepted else None

    def _start_generation(self) -> None:
        skus = self._report_skus()
        if not skus:
            messagebox.showwarning("缺少 SKU", "请输入至少一个 SKU。")
            return
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请先刷新并选择店铺，以加载对应的 D01 设置。")
            return
        output_dir = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if output_dir is None:
            messagebox.showwarning("缺少输出文件夹", "请选择输出文件夹。")
            return
        existing = Path(self.existing_var.get()) if self.existing_var.get().strip() else None
        confirmed = self._confirmed_d01_for_generation(seller, skus, existing)
        if confirmed is None:
            return
        d01_dates, confirmed_d01_changes = confirmed
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_names = "_".join(skus)
        if len(safe_names) > 100:
            safe_names = f"{skus[0]}_等{len(skus)}个SKU"
        output_path = output_dir / f"新品推广周报_{safe_names}_{timestamp}.xlsx"
        template_path = Path(self.template_var.get())
        data_path = Path(self.data_var.get())
        mapping_path = Path(self.mapping_var.get())

        def generate():
            result = generate_reports(
                template_path=template_path,
                data_folder=data_path,
                mapping_path=mapping_path,
                skus=skus,
                output_path=output_path,
                d01_dates=d01_dates,
                d01_confirmed=True,
                confirmed_d01_changes=confirmed_d01_changes,
                existing_report_path=existing,
            )
            ok = sum(item.status == "success" for item in result.sku_results)
            partial = sum(item.status == "partial" for item in result.sku_results)
            failed = sum(item.status == "failed" for item in result.sku_results)
            message = (
                f"完成：完整 {ok} 个 SKU，销售已生成但广告总花费未完整核验 {partial} 个，失败 {failed} 个。\n"
                f"未完整核验的总广告花费与 ACOS 留空；详情见校验日志。\n"
                f"周报：{result.output_path}\n校验日志：{result.audit_path}"
            )
            return {"message": message, "output": str(result.output_path)}

        self._start_task(generate, self.report_status, "周报生成完成。")

    def _start_api_generation(self) -> None:
        """Query real OpenAPI sales and send only validated rows to the workbook writer."""
        if self._busy:
            return
        seller = self._selected_report_seller()
        if seller is None:
            messagebox.showwarning("请选择店铺", "请刷新并选择领星授权店铺。")
            return
        skus = list(dict.fromkeys(
            value.strip()
            for value in self.skus_var.get().replace("，", ",").split(",")
            if value.strip()
        ))
        if not 1 <= len(skus) <= 50:
            messagebox.showwarning("SKU 数量无效", "请输入 1–50 个不重复的本地 SKU，以逗号分隔。")
            return
        existing = Path(self.existing_var.get()) if self.existing_var.get().strip() else None
        confirmed = self._confirmed_d01_for_generation(seller, skus, existing)
        if confirmed is None:
            return
        d01_dates, confirmed_d01_changes = confirmed

        start_text = self.query_start_date_var.get().strip()
        end_text = self.query_end_date_var.get().strip()
        try:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_text) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_text):
                raise ValueError
            query_start = date.fromisoformat(start_text)
            query_end = date.fromisoformat(end_text)
        except ValueError:
            messagebox.showwarning("日期格式错误", "OpenAPI 周报查询请输入有效的 YYYY-MM-DD 站点日期范围。")
            return
        if query_end < query_start:
            messagebox.showwarning("日期范围错误", "查询结束日不能早于开始日。")
            return
        earliest_d01 = min(d01_dates[sku] for sku in skus)
        latest_d01 = max(d01_dates[sku] for sku in skus)
        if query_start > earliest_d01 or query_end < latest_d01:
            messagebox.showwarning("查询范围不完整", "查询开始日须不晚于最早 D01，查询结束日须不早于每个 SKU 的 D01。")
            return
        if not self.template_var.get().strip() or not Path(self.template_var.get()).is_file():
            messagebox.showwarning("缺少模板", "请选择现有《新品推广周报.xlsx》模板。")
            return
        output_dir = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if output_dir is None:
            messagebox.showwarning("缺少输出文件夹", "请选择周报输出文件夹。")
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_dir / f"新品推广周报_OpenAPI_{timestamp}.xlsx"
        template_path = Path(self.template_var.get())

        def generate_from_api() -> dict[str, Any]:
            with LingxingOpenApiClient(app_id, app_secret) as client:
                token_result = client.get_access_token()
                windows = client.query_daily_sales_history(
                    access_token=token_result.token,
                    store_sid=seller.sid,
                    local_skus=skus,
                    start_date=query_start,
                    end_date=query_end,
                )
                aliases_by_sku: dict[str, dict[str, set[str]]] = {}
                for sales_window in windows:
                    for mapping in sales_window.mappings:
                        key = mapping.local_sku.casefold()
                        aliases = aliases_by_sku.setdefault(key, {"seller_skus": set(), "asins": set()})
                        aliases["seller_skus"].update(mapping.seller_skus)
                        aliases["asins"].update(mapping.asins)
                ad_mappings = [
                    SkuProductMapping(
                        local_sku=next(sku for sku in skus if sku.casefold() == key),
                        seller_skus=tuple(sorted(values["seller_skus"])),
                        asins=tuple(sorted(values["asins"])),
                    )
                    for key, values in aliases_by_sku.items()
                ]
                ad_windows: dict[str, list[Any]] = {"SP": [], "SD": []}
                ad_errors: dict[str, str] = {}
                for kind, query_method in (
                    ("SP", client.query_sp_ad_spend_history),
                    ("SD", client.query_sd_ad_spend_history),
                ):
                    try:
                        arguments = {
                            "access_token": token_result.token,
                            "store_sid": seller.sid,
                            "start_date": query_start,
                            "end_date": query_end,
                        }
                        arguments["mappings"] = ad_mappings
                        if kind == "SD":
                            arguments["include_catalog"] = False
                        ad_windows[kind] = query_method(**arguments)
                    except LingxingApiError as exc:
                        # Keep sales generation available, but carry a visible
                        # failure state so missing ad data can never become zero.
                        ad_errors[kind] = f"查询失败：{exc}"
                sales_currency_codes = {
                    row.currency_code.strip().upper()
                    for sales_window in windows
                    for row in sales_window.rows
                    if row.sid == seller.sid and row.currency_code.strip()
                }
                ad_currency_basis_verified = len(sales_currency_codes) == 1
            result = generate_reports(
                template_path=template_path,
                data_folder=None,
                mapping_path=None,
                skus=skus,
                output_path=output_path,
                d01_dates=d01_dates,
                d01_confirmed=True,
                confirmed_d01_changes=confirmed_d01_changes,
                existing_report_path=existing,
                api_sales_windows=windows,
                api_store_sid=seller.sid,
                api_store_name=seller.name,
                api_ad_spend_windows_by_type=ad_windows,
                api_ad_sku_mappings=ad_mappings,
                api_ad_query_errors=ad_errors,
                # Ads APIs have no currency field; use the unique currency from
                # the same selected SID's daily sales response as the store currency.
                api_ad_currency_and_date_basis_verified=ad_currency_basis_verified,
            )
            complete = sum(item.status == "success" for item in result.sku_results)
            partial = sum(item.status == "partial" for item in result.sku_results)
            failed = sum(item.status == "failed" for item in result.sku_results)
            state = "部分完成" if partial or failed else "完成"
            message = (
                f"{state}：完整 {complete} 个、部分数据 {partial} 个、失败 {failed} 个。\n"
                f"查询日期：{query_start.isoformat()} 至 {query_end.isoformat()}（站点日期，按不超过 92 天分段）\n"
                f"广告统计范围：SP+SD。两类报表分别按商品行与活动总额核对；任一类型失败、缺失、未归属、币种不唯一或未核平的日期，广告花费保持空白。\n"
                f"周报：{result.output_path}\n校验日志：{result.audit_path}"
            )
            return {"message": message, "output": str(result.output_path)}

        self._start_task(generate_from_api, self.report_status, "OpenAPI 周报生成已完成。")

    def _open_output_folder(self) -> None:
        path = Path(self.output_var.get().strip())
        if not path.is_dir():
            messagebox.showwarning("文件夹不存在", "请先选择已存在的输出文件夹。")
            return
        try:
            import os
            import subprocess
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            messagebox.showerror("无法打开", "系统未能打开输出文件夹。")


def main() -> int:
    try:
        WeeklyReportApp().mainloop()
    except tk.TclError as exc:
        raise SystemExit("无法启动图形界面。请在 Windows 10/11 本机运行，云端无桌面显示。") from exc
    return 0
