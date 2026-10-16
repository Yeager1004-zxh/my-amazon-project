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

from .credential_store import (
    CredentialStoreError,
    LingxingOpenApiCredentialStore,
    WindowsCredentialStore,
)
from .d01_settings import D01SettingsError, get_d01_dates, save_d01_dates
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
    SpAdSpendWindow,
    SkuDailySales,
    SkuDailySalesWindow,
    SkuProductMapping,
)
from .report_generation import ReportGenerationError, generate_reports, read_existing_d01_dates


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

        query_box = ttk.LabelFrame(parent, text="领星 SKU 与每日销量预览（只读，不写入周报）", padding=8)
        query_box.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=6)
        ttk.Label(query_box, text="查询开始日（站点时间）").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(query_box, textvariable=self.query_start_date_var, width=13).grid(row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Label(query_box, text="查询结束日").grid(row=0, column=2, sticky="e", padx=4, pady=4)
        ttk.Entry(query_box, textvariable=self.query_end_date_var, width=13).grid(row=0, column=3, sticky="w", padx=4, pady=4)
        self.report_query_button = ttk.Button(query_box, text="查询并预览真实数据", command=self._query_report_sku_data)
        self.report_query_button.grid(row=0, column=4, sticky="w", padx=6, pady=4)
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
        query_box.columnconfigure(4, weight=1)
        query_box.rowconfigure(3, weight=1)
        self.report_ad_query_button = ttk.Button(
            query_box,
            text="查询 SP 广告花费并预览",
            command=self._query_report_sp_ad_data,
        )
        self.report_ad_query_button.grid(row=4, column=0, sticky="w", padx=4, pady=(8, 3))
        ttk.Label(query_box, text="SP 广告商品明细（按 campaign 与活动报表核对；仅预览，不写入周报）").grid(
            row=4, column=1, columnspan=4, sticky="w", padx=4, pady=(8, 3)
        )
        self.report_ad_status = tk.StringVar(value="仅查询 SP；未接入的 SB/SD 不视为零。")
        ttk.Label(query_box, textvariable=self.report_ad_status, wraplength=900).grid(
            row=5, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 3)
        )
        ad_columns = ("report_date", "local_sku", "seller_sku", "campaign_id", "ad_group_id", "cost", "check")
        self.report_ad_tree = ttk.Treeview(query_box, columns=ad_columns, show="headings", height=4)
        for column, heading, width in (
            ("report_date", "报表日期", 90),
            ("local_sku", "本地 SKU", 105),
            ("seller_sku", "报表 SKU/MSKU", 120),
            ("campaign_id", "Campaign ID", 145),
            ("ad_group_id", "Ad Group ID", 135),
            ("cost", "SP cost", 90),
            ("check", "核对状态", 205),
        ):
            self.report_ad_tree.heading(column, text=heading)
            self.report_ad_tree.column(column, width=width, anchor="w")
        self.report_ad_tree.grid(row=6, column=0, columnspan=5, sticky="nsew", padx=4, pady=4)
        ad_scrollbar = ttk.Scrollbar(query_box, orient="vertical", command=self.report_ad_tree.yview)
        ad_scrollbar.grid(row=6, column=5, sticky="ns", pady=4)
        self.report_ad_tree.configure(yscrollcommand=ad_scrollbar.set)
        query_box.rowconfigure(6, weight=1)

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
        self._populate_d01_editor(preserve_current=False)

    def _on_report_skus_changed(self, *_args: Any) -> None:
        if not self._d01_refreshing:
            self._populate_d01_editor(preserve_current=True)

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

    def _query_report_sp_ad_data(self) -> None:
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
            messagebox.showwarning("日期格式错误", "SP 报表日期请输入有效的 YYYY-MM-DD。")
            return
        if end_date < start_date:
            messagebox.showwarning("日期范围错误", "结束日期不能早于开始日期。")
            return
        if (end_date - start_date).days + 1 > 92:
            messagebox.showwarning("日期范围过长", "为控制请求量，SP 预览一次最多查询 92 个日期；遇到 429 会停止且不自动重试。")
            return
        try:
            app_id, app_secret = self._openapi_credentials()
        except CredentialStoreError as exc:
            messagebox.showerror("OpenAPI 配置错误", str(exc))
            return
        for item_id in self.report_ad_tree.get_children():
            self.report_ad_tree.delete(item_id)
        self.report_ad_status.set("正在读取 SP 活动及商品报表并逐日核对……")

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
                missing = [sku for sku in local_skus if sku.casefold() not in by_sku or not by_sku[sku.casefold()].seller_skus]
                if missing:
                    raise LingxingApiError("以下 SKU 没有完整 MSKU 映射，未查询广告费用：" + ", ".join(missing))
                windows = client.query_sp_ad_spend_history(
                    access_token=token_result.token,
                    store_sid=seller.sid,
                    mappings=[by_sku[sku.casefold()] for sku in local_skus],
                    start_date=start_date,
                    end_date=end_date,
                )
            return {
                "_sp_ad_windows": windows,
                "_sp_ad_mappings": mappings,
                "_sp_ad_store_name": seller.name,
                "_sp_ad_start": start_date,
                "_sp_ad_end": end_date,
                "message": "SP 广告费用预览查询完成；该预览未写入周报。",
            }

        self._start_task(query, self.report_ad_status, "SP 广告费用查询完成。")

    def _apply_sp_ad_preview(self, result: dict[str, Any]) -> None:
        windows: list[SpAdSpendWindow] = result["_sp_ad_windows"]
        mappings: list[SkuProductMapping] = result["_sp_ad_mappings"]
        local_by_msku = {
            msku.casefold(): mapping.local_sku
            for mapping in mappings
            for msku in mapping.seller_skus
        }
        discrepancy_count = sum(len(window.discrepancies) for window in windows)
        product_total = sum(
            (row.cost for window in windows for row in window.product_rows),
            start=Decimal("0"),
        )
        campaign_total = sum(
            (row.cost for window in windows for row in window.campaign_rows),
            start=Decimal("0"),
        )
        sku_total_by_key: dict[str, Any] = {}
        for window in windows:
            blocked_dates = {item.report_date for item in window.discrepancies}
            for spend in window.sku_daily_spends:
                if spend.report_date not in blocked_dates:
                    key = spend.local_sku.casefold()
                    sku_total_by_key[key] = sku_total_by_key.get(key, 0) + spend.actual_spend
            state = "日期内成本不一致；SKU 花费被阻止" if blocked_dates else "活动/商品总成本相等（待后台验收）"
            for row in window.product_rows:
                self.report_ad_tree.insert(
                    "",
                    "end",
                    values=(
                        row.report_date.isoformat(),
                        local_by_msku.get(row.sku.casefold(), "未映射"),
                        row.sku,
                        row.campaign_id,
                        row.ad_group_id,
                        format(row.cost, "f"),
                        state,
                    ),
                )
        summary = [
            f"店铺 {result['_sp_ad_store_name']}：{result['_sp_ad_start'].isoformat()} 至 {result['_sp_ad_end'].isoformat()}（报表日期）",
            f"campaign cost 合计 {format(campaign_total, 'f')}；商品 cost 合计 {format(product_total, 'f')}；不一致 campaign {discrepancy_count} 个",
            "SP cost 币种字段未在本次官方 Schema 中提供，金额只供对照；只有逐日活动/商品成本一致的日期才显示 SKU 汇总。",
            "报表日期时区也未在已提供的官方 Schema 中说明；按返回的 YYYY-MM-DD 对照，不推断为店铺站点日期。",
            "本结果仅供预览，尚未写入周报；SB/SD 未查询且不视为零。",
        ]
        incomplete_days = [
            window for window in windows
            if not window.campaign_rows or not window.product_rows
        ]
        if incomplete_days:
            summary.append(
                "缺少完整双报表记录的日期（未知，不当作零花费）：" + "、".join(
                    f"{window.start_date.isoformat()}["
                    f"{'活动报表缺记录' if not window.campaign_rows else ''}"
                    f"{'；' if not window.campaign_rows and not window.product_rows else ''}"
                    f"{'商品报表缺记录' if not window.product_rows else ''}]"
                    for window in incomplete_days
                )
            )
        for local_key, amount in sorted(sku_total_by_key.items()):
            summary.append(f"{local_key} SP SKU cost={format(amount, 'f')}")
        if discrepancy_count:
            sample = [item for window in windows for item in window.discrepancies][:5]
            summary.append("差异示例：" + "；".join(
                f"{item.report_date.isoformat()} campaign {item.campaign_id}：活动={item.campaign_cost if item.campaign_cost is not None else '缺失'}，商品={item.product_cost if item.product_cost is not None else '缺失'}"
                for item in sample
            ))
        self.report_ad_status.set(" ".join(summary))

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
                    elif status_var is self.report_ad_status and isinstance(result, dict) and result.get("_sp_ad_windows") is not None:
                        self._apply_sp_ad_preview(result)
                    elif isinstance(result, dict):
                        status_var.set(result.get("message", title))
                        if result.get("output"):
                            messagebox.showinfo("周报生成完成", result["message"])
                else:
                    status_var, error = payload
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
            failed = sum(item.status == "failed" for item in result.sku_results)
            message = f"完成：{ok} 个 SKU 成功，{failed} 个失败。\n周报：{result.output_path}\n校验日志：{result.audit_path}"
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
            )
            complete = sum(item.status == "success" for item in result.sku_results)
            partial = sum(item.status == "partial" for item in result.sku_results)
            failed = sum(item.status == "failed" for item in result.sku_results)
            state = "部分完成" if partial or failed else "完成"
            message = (
                f"{state}：完整 {complete} 个、仅销售数据 {partial} 个、失败 {failed} 个。\n"
                f"查询日期：{query_start.isoformat()} 至 {query_end.isoformat()}（站点日期，按不超过 92 天分段）\n"
                f"广告活动实际花费尚未接入，花费单元格留空。\n周报：{result.output_path}\n校验日志：{result.audit_path}"
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
