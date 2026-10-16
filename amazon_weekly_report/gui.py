from __future__ import annotations

import json
import queue
import threading
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import Any

from .credential_store import (
    CredentialStoreError,
    LingxingOpenApiCredentialStore,
    WindowsCredentialStore,
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
from .providers.lingxing_api import LingxingApiError, LingxingOpenApiClient, LingxingSeller
from .report_generation import ReportGenerationError, generate_reports


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
        self._build_report_tab(report_tab)
        self._build_mcp_tab(mcp_tab)
        self._build_openapi_tab(openapi_tab)

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
        self.d01_mode = tk.StringVar(value="first_sale")
        self.launch_dates_var = tk.StringVar()
        self._path_row(parent, 0, "Excel 格式模板", self.template_var)
        self._path_row(parent, 1, "领星导出数据文件夹", self.data_var, directory=True)
        self._path_row(parent, 2, "字段映射 JSON", self.mapping_var)
        self._path_row(parent, 3, "输出文件夹", self.output_var, directory=True)
        self._path_row(parent, 4, "已有周报（可选）", self.existing_var)
        ttk.Label(parent, text="领星店铺", width=20).grid(row=5, column=0, sticky="w", pady=5)
        self.report_store_combo = ttk.Combobox(parent, textvariable=self.report_store_var, state="readonly")
        self.report_store_combo.grid(row=5, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="刷新店铺列表", command=self._openapi_seller_list).grid(row=5, column=2, sticky="w", padx=6, pady=5)
        ttk.Label(parent, text="选择项会自动绑定接口返回的 SID；本地 Excel/CSV 模式仍按字段映射读取。", wraplength=850).grid(
            row=6, column=1, columnspan=2, sticky="w", padx=6, pady=(0, 4)
        )
        ttk.Label(parent, text="SKU（逗号分隔）", width=20).grid(row=7, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.skus_var).grid(row=7, column=1, columnspan=2, sticky="ew", padx=6, pady=5)

        mode_box = ttk.LabelFrame(parent, text="D01 规则", padding=8)
        mode_box.grid(row=8, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Radiobutton(mode_box, text="SKU 首个有销量日期", variable=self.d01_mode, value="first_sale").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Radiobutton(mode_box, text="正式开售日期", variable=self.d01_mode, value="launch_date").grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(mode_box, text="选择正式开售日期时，每行填写 SKU=YYYY-MM-DD（不同 SKU 可用不同日期）").grid(row=1, column=0, columnspan=2, sticky="w", padx=4, pady=(5, 2))
        self.launch_dates_text = tk.Text(mode_box, height=3, wrap="none")
        self.launch_dates_text.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)
        mode_box.columnconfigure(1, weight=1)

        self.progress = ttk.Progressbar(parent, mode="indeterminate")
        self.progress.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        self.report_status = tk.StringVar(value="待处理。字段映射必须基于真实报表表头配置。")
        ttk.Label(parent, textvariable=self.report_status, wraplength=880).grid(row=10, column=0, columnspan=3, sticky="w", pady=4)
        buttons = ttk.Frame(parent)
        buttons.grid(row=11, column=0, columnspan=3, sticky="ew", pady=10)
        self.generate_button = ttk.Button(buttons, text="开始生成周报", command=self._start_generation)
        self.generate_button.pack(side="left")
        ttk.Button(buttons, text="打开输出文件夹", command=self._open_output_folder).pack(side="left", padx=8)
        ttk.Label(
            parent,
            text="程序不会自动推断导出字段；无法唯一匹配或缺少来源的数据留空，并写入校验日志。\n"
                 "新增周次按原模板七天一周延续。A2 推广阶段、G2 发货时间、M 列日目标销量由你手动维护。",
            justify="left",
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=8)

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
        parent.rowconfigure(9, weight=1)
        self.openapi_app_id_var = tk.StringVar()
        self.openapi_app_secret_var = tk.StringVar()
        self.openapi_sku_var = tk.StringVar()
        self.openapi_start_var = tk.StringVar()
        self.openapi_end_var = tk.StringVar()

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
        ttk.Label(parent, text="SKU", width=22).grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.openapi_sku_var).grid(row=5, column=1, sticky="ew", padx=6, pady=5)
        ttk.Label(parent, text="开始日期 (YYYY-MM-DD)").grid(row=5, column=2, sticky="e", padx=4)
        ttk.Entry(parent, textvariable=self.openapi_start_var, width=14).grid(row=5, column=3, sticky="ew", padx=6, pady=5)
        ttk.Label(parent, text="结束日期 (YYYY-MM-DD)").grid(row=6, column=2, sticky="e", padx=4)
        ttk.Entry(parent, textvariable=self.openapi_end_var, width=14).grid(row=6, column=3, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="测试 SKU 数据", command=self._openapi_sku_query).grid(row=6, column=1, sticky="w", padx=6, pady=5)

        self.openapi_status = tk.StringVar(value="OpenAPI 地址固定为官方域名；密钥仅保存在 Windows Credential Manager。")
        ttk.Label(parent, textvariable=self.openapi_status, wraplength=900).grid(row=7, column=0, columnspan=4, sticky="w", pady=6)
        ttk.Label(parent, text="结果 / 说明（不显示 Token，不写入日志）").grid(row=8, column=0, columnspan=4, sticky="w", pady=(8, 2))
        self.openapi_output = tk.Text(parent, height=14, wrap="word", state="disabled")
        self.openapi_output.grid(row=9, column=0, columnspan=4, sticky="nsew", pady=4)
        ttk.Label(
            parent,
            text="已启用 Token 身份验证和只读店铺列表请求。列表需通过业务码及字段类型校验后才会更新；status、has_ads_setting 显示原值，SKU 查询尚未接入。",
            wraplength=900,
        ).grid(row=10, column=0, columnspan=4, sticky="w", pady=6)

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

    def _openapi_sku_query(self) -> None:
        self.openapi_status.set("未发送查询：销售接口路径、参数和字段尚待依据官方文档核验。")
        self._show_openapi_result(self.openapi_status.get())

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

    def _parse_launch_dates(self) -> dict[str, date]:
        result: dict[str, date] = {}
        content = self.launch_dates_text.get("1.0", "end").strip()
        for line_number, line in enumerate(content.splitlines(), start=1):
            if not line.strip():
                continue
            if "=" not in line:
                raise ReportGenerationError(f"正式开售日期第 {line_number} 行格式应为 SKU=YYYY-MM-DD。")
            sku, day_text = (part.strip() for part in line.split("=", 1))
            try:
                result[sku] = date.fromisoformat(day_text)
            except ValueError as exc:
                raise ReportGenerationError(f"正式开售日期第 {line_number} 行不是 YYYY-MM-DD。") from exc
        return result

    def _start_generation(self) -> None:
        skus = [item.strip() for item in self.skus_var.get().replace("，", ",").split(",") if item.strip()]
        if not skus:
            messagebox.showwarning("缺少 SKU", "请输入至少一个 SKU。")
            return
        output_dir = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if output_dir is None:
            messagebox.showwarning("缺少输出文件夹", "请选择输出文件夹。")
            return
        try:
            launch_dates = self._parse_launch_dates()
        except ReportGenerationError as exc:
            messagebox.showerror("日期格式错误", str(exc))
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_names = "_".join(skus)
        if len(safe_names) > 100:
            safe_names = f"{skus[0]}_等{len(skus)}个SKU"
        output_path = output_dir / f"新品推广周报_{safe_names}_{timestamp}.xlsx"
        existing = Path(self.existing_var.get()) if self.existing_var.get().strip() else None
        template_path = Path(self.template_var.get())
        data_path = Path(self.data_var.get())
        mapping_path = Path(self.mapping_var.get())
        d01_mode = self.d01_mode.get()

        def generate():
            result = generate_reports(
                template_path=template_path,
                data_folder=data_path,
                mapping_path=mapping_path,
                skus=skus,
                output_path=output_path,
                d01_mode=d01_mode,
                launch_dates=launch_dates,
                existing_report_path=existing,
            )
            ok = sum(item.status == "success" for item in result.sku_results)
            failed = sum(item.status == "failed" for item in result.sku_results)
            message = f"完成：{ok} 个 SKU 成功，{failed} 个失败。\n周报：{result.output_path}\n校验日志：{result.audit_path}"
            return {"message": message, "output": str(result.output_path)}

        self._start_task(generate, self.report_status, "周报生成完成。")

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
