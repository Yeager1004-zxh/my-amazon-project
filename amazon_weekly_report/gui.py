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
from .mcp_client import LingxingMcpClient, McpClientError, McpTool, redact_sensitive
from .providers.lingxing_api import LingxingOpenApiClient
from .report_generation import ReportGenerationError, generate_reports


class WeeklyReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("亚马逊新品推广周报工具")
        self.geometry("980x760")
        self.minsize(850, 640)
        self._events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._busy = False
        self._tools: list[McpTool] = []
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
        self.d01_mode = tk.StringVar(value="first_sale")
        self.launch_dates_var = tk.StringVar()
        self._path_row(parent, 0, "Excel 格式模板", self.template_var)
        self._path_row(parent, 1, "领星导出数据文件夹", self.data_var, directory=True)
        self._path_row(parent, 2, "字段映射 JSON", self.mapping_var)
        self._path_row(parent, 3, "输出文件夹", self.output_var, directory=True)
        self._path_row(parent, 4, "已有周报（可选）", self.existing_var)
        ttk.Label(parent, text="SKU（逗号分隔）", width=20).grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(parent, textvariable=self.skus_var).grid(row=5, column=1, columnspan=2, sticky="ew", padx=6, pady=5)

        mode_box = ttk.LabelFrame(parent, text="D01 规则", padding=8)
        mode_box.grid(row=6, column=0, columnspan=3, sticky="ew", pady=8)
        ttk.Radiobutton(mode_box, text="SKU 首个有销量日期", variable=self.d01_mode, value="first_sale").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Radiobutton(mode_box, text="正式开售日期", variable=self.d01_mode, value="launch_date").grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(mode_box, text="选择正式开售日期时，每行填写 SKU=YYYY-MM-DD（不同 SKU 可用不同日期）").grid(row=1, column=0, columnspan=2, sticky="w", padx=4, pady=(5, 2))
        self.launch_dates_text = tk.Text(mode_box, height=3, wrap="none")
        self.launch_dates_text.grid(row=2, column=0, columnspan=2, sticky="ew", padx=4)
        mode_box.columnconfigure(1, weight=1)

        self.progress = ttk.Progressbar(parent, mode="indeterminate")
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        self.report_status = tk.StringVar(value="待处理。字段映射必须基于真实报表表头配置。")
        ttk.Label(parent, textvariable=self.report_status, wraplength=880).grid(row=8, column=0, columnspan=3, sticky="w", pady=4)
        buttons = ttk.Frame(parent)
        buttons.grid(row=9, column=0, columnspan=3, sticky="ew", pady=10)
        self.generate_button = ttk.Button(buttons, text="开始生成周报", command=self._start_generation)
        self.generate_button.pack(side="left")
        ttk.Button(buttons, text="打开输出文件夹", command=self._open_output_folder).pack(side="left", padx=8)
        ttk.Label(
            parent,
            text="程序不会自动推断导出字段；无法唯一匹配或缺少来源的数据留空，并写入校验日志。\n"
                 "新增周次按原模板七天一周延续。A2 推广阶段、G2 发货时间、M 列日目标销量由你手动维护。",
            justify="left",
        ).grid(row=10, column=0, columnspan=3, sticky="w", pady=8)

    def _build_mcp_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(8, weight=1)
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
        ttk.Label(parent, text="只读工具", width=25).grid(row=4, column=0, sticky="w", pady=5)
        self.tool_combo = ttk.Combobox(parent, textvariable=self.tool_var, state="readonly")
        self.tool_combo.grid(row=4, column=1, columnspan=3, sticky="ew", padx=6, pady=5)
        ttk.Label(parent, text="查询参数 JSON", width=25).grid(row=5, column=0, sticky="nw", pady=5)
        self.arguments_text = tk.Text(parent, height=5, wrap="none")
        self.arguments_text.insert("1.0", "{}")
        self.arguments_text.grid(row=5, column=1, columnspan=3, sticky="ew", padx=6, pady=5)
        ttk.Button(parent, text="运行只读查询测试", command=self._query_tool).grid(row=6, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(parent, text="查询结果（仅显示，不写入日志或文件）").grid(row=7, column=0, columnspan=4, sticky="w", pady=(8, 2))
        self.mcp_output = tk.Text(parent, height=14, wrap="word", state="disabled")
        self.mcp_output.grid(row=8, column=0, columnspan=4, sticky="nsew", pady=4)
        ttk.Label(
            parent,
            text="密钥只写入 Windows Credential Manager；配置文件只存服务地址。工具必须被 MCP 标记 readOnlyHint=true 才允许调用。",
            wraplength=880,
        ).grid(row=9, column=0, columnspan=4, sticky="w", pady=6)

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
        ttk.Button(actions, text="获取店铺列表", command=self._openapi_seller_list).pack(side="left", padx=8)

        ttk.Label(parent, text="店铺名称 / SID / 站点 / 广告授权状态").grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 2))
        self.openapi_seller_tree = ttk.Treeview(
            parent,
            columns=("seller_name", "sid", "site", "ad_auth"),
            show="headings",
            height=3,
        )
        for column, heading, width in (
            ("seller_name", "店铺名称", 260),
            ("sid", "SID", 150),
            ("site", "站点", 160),
            ("ad_auth", "广告授权", 180),
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
            text="目前仅启用 Token 身份验证测试。店铺列表和 SKU 查询依赖 app_key/timestamp/sign 规则及官方字段定义；在文档核实前，按钮不会发送查询请求。广告接口未接入。",
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

    def _openapi_seller_list(self) -> None:
        self.openapi_status.set("未发送查询：官方 app_key/timestamp/sign 签名规则尚待核验。")
        self._show_openapi_result(self.openapi_status.get())

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
                        if isinstance(result, dict) and result.get("_tool_names") is not None:
                            self.tool_combo.configure(values=result["_tool_names"])
                            if result["_tool_names"]:
                                self.tool_combo.current(0)
                        self._show_mcp_result(result)
                    elif status_var is self.openapi_status and isinstance(result, dict):
                        status_var.set(result.get("message", title))
                        self._show_openapi_result(status_var.get())
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
    ):
        client = LingxingMcpClient(endpoint, key)
        try:
            init = client.connect()
            if action == "connect":
                return {"_status": f"连接成功；协议版本 {init.get('protocolVersion')}。", "_output_text": "MCP 初始化握手成功。"}
            tools = client.list_tools()
            self._tools = tools
            if action == "tools":
                display = [
                    {"name": tool.name, "read_only": tool.read_only, "description": tool.description, "input_schema": tool.input_schema}
                    for tool in tools
                ]
                read_only_names = tuple(f"{tool.name} — {tool.description[:100]}" for tool in tools if tool.read_only)
                return {
                    "_status": f"已获取 {len(tools)} 个工具；只读工具 {sum(tool.read_only for tool in tools)} 个。",
                    "_output_text": json.dumps(display, ensure_ascii=False, indent=2),
                    "_tool_names": read_only_names,
                }
            selected = next((tool for tool in tools if tool.name == tool_name), None)
            if selected is None:
                raise McpClientError("工具列表中找不到所选工具，请先刷新列表。")
            result = client.call_read_only_tool(selected, arguments or {})
            return {"_status": f"只读查询完成：{selected.name}。结果仅显示在本机界面。", "_output_text": json.dumps(result, ensure_ascii=False, indent=2, default=str)}
        finally:
            client.close()

    def _start_mcp_action(self, action: str, *, tool_name: str | None = None, arguments: dict[str, Any] | None = None) -> None:
        try:
            endpoint, key = self._credentials()
        except (CredentialStoreError, LocalSettingsError) as exc:
            messagebox.showerror("MCP 配置错误", str(exc))
            return
        self._start_task(
            lambda: self._run_mcp(endpoint, key, action, tool_name=tool_name, arguments=arguments),
            self.mcp_status,
            "MCP 操作完成。",
        )

    def _test_mcp_connection(self) -> None:
        self._start_mcp_action("connect")

    def _refresh_tools(self) -> None:
        self._start_mcp_action("tools")

    def _query_tool(self) -> None:
        selected = self.tool_var.get().split(" — ", 1)[0].strip()
        if not selected:
            messagebox.showwarning("请选择工具", "请先获取工具列表并选择标记为只读的工具。")
            return
        try:
            args = json.loads(self.arguments_text.get("1.0", "end").strip() or "{}")
        except json.JSONDecodeError as exc:
            messagebox.showerror("参数格式错误", f"JSON 第 {exc.lineno} 行：{exc.msg}")
            return
        if not isinstance(args, dict):
            messagebox.showerror("参数格式错误", "查询参数必须是 JSON 对象。")
            return
        self._start_mcp_action("query", tool_name=selected, arguments=args)

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
