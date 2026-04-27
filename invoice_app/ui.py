from __future__ import annotations

import ctypes
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from .classification import DEFAULT_CATEGORIES
from .models import InvoiceRecord
from .parser import iter_pdf_files, parse_invoice, resolve_attachment_metadata
from .service import (
    apply_category_to_records,
    clear_categories,
    ProcessOptions,
    add_category,
    execute_records,
    preview_records,
    remove_category,
    rename_category,
)


def get_workspace_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


class InvoiceDesktopApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.workspace_root = get_workspace_root()
        self.root.title("发票归类工具")
        self.root.geometry("1480x900")
        self.root.minsize(1180, 760)
        self.root.configure(bg="#edf2f7")

        self.input_paths: list[str] = []
        self.records: list[InvoiceRecord] = []
        self.categories: list[str] = list(DEFAULT_CATEGORIES)
        self.target_category_var = tk.StringVar(value=self.categories[0])
        self.busy = tk.BooleanVar(value=False)
        self.recursive_var = tk.BooleanVar(value=False)
        self.excel_output_var = tk.StringVar(
            value=str(self.workspace_root / "发票信息汇总.xlsx")
        )
        self.status_var = tk.StringVar(
            value=f"默认扫描目录: {self.workspace_root}"
        )
        self.paths_var = tk.StringVar(
            value=f"输入路径 1 个，默认扫描 {self.workspace_root}"
        )
        self.input_paths = [str(self.workspace_root)]

        self._build_style()
        self._build_layout()
        self._refresh_category_list()
        self._refresh_summary([])
        self.root.after(200, self.parse_inputs)

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#edf2f7")
        style.configure("TLabel", background="#edf2f7", foreground="#22313f")
        style.configure("Header.TFrame", background="#20324b")
        style.configure("Header.TLabel", background="#20324b", foreground="white")
        style.configure("Muted.TLabel", background="#edf2f7", foreground="#617286")
        style.configure("Accent.TButton", padding=(12, 8))
        style.configure("TButton", padding=(10, 6))
        style.configure(
            "Treeview",
            rowheight=28,
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#22313f",
        )
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_layout(self) -> None:
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(18, 14))
        header.pack(fill="x")
        ttk.Label(
            header,
            text="发票归类与批处理",
            style="Header.TLabel",
            font=("Microsoft YaHei UI", 18, "bold"),
        ).pack(anchor="w")
        ttk.Label(
            header,
            text="单票多分类、分类汇总、按类别目录复制、Excel 导出",
            style="Header.TLabel",
            font=("Microsoft YaHei UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        toolbar = ttk.Frame(self.root, padding=(14, 12))
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="添加文件", command=self.add_files).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(toolbar, text="添加目录", command=self.add_folder).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(toolbar, text="重新解析", command=self.parse_inputs).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(toolbar, text="归类所选", command=self.append_selected_category).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(toolbar, text="开始执行", command=self.execute_current).pack(
            side="right"
        )

        options = ttk.Frame(self.root, padding=(14, 0, 14, 10))
        options.pack(fill="x")
        ttk.Checkbutton(
            options,
            text="递归扫描子目录",
            variable=self.recursive_var,
        ).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(options, text="Excel 输出").grid(row=0, column=1, sticky="w")
        ttk.Entry(options, textvariable=self.excel_output_var, width=40).grid(
            row=0, column=2, sticky="ew", padx=(6, 10)
        )
        ttk.Button(options, text="浏览", command=self.choose_excel_output).grid(
            row=0, column=3
        )
        options.columnconfigure(2, weight=1)

        paths = ttk.Frame(self.root, padding=(14, 0, 14, 8))
        paths.pack(fill="x")
        ttk.Label(paths, textvariable=self.paths_var, style="Muted.TLabel").pack(
            anchor="w"
        )

        body = ttk.Panedwindow(self.root, orient="horizontal")
        body.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        left = ttk.Frame(body, padding=12)
        center = ttk.Frame(body, padding=12)
        right = ttk.Frame(body, padding=12)
        body.add(left, weight=1)
        body.add(center, weight=3)
        body.add(right, weight=2)

        self._build_category_panel(left)
        self._build_invoice_panel(center)
        self._build_detail_panel(right)

        footer = ttk.Frame(self.root, padding=(14, 0, 14, 12))
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_var, style="Muted.TLabel").pack(
            anchor="w"
        )

    def _build_category_panel(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="类别", font=("Microsoft YaHei UI", 12, "bold")).pack(
            anchor="w"
        )
        ttk.Label(
            parent,
            text="点击类别确定当前操作对象，再用下方按钮直接替换、追加或一键归类全部发票。",
            style="Muted.TLabel",
            wraplength=280,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))
        self.category_list = tk.Listbox(
            parent,
            height=16,
            activestyle="dotbox",
            exportselection=False,
        )
        self.category_list.pack(fill="both", expand=True, pady=(8, 8))
        self.category_list.bind("<<ListboxSelect>>", self._on_category_select)

        control = ttk.Frame(parent)
        control.pack(fill="x", pady=(4, 8))
        ttk.Label(control, text="当前类别").grid(row=0, column=0, sticky="w")
        self.target_category_combo = ttk.Combobox(
            control,
            textvariable=self.target_category_var,
            state="readonly",
            values=self.categories,
            width=18,
        )
        self.target_category_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.target_category_combo.bind("<<ComboboxSelected>>", self._on_target_category_select)
        control.columnconfigure(1, weight=1)

        action_row = ttk.Frame(parent)
        action_row.pack(fill="x")
        ttk.Button(action_row, text="替换所选", command=self.replace_selected_category).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(action_row, text="追加所选", command=self.append_selected_category).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(action_row, text="归类全部", command=self.apply_category_to_all).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(action_row, text="清空所选", command=self.clear_selected_category).pack(
            side="left"
        )

        button_row = ttk.Frame(parent)
        button_row.pack(fill="x", pady=(8, 0))
        ttk.Button(button_row, text="新增", command=self.add_category_prompt).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="重命名", command=self.rename_category_prompt).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="删除", command=self.delete_category_prompt).pack(
            side="left"
        )

        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=12)
        ttk.Label(
            parent,
            text="分类汇总",
            font=("Microsoft YaHei UI", 12, "bold"),
        ).pack(anchor="w")
        self.summary_tree = ttk.Treeview(
            parent,
            columns=("count", "amount"),
            show="headings",
            height=10,
        )
        self.summary_tree.heading("count", text="票数")
        self.summary_tree.heading("amount", text="金额")
        self.summary_tree.column("count", width=70, anchor="center")
        self.summary_tree.column("amount", width=110, anchor="e")
        self.summary_tree.pack(fill="both", expand=True, pady=(8, 0))

    def _build_invoice_panel(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="发票列表", font=("Microsoft YaHei UI", 12, "bold")).pack(
            anchor="w"
        )
        ttk.Label(
            parent,
            text="支持多分类；选中发票后，直接用当前类别按钮替换或追加分类。",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 0))
        columns = ("source", "invoice", "seller", "amount", "categories", "status")
        self.invoice_tree = ttk.Treeview(
            parent,
            columns=columns,
            show="headings",
            selectmode="extended",
        )
        headings = {
            "source": "来源文件",
            "invoice": "发票号",
            "seller": "销售方",
            "amount": "金额",
            "categories": "类别",
            "status": "状态",
        }
        widths = {
            "source": 220,
            "invoice": 120,
            "seller": 220,
            "amount": 90,
            "categories": 180,
            "status": 110,
        }
        for column in columns:
            self.invoice_tree.heading(column, text=headings[column])
            self.invoice_tree.column(column, width=widths[column], anchor="w")
        self.invoice_tree.column("amount", anchor="e")
        self.invoice_tree.pack(fill="both", expand=True, pady=(8, 8))
        self.invoice_tree.bind("<<TreeviewSelect>>", self._on_invoice_select)

        button_row = ttk.Frame(parent)
        button_row.pack(fill="x")
        ttk.Button(button_row, text="替换所选", command=self.replace_selected_category).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="追加所选", command=self.append_selected_category).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="归类全部", command=self.apply_category_to_all).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="清空所选", command=self.clear_selected_category).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(button_row, text="刷新预览", command=self.refresh_preview).pack(
            side="right"
        )

    def _build_detail_panel(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="详情", font=("Microsoft YaHei UI", 12, "bold")).pack(
            anchor="w"
        )
        self.detail_text = tk.Text(parent, height=28, wrap="word")
        self.detail_text.pack(fill="both", expand=True, pady=(8, 0))
        self.detail_text.configure(state="disabled")

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _set_busy(self, busy: bool, message: str | None = None) -> None:
        self.busy.set(busy)
        if message is not None:
            self._set_status(message)

    def _refresh_paths_label(self) -> None:
        if not self.input_paths:
            self.paths_var.set("未添加输入路径")
        else:
            default_path = self.input_paths[0]
            self.paths_var.set(
                f"默认扫描: {default_path} | 输入路径 {len(self.input_paths)} 个 | 当前类别 {len(self.categories)} 个"
            )

    def _refresh_category_list(self) -> None:
        self.category_list.delete(0, tk.END)
        for category in self.categories:
            self.category_list.insert(tk.END, category)
        self.target_category_combo["values"] = self.categories
        current = self.target_category_var.get().strip()
        if current not in self.categories:
            current = self.categories[0] if self.categories else ""
            self.target_category_var.set(current)
        if current and current in self.categories:
            index = self.categories.index(current)
            self.category_list.selection_clear(0, tk.END)
            self.category_list.selection_set(index)
            self.category_list.see(index)
        self._refresh_paths_label()

    def _refresh_invoice_tree(self) -> None:
        self.invoice_tree.delete(*self.invoice_tree.get_children())
        for index, record in enumerate(self.records):
            amount = f"{record.total_amount:.2f}" if record.total_amount is not None else ""
            self.invoice_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    record.source.name if record.source is not None else "",
                    record.invoice_number or "",
                    record.seller_name or "",
                    amount,
                    "、".join(record.categories),
                    record.status,
                ),
            )
        if self.records:
            self.invoice_tree.selection_set("0")
            self._show_record_details(self.records[0])
        else:
            self._show_record_details(None)

    def _refresh_summary(self, summaries) -> None:
        self.summary_tree.delete(*self.summary_tree.get_children())
        for summary in summaries:
            self.summary_tree.insert(
                "",
                tk.END,
                values=(
                    summary.name,
                    summary.invoice_count,
                    f"{summary.total_amount:.2f}",
                ),
            )

    def _show_record_details(self, record: InvoiceRecord | None) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", tk.END)
        if record is None:
            self.detail_text.insert(tk.END, "未选择记录")
        else:
            lines = [
                f"来源文件: {record.source.name if record.source is not None else ''}",
                f"发票号码: {record.invoice_number or ''}",
                f"开票日期: {record.invoice_date or ''}",
                f"购买方: {record.buyer_name or ''}",
                f"销售方: {record.seller_name or ''}",
                f"金额: {record.amount if record.amount is not None else ''}",
                f"税额: {record.tax_amount if record.tax_amount is not None else ''}",
                f"价税合计: {record.total_amount if record.total_amount is not None else ''}",
                f"类别: {'、'.join(record.categories)}",
                f"状态: {record.status}",
                "",
                "提醒:",
                *record.warnings,
            ]
            self.detail_text.insert(tk.END, "\n".join(lines).strip())
        self.detail_text.configure(state="disabled")

    def _selected_record_indices(self) -> list[int]:
        indices: list[int] = []
        for item_id in self.invoice_tree.selection():
            try:
                indices.append(int(item_id))
            except ValueError:
                continue
        return indices

    def _on_category_select(self, _event) -> None:
        selection = self.category_list.curselection()
        if not selection:
            return
        category = self.categories[selection[0]]
        self.target_category_var.set(category)

    def _on_target_category_select(self, _event) -> None:
        category = self.target_category_var.get().strip()
        if category in self.categories:
            index = self.categories.index(category)
            self.category_list.selection_clear(0, tk.END)
            self.category_list.selection_set(index)
            self.category_list.see(index)

    def _on_invoice_select(self, _event) -> None:
        indices = self._selected_record_indices()
        if not indices:
            self._show_record_details(None)
            return
        self._show_record_details(self.records[indices[0]])

    def _resolve_target_category(self) -> str | None:
        category = self.target_category_var.get().strip()
        if category in self.categories:
            return category
        if self.categories:
            category = self.categories[0]
            self.target_category_var.set(category)
            return category
        return None

    def _apply_category_to_records(
        self,
        *,
        target_indices: list[int] | None = None,
        replace: bool,
        apply_all: bool = False,
    ) -> None:
        if self.busy.get():
            return
        if apply_all and not self.records:
            messagebox.showinfo("提示", "请先解析发票。")
            return
        category = self._resolve_target_category()
        if not category:
            messagebox.showinfo("提示", "请先新增一个类别。")
            return
        if apply_all:
            target_indices = list(range(len(self.records)))
        elif target_indices is None:
            target_indices = self._selected_record_indices()
        if not target_indices:
            messagebox.showinfo("提示", "请先选择一条或多条发票记录。")
            return

        apply_category_to_records(
            self.records,
            category,
            target_indices=target_indices,
            replace=replace,
        )
        self._refresh_invoice_tree()
        self.refresh_preview()
        if apply_all:
            self._set_status(f"已将全部记录归类到「{category}」")
        elif replace:
            self._set_status(f"已替换所选记录的类别为「{category}」")
        else:
            self._set_status(f"已追加类别「{category}」到所选记录")

    def replace_selected_category(self) -> None:
        self._apply_category_to_records(replace=True)

    def append_selected_category(self) -> None:
        self._apply_category_to_records(replace=False)

    def apply_category_to_all(self) -> None:
        self._apply_category_to_records(replace=True, apply_all=True)

    def clear_selected_category(self) -> None:
        if self.busy.get():
            return
        indices = self._selected_record_indices()
        if not indices:
            messagebox.showinfo("提示", "请先选择一条或多条发票记录。")
            return
        clear_categories(self.records, target_indices=indices)
        self._refresh_invoice_tree()
        self.refresh_preview()
        self._set_status("已清空所选记录的类别")

    def choose_excel_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="选择 Excel 输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.excel_output_var.set(path)

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择发票文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")],
        )
        if not paths:
            return
        self.input_paths.extend([path for path in paths if path not in self.input_paths])
        self._refresh_paths_label()
        self.parse_inputs()

    def add_folder(self) -> None:
        path = filedialog.askdirectory(title="选择发票目录")
        if not path:
            return
        if path not in self.input_paths:
            self.input_paths.append(path)
        self._refresh_paths_label()
        self.parse_inputs()

    def parse_inputs(self) -> None:
        if self.busy.get():
            return
        if not self.input_paths:
            messagebox.showinfo("提示", "请先添加文件或目录。")
            return
        self._set_busy(True, "正在解析发票...")
        threading.Thread(target=self._parse_worker, daemon=True).start()

    def _parse_worker(self) -> None:
        try:
            files = iter_pdf_files(self.input_paths, self.recursive_var.get())
            records = []
            for file in files:
                try:
                    records.append(parse_invoice(file))
                except Exception as exc:  # noqa: BLE001
                    records.append(
                        InvoiceRecord(
                            source=file,
                            is_invoice_like=False,
                            status="失败",
                            warnings=[str(exc)],
                        )
                    )
            resolve_attachment_metadata(records)
            options = ProcessOptions(
                output_root=self.workspace_root,
                category_names=tuple(self.categories),
                excel_output=None,
            )
            preview = preview_records(records, options)
            self.root.after(0, lambda: self._finish_parse(preview.records, preview.category_summaries, preview.review_notes))
        except Exception as exc:  # noqa: BLE001
            self.root.after(0, lambda: self._fail_operation(f"解析失败: {exc}"))

    def _finish_parse(self, records: list[InvoiceRecord], summaries, review_notes: list[str]) -> None:
        self.records = records
        self._refresh_invoice_tree()
        self._refresh_summary(summaries)
        if review_notes:
            self._set_status("；".join(review_notes[:3]))
        else:
            self._set_status(f"解析完成，共 {len(records)} 条记录")
        self._set_busy(False)

    def refresh_preview(self) -> None:
        if self.busy.get():
            return
        options = ProcessOptions(
            output_root=self.workspace_root,
            category_names=tuple(self.categories),
            excel_output=None,
        )
        preview = preview_records(self.records, options)
        self.records = preview.records
        self._refresh_invoice_tree()
        self._refresh_summary(preview.category_summaries)
        self._set_status("预览已刷新")

    def add_category_prompt(self) -> None:
        if self.busy.get():
            return
        value = simpledialog.askstring("新增类别", "请输入类别名称", parent=self.root)
        if value is None:
            return
        updated = add_category(self.categories, value)
        if updated == self.categories:
            return
        self.categories = updated
        self._refresh_category_list()
        self.refresh_preview()

    def rename_category_prompt(self) -> None:
        if self.busy.get():
            return
        selection = self.category_list.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个类别。")
            return
        old_name = self.categories[selection[0]]
        new_name = simpledialog.askstring("重命名类别", "请输入新类别名称", initialvalue=old_name, parent=self.root)
        if new_name is None:
            return
        self.categories = rename_category(self.categories, old_name, new_name, self.records)
        self._refresh_category_list()
        self.refresh_preview()

    def delete_category_prompt(self) -> None:
        if self.busy.get():
            return
        selection = self.category_list.curselection()
        if not selection:
            messagebox.showinfo("提示", "请先选择一个类别。")
            return
        target = self.categories[selection[0]]
        if not messagebox.askyesno("删除类别", f"确认删除类别「{target}」？"):
            return
        self.categories = remove_category(self.categories, target, self.records)
        self._refresh_category_list()
        self.refresh_preview()

    def execute_current(self) -> None:
        if self.busy.get():
            return
        if not self.records:
            messagebox.showinfo("提示", "请先解析发票。")
            return
        self._set_busy(True, "正在执行归档与导出...")
        threading.Thread(target=self._execute_worker, daemon=True).start()

    def _execute_worker(self) -> None:
        try:
            options = ProcessOptions(
                output_root=self.workspace_root,
                category_names=tuple(self.categories),
                excel_output=Path(self.excel_output_var.get()).expanduser(),
            )
            preview = preview_records(self.records, options)
            result = execute_records(preview.records, options)
            self.root.after(0, lambda: self._finish_execute(result))
        except Exception as exc:  # noqa: BLE001
            self.root.after(0, lambda: self._fail_operation(f"执行失败: {exc}"))

    def _finish_execute(self, result) -> None:
        self.records = result.records
        self._refresh_invoice_tree()
        self._refresh_summary(
            preview_records(
                self.records,
                ProcessOptions(output_root=self.workspace_root, category_names=tuple(self.categories)),
            ).category_summaries
        )
        message = (
            f"执行完成，成功 {result.success_count} 条，"
            f"跳过 {result.skipped_count} 条，失败 {result.failure_count} 条"
        )
        if result.excel_output_path is not None:
            message += f"，Excel 已导出到 {result.excel_output_path}"
        self._set_status(message)
        if result.failures:
            messagebox.showwarning("执行结果", "\n".join(result.failures[:10]))
        self._set_busy(False)

    def _fail_operation(self, message: str) -> None:
        self._set_busy(False, message)
        messagebox.showerror("错误", message)


def run_ui() -> None:
    if hasattr(ctypes, "windll"):
        try:
            handle = ctypes.windll.kernel32.GetConsoleWindow()
            if handle:
                ctypes.windll.user32.ShowWindow(handle, 0)
        except Exception:  # noqa: BLE001
            pass
    root = tk.Tk()
    app = InvoiceDesktopApp(root)
    root.mainloop()
