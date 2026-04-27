from __future__ import annotations

import ctypes
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from .app_icon import apply_window_icon, configure_windows_app_id
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


class InvoiceRowFrame(ttk.Frame):
    """单个发票行的框架，包含发票信息和类别复选框"""
    def __init__(self, parent, record: InvoiceRecord, index: int, categories: list[str], on_change_callback):
        super().__init__(parent, relief="solid", borderwidth=2, padding=15)
        self.record = record
        self.index = index
        self.categories = categories
        self.on_change_callback = on_change_callback
        self.category_vars = {}

        self._build_ui()

    def _build_ui(self):
        # 发票信息行
        info_frame = ttk.Frame(self)
        info_frame.pack(fill="x", pady=(0, 10))

        # 文件名
        file_label = ttk.Label(
            info_frame,
            text=f"📄 {self.record.source.name if self.record.source else '未知文件'}",
            font=("Microsoft YaHei UI", 11, "bold")
        )
        file_label.pack(side="left", padx=(0, 15))

        # 发票号
        if self.record.invoice_number:
            ttk.Label(
                info_frame,
                text=f"发票号: {self.record.invoice_number[-8:]}",
                foreground="gray",
                font=("Microsoft YaHei UI", 10)
            ).pack(side="left", padx=(0, 15))

        # 销售方
        if self.record.seller_name:
            ttk.Label(
                info_frame,
                text=f"销售方: {self.record.seller_name[:20]}",
                foreground="gray",
                font=("Microsoft YaHei UI", 10)
            ).pack(side="left", padx=(0, 15))

        # 金额
        if self.record.total_amount is not None:
            ttk.Label(
                info_frame,
                text=f"金额: ¥{self.record.total_amount:.2f}",
                foreground="green",
                font=("Microsoft YaHei UI", 11, "bold")
            ).pack(side="left", padx=(0, 15))

        # 状态
        status_color = "green" if self.record.status == "已就绪" else "orange"
        ttk.Label(
            info_frame,
            text=self.record.status,
            foreground=status_color,
            font=("Microsoft YaHei UI", 10)
        ).pack(side="right")

        # 类别复选框行
        category_frame = ttk.Frame(self)
        category_frame.pack(fill="x")

        ttk.Label(
            category_frame,
            text="类别:",
            foreground="gray",
            font=("Microsoft YaHei UI", 10, "bold")
        ).pack(side="left", padx=(0, 15))

        for category in self.categories:
            var = tk.BooleanVar(value=category in self.record.categories)
            self.category_vars[category] = var

            cb = ttk.Checkbutton(
                category_frame,
                text=category,
                variable=var,
                command=lambda c=category: self._on_category_change(c)
            )
            cb.pack(side="left", padx=10)

    def _on_category_change(self, category: str):
        """类别复选框变化时的回调"""
        if self.category_vars[category].get():
            if category not in self.record.categories:
                self.record.categories.append(category)
        else:
            if category in self.record.categories:
                self.record.categories.remove(category)

        if self.on_change_callback:
            self.on_change_callback()

    def update_categories(self, categories: list[str]):
        """更新可用类别列表"""
        self.categories = categories
        self._build_ui()


class ImprovedInvoiceApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.workspace_root = get_workspace_root()
        self.root.title("发票归类工具")
        apply_window_icon(self.root)
        self.root.geometry("1400x850")
        self.root.minsize(1200, 700)

        self.input_paths: list[str] = [str(self.workspace_root)]
        self.records: list[InvoiceRecord] = []
        self.categories: list[str] = list(DEFAULT_CATEGORIES)
        self.busy = tk.BooleanVar(value=False)
        self.recursive_var = tk.BooleanVar(value=False)
        self.batch_category_var = tk.StringVar(value="")
        self.excel_output_var = tk.StringVar(
            value=str(self.workspace_root / "发票信息汇总.xlsx")
        )
        self.invoice_frames: list[InvoiceRowFrame] = []

        self._build_ui()
        self.root.after(200, self.parse_inputs)

    def _build_ui(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        # 配置样式 - 增大按钮和复选框
        style.configure("TButton", padding=(15, 10), font=("Microsoft YaHei UI", 10))
        style.configure("TCheckbutton", padding=5, font=("Microsoft YaHei UI", 10))
        style.configure("TLabel", font=("Microsoft YaHei UI", 10))
        style.configure("TRadiobutton", padding=5, font=("Microsoft YaHei UI", 10))

        main_container = ttk.Frame(self.root, padding=15)
        main_container.pack(fill="both", expand=True)

        self._build_toolbar(main_container)
        self._build_content(main_container)
        self._build_statusbar(main_container)

    def _build_toolbar(self, parent: ttk.Frame) -> None:
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill="x", pady=(0, 15))

        # 第一行：文件操作
        top_row = ttk.Frame(toolbar)
        top_row.pack(fill="x", pady=(0, 10))

        left_frame = ttk.Frame(top_row)
        left_frame.pack(side="left", fill="x", expand=True)

        ttk.Button(left_frame, text="📁 添加文件", command=self.add_files, width=14).pack(side="left", padx=5)
        ttk.Button(left_frame, text="📂 添加目录", command=self.add_folder, width=14).pack(side="left", padx=5)
        ttk.Button(left_frame, text="🔄 重新解析", command=self.parse_inputs, width=14).pack(side="left", padx=5)

        ttk.Separator(left_frame, orient="vertical").pack(side="left", fill="y", padx=15)

        ttk.Checkbutton(left_frame, text="递归扫描子目录", variable=self.recursive_var).pack(side="left", padx=10)

        right_frame = ttk.Frame(top_row)
        right_frame.pack(side="right")

        ttk.Button(right_frame, text="⚙️ 管理类别", command=self.manage_categories, width=14).pack(side="left", padx=5)
        ttk.Button(right_frame, text="✅ 开始执行", command=self.execute_current, width=14).pack(side="left", padx=5)

        # 第二行：批量归类
        bottom_row = ttk.Frame(toolbar)
        bottom_row.pack(fill="x")

        ttk.Label(bottom_row, text="批量归类:", font=("Microsoft YaHei UI", 10, "bold")).pack(side="left", padx=(0, 10))

        self.batch_category_var = tk.StringVar(value="")

        for category in self.categories:
            ttk.Radiobutton(
                bottom_row,
                text=category,
                variable=self.batch_category_var,
                value=category,
                command=self.apply_batch_category
            ).pack(side="left", padx=8)

        ttk.Button(bottom_row, text="清除批量归类", command=self.clear_batch_category, width=14).pack(side="left", padx=10)

    def _build_content(self, parent: ttk.Frame) -> None:
        content = ttk.PanedWindow(parent, orient="horizontal")
        content.pack(fill="both", expand=True)

        left_panel = self._build_invoice_panel(content)
        right_panel = self._build_summary_panel(content)

        content.add(left_panel, weight=7)
        content.add(right_panel, weight=3)

    def _build_invoice_panel(self, parent: ttk.PanedWindow) -> ttk.Frame:
        panel = ttk.Frame(parent)

        header = ttk.Frame(panel)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="发票列表", font=("Microsoft YaHei UI", 12, "bold")).pack(side="left")
        ttk.Label(header, text="（直接勾选类别复选框）", foreground="gray", font=("Microsoft YaHei UI", 10)).pack(side="left", padx=15)

        # 创建滚动区域
        canvas = tk.Canvas(panel, highlightthickness=0)
        scrollbar = ttk.Scrollbar(panel, orient="vertical", command=canvas.yview)
        self.invoice_container = ttk.Frame(canvas)

        self.invoice_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.invoice_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 鼠标滚轮支持
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        return panel

    def _build_summary_panel(self, parent: ttk.PanedWindow) -> ttk.Frame:
        panel = ttk.Frame(parent)

        header = ttk.Frame(panel)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="分类汇总", font=("Microsoft YaHei UI", 12, "bold")).pack(side="left")

        tree_frame = ttk.Frame(panel)
        tree_frame.pack(fill="both", expand=True, pady=(0, 15))

        style = ttk.Style()
        style.configure("Treeview", rowheight=30, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

        self.summary_tree = ttk.Treeview(
            tree_frame,
            columns=("category", "count", "amount"),
            show="headings",
            height=15
        )

        self.summary_tree.heading("category", text="类别")
        self.summary_tree.heading("count", text="票数")
        self.summary_tree.heading("amount", text="金额")

        self.summary_tree.column("category", width=140)
        self.summary_tree.column("count", width=70, anchor="center")
        self.summary_tree.column("amount", width=120, anchor="e")

        self.summary_tree.pack(fill="both", expand=True)

        excel_frame = ttk.LabelFrame(panel, text="Excel 输出", padding=15)
        excel_frame.pack(fill="x", pady=(15, 0))

        ttk.Entry(excel_frame, textvariable=self.excel_output_var, font=("Microsoft YaHei UI", 10)).pack(fill="x", pady=(0, 8))
        ttk.Button(excel_frame, text="选择路径", command=self.choose_excel_output).pack(fill="x")

        return panel

    def _build_statusbar(self, parent: ttk.Frame) -> None:
        self.status_var = tk.StringVar(value=f"就绪 | 工作目录: {self.workspace_root}")
        statusbar = ttk.Label(
            parent,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            font=("Microsoft YaHei UI", 10),
            padding=8
        )
        statusbar.pack(fill="x", pady=(15, 0))

    def _refresh_invoice_list(self) -> None:
        """刷新发票列表显示"""
        # 清空现有框架
        for frame in self.invoice_frames:
            frame.destroy()
        self.invoice_frames.clear()

        # 创建新的发票框架
        for index, record in enumerate(self.records):
            frame = InvoiceRowFrame(
                self.invoice_container,
                record,
                index,
                self.categories,
                self.refresh_preview
            )
            frame.pack(fill="x", pady=8, padx=10)
            self.invoice_frames.append(frame)

    def _refresh_toolbar_categories(self) -> None:
        """刷新工具栏的批量归类单选按钮"""
        # 重新构建整个UI以更新类别列表
        # 这是一个简化的方法，实际上应该只更新工具栏部分
        # 但为了简单起见，我们保存当前状态后重建
        pass  # 暂时不实现，因为需要重构工具栏

    def _refresh_summary(self, summaries) -> None:
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)

        for summary in summaries:
            self.summary_tree.insert(
                "",
                tk.END,
                values=(
                    summary.name,
                    summary.invoice_count,
                    f"{summary.total_amount:.2f}"
                )
            )

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择发票文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if paths:
            self.input_paths.extend([p for p in paths if p not in self.input_paths])
            self.parse_inputs()

    def add_folder(self) -> None:
        path = filedialog.askdirectory(title="选择发票目录")
        if path and path not in self.input_paths:
            self.input_paths.append(path)
            self.parse_inputs()

    def choose_excel_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="选择 Excel 输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if path:
            self.excel_output_var.set(path)

    def manage_categories(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("管理类别")
        apply_window_icon(dialog)
        dialog.geometry("400x500")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="类别管理", font=("", 11, "bold")).pack(pady=10)

        listbox = tk.Listbox(dialog, height=15)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        def refresh_list():
            listbox.delete(0, tk.END)
            for cat in self.categories:
                listbox.insert(tk.END, cat)

        refresh_list()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=10, pady=10)

        def add_cat():
            name = simpledialog.askstring("新增类别", "请输入类别名称:", parent=dialog)
            if name and name.strip():
                self.categories = add_category(self.categories, name)
                refresh_list()
                self._refresh_toolbar_categories()
                self._refresh_invoice_list()
                self.refresh_preview()

        def rename_cat():
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("提示", "请先选择一个类别", parent=dialog)
                return
            old = self.categories[sel[0]]
            new = simpledialog.askstring("重命名类别", "请输入新名称:", initialvalue=old, parent=dialog)
            if new:
                self.categories = rename_category(self.categories, old, new, self.records)
                refresh_list()
                self._refresh_toolbar_categories()
                self._refresh_invoice_list()
                self.refresh_preview()

        def delete_cat():
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("提示", "请先选择一个类别", parent=dialog)
                return
            cat = self.categories[sel[0]]
            if messagebox.askyesno("确认", f"确定删除类别「{cat}」？", parent=dialog):
                self.categories = remove_category(self.categories, cat, self.records)
                refresh_list()
                self._refresh_toolbar_categories()
                self._refresh_invoice_list()
                self.refresh_preview()

        ttk.Button(button_frame, text="新增", command=add_cat).pack(side="left", padx=2)
        ttk.Button(button_frame, text="重命名", command=rename_cat).pack(side="left", padx=2)
        ttk.Button(button_frame, text="删除", command=delete_cat).pack(side="left", padx=2)
        ttk.Button(button_frame, text="关闭", command=dialog.destroy).pack(side="right", padx=2)

    def apply_batch_category(self) -> None:
        """应用批量归类"""
        category = self.batch_category_var.get()
        if not category:
            return

        for record in self.records:
            if record.is_invoice_like:
                record.categories = [category]

        self._refresh_invoice_list()
        self.refresh_preview()
        self.status_var.set(f"已将所有发票归类到「{category}」")

    def clear_batch_category(self) -> None:
        """清除批量归类"""
        self.batch_category_var.set("")
        self.status_var.set("已清除批量归类选择")

    def batch_categorize(self) -> None:
        """批量归类对话框（已废弃，保留兼容性）"""
        pass

    def parse_inputs(self) -> None:
        if self.busy.get():
            return
        if not self.input_paths:
            messagebox.showinfo("提示", "请先添加文件或目录")
            return

        self.busy.set(True)
        self.status_var.set("正在解析发票...")
        threading.Thread(target=self._parse_worker, daemon=True).start()

    def _parse_worker(self) -> None:
        try:
            files = iter_pdf_files(self.input_paths, self.recursive_var.get())
            records = []
            for file in files:
                try:
                    records.append(parse_invoice(file))
                except Exception as exc:
                    records.append(
                        InvoiceRecord(
                            source=file,
                            is_invoice_like=False,
                            status="失败",
                            warnings=[str(exc)]
                        )
                    )

            resolve_attachment_metadata(records)
            options = ProcessOptions(
                output_root=self.workspace_root,
                category_names=tuple(self.categories),
                excel_output=None
            )
            preview = preview_records(records, options)

            self.root.after(0, lambda: self._finish_parse(preview.records, preview.category_summaries))
        except Exception as exc:
            self.root.after(0, lambda: self._show_error(f"解析失败: {exc}"))

    def _finish_parse(self, records: list[InvoiceRecord], summaries) -> None:
        self.records = records
        self._refresh_invoice_list()
        self._refresh_summary(summaries)
        self.status_var.set(f"解析完成 | 共 {len(records)} 条记录")
        self.busy.set(False)

    def refresh_preview(self) -> None:
        if self.busy.get():
            return

        options = ProcessOptions(
            output_root=self.workspace_root,
            category_names=tuple(self.categories),
            excel_output=None
        )
        preview = preview_records(self.records, options)
        self.records = preview.records
        self._refresh_summary(preview.category_summaries)

    def execute_current(self) -> None:
        if self.busy.get():
            return
        if not self.records:
            messagebox.showinfo("提示", "请先解析发票")
            return

        if not messagebox.askyesno("确认", "确定要执行归档操作吗？\n\n注意：原文件将被移动到分类目录"):
            return

        self.busy.set(True)
        self.status_var.set("正在执行归档...")
        threading.Thread(target=self._execute_worker, daemon=True).start()

    def _execute_worker(self) -> None:
        try:
            options = ProcessOptions(
                output_root=self.workspace_root,
                category_names=tuple(self.categories),
                excel_output=Path(self.excel_output_var.get()).expanduser()
            )
            preview = preview_records(self.records, options)
            result = execute_records(preview.records, options)

            self.root.after(0, lambda: self._finish_execute(result))
        except Exception as exc:
            self.root.after(0, lambda: self._show_error(f"执行失败: {exc}"))

    def _finish_execute(self, result) -> None:
        self.records = result.records
        self._refresh_invoice_list()

        options = ProcessOptions(
            output_root=self.workspace_root,
            category_names=tuple(self.categories)
        )
        preview = preview_records(self.records, options)
        self._refresh_summary(preview.category_summaries)

        message = f"执行完成 | 成功: {result.success_count}, 跳过: {result.skipped_count}, 失败: {result.failure_count}"
        if result.excel_output_path:
            message += f" | Excel: {result.excel_output_path.name}"

        self.status_var.set(message)
        self.busy.set(False)

        if result.failures:
            messagebox.showwarning("执行结果", "\n".join(result.failures[:10]))
            # 有失败的情况下不自动关闭，让用户查看错误
        else:
            # 成功完成，显示提示后自动关闭
            messagebox.showinfo("成功", f"归档完成！\n\n成功处理 {result.success_count} 个文件\n\n程序将自动关闭")
            self.root.after(500, self.root.destroy)  # 延迟500ms后关闭窗口

    def _show_error(self, message: str) -> None:
        self.status_var.set(f"错误: {message}")
        self.busy.set(False)
        messagebox.showerror("错误", message)


def run_improved_ui() -> None:
    configure_windows_app_id()
    if hasattr(ctypes, "windll"):
        try:
            handle = ctypes.windll.kernel32.GetConsoleWindow()
            if handle:
                ctypes.windll.user32.ShowWindow(handle, 0)
        except Exception:
            pass

    root = tk.Tk()
    apply_window_icon(root)
    app = ImprovedInvoiceApp(root)
    root.mainloop()
