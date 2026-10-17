# -*- coding: utf-8 -*-
"""不依赖浏览器的论文格式检查桌面版。"""

from __future__ import annotations

import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from utils.format_checker import FORMAT_TEMPLATES, FormatChecker


class PaperFormatDesktop(tk.Tk):
    """Native front end for environments that block localhost connections."""

    def __init__(self):
        super().__init__()
        self.title("论文格式检查工作台")
        self.geometry("1040x700")
        self.minsize(860, 560)
        self.configure(bg="#f4f8fb")
        self.file_path: Path | None = None
        self.issues = []
        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f8fb")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#0b3b70", foreground="#ffffff", font=("Microsoft YaHei UI", 22, "bold"))
        style.configure("Subtitle.TLabel", background="#0b3b70", foreground="#d6e9ff", font=("Microsoft YaHei UI", 10))
        style.configure("TLabel", background="#f4f8fb", foreground="#19324a", font=("Microsoft YaHei UI", 10))
        style.configure("Card.TLabel", background="#ffffff", foreground="#19324a", font=("Microsoft YaHei UI", 10))
        style.configure("Accent.TButton", background="#1769e0", foreground="#ffffff", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 9))
        style.map("Accent.TButton", background=[("active", "#0e55bc")])
        style.configure("Safe.TButton", background="#0d9488", foreground="#ffffff", font=("Microsoft YaHei UI", 10, "bold"), padding=(14, 9))
        style.map("Safe.TButton", background=[("active", "#087a70")])
        style.configure("Treeview", rowheight=29, font=("Microsoft YaHei UI", 9), background="#ffffff", fieldbackground="#ffffff", foreground="#19324a")
        style.configure("Treeview.Heading", background="#eaf2fb", foreground="#19324a", font=("Microsoft YaHei UI", 9, "bold"))

    def _build_ui(self):
        header = ttk.Frame(self, style="Card.TFrame", height=112)
        header.pack(fill="x")
        header.pack_propagate(False)
        banner = tk.Frame(header, bg="#0b3b70")
        banner.pack(fill="both", expand=True)
        ttk.Label(banner, text="论文格式检查工作台", style="Title.TLabel").pack(anchor="w", padx=34, pady=(22, 1))
        ttk.Label(banner, text="桌面本地版 · 无需浏览器 · 文档不离开本机", style="Subtitle.TLabel").pack(anchor="w", padx=35)

        content = ttk.Frame(self, padding=(28, 22))
        content.pack(fill="both", expand=True)
        controls = ttk.Frame(content, style="Card.TFrame", padding=20)
        controls.pack(fill="x")
        ttk.Label(controls, text="1. 选择 Word 文档", style="Card.TLabel", font=("Microsoft YaHei UI", 11, "bold")).grid(row=0, column=0, sticky="w")
        self.file_label = ttk.Label(controls, text="尚未选择文件（仅支持 .docx）", style="Card.TLabel")
        self.file_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(9, 12))
        ttk.Button(controls, text="选择文件", command=self.choose_file).grid(row=2, column=0, sticky="w")
        ttk.Label(controls, text="检查规则", style="Card.TLabel").grid(row=2, column=1, padx=(28, 8), sticky="e")
        self.template_box = ttk.Combobox(controls, values=[f"{key} | {value['name']}" for key, value in FORMAT_TEMPLATES.items()], state="readonly", width=25)
        self.template_box.current(0)
        self.template_box.grid(row=2, column=2, sticky="w")
        controls.columnconfigure(0, weight=1)

        actions = ttk.Frame(content)
        actions.pack(fill="x", pady=16)
        ttk.Button(actions, text="检查格式", style="Accent.TButton", command=self.check_document).pack(side="left")
        ttk.Button(actions, text="备份并自动修复", style="Safe.TButton", command=self.fix_document).pack(side="left", padx=10)
        ttk.Label(actions, text="修复会先在原文件夹创建备份，再生成“_已修复.docx”。", foreground="#4f6478").pack(side="left", padx=8)

        report = ttk.Frame(content, style="Card.TFrame", padding=(1, 1, 1, 12))
        report.pack(fill="both", expand=True)
        self.summary_var = tk.StringVar(value="请选择 .docx 文件后开始检查。")
        ttk.Label(report, textvariable=self.summary_var, style="Card.TLabel", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w", padx=16, pady=13)
        columns = ("category", "location", "description", "suggestion", "fixable")
        self.tree = ttk.Treeview(report, columns=columns, show="headings")
        headings = {"category": "类别", "location": "位置", "description": "发现的问题", "suggestion": "建议", "fixable": "可修复"}
        widths = {"category": 118, "location": 82, "description": 290, "suggestion": 300, "fixable": 75}
        for key in columns:
            self.tree.heading(key, text=headings[key])
            self.tree.column(key, width=widths[key], minwidth=60, anchor="w")
        scrollbar = ttk.Scrollbar(report, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(12, 0))
        scrollbar.pack(side="right", fill="y", padx=(0, 10))

    def _template_id(self):
        return self.template_box.get().split(" | ", 1)[0] or "general"

    def choose_file(self):
        selected = filedialog.askopenfilename(title="选择需要检查的 Word 文档", filetypes=[("Word 文档", "*.docx")])
        if selected:
            self.file_path = Path(selected)
            self.file_label.configure(text=str(self.file_path))
            self.summary_var.set("文件已选择。可先检查，再决定是否修复。")

    def _require_file(self):
        if not self.file_path or not self.file_path.is_file():
            messagebox.showwarning("尚未选择文件", "请先选择一个有效的 .docx 文档。")
            return False
        return True

    def _show_issues(self, issues, stats):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for issue in issues:
            self.tree.insert("", "end", values=(issue.category, issue.location, issue.description, issue.suggestion, "是" if issue.fixable else "否"))
        fixable = sum(issue.fixable for issue in issues)
        self.summary_var.set(f"检查完成：发现 {len(issues)} 项；其中 {fixable} 项可自动修复。 段落 {stats.get('total_paragraphs', 0)}，表格 {stats.get('table_count', 0)}。")

    def check_document(self):
        if not self._require_file():
            return
        try:
            checker = FormatChecker(template_id=self._template_id())
            self.issues, stats = checker.check_file(self.file_path)
            self._show_issues(self.issues, stats)
        except Exception as error:
            messagebox.showerror("检查失败", f"无法读取该文档。请确认文件未损坏且不是受保护文档。\n\n{error}")

    def fix_document(self):
        if not self._require_file():
            return
        if not messagebox.askyesno("确认修复", "将先在原文件夹创建 .bak.docx 备份，再生成新的修复稿。\n\n是否继续？"):
            return
        try:
            backup = self.file_path.with_name(f"{self.file_path.stem}.bak.docx")
            output = self.file_path.with_name(f"{self.file_path.stem}_已修复.docx")
            if backup.exists():
                backup = self.file_path.with_name(f"{self.file_path.stem}.bak_副本.docx")
            if output.exists():
                output = self.file_path.with_name(f"{self.file_path.stem}_已修复_副本.docx")
            shutil.copy2(self.file_path, backup)
            records = FormatChecker(template_id=self._template_id()).fix_file(self.file_path, output)
            issues_after, stats_after = FormatChecker(template_id=self._template_id()).check_file(output)
            self._show_issues(issues_after, stats_after)
            messagebox.showinfo("修复完成", f"已修复 {len(records)} 处。\n\n备份：{backup.name}\n修复稿：{output.name}\n\n仍有 {len(issues_after)} 项建议人工确认。")
        except Exception as error:
            messagebox.showerror("修复失败", f"未能完成修复，原文件未被改动。\n\n{error}")


if __name__ == "__main__":
    PaperFormatDesktop().mainloop()
