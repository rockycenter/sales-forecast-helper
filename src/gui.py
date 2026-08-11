"""销售预测助手 - 图形界面"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np

from .excel_io import load_workbook, get_salespeople, run_forecast, save_result, parse_forecast_months
from .config import OUTPUT_COLUMNS
from . import __version__


class ForecastApp:
    def __init__(self, file_path=None):
        self.root = tk.Tk()
        self.root.title(f'销售预测助手 v{__version__}')
        self.root.geometry("1000x680")
        self.root.minsize(960, 550)

        self.df = None
        self.result_df = None
        self.file_path = file_path
        self.salesperson = None
        self.sheet_name = None
        self.forecast_months = None

        self._build_ui()

        if file_path and os.path.exists(file_path):
            self._load_file(file_path)

        self.root.mainloop()

    # ── UI 构建 ──

    def _build_ui(self):
        # 顶部标题
        header = tk.Frame(self.root, bg="#2196F3", height=50)
        header.pack(fill=tk.X)
        hf = tk.Frame(header, bg="#2196F3")
        hf.pack(pady=10)
        tk.Label(hf, text=f"📊  销售预测助手 v{__version__}",
                 bg="#2196F3", fg="white",
                 font=("Microsoft YaHei", 16, "bold")).pack(side=tk.LEFT, padx=(0, 15))
        tk.Button(hf, text="📖 ABC分类说明", command=self._show_abc_help,
                  bg="#1976D2", fg="white",
                  font=("Microsoft YaHei", 9), padx=10).pack(side=tk.LEFT)

        # 文件选择行
        f1 = tk.Frame(self.root)
        f1.pack(fill=tk.X, padx=20, pady=(15, 5))
        tk.Label(f1, text="📂 Excel 文件：", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        tk.Entry(f1, textvariable=self.file_var, width=60,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=5)
        tk.Button(f1, text="浏览...", command=self._browse_file,
                  font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)

        # 销售员选择行
        f2 = tk.Frame(self.root)
        f2.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(f2, text="👤 销售员：", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.sales_var = tk.StringVar()
        self.sales_combo = ttk.Combobox(f2, textvariable=self.sales_var,
                                         state="disabled", width=40,
                                         font=("Microsoft YaHei", 10))
        self.sales_combo.pack(side=tk.LEFT, padx=5)
        self.sales_combo.bind('<<ComboboxSelected>>', self._on_sales_selected)

        self.run_btn = tk.Button(f2, text="生成预测", command=self._run_forecast,
                                  state="disabled", bg="#4CAF50", fg="white",
                                  font=("Microsoft YaHei", 10, "bold"), padx=15)
        self.run_btn.pack(side=tk.LEFT, padx=10)

        # 表格区域
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 列定义（月份列在 _run_forecast 后动态重建）
        self.tree = ttk.Treeview(table_frame, show="headings",
                                  selectmode="browse", height=15)
        # 滚动条
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        # 双击编辑
        self.tree.bind("<Double-1>", self._on_cell_edit)

        # 底部操作栏（固定高度）
        f3 = tk.Frame(self.root, height=40)
        f3.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=(0, 5))
        f3.pack_propagate(False)

        # 底部状态栏（固定高度）
        self.status_var = tk.StringVar(value="就绪")
        status = tk.Label(self.root, textvariable=self.status_var,
                          relief=tk.SUNKEN, anchor=tk.W, padx=10,
                          font=("Microsoft YaHei", 9), height=1)
        status.pack(fill=tk.X, side=tk.BOTTOM)

        self.summary_var = tk.StringVar(value="")
        tk.Label(f3, textvariable=self.summary_var,
                 font=("Microsoft YaHei", 10), fg="#555").pack(side=tk.LEFT)

        self.export_btn = tk.Button(f3, text="💾 导出 Excel", command=self._export,
                                     state="disabled", bg="#FF9800", fg="white",
                                     font=("Microsoft YaHei", 10, "bold"), padx=15)
        self.export_btn.pack(side=tk.RIGHT)

    # ── 文件操作 ──

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self._load_file(path)

    def _load_file(self, path):
        try:
            self.df, self.sheet_name, _ = load_workbook(path)
            self.forecast_months = parse_forecast_months(self.sheet_name)
            self.file_path = path
            self.file_var.set(path)

            names = get_salespeople(self.df)
            self.sales_combo['values'] = names
            self.sales_combo['state'] = 'readonly'
            self.run_btn['state'] = 'normal'
            self.status_var.set(f"已加载: {os.path.basename(path)} | Sheet: {self.sheet_name} | 共 {len(names)} 位销售员")
        except Exception as e:
            messagebox.showerror("加载失败", str(e))

    # ── 预测 ──

    def _on_sales_selected(self, event=None):
        pass

    def _run_forecast(self):
        person = self.sales_var.get()
        if not person:
            messagebox.showwarning("提示", "请先选择销售员")
            return

        self.status_var.set(f"正在为 {person} 生成预测...")
        self.root.update()

        try:
            self.result_df, warnings = run_forecast(self.df, person, self.forecast_months)
            self.salesperson = person
        except Exception as e:
            messagebox.showerror("预测失败", str(e))
            self.status_var.set("预测失败")
            return

        self._populate_table()
        self._update_summary(warnings)
        self.export_btn['state'] = 'normal'
        self.status_var.set(f"预测完成 — {person} — {len(self.result_df)} 条记录")

    def _populate_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 动态重建列
        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        columns = ("#", "SPEC料号", "Legacy Item", "类型", "Open SO") + tuple(months)
        col_widths = [40, 120, 120, 50, 80] + [90] * len(months)

        self.tree['columns'] = columns
        for col, w in zip(columns, col_widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")

        for i, (_, row) in enumerate(self.result_df.iterrows()):
            vals = [
                i + 1,
                row['SPEC料号'],
                row['Legacy Item'],
                row['产品类型'],
                f"{row['Open_SO']:,}",
            ]
            for m in months:
                vals.append(f"{row[f'推荐_{m}']:,}")
            tag = row['产品类型']
            self.tree.insert("", "end", values=vals, tags=(tag,))

        self.tree.tag_configure('A', background='#E8F5E9')
        self.tree.tag_configure('B', background='#FFF8E1')
        self.tree.tag_configure('C', background='#F5F5F5')

    def _update_summary(self, warnings):
        counts = self.result_df['产品类型'].value_counts()
        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        total = sum(self.result_df[f'推荐_{m}'].sum() for m in months)
        summary = (f"A类: {counts.get('A', 0)}  |  "
                   f"B类: {counts.get('B', 0)}  |  "
                   f"C类: {counts.get('C', 0)}")
        self.summary_var.set(summary)

        if warnings:
            self.status_var.set(f"⚠️ 有 {len(warnings)} 条 Open SO 警告")

    # ── 编辑 ──

    def _on_cell_edit(self, event):
        item = self.tree.selection()
        if not item:
            return
        item = item[0]
        col = self.tree.identify_column(event.x)
        col_idx = int(col.replace("#", "")) - 1

        # 只允许编辑月份列 (列索引 5+)
        if col_idx < 5:
            return

        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        month = months[col_idx - 5]  # skip #, SPEC, Legacy, 类型, Open SO
        values = self.tree.item(item, 'values')
        old_val = values[col_idx]
        row_idx = int(values[0]) - 1  # first column is row number

        # 弹出输入框
        top = tk.Toplevel(self.root)
        top.title("微调预测")
        top.geometry("300x150")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        spec = self.result_df.at[row_idx, 'SPEC料号']
        legacy = self.result_df.at[row_idx, 'Legacy Item']
        tk.Label(top, text=f"产品: {spec} / {legacy}", font=("Microsoft YaHei", 10)).pack(pady=(15, 5))
        tk.Label(top, text=f"{month}推荐值:", font=("Microsoft YaHei", 10)).pack()

        entry_var = tk.StringVar(value=old_val.replace(",", ""))
        entry = tk.Entry(top, textvariable=entry_var, font=("Microsoft YaHei", 12),
                         justify="center", width=15)
        entry.pack(pady=5)
        entry.select_range(0, tk.END)
        entry.focus()

        columns = self.tree['columns']

        def confirm():
            try:
                new_val = int(float(entry_var.get()))
                col_name = f"推荐_{month}"
                self.result_df.at[row_idx, col_name] = new_val
                self.tree.set(item, columns[col_idx], f"{new_val:,}")
                self._update_summary([])
                top.destroy()
            except ValueError:
                messagebox.showwarning("格式错误", "请输入整数", parent=top)

        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="确定", command=confirm,
                  bg="#4CAF50", fg="white", padx=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=top.destroy, padx=15).pack(side=tk.LEFT)

    # ── 导出 ──

    def _export(self):
        if self.result_df is None:
            return

        path = filedialog.asksaveasfilename(
            title="保存预测结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=f"预测结果_{self.salesperson.replace(' ', '_')}.xlsx"
        )
        if not path:
            return

        try:
            save_result(self.result_df, self.salesperson, path)
            self.status_var.set(f"✅ 已导出: {path}")
            messagebox.showinfo("导出成功", f"文件已保存到:\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    # ── ABC 说明 ──

    def _show_abc_help(self):
        top = tk.Toplevel(self.root)
        top.title("ABC 分类说明")
        top.geometry("560x440")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        text = tk.Text(top, wrap=tk.WORD, padx=20, pady=15,
                       font=("Microsoft YaHei", 10), relief=tk.FLAT,
                       bg="#FAFAFA")
        text.pack(fill=tk.BOTH, expand=True)

        help_text = """A B C  分 类 说 明
━━━━━━━━━━━━━━━━━━━━━━━━━━━

程序根据每个产品过去 12 个月的历史销售数据，
自动将其分为 A / B / C 三类：

🟢 A 类 — 稳定产品
   判定：有 ≥8 个月数据、波动系数 <0.8、月均 ≥5,000
   算法：近期趋势 + 去年同期 + 全年平均，加权融合
   典型：常年稳定出货的大单品

🟡 B 类 — 波动/季节性产品
   判定：有 ≥6 个月数据，但波动剧烈或季节性明显
   算法：中位数法 + 上下限控制，避免追高踩低
   典型：大单拉动型、淡旺季明显的产品

⚪ C 类 — 稀疏/微量产品
   判定：数据不足 6 个月，或月均 <3,000
   算法：保守估计，有 Open SO 取 Open SO，否则打折
   典型：新品、尾货、微量零星出货

━━━━━━━━━━━━━━━━━━━━━━━━━━━
💡 算法建议 ≠ 最终决定
   如果你知道客户已确认大单、促销计划等额外信息，
   请双击表格单元格手动微调后再导出。

📌 颜色标注：A 绿色 | B 黄色 | C 灰色"""

        text.insert("1.0", help_text)
        text.configure(state="disabled")

        btn = tk.Button(top, text="知道了", command=top.destroy,
                        bg="#2196F3", fg="white",
                        font=("Microsoft YaHei", 10, "bold"), padx=30)
        btn.pack(pady=10)


def launch_gui(file_path=None):
    ForecastApp(file_path)
