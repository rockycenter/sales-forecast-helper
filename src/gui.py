"""销售预测助手 - 图形界面"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np

from .excel_io import load_workbook, get_salespeople, run_forecast, save_result, parse_forecast_months, parse_forecast_year, compute_all_quarters
from .predictors import smart_round
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
        self.history_count = 12
        self.forecast_year = None
        self.manual_overrides = {}

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

        # 档位增量调控行
        f2b = tk.Frame(self.root)
        f2b.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(f2b, text="🔧 档位增量：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        self.tier_adjust_vars = {}
        for ptype, bgc in [('A', '#FFEBEE'), ('B', '#FFF8E1'), ('C', '#FFFFFF')]:
            tk.Label(f2b, text=f"{ptype}类", font=("Microsoft YaHei", 10), bg=bgc,
                     relief=tk.GROOVE, padx=5).pack(side=tk.LEFT, padx=(8, 3))
            var = tk.DoubleVar(value=0)
            sp = tk.Spinbox(f2b, from_=-100, to=100, increment=5, textvariable=var,
                            width=6, font=("Microsoft YaHei", 10), justify="center")
            sp.pack(side=tk.LEFT, padx=2)
            tk.Label(f2b, text="%", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
            self.tier_adjust_vars[ptype] = var
        tk.Button(f2b, text="应用增量", command=self._apply_tier_adjust,
                  bg="#FF5722", fg="white",
                  font=("Microsoft YaHei", 9, "bold"), padx=10).pack(side=tk.LEFT, padx=12)
        tk.Label(f2b, text="（负数=下调，正数=上调）", font=("Microsoft YaHei", 9), fg="#777").pack(side=tk.LEFT)

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
            self.df, self.sheet_name, _, self.history_count = load_workbook(path)
            self.forecast_months = parse_forecast_months(self.sheet_name)
            self.forecast_year = parse_forecast_year(self.sheet_name)
            self.file_path = path
            self.file_var.set(path)

            names = get_salespeople(self.df)
            self.sales_combo['values'] = names
            self.sales_combo['state'] = 'readonly'
            self.run_btn['state'] = 'normal'
            self.status_var.set(f"已加载: {os.path.basename(path)} | Sheet: {self.sheet_name} | 历史: {self.history_count}个月 | 销售员: {len(names)} 人")
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
            self.result_df, warnings = run_forecast(self.df, person, self.forecast_months, self.history_count, self.forecast_year)
            self.salesperson = person
        except Exception as e:
            messagebox.showerror("预测失败", str(e))
            self.status_var.set("预测失败")
            return


        self._populate_table()
        self._update_summary(warnings)
        self.export_btn['state'] = 'normal'
        self.status_var.set(f"预测完成 — {person} — {len(self.result_df)} 条记录")

    def _get_quarter_labels(self):
        """从 result_df 列名提取季度标签，如 ['2026Q3','2026Q4']"""
        if self.result_df is None:
            return []
        labels = []
        for col in self.result_df.columns:
            if col.endswith('_今年'):
                labels.append(col[:-3])
        return labels

    def _populate_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 动态重建列
        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        q_labels = self._get_quarter_labels()
        
        columns = ["#", "SPEC料号", "Legacy Item", "类型", "Open SO"] + list(months)
        for q in q_labels:
            columns += [f"{q}_今年", f"{q}_去年", f"{q}_同比"]
        col_widths = [40, 100, 100, 50, 80] + [80] * len(months) + [90, 90, 70] * len(q_labels)

        self.tree['columns'] = tuple(columns)
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
            for q in q_labels:
                vals.append(f"{int(row[f'{q}_今年']):,}")
                vals.append(f"{int(row[f'{q}_去年']):,}")
                pct = row.get(f'{q}_同比', '')
                vals.append(f"{pct}%" if pct != '' else 'N/A')
            tag = row['产品类型']
            self.tree.insert("", "end", values=vals, tags=(tag,))

        self.tree.tag_configure('A', background='#FFEBEE')
        self.tree.tag_configure('B', background='#FFF8E1')
        self.tree.tag_configure('C', background='#FFFFFF')

    def _update_summary(self, warnings):
        counts = self.result_df['产品类型'].value_counts()
        parts = [f"A类: {counts.get('A', 0)}", f"B类: {counts.get('B', 0)}", f"C类: {counts.get('C', 0)}"]
        for q in self._get_quarter_labels():
            this_col = f"{q}_今年"
            last_col = f"{q}_去年"
            if this_col in self.result_df.columns and last_col in self.result_df.columns:
                q_total = int(self.result_df[this_col].sum())
                q_last = int(self.result_df[last_col].sum())
                if q_last > 0:
                    q_pct = round((q_total - q_last) / q_last * 100, 1)
                    arrow = '↑' if q_pct >= 0 else '↓'
                    parts.append(f"{q}: {arrow}{abs(q_pct)}%")
                else:
                    parts.append(f"{q}: N/A")
        self.summary_var.set("  |  ".join(parts))

        if warnings:
            self.status_var.set(f"⚠️ 有 {len(warnings)} 条 Open SO 警告")

    def _recompute_all_quarters(self):
        """根据当前 推荐_X月 值，重算所有行的季度同比"""
        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        first_month_num = int(months[0].replace('月', ''))
        for row_idx in range(len(self.result_df)):
            row = self.result_df.iloc[row_idx]
            history = row['_历史']
            forecasts = [row[f'推荐_{m}'] for m in months]
            quarter_results, _ = compute_all_quarters(
                history, forecasts, first_month_num, self.forecast_year
            )
            for qr in quarter_results:
                q = qr['label']
                self.result_df.at[row_idx, f'{q}_今年'] = qr['this']
                self.result_df.at[row_idx, f'{q}_去年'] = qr['last']
                self.result_df.at[row_idx, f'{q}_同比'] = qr['pct'] if qr['pct'] is not None else ''
                self.result_df.at[row_idx, f'{q}_有效月'] = qr['valid']

    def _apply_tier_adjust(self):
        """按 A/B/C 档位百分比增量，从原始预测重算所有月份值"""
        if self.result_df is None:
            messagebox.showwarning("提示", "请先生成预测")
            return

        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        for row_idx in range(len(self.result_df)):
            ptype = self.result_df.at[row_idx, '产品类型']
            factor = 1 + self.tier_adjust_vars[ptype].get() / 100
            base_forecasts = self.result_df.at[row_idx, '_预测']
            for i, m in enumerate(months):
                col = f'推荐_{m}'
                if (row_idx, col) in self.manual_overrides:
                    new_val = self.manual_overrides[(row_idx, col)]
                else:
                    new_val = smart_round(base_forecasts[i] * factor)
                self.result_df.at[row_idx, col] = new_val

        self._recompute_all_quarters()
        self._populate_table()
        self._update_summary([])
        self.status_var.set(f"已应用档位增量：A={self.tier_adjust_vars['A'].get():.0f}%  "
                            f"B={self.tier_adjust_vars['B'].get():.0f}%  "
                            f"C={self.tier_adjust_vars['C'].get():.0f}%")

    # ── 编辑 ──

    def _on_cell_edit(self, event):
        item = self.tree.selection()
        if not item:
            return
        item = item[0]
        col = self.tree.identify_column(event.x)
        col_idx = int(col.replace("#", "")) - 1

        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        # 只允许编辑月份列 (列索引 5 到 5+len(months)-1)
        if col_idx < 5 or col_idx >= 5 + len(months):
            return

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
                self.manual_overrides[(row_idx, col_name)] = new_val
                self.tree.set(item, columns[col_idx], f"{new_val:,}")
                # 实时重算该行季度同比
                self._refresh_quarter_row(row_idx, item, columns)
                top.destroy()
            except ValueError:
                messagebox.showwarning("格式错误", "请输入整数", parent=top)

        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="确定", command=confirm,
                  bg="#4CAF50", fg="white", padx=15).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=top.destroy, padx=15).pack(side=tk.LEFT)

    # ── 季度同比实时更新 ──

    def _refresh_quarter_row(self, row_idx, item, columns):
        """编辑某个预测值后，重算该行的季度同比并刷新显示"""
        row = self.result_df.iloc[row_idx]
        history = row['_历史']
        months = self.forecast_months or ["8月", "9月", "10月", "11月"]
        forecasts = [row[f'推荐_{m}'] for m in months]
        first_month_num = int(months[0].replace('月', ''))

        quarter_results, _ = compute_all_quarters(
            history, forecasts, first_month_num, self.forecast_year
        )
        for qr in quarter_results:
            q = qr['label']
            self.result_df.at[row_idx, f'{q}_今年'] = qr['this']
            self.result_df.at[row_idx, f'{q}_去年'] = qr['last']
            self.result_df.at[row_idx, f'{q}_同比'] = qr['pct'] if qr['pct'] is not None else ''
            self.result_df.at[row_idx, f'{q}_有效月'] = qr['valid']

        # 更新表格显示（季度列在月份列之后）
        q_labels = self._get_quarter_labels()
        base = 5 + len(months)
        for j, q in enumerate(q_labels):
            row2 = self.result_df.iloc[row_idx]
            self.tree.set(item, columns[base + j*3], f"{int(row2[f'{q}_今年']):,}")
            self.tree.set(item, columns[base + j*3 + 1], f"{int(row2[f'{q}_去年']):,}")
            pct = row2.get(f'{q}_同比', '')
            self.tree.set(item, columns[base + j*3 + 2], f"{pct}%" if pct != '' else 'N/A')

        # 刷新汇总
        self._update_summary([])

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
