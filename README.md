# 销售预测助手 (Sales Forecast Helper)

自动分析历史销售数据，按 ABC 分类法为每个产品生成未来 4 个月的销量预测推荐值。

---

## ✨ 功能特点

- **图形界面**：可视化操作，浏览选择文件、下拉选择销售员、双击表格编辑
- **智能分类**：A(稳定) / B(波动) / C(稀疏) 三分类算法
- **Open SO 检查**：自动校验 8 月预测 ≥ 在手订单
- **实时微调**：双击预测值直接修改，汇总自动更新
- **一键导出**：确认后输出标准 Excel

---

## 📦 安装与运行

### 方式一：双击 EXE（推荐，无需安装任何东西）

从 [Releases](https://github.com/rockycenter/sales-forecast-helper/releases) 或 [Actions](https://github.com/rockycenter/sales-forecast-helper/actions) 下载 `销售预测助手.exe`，双击运行。

### 方式二：有 Python 环境

```bash
git clone git@github.com:rockycenter/sales-forecast-helper.git
cd sales-forecast-helper
pip install -r requirements.txt

# 图形界面（默认）
python forecast.py

# 或直接传文件
python forecast.py -i "文件路径.xlsx"

# 命令行模式
python forecast.py --cli -i "文件路径.xlsx"
```

| 系统 | 快捷启动 |
|------|---------|
| **Windows** | 双击 `run.bat` 或拖拽 Excel 到 bat 上 |
| **macOS** | 双击 `run.sh` |

---

## 🖥️ 使用流程

1. **浏览/拖入** Excel 文件
2. 下拉选择**销售员**
3. 点击 **"生成预测"**
4. 用 A/B/C **档位增量** 整体调控（负数=下调，正数=上调）
5. 双击表格中任意月份**微调数值**
6. 点击 **"导出 Excel"** 保存最终结果

---

## 📁 项目结构

```
sales-forecast-helper/
├── forecast.py           # 入口（默认 GUI，--cli 切换命令行）
├── src/
│   ├── config.py         # 配置中心
│   ├── classifier.py     # ABC 分类
│   ├── predictors.py     # 预测算法
│   ├── excel_io.py       # Excel 读写
│   ├── gui.py            # 图形界面 (tkinter)
│   └── cli.py            # 命令行界面
├── .github/workflows/    # GitHub Actions 自动打包 EXE
├── build_exe.bat         # 本地打包 EXE 脚本
├── run.bat / run.sh      # 快捷启动
├── requirements.txt
└── README.md
```

---

## 📜 License

[MIT License](LICENSE)
