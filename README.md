# 销售预测助手 (Sales Forecast Helper)

自动分析历史销售数据，按 ABC 分类法为每个产品生成未来 4 个月的销量预测推荐值。

---

## ✨ 功能特点

- **智能分类**：自动将产品分为 A(稳定) / B(波动) / C(稀疏) 三类
- **分层算法**：针对不同产品类型使用最适合的预测模型
- **Open SO 检查**：自动校验 8 月预测是否覆盖在手订单
- **交互式选择**：从 Excel 中自动列出所有销售员供选择，无需手动输入
- **终端预览**：预测结果直接在终端展示，无需打开中间文件
- **即时微调**：可在终端内逐条修改推荐值，确认后一次性输出最终 Excel

---

## 📦 安装

### 方式一：命令行运行

```bash
# 1. 克隆仓库
git clone https://github.com/rockycenter/sales-forecast-helper.git
cd sales-forecast-helper

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行（交互模式）
python forecast.py -i "你的Excel文件路径.xlsx"
```

### 方式二：双击启动

| 系统 | 文件 |
|------|------|
| **Windows** | 双击 `run.bat` |
| **macOS / Linux** | 双击 `run.sh` |

---

## 🚀 使用流程

```
📂 拖入 Excel → 👤 选择销售员 → 📊 终端预览预测 → 🔧 逐条微调 → 📁 输出最终 Excel
```

### 交互示例

```
============================================================
         📊 销售预测助手 v2.0
============================================================

📂 读取文件: 2026经销商预测.xlsx
📊 使用Sheet: 销售预测收集26年8-11月

👥 共发现 5 位销售员:
     1. ALICE WANG (45 个产品)
     2. BOB LI (62 个产品)
     3. ROCKY JIANG (81 个产品)
   ...

🔍 请选择销售员 (1-5) > 3

⏳ 正在为 ROCKY JIANG 生成预测...

📈 预测结果统计:
   A类(稳定): 31 个
   B类(波动): 18 个
   C类(稀疏): 32 个

🔧 是否需要微调？(y/n) > y

🔧 微调模式
   输入格式: <行号> <月份(8/9/10/11)> <新值>
   示例: 3 8 50000  (将第3行8月推荐改为50000)

✏️  微调 > 5 8 60000
   ✅ #5 ABC-123 8月: 45,000 → 60,000

✏️  微调 > q

💾 确认输出最终 Excel？(y/n) > y
✅ 结果已保存: 预测结果_ROCKY_JIANG_0810_2100.xlsx
```

---

## 📊 ABC 分类逻辑

| 类型 | 判定标准 | 预测方法 |
|------|---------|---------|
| **A - 稳定产品** | 数据≥8月，CV<0.8，平均≥5000 | 加权移动平均 + 趋势修正 |
| **B - 波动/季节性** | 数据≥6月，高波动或季节性明显 | 中位数法 + 峰值/谷底控制 |
| **C - 稀疏/微量** | 数据<6月 或 平均<3000 | 保守估计，避免虚高 |

---

## 📁 项目结构

```
sales-forecast-helper/
├── forecast.py           # 程序入口
├── src/
│   ├── __init__.py
│   ├── config.py         # 配置中心（阈值、列映射、权重）
│   ├── classifier.py     # ABC 产品分类
│   ├── predictors.py     # A/B/C 三类预测算法
│   ├── excel_io.py       # Excel 读写 + 预测编排
│   └── cli.py            # 交互式命令行界面
├── run.bat               # Windows 一键启动
├── run.sh                # macOS/Linux 一键启动
├── requirements.txt      # Python 依赖
├── README.md             # 本文件
├── LICENSE               # MIT 许可证
└── .gitignore
```

---

## ⚠️ 注意事项

1. **Open SO 规则**：8 月预测值 **必须 ≥** Open SO，程序会自动检查并警告
2. **算法 vs 业务判断**：算法基于纯历史数据，微调阶段请根据业务实际情况调整
3. **Excel 格式要求**：需包含含"销售预测收集"关键字的 Sheet 名

---

## 📜 License

[MIT License](LICENSE)

---

## 🪟 Windows EXE 下载

每次推送代码，GitHub Actions 会自动打包生成 `.exe` 文件：

1. 打开 [Actions](https://github.com/rockycenter/sales-forecast-helper/actions) 页面
2. 点击最新的 workflow run
3. 在 **Artifacts** 区域下载 `销售预测助手`
4. 解压后双击 `销售预测助手.exe` 即可运行（无需安装 Python）

> 发布正式版本时，打一个 `v` 开头的 tag（如 `v2.0`），GitHub 会自动创建 Release 并附带 exe。
