"""配置中心：列映射、分类阈值、算法权重"""

# ---- Excel 列映射 (0-based) ----
COL_SALESPERSON = 3        # 销售员
COL_HISTORY_START = 5      # 12个月历史起始列
COL_HISTORY_END = 17       # 12个月历史结束列（不含）
COL_OPEN_SO = 21           # Open SO列
COL_REGION = 0             # 区域
COL_SPEC = 1               # SPEC料号
COL_LEGACY = 2             # Legacy Item
COL_UNIT = 4               # 单位

# Sheet 匹配关键字
SHEET_KEYWORD = "销售预测收集"

# ---- ABC 分类阈值 ----
CLASS_A_MIN_MONTHS = 8
CLASS_A_MAX_CV = 0.8
CLASS_A_MIN_AVG = 5000
CLASS_B_MIN_MONTHS = 6
CLASS_C_MAX_AVG = 3000
CLASS_C_MAX_MONTHS = 6
HIGH_VOLATILITY_CV = 1.0
HIGH_VOLATILITY_MIN_AVG = 10000
SEASONAL_RATIO_UPPER = 2.0
SEASONAL_RATIO_LOWER = 0.5

# ---- A类预测权重 ----
A_WEIGHT_MA3 = 0.40
A_WEIGHT_LY = 0.30
A_WEIGHT_YEARLY = 0.20
A_WEIGHT_TREND = 0.10
A_OPEN_SO_MULTIPLIER = 1.05

# ---- B类预测权重 ----
B_WEIGHT_MEDIAN = 0.35
B_WEIGHT_RECENT = 0.35
B_WEIGHT_LY_CAP = 0.30
B_PERCENTILE_CAP = 80
B_PERCENTILE_FLOOR = 20
B_CAP_MULTIPLIER = 1.05
B_FLOOR_MULTIPLIER = 0.8

# ---- C类预测阈值 ----
C_TINY_THRESHOLD = 1000
C_RECENT_MULTIPLIER = 0.5

# ---- 智能取整 ----
ROUND_100K = 10000   # >=100000 按万取整
ROUND_10K = 1000     # >=10000 按千取整
ROUND_1K = 1000      # >=1000 按千取整

# ---- 输出列定义 ----
OUTPUT_COLUMNS = [
    "行号", "区域", "SPEC料号", "Legacy Item", "销售员", "单位",
    "产品类型", "分类原因", "Open_SO", "历史平均",
    "推荐_8月", "推荐_9月", "推荐_10月", "推荐_11月"
]

FORECAST_MONTHS = ["8月", "9月", "10月", "11月"]
