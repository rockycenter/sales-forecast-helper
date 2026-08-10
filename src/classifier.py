"""ABC 产品分类模块"""

import numpy as np
import pandas as pd
from .config import (
    CLASS_A_MIN_MONTHS, CLASS_A_MAX_CV, CLASS_A_MIN_AVG,
    CLASS_B_MIN_MONTHS, CLASS_C_MAX_AVG, CLASS_C_MAX_MONTHS,
    HIGH_VOLATILITY_CV, HIGH_VOLATILITY_MIN_AVG,
    SEASONAL_RATIO_UPPER, SEASONAL_RATIO_LOWER,
)


def classify_product(history_list):
    """
    产品三分类：
    A - 稳定产品：数据完整(≥8月)，波动小(CV<0.8)，平均≥5000
    B - 波动/季节性：数据≥6月，但高波动(CV≥0.8)或季节性明显
    C - 稀疏/微量：数据<6月 或 平均<3000

    Returns: (类型, 原因)
    """
    valid = [float(x) for x in history_list if pd.notna(x) and x > 0]
    avg = np.mean(valid) if valid else 0

    if len(valid) < CLASS_C_MAX_MONTHS or avg < CLASS_C_MAX_AVG:
        return 'C', "数据稀疏/微量"

    cv = np.std(valid) / avg if avg > 0 else float('inf')

    # 季节性检查
    recent_6m = [x for x in history_list[-6:] if pd.notna(x) and x > 0]
    recent_avg = np.mean(recent_6m) if recent_6m else 0
    same_period = [x for x in history_list[:4] if pd.notna(x) and x > 0]
    same_period_avg = np.mean(same_period) if same_period else 0
    seasonal_ratio = same_period_avg / recent_avg if recent_avg > 0 else 0
    is_seasonal = (seasonal_ratio > SEASONAL_RATIO_UPPER
                   or seasonal_ratio < SEASONAL_RATIO_LOWER) if recent_avg > 0 else False

    if is_seasonal:
        return 'B', "季节性明显"
    if cv >= HIGH_VOLATILITY_CV and avg > HIGH_VOLATILITY_MIN_AVG:
        return 'B', "高波动产品"
    if cv < CLASS_A_MAX_CV and len(valid) >= CLASS_A_MIN_MONTHS and avg >= CLASS_A_MIN_AVG:
        return 'A', "稳定产品"
    if len(valid) >= CLASS_B_MIN_MONTHS:
        return 'B', "中等波动"
    return 'C', "数据不足"
