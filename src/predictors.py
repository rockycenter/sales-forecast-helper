"""A/B/C 三类预测算法"""

import numpy as np
import pandas as pd
from .config import (
    A_WEIGHT_MA3, A_WEIGHT_LY, A_WEIGHT_YEARLY, A_WEIGHT_TREND,
    A_OPEN_SO_MULTIPLIER,
    B_WEIGHT_MEDIAN, B_WEIGHT_RECENT, B_WEIGHT_LY_CAP,
    B_PERCENTILE_CAP, B_PERCENTILE_FLOOR, B_CAP_MULTIPLIER, B_FLOOR_MULTIPLIER,
    C_TINY_THRESHOLD, C_RECENT_MULTIPLIER,
    ROUND_100K, ROUND_10K, ROUND_1K,
)


def smart_round(value):
    """智能取整"""
    value = max(value, 0)
    if value >= 100000:
        return round(value / ROUND_100K) * ROUND_100K
    elif value >= 10000:
        return round(value / ROUND_10K) * ROUND_10K
    elif value >= 1000:
        return round(value / ROUND_1K) * ROUND_1K
    return round(value)


def forecast_A_stable(history, same_month_ly, month_idx, open_so=None):
    """A类-稳定产品：加权融合近期趋势 + 去年同期 + 季节性"""
    valid = [x for x in history if pd.notna(x) and x > 0]
    recent_3m = [x for x in history[-3:] if pd.notna(x)]
    ma3 = np.mean(recent_3m) if recent_3m else np.mean(valid)
    ly_val = same_month_ly if pd.notna(same_month_ly) else ma3
    yearly_avg = np.mean(valid)

    # 线性趋势
    recent_6m = [x for x in history[-6:] if pd.notna(x)]
    slope = 0
    if len(recent_6m) >= 3:
        x = list(range(len(recent_6m)))
        y = recent_6m
        n = len(x)
        try:
            slope = (n * sum(x[i] * y[i] for i in range(n)) - sum(x) * sum(y)) / \
                    (n * sum(xi * xi for xi in x) - sum(x) ** 2)
        except ZeroDivisionError:
            pass

    forecast = (ma3 * A_WEIGHT_MA3 + ly_val * A_WEIGHT_LY
                + yearly_avg * A_WEIGHT_YEARLY
                + slope * (month_idx + 1) * A_WEIGHT_TREND)

    if month_idx == 0 and pd.notna(open_so) and open_so > 0:
        forecast = max(forecast, open_so * A_OPEN_SO_MULTIPLIER)

    return max(forecast, 0)


def forecast_B_volatile(history, same_month_ly, month_idx, open_so=None):
    """B类-波动产品：中位数法 + 峰值控制"""
    valid = [x for x in history if pd.notna(x) and x > 0]
    median_val = np.median(valid)
    q75 = np.percentile(valid, 75)
    recent_3m = [x for x in history[-3:] if pd.notna(x)]
    recent_avg = np.mean(recent_3m) if recent_3m else median_val
    ly_val = same_month_ly if pd.notna(same_month_ly) else median_val

    forecast = (median_val * B_WEIGHT_MEDIAN
                + recent_avg * B_WEIGHT_RECENT
                + min(ly_val, q75) * B_WEIGHT_LY_CAP)

    p80 = np.percentile(valid, B_PERCENTILE_CAP)
    p20 = np.percentile(valid, B_PERCENTILE_FLOOR)
    forecast = min(forecast, p80 * B_CAP_MULTIPLIER)
    forecast = max(forecast, p20 * B_FLOOR_MULTIPLIER)

    if month_idx == 0 and pd.notna(open_so) and open_so > 0:
        forecast = max(forecast, open_so)

    return max(forecast, 0)


def forecast_C_sparse(history, same_month_ly, month_idx, open_so=None):
    """C类-稀疏产品：保守估计，避免虚高"""
    valid = [x for x in history if pd.notna(x) and x > 0]
    if not valid:
        return open_so if (month_idx == 0 and pd.notna(open_so)) else 0

    avg = np.mean(valid)
    if avg < C_TINY_THRESHOLD:
        return open_so if (month_idx == 0 and pd.notna(open_so) and open_so > 0) else 0

    if month_idx == 0 and pd.notna(open_so) and open_so > 0:
        return open_so

    recent = [x for x in history[-6:] if pd.notna(x) and x > 0]
    if recent:
        return np.mean(recent[-3:]) * C_RECENT_MULTIPLIER
    return 0
