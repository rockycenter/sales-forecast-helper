"""Excel 读写操作"""

import os
import pandas as pd
import numpy as np
from datetime import datetime
from .config import (
    COL_SALESPERSON, COL_HISTORY_START, COL_HISTORY_END,
    COL_OPEN_SO, COL_REGION, COL_SPEC, COL_LEGACY, COL_UNIT,
    SHEET_KEYWORD, OUTPUT_COLUMNS, FORECAST_MONTHS,
)
from .classifier import classify_product
from .predictors import (
    forecast_A_stable, forecast_B_volatile, forecast_C_sparse, smart_round,
)



def parse_forecast_months(sheet_name):
    """从Sheet名解析预测月份，如 '销售预测收集26年8-11月' → ['8月','9月','10月','11月']"""
    import re
    # 匹配 "8-11月" 或 "9-12" 这类起止月份
    patterns = [
        r'(\d{1,2})-(\d{1,2})\s*月',   # "8-11月"
        r'(\d{1,2})-(\d{1,2})',         # "8-11"
    ]
    for pat in patterns:
        match = re.search(pat, sheet_name)
        if match:
            start = int(match.group(1))
            end = int(match.group(2))
            if start <= end:
                months = list(range(start, end + 1))
            else:
                months = list(range(start, 13)) + list(range(1, end + 1))
            return [f"{m}月" for m in months]
    # Fallback: default 4 months from current
    from datetime import datetime
    current = datetime.now().month
    return [f"{(current + i - 1) % 12 + 1}月" for i in range(4)]


def load_workbook(file_path):
    """加载 Excel 并找到目标 sheet"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"找不到文件: {file_path}")

    excel = pd.ExcelFile(file_path)
    sheet_names = excel.sheet_names

    target_sheet = None
    for name in sheet_names:
        if SHEET_KEYWORD in name:
            target_sheet = name
            break

    if not target_sheet:
        raise ValueError(
            f"找不到包含'{SHEET_KEYWORD}'的Sheet\n"
            f"可用Sheet: {', '.join(sheet_names)}"
        )

    df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
    return df, target_sheet, sheet_names


def get_salespeople(df):
    """从 DataFrame 提取销售员列表"""
    names = df.iloc[3:, COL_SALESPERSON].dropna().unique()
    return sorted(names)


def run_forecast(df, salesperson, forecast_months=None):
    """对指定销售员运行完整预测"""
    if forecast_months is None:
        forecast_months = ["8月", "9月", "10月", "11月"]
    user_data = df[df[COL_SALESPERSON] == salesperson].copy()

    if len(user_data) == 0:
        all_names = get_salespeople(df)
        raise ValueError(
            f"找不到销售员 '{salesperson}'\n"
            f"可用销售员: {', '.join(all_names[:20])}"
        )

    results = []
    warnings = []

    for idx, row in user_data.iterrows():
        spec = row[COL_SPEC]
        legacy = row[COL_LEGACY]
        history = [row[c] for c in range(COL_HISTORY_START, COL_HISTORY_END)]
        open_so = row[COL_OPEN_SO] if pd.notna(row[COL_OPEN_SO]) else 0

        ptype, reason = classify_product(history)

        forecasts = []
        for month_idx in range(4):
            same_month_ly = history[month_idx] if month_idx < len(history) else None

            if ptype == 'A':
                f = forecast_A_stable(history, same_month_ly, month_idx,
                                      open_so if month_idx == 0 else None)
            elif ptype == 'B':
                f = forecast_B_volatile(history, same_month_ly, month_idx,
                                        open_so if month_idx == 0 else None)
            else:
                f = forecast_C_sparse(history, same_month_ly, month_idx,
                                      open_so if month_idx == 0 else None)

            if month_idx == 0 and pd.notna(open_so) and open_so > 0:
                f = max(f, open_so)

            forecasts.append(smart_round(f))

        # Open SO 强制修正
        if pd.notna(open_so) and open_so > 0 and forecasts[0] < open_so:
            warnings.append(
                f"⚠️ {spec}({legacy}): 8月推荐{forecasts[0]:,} < Open SO {open_so:,}，已强制修正"
            )
            forecasts[0] = smart_round(open_so)

        result_row = {
            '行号': idx + 1,
            '区域': row[COL_REGION],
            'SPEC料号': spec,
            'Legacy Item': legacy,
            '销售员': row[COL_SALESPERSON],
            '单位': row[COL_UNIT],
            '产品类型': ptype,
            '分类原因': reason,
            'Open_SO': int(open_so) if pd.notna(open_so) else 0,
            '历史平均': round(np.mean([x for x in history if pd.notna(x) and x > 0]))
            if [x for x in history if pd.notna(x) and x > 0] else 0,
        }
        for i, m in enumerate(forecast_months):
            result_row[f'推荐_{m}'] = forecasts[i]
        results.append(result_row)

    result_df = pd.DataFrame(results)
    return result_df, warnings


def save_result(result_df, salesperson, output_path=None):
    """保存结果到 Excel"""
    if output_path:
        path = output_path
    else:
        timestamp = datetime.now().strftime("%m%d_%H%M")
        path = f"预测结果_{salesperson.replace(' ', '_')}_{timestamp}.xlsx"

    result_df.to_excel(path, index=False)
    return os.path.abspath(path)
