"""Excel 读写操作"""

import os
import pandas as pd
import numpy as np
from datetime import datetime
from .config import (
    COL_SALESPERSON, COL_HISTORY_START,
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


def get_quarter_months(month):
    """返回指定月份所在季度的三个月，如 8 → [7,8,9]"""
    q_start = ((month - 1) // 3) * 3 + 1
    return [q_start, q_start + 1, q_start + 2]


def compute_quarter_comparison(history, forecasts, first_month, date_map=None):
    """
    计算季度同比（仅比较有完整两年数据的月份）
    history: N个月实际数据（列表），覆盖 [first_month-N, first_month-1] 月
    forecasts: 4个月预测值 [f1, f2, f3, f4]
    first_month: 第一个预测月的月份 (1-12)
    date_map: 保留参数，暂不使用

    返回: (本季合计, 去年同季合计, 同比%, 有效月数, 季度总月数)
    """
    import pandas as pd
    q_months = get_quarter_months(first_month)
    N = len(history)

    this_q = 0
    last_q = 0
    valid_months = 0

    for qm in q_months:
        # 去年同月在 history 中的索引
        # history[0] = first_month - N 月
        # qm 的去年 = qm - 12
        last_idx = (qm - 12) - (first_month - N)

        if last_idx < 0 or last_idx >= N:
            continue
        if not (pd.notna(history[last_idx]) and history[last_idx] > 0):
            continue

        valid_months += 1

        # 今年该月
        if qm >= first_month:
            this_val = forecasts[qm - first_month]
        else:
            this_idx = qm - (first_month - N)
            raw = history[this_idx] if 0 <= this_idx < N else 0
            this_val = raw if pd.notna(raw) and raw > 0 else 0

        last_val = history[last_idx]

        this_q += this_val
        last_q += last_val

    if last_q > 0:
        pct = round((this_q - last_q) / last_q * 100, 1)
    else:
        pct = None

    return this_q, last_q, pct, valid_months, len(q_months)



def _detect_history_meta(df):
    """从表头自动检测历史数据列数及年月映射。
    返回: (history_count, date_map)
    date_map 为 {列索引: (年, 月)}，如 {5: (2025, 7)}"""
    import re
    from .config import COL_HISTORY_START
    
    # 优先匹配 "25年7月" 格式
    full_pat = re.compile(r'(\d{2,4})\s*年\s*(\d{1,2})\s*月')
    short_pat = re.compile(r'(\d{1,2})\s*月')
    
    for header_row_idx in range(min(4, len(df))):
        row = df.iloc[header_row_idx]
        date_map = {}
        count = 0
        for c in range(COL_HISTORY_START, len(df.columns)):
            val = row[c]
            if pd.isna(val):
                break
            text = str(val).strip()
            
            m = full_pat.search(text)
            if m:
                y = int(m.group(1))
                if y < 100:
                    y += 2000
                mo = int(m.group(2))
                if 1 <= mo <= 12:
                    date_map[c] = (y, mo)
                    count += 1
                    continue
            
            m = short_pat.search(text)
            if m:
                mo = int(m.group(1))
                if 1 <= mo <= 12:
                    date_map[c] = (None, mo)  # 年份未知
                    count += 1
                    continue
            break
        if count >= 6:
            # 补全年份：按月份序列推断
            date_map = _infer_years(date_map)
            return count, date_map
    
    # 兜底：扫数值
    if len(df) < 4:
        return 12, {}
    row = df.iloc[3]
    count = 0
    for c in range(COL_HISTORY_START, len(df.columns)):
        val = row[c]
        if pd.notna(val) and isinstance(val, (int, float, np.integer, np.floating)) and val >= 0:
            count += 1
        else:
            break
    return max(count, 6), {}


def _infer_years(date_map):
    """补全 date_map 中缺失的年份（通过月份递增/递减推断）"""
    items = sorted(date_map.items())
    if not items:
        return date_map
    # 找第一个有年份的
    base_year = None
    base_col, base_month = None, None
    for col, (y, m) in items:
        if y is not None:
            base_year = y
            base_col = col
            base_month = m
            break
    if base_year is None:
        return date_map  # 全都没有年份
    
    result = {}
    for col, (y, m) in sorted(date_map.items()):
        if y is not None:
            result[col] = (y, m)
        else:
            # 根据月份变化推断年份
            if m >= base_month:
                result[col] = (base_year, m)
            else:
                result[col] = (base_year + 1, m)
        base_month = m
    return result




def load_workbook(file_path):
    """加载 Excel 并找到目标 sheet，自动检测历史数据月数及月份映射"""
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

    # 自动检测历史数据列数及月份映射
    history_count, date_map = _detect_history_meta(df)
    
    return df, target_sheet, sheet_names, history_count, date_map


def get_salespeople(df):
    """从 DataFrame 提取销售员列表"""
    names = df.iloc[3:, COL_SALESPERSON].dropna().unique()
    return sorted(names)


def run_forecast(df, salesperson, forecast_months=None, history_count=12, date_map=None):
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
        history = [row[c] for c in range(COL_HISTORY_START, COL_HISTORY_START + history_count)]
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

        # 季度同比计算
        first_month_num = int(forecast_months[0].replace('月', ''))
        q_this, q_last, q_pct, q_valid, q_total = compute_quarter_comparison(history, forecasts, first_month_num, date_map)

        
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
            '本季合计': int(q_this),
            '去年同季': int(q_last),
            '同比%': q_pct if q_pct is not None else '',
            '有效对比月': q_valid,
            '季度月数': q_total,
            '_历史': history,       # 原始12月数据，供GUI实时重算同比
            '_预测': forecasts,     # 原始4月预测，供GUI实时重算同比
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
