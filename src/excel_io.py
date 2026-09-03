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


def parse_forecast_year(sheet_name):
    """从Sheet名解析预测年份，如 '26年9-12月' → 2026"""
    import re
    m = re.search(r'(\d{4})\s*年', sheet_name)
    if m:
        return int(m.group(1))
    m = re.search(r'(\d{2})\s*年', sheet_name)
    if m:
        y = int(m.group(1))
        return 2000 + y if y < 100 else y
    from datetime import datetime
    return datetime.now().year


def get_quarter_months(month):
    """返回指定月份所在季度的三个月，如 8 → [7,8,9]"""
    q_start = ((month - 1) // 3) * 3 + 1
    return [q_start, q_start + 1, q_start + 2]


def compute_all_quarters(history, forecasts, first_month, forecast_year):
    """
    计算预测月覆盖的所有季度对比（每次4个月必跨2个季度）
    history: 历史实际数据列表，[first_month-N, first_month-1] 月
    forecasts: 预测值列表（与 forecast_months 对应）
    first_month: 第一个预测月 (1-12)
    forecast_year: 第一个预测月所在年份
    返回: (quarter_results, quarter_labels)
        quarter_results: [{'label':'Q3','this':..,'last':..,'pct':..,'valid':..}, ...]
        quarter_labels: ['Q3','Q4'] 等按时间顺序
    """
    import pandas as pd
    N = len(history)

    # 实际数据映射 {(年, 月): 值}
    actual = {}
    for i in range(N):
        cum = first_month - N + i  # 相对 forecast_year 的累计月（1月起算）
        y = forecast_year + (cum - 1) // 12
        m = ((cum - 1) % 12) + 1
        val = history[i]
        if pd.notna(val) and val > 0:
            actual[(y, m)] = val

    # 预测映射
    forecast_map = {}
    for i, f in enumerate(forecasts):
        cum = first_month + i
        y = forecast_year + (cum - 1) // 12
        m = ((cum - 1) % 12) + 1
        forecast_map[(y, m)] = f

    # 覆盖的季度（预测月涉及的所有 (年, 季度)）
    quarters = sorted(set((y, (m - 1) // 3 + 1) for (y, m) in forecast_map.keys()))
    results = []
    labels = []
    for (y, q) in quarters:
        q_months = [3 * q - 2, 3 * q - 1, 3 * q]
        this_q = 0
        last_q = 0
        valid = 0
        for m in q_months:
            last_val = actual.get((y - 1, m))
            if last_val is None:
                continue
            this_val = forecast_map.get((y, m), actual.get((y, m)))
            if this_val is None:
                continue
            this_q += this_val
            last_q += last_val
            valid += 1
        pct = round((this_q - last_q) / last_q * 100, 1) if last_q > 0 else None
        label = f"{y}Q{q}"
        results.append({'label': label, 'year': y, 'this': int(this_q),
                        'last': int(last_q), 'pct': pct, 'valid': valid})
        labels.append(label)
    return results, labels



def _detect_history_count(df):
    """扫描数据行，从 COL_HISTORY_START 开始数连续数值列"""
    from .config import COL_HISTORY_START
    for row_idx in range(3, min(len(df), 20)):
        row = df.iloc[row_idx]
        count = 0
        for c in range(COL_HISTORY_START, len(df.columns)):
            val = row[c]
            if pd.notna(val) and isinstance(val, (int, float, np.integer, np.floating)) and val >= 0:
                count += 1
            else:
                break
        if count >= 6:
            return count
    return 12



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
    history_count = _detect_history_count(df)
    
    return df, target_sheet, sheet_names, history_count


def get_salespeople(df):
    """从 DataFrame 提取销售员列表"""
    names = df.iloc[3:, COL_SALESPERSON].dropna().unique()
    return sorted(names)


def run_forecast(df, salesperson, forecast_months=None, history_count=12, forecast_year=None):
    """对指定销售员运行完整预测"""
    if forecast_months is None:
        forecast_months = ["8月", "9月", "10月", "11月"]
    if forecast_year is None:
        forecast_year = datetime.now().year
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

        # 季度同比计算（覆盖所有预测季度）
        first_month_num = int(forecast_months[0].replace('月', ''))
        quarter_results, _ = compute_all_quarters(
            history, forecasts, first_month_num, forecast_year
        )

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
            '_历史': history,       # 原始历史数据，供GUI实时重算
            '_预测': forecasts,     # 原始预测值，供GUI实时重算
        }
        for i, m in enumerate(forecast_months):
            result_row[f'推荐_{m}'] = forecasts[i]
        for qr in quarter_results:
            q = qr['label']
            result_row[f'{q}_今年'] = qr['this']
            result_row[f'{q}_去年'] = qr['last']
            result_row[f'{q}_同比'] = qr['pct'] if qr['pct'] is not None else ''
            result_row[f'{q}_有效月'] = qr['valid']
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

    export_df = result_df[[c for c in result_df.columns if not c.startswith('_')]]
    export_df.to_excel(path, index=False)
    return os.path.abspath(path)
