import pandas as pd
import numpy as np
import os
import argparse
from datetime import datetime

"""
销售预测助手 - Sales Forecast Helper
=====================================
功能：自动分析历史销售数据，按ABC分类给出 W-Z 列推荐值
版本：1.0.0
作者：AI Assistant
"""


def classify_product(history_list):
    """
    产品三分类：
    A - 稳定产品：数据完整(≥8月)，波动小(CV<0.8)，平均≥5000
    B - 波动/季节性：数据≥6月，但高波动(CV≥0.8)或季节性明显
    C - 稀疏/微量：数据<6月 或 平均<3000
    """
    valid = [float(x) for x in history_list if pd.notna(x) and x > 0]
    avg = np.mean(valid) if valid else 0

    if len(valid) < 6 or avg < 3000:
        return 'C', "数据稀疏/微量"

    cv = np.std(valid) / avg if avg > 0 else float('inf')
    recent_6m = [x for x in history_list[-6:] if pd.notna(x) and x > 0]
    recent_avg = np.mean(recent_6m) if recent_6m else 0
    same_period = [x for x in history_list[:4] if pd.notna(x) and x > 0]
    same_period_avg = np.mean(same_period) if same_period else 0
    seasonal_ratio = same_period_avg / recent_avg if recent_avg > 0 else 0
    is_seasonal = seasonal_ratio > 2.0 or seasonal_ratio < 0.5 if recent_avg > 0 else False

    if is_seasonal:
        return 'B', "季节性明显"
    if cv >= 1.0 and avg > 10000:
        return 'B', "高波动产品"
    if cv < 0.8 and len(valid) >= 8 and avg >= 5000:
        return 'A', "稳定产品"
    if len(valid) >= 6:
        return 'B', "中等波动"
    return 'C', "数据不足"


def forecast_A_stable(history, same_month_ly, month_idx, open_so=None):
    """A类-稳定产品：加权融合近期趋势 + 去年同期 + 季节性"""
    valid = [x for x in history if pd.notna(x) and x > 0]
    recent_3m = [x for x in history[-3:] if pd.notna(x)]
    ma3 = np.mean(recent_3m) if recent_3m else np.mean(valid)
    ly_val = same_month_ly if pd.notna(same_month_ly) else ma3
    yearly_avg = np.mean(valid)

    recent_6m = [x for x in history[-6:] if pd.notna(x)]
    slope = 0
    if len(recent_6m) >= 3:
        x = list(range(len(recent_6m)))
        y = recent_6m
        n = len(x)
        try:
            slope = (n * sum(x[i]*y[i] for i in range(n)) - sum(x)*sum(y)) / \
                    (n*sum(xi*xi for xi in x) - sum(x)**2)
        except ZeroDivisionError:
            pass

    forecast = ma3 * 0.40 + ly_val * 0.30 + yearly_avg * 0.20 + slope * (month_idx + 1) * 0.10

    if month_idx == 0 and pd.notna(open_so) and open_so > 0:
        forecast = max(forecast, open_so * 1.05)

    return max(forecast, 0)


def forecast_B_volatile(history, same_month_ly, month_idx, open_so=None):
    """B类-波动产品：中位数法 + 峰值控制"""
    valid = [x for x in history if pd.notna(x) and x > 0]
    median_val = np.median(valid)
    q75 = np.percentile(valid, 75)
    recent_3m = [x for x in history[-3:] if pd.notna(x)]
    recent_avg = np.mean(recent_3m) if recent_3m else median_val
    ly_val = same_month_ly if pd.notna(same_month_ly) else median_val

    forecast = median_val * 0.35 + recent_avg * 0.35 + min(ly_val, q75) * 0.30
    p80 = np.percentile(valid, 80)
    p20 = np.percentile(valid, 20)
    forecast = min(forecast, p80 * 1.05)
    forecast = max(forecast, p20 * 0.8)

    if month_idx == 0 and pd.notna(open_so) and open_so > 0:
        forecast = max(forecast, open_so)

    return max(forecast, 0)


def forecast_C_sparse(history, same_month_ly, month_idx, open_so=None):
    """C类-稀疏产品：保守估计，避免虚高"""
    valid = [x for x in history if pd.notna(x) and x > 0]
    if not valid:
        return open_so if (month_idx == 0 and pd.notna(open_so)) else 0

    avg = np.mean(valid)
    if avg < 1000:
        return open_so if (month_idx == 0 and pd.notna(open_so) and open_so > 0) else 0

    if month_idx == 0 and pd.notna(open_so) and open_so > 0:
        return open_so

    recent = [x for x in history[-6:] if pd.notna(x) and x > 0]
    if recent:
        return np.mean(recent[-3:]) * 0.5
    return 0


def smart_round(value):
    """智能取整"""
    if value >= 100000:
        return round(value / 10000) * 10000
    elif value >= 10000:
        return round(value / 1000) * 1000
    elif value >= 1000:
        return round(value / 1000) * 1000
    return round(value)


def main():
    parser = argparse.ArgumentParser(description="销售预测助手")
    parser.add_argument("--input", "-i", required=True, help="输入Excel文件路径")
    parser.add_argument("--salesperson", "-s", required=True, help="销售员英文名")
    parser.add_argument("--output", "-o", default="", help="输出文件路径（可选）")
    args = parser.parse_args()

    file_path = args.input
    salesperson = args.salesperson

    if not os.path.exists(file_path):
        print(f"\n❌ 错误：找不到文件！\n   路径: {file_path}")
        return 1

    print("=" * 70)
    print("           销售预测助手 v1.0")
    print("=" * 70)
    print(f"\n📂 读取文件: {os.path.basename(file_path)}")

    # 读取Excel
    excel = pd.ExcelFile(file_path)
    sheet_names = excel.sheet_names
    target_sheet = None
    for name in sheet_names:
        if "销售预测收集" in name:
            target_sheet = name
            break

    if not target_sheet:
        print("\n❌ 错误：找不到包含'销售预测收集'的Sheet")
        print(f"   可用Sheet: {', '.join(sheet_names)}")
        return 1

    print(f"📊 使用Sheet: {target_sheet}")

    # 读取数据（无header，手动解析）
    df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
    print(f"   总行数: {len(df)}")

    # 过滤出目标销售员数据
    salesperson_col = 3
    user_data = df[df[salesperson_col] == salesperson].copy()

    if len(user_data) == 0:
        print(f"\n❌ 错误：找不到销售员 '{salesperson}'")
        all_names = df.iloc[3:, salesperson_col].dropna().unique()
        print(f"   可用销售员（前20个）: {', '.join(sorted(all_names)[:20])}")
        return 1

    print(f"\n👤 销售员: {salesperson}")
    print(f"   产品数: {len(user_data)}")

    # 处理每一行
    results = []
    warnings = []

    for idx, row in user_data.iterrows():
        spec = row[1]
        legacy = row[2]
        history = [row[c] for c in range(5, 17)]  # E-P列：12个月历史
        open_so = row[21] if pd.notna(row[21]) else 0  # V列

        ptype, reason = classify_product(history)

        # 计算4个月推荐值
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

        # Open SO 检查
        if pd.notna(open_so) and open_so > 0 and forecasts[0] < open_so:
            warnings.append(
                f"⚠️ {spec}({legacy}): 8月推荐{forecasts[0]:,} < Open SO {open_so:,}，已强制修正")
            forecasts[0] = smart_round(open_so)

        # 收集结果
        results.append({
            '行号': idx + 1,
            '区域': row[0],
            'SPEC料号': spec,
            'Legacy Item': legacy,
            '销售员': row[3],
            '单位': row[4],
            '产品类型': ptype,
            '分类原因': reason,
            'Open_SO': int(open_so) if pd.notna(open_so) else 0,
            '历史平均': round(np.mean([x for x in history if pd.notna(x) and x > 0]))
            if [x for x in history if pd.notna(x) and x > 0] else 0,
            '推荐_8月': forecasts[0],
            '推荐_9月': forecasts[1],
            '推荐_10月': forecasts[2],
            '推荐_11月': forecasts[3],
        })

    # 创建结果DataFrame
    result_df = pd.DataFrame(results)

    # 统计
    type_counts = result_df['产品类型'].value_counts()
    total_forecast = (result_df['推荐_8月'].sum() + result_df['推荐_9月'].sum() +
                      result_df['推荐_10月'].sum() + result_df['推荐_11月'].sum())

    print(f"\n📈 预测结果统计:")
    print(f"   A类(稳定): {type_counts.get('A', 0)} 个")
    print(f"   B类(波动): {type_counts.get('B', 0)} 个")
    print(f"   C类(稀疏): {type_counts.get('C', 0)} 个")
    print(f"   4个月预测总量: {total_forecast:,.0f} M2")

    if warnings:
        print(f"\n⚠️  警告 ({len(warnings)}条):")
        for w in warnings:
            print(f"   {w}")

    # 保存结果
    if args.output:
        output_path = args.output
    else:
        output_dir = os.path.dirname(file_path) if os.path.dirname(file_path) else "."
        timestamp = datetime.now().strftime("%m%d_%H%M")
        output_path = os.path.join(
            output_dir,
            f"预测结果_{salesperson.replace(' ', '_')}_{timestamp}.xlsx"
        )

    try:
        result_df.to_excel(output_path, index=False)
        print(f"\n✅ 结果已保存: {output_path}")
    except Exception as e:
        fallback = f"预测结果_{salesperson.replace(' ', '_')}.xlsx"
        result_df.to_excel(fallback, index=False)
        print(f"\n⚠️  保存到指定路径失败: {e}")
        print(f"   已保存到: {os.path.abspath(fallback)}")

    print("\n" + "=" * 70)
    print("完成！请打开输出的Excel文件查看推荐值。")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    exit(main())
