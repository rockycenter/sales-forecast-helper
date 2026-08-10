"""交互式命令行界面"""

import os
import numpy as np
from .excel_io import load_workbook, get_salespeople, run_forecast, save_result
from .config import FORECAST_MONTHS


def print_banner():
    print()
    print("=" * 60)
    print("         📊 销售预测助手 v2.0")
    print("=" * 60)
    print()


def print_table(result_df, start=0, show=10):
    """格式化打印预测结果表格"""
    total = len(result_df)
    subset = result_df.iloc[start:start + show]

    # 表头
    header = (f"{'#':>4}  {'SPEC料号':<20} {'类型':<4} {'Open SO':>10}  "
              f"{'8月':>10}  {'9月':>10}  {'10月':>10}  {'11月':>10}")
    print(header)
    print("-" * len(header))

    for i, (_, row) in enumerate(subset.iterrows()):
        idx = start + i + 1
        print(
            f"{idx:>4}  {str(row['SPEC料号']):<20} "
            f"{row['产品类型']:<4} {row['Open_SO']:>10,}  "
            f"{row['推荐_8月']:>10,}  {row['推荐_9月']:>10,}  "
            f"{row['推荐_10月']:>10,}  {row['推荐_11月']:>10,}"
        )

    if total > show:
        print(f"\n  ... 共 {total} 条（当前显示第 {start + 1}-{min(start + show, total)} 条）")


def print_summary(result_df, warnings):
    """打印汇总统计"""
    type_counts = result_df['产品类型'].value_counts()
    total_forecast = sum(
        result_df[f'推荐_{m}'].sum() for m in FORECAST_MONTHS
    )

    print()
    print("📈 预测结果统计:")
    print(f"   A类(稳定): {type_counts.get('A', 0)} 个")
    print(f"   B类(波动): {type_counts.get('B', 0)} 个")
    print(f"   C类(稀疏): {type_counts.get('C', 0)} 个")
    print(f"   4个月预测总量: {total_forecast:,.0f}")

    if warnings:
        print(f"\n⚠️  警告 ({len(warnings)}条):")
        for w in warnings:
            print(f"   {w}")


def adjust_forecast(result_df):
    """交互式微调循环"""
    print()
    print("=" * 60)
    print("🔧 微调模式")
    print("   输入格式: <行号> <月份(8/9/10/11)> <新值>")
    print("   示例: 3 8 50000  (将第3行8月推荐改为50000)")
    print("   输入 q 或 0 退出微调")
    print("=" * 60)

    while True:
        try:
            cmd = input("\n✏️  微调 > ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if not cmd:
            continue
        if cmd.lower() in ('q', 'quit', 'exit', '0'):
            break

        parts = cmd.split()
        if len(parts) < 3:
            print("   ⚠️  格式错误，请用: <行号> <月份> <新值>")
            continue

        try:
            row_idx = int(parts[0]) - 1  # 转为 0-based
            month = parts[1]
            new_val = float(parts[2])

            if row_idx < 0 or row_idx >= len(result_df):
                print(f"   ⚠️  行号超出范围 (1-{len(result_df)})")
                continue

            col_name = f"推荐_{month}月"
            if col_name not in result_df.columns:
                print(f"   ⚠️  月份无效，请输入 8/9/10/11")
                continue

            old_val = result_df.at[row_idx, col_name]
            result_df.at[row_idx, col_name] = int(new_val)

            spec = result_df.at[row_idx, 'SPEC料号']
            print(f"   ✅ #{row_idx + 1} {spec} {month}月: "
                  f"{old_val:,} → {int(new_val):,}")

        except (ValueError, IndexError):
            print("   ⚠️  输入格式错误，请重新输入")

    return result_df


def run_interactive(file_path):
    """主交互流程"""
    print_banner()

    # 1. 加载文件
    print(f"📂 读取文件: {os.path.basename(file_path)}")
    try:
        df, sheet_name, _ = load_workbook(file_path)
    except (FileNotFoundError, ValueError) as e:
        print(f"\n❌ {e}")
        return 1

    print(f"📊 使用Sheet: {sheet_name}")
    print(f"   总行数: {len(df)}")

    # 2. 选择销售员
    all_names = get_salespeople(df)
    if not all_names:
        print("\n❌ 未找到任何销售员数据")
        return 1

    print(f"\n👥 共发现 {len(all_names)} 位销售员:")
    for i, name in enumerate(all_names, 1):
        count = len(df[df[3] == name])
        print(f"   {i:>3}. {name} ({count} 个产品)")

    # 选择
    while True:
        try:
            choice = input(f"\n🔍 请选择销售员 (1-{len(all_names)}) > ").strip()
            idx = int(choice) - 1
            if 0 <= idx < len(all_names):
                salesperson = all_names[idx]
                break
            print(f"   ⚠️  请输入 1-{len(all_names)} 之间的数字")
        except ValueError:
            print("   ⚠️  请输入数字")
        except (EOFError, KeyboardInterrupt):
            print("\n   👋 已取消")
            return 0

    # 3. 运行预测
    print(f"\n⏳ 正在为 {salesperson} 生成预测...")
    try:
        result_df, warnings = run_forecast(df, salesperson)
    except ValueError as e:
        print(f"\n❌ {e}")
        return 1

    # 4. 初版预览
    print()
    print("=" * 60)
    print(f"   📋 初版预测结果 — {salesperson} ({len(result_df)} 条)")
    print("=" * 60)
    print_table(result_df)
    print_summary(result_df, warnings)

    # 5. 微调
    print()
    while True:
        cmd = input("🔧 是否需要微调？(y/n) > ").strip().lower()
        if cmd in ('n', 'no', 'q', ''):
            break
        if cmd in ('y', 'yes'):
            result_df = adjust_forecast(result_df)
            print("\n📋 当前预测（已更新）:")
            print_table(result_df)
            break
        print("   请输入 y(是) 或 n(否)")

    # 6. 确认输出
    print()
    cmd = input("💾 确认输出最终 Excel？(y/n) > ").strip().lower()
    if cmd not in ('y', 'yes', ''):
        print("   👋 已取消输出")
        return 0

    try:
        output_path = save_result(result_df, salesperson)
        print(f"\n✅ 结果已保存: {output_path}")
    except Exception as e:
        fallback = f"预测结果_{salesperson.replace(' ', '_')}.xlsx"
        result_df.to_excel(fallback, index=False)
        print(f"\n⚠️  保存失败: {e}")
        print(f"   已保存到: {os.path.abspath(fallback)}")

    # 最终汇总
    print_summary(result_df, warnings if 'warnings' in dir() else [])
    print("\n" + "=" * 60)
    print("   ✅ 完成！")
    print("=" * 60)
    return 0
