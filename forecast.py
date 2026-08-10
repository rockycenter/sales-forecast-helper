"""销售预测助手 - 入口"""

import argparse
import os
import sys
from src.cli import run_interactive


def main():
    parser = argparse.ArgumentParser(description="销售预测助手 v2.0")
    parser.add_argument("--input", "-i", help="输入 Excel 文件路径（可选，不传则拖入）")
    args = parser.parse_args()

    file_path = args.input

    if not file_path:
        # 无参数：提示拖入文件
        print("=" * 60)
        print("         📊 销售预测助手 v2.0")
        print("=" * 60)
        print()
        print("使用方法:")
        print("   python forecast.py -i \"文件路径.xlsx\"")
        print()
        file_path = input("📂 请拖入 Excel 文件或输入路径 > ").strip().strip('"').strip("'")

    if not file_path:
        print("❌ 未提供文件路径，退出。")
        return 1

    if not os.path.exists(file_path):
        print(f"❌ 找不到文件: {file_path}")
        return 1

    return run_interactive(file_path)


if __name__ == "__main__":
    sys.exit(main())
