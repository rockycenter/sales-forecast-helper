"""销售预测助手 v3.0 - 入口"""

import argparse
import os
import sys


def main():
    parser = argparse.ArgumentParser(description="销售预测助手 v3.0")
    parser.add_argument("--input", "-i", help="输入 Excel 文件路径（可选）")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式（默认为图形界面）")
    args = parser.parse_args()

    if args.cli:
        from src.cli import run_interactive
        file_path = args.input
        if not file_path:
            file_path = input("📂 请拖入 Excel 文件或输入路径 > ").strip().strip('"').strip("'")
        if not file_path:
            print("❌ 未提供文件路径，退出。")
            return 1
        if not os.path.exists(file_path):
            print(f"❌ 找不到文件: {file_path}")
            return 1
        return run_interactive(file_path)
    else:
        from src.gui import launch_gui
        launch_gui(args.input)
        return 0


if __name__ == "__main__":
    sys.exit(main())
