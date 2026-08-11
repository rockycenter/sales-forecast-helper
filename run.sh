#!/bin/bash
# 销售预测助手 - macOS/Linux GUI 启动脚本

echo ""
echo "========================================"
echo "      销售预测助手 v3.0"
echo "========================================"
echo ""
echo "🖥️  正在启动图形界面..."
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误：未检测到 Python 3！"
    echo "   请先安装 Python: brew install python3"
    exit 1
fi

# 检查依赖
python3 -c "import pandas, numpy, openpyxl, tkinter" &> /dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  缺少依赖库，正在安装..."
    pip3 install -r requirements.txt
fi

# 启动 GUI
if [ -n "$1" ]; then
    python3 forecast.py -i "$1" &
else
    python3 forecast.py &
fi

echo ""

