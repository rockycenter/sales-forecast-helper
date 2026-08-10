#!/bin/bash
# 销售预测助手 - macOS/Linux 启动脚本

echo ""
echo "========================================"
echo "      销售预测助手 v2.0"
echo "========================================"
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误：未检测到 Python 3！"
    echo "   请先安装 Python: brew install python3"
    exit 1
fi

echo "✅ Python: $(python3 --version)"
echo ""

# 检查依赖
echo "检查依赖库..."
python3 -c "import pandas, numpy, openpyxl" &> /dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  缺少依赖库，正在安装..."
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "❌ 依赖安装失败"
        exit 1
    fi
    echo "✅ 依赖安装完成"
else
    echo "✅ 依赖库已就绪"
fi

echo ""

# 运行程序
echo "📂 请拖入 Excel 文件或输入路径:"
read -r FILE_PATH

if [ -z "$FILE_PATH" ]; then
    echo "❌ 未提供文件路径，退出。"
    exit 1
fi

# 去掉拖入文件时产生的引号和转义空格
FILE_PATH=$(echo "$FILE_PATH" | sed 's/^["'"'"']*//;s/["'"'"']*$//' | sed 's/\\//g')

echo ""
echo "正在启动..."
echo ""

python3 forecast.py -i "$FILE_PATH"

echo ""
echo "========================================"
echo "程序运行完毕！"
echo "========================================"
read -p "按 Enter 键关闭..."
