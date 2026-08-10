@echo off
chcp 65001 >nul
echo.
echo ========================================
echo      销售预测助手
echo ========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误：未检测到 Python！
    echo.
    echo 请先安装 Python：
    echo   1. 访问 https://www.python.org/downloads/
    echo   2. 下载并安装 Python 3.x
    echo   3. 安装时勾选 "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo ✅ Python 已安装
python --version
echo.

:: Check dependencies
echo 检查依赖库...
python -c "import pandas, numpy, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  缺少依赖库，正在安装...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ❌ 依赖安装失败，请手动运行：pip install -r requirements.txt
        pause
        exit /b 1
    )
    echo ✅ 依赖安装完成
) else (
    echo ✅ 依赖库已就绪
)

echo.

:: Run the program with interactive input
echo 请按提示输入信息...
echo.

set /p FILE_PATH="Excel文件路径: "
set /p SALESPERSON="销售员英文名: "

echo.
echo 正在运行预测...
echo.

python forecast.py -i "%FILE_PATH%" -s "%SALESPERSON%"

echo.
echo ========================================
echo 程序运行完毕！
echo ========================================
pause
