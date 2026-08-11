@echo off
chcp 65001 >nul

:: Read version
set /p VERSION=<VERSION

echo.
echo ========================================
echo      销售预测助手 v%VERSION%
echo ========================================
echo.
echo 🖥️  正在启动图形界面...
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误：未检测到 Python！
    echo   请先安装 Python 3.x
    echo   https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Check deps
python -c "import pandas, numpy, openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  缺少依赖库，正在安装...
    pip install -r requirements.txt
)

:: Launch GUI (pass dragged file if any)
if not "%~1"=="" (
    pythonw forecast.py -i "%~1"
) else (
    pythonw forecast.py
)

endlocal
