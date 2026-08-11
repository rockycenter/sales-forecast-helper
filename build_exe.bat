@echo off
chcp 65001 >nul

:: Read version
set /p VERSION=<VERSION

echo.
echo ============================================
echo   销售预测助手 - 打包 EXE v%VERSION%
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未检测到 Python，请先安装 Python 3.x
    pause
    exit /b 1
)

echo [1/3] 安装 PyInstaller...
pip install pyinstaller
if errorlevel 1 (
    echo ❌ 安装失败
    pause
    exit /b 1
)

echo.
echo [2/3] 打包中（可能需要 1-2 分钟）...
pyinstaller --onefile --noconsole --name "销售预测助手_v%VERSION%" --clean --noconfirm forecast.py

if errorlevel 1 (
    echo ❌ 打包失败
    pause
    exit /b 1
)

echo.
echo [3/3] 打包完成！
echo.
echo 📁 EXE 文件: dist\销售预测助手_v%VERSION%.exe
echo.
echo ============================================
pause
