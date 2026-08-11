@echo off
chcp 65001 >nul
echo.
echo ============================================
echo   销售预测助手 - 打包为 Windows EXE (GUI)
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
pyinstaller --onefile --noconsole --name "销售预测助手" --clean --noconfirm forecast.py

if errorlevel 1 (
    echo ❌ 打包失败
    pause
    exit /b 1
)

echo.
echo [3/3] 打包完成！
echo.
echo 📁 EXE 文件位置: dist\销售预测助手.exe
echo.
echo 直接双击运行，无需安装 Python，无命令行窗口！
echo.
echo ============================================
pause
