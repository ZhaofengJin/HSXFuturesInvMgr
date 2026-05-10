@echo off
chcp 65001 >nul
echo ========================================
echo    期货库存数据处理工具 - 快速启动
echo ========================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python安装！
    echo.
    echo 请先安装Python:
    echo   1. 访问 https://www.python.org/downloads/
    echo   2. 下载并安装 Python 3.8 或更高版本
    echo   3. 安装时勾选 "Add Python to PATH"
    echo.
    echo 安装完成后，请重新运行此脚本。
    pause
    exit /b 1
)

:: 检查openpyxl是否安装
python -c "import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖包 openpyxl...
    pip install openpyxl
    if %errorlevel% neq 0 (
        echo [错误] 安装openpyxl失败！
        echo 请手动运行: pip install openpyxl
        pause
        exit /b 1
    )
)

echo [OK] 环境检查通过
echo.
echo 正在启动数据处理工具...
echo.
python "%~dp0期货库存工具.py"

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    处理完成！请查看输出文件。
    echo ========================================
) else (
    echo.
    echo [错误] 处理过程中出现错误。
)

pause
