@echo off
chcp 65001 >nul
title 期货库存数据处理工具 v5.0
echo ============================================
echo  期货库存数据处理工具 v5.0
echo ============================================
echo.

REM 查找 Python
set "PYTHON_CMD="

REM 1. 优先查找 .workbuddy Python
if exist "C:\Users\77188\.workbuddy\binaries\python\versions\3.13.12\python.exe" (
    set "PYTHON_CMD=C:\Users\77188\.workbuddy\binaries\python\versions\3.13.12\python.exe"
    goto :found_python
)

REM 2. 查找 PATH 中的 python
for %%P in (python3.exe python.exe) do (
    for %%I in (%%P) do (
        set "PYTHON_CMD=%%~$PATH:I"
        if not "!PYTHON_CMD!"=="" goto :found_python
    )
)

REM 3. 未找到
if "!PYTHON_CMD!"=="" (
    echo [错误] 未找到 Python，请先安装 Python 3.12+
    echo.
    pause
    exit /b 1
)

:found_python
echo 使用 Python: %PYTHON_CMD%
echo.

REM 运行主程序
"%PYTHON_CMD%" "%~dp0cli.py" %*
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if %EXIT_CODE%==0 (
    echo [完成] 处理成功！
) else if %EXIT_CODE%==2 (
    echo [完成] 处理完成，部分订单未匹配（警告）
) else (
    echo [错误] 处理失败，请检查上方错误信息
)
echo.
pause
