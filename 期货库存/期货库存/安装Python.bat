@echo off
chcp 65001 >nul
echo ========================================
echo    Python 自动安装脚本
echo ========================================
echo.

set INSTALLER_PATH="%~dp0python-installer.exe"

if not exist %INSTALLER_PATH% (
    echo [错误] 未找到Python安装包！
    echo 请确保 python-installer.exe 在同一目录下。
    pause
    exit /b 1
)

echo [提示] 即将安装Python 3.12
echo.
echo 安装参数:
echo   - 添加到系统PATH
echo   - 安装pip
echo   - 静默模式安装
echo.
echo 按任意键继续安装...
pause >nul

echo.
echo 正在安装Python，请稍候...
echo (安装过程可能需要几分钟)

:: 静默安装Python，添加PATH，安装pip
%INSTALLER_PATH% /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1

:: 等待安装完成
timeout /t 30 /nobreak >nul

:: 刷新环境变量
set PATH=%PATH%;C:\Users\77188\AppData\Local\Programs\Python\Python312
set PATH=%PATH%;C:\Users\77188\AppData\Local\Programs\Python\Python312\Scripts

:: 等待PATH更新
timeout /t 5 /nobreak >nul

echo.
echo [提示] 正在安装openpyxl依赖包...
python -m pip install openpyxl --quiet

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    Python 安装完成！
    echo ========================================
    echo.
    echo 现在可以运行: 启动工具.bat
    echo.
) else (
    echo.
    echo [提示] openpyxl安装可能需要手动执行:
    echo   pip install openpyxl
    echo.
)

echo 按任意键退出...
pause >nul
