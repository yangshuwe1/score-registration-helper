@echo off
chcp 65001 > nul
echo ========================================
echo 启动登分助手
echo ========================================
echo.

REM 检查虚拟环境是否存在
if not exist "venv\Scripts\activate.bat" (
    echo [错误] 虚拟环境不存在
    echo 请先运行 setup_python312.bat 设置环境
    echo.
    pause
    exit /b 1
)

REM 激活虚拟环境
call venv\Scripts\activate.bat

REM 显示Python版本
echo 使用Python版本:
python --version
echo.

REM 运行程序
echo 正在启动程序...
python gui.py

REM 如果程序异常退出，暂停以查看错误信息
if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出
    pause
)
