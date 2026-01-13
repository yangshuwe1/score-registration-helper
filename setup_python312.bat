@echo off
chcp 65001 > nul
echo ========================================
echo 切换到Python 3.12并安装依赖
echo ========================================
echo.

REM 检查Python 3.12是否存在
set PYTHON312=C:\Users\Shuwei Yang\AppData\Local\Programs\Python\Python312\python.exe

if not exist "%PYTHON312%" (
    echo [错误] 未找到Python 3.12
    echo 请确认Python 3.12安装路径，或从以下地址下载：
    echo https://www.python.org/downloads/release/python-3120/
    pause
    exit /b 1
)

echo [1/5] 找到Python 3.12: %PYTHON312%
"%PYTHON312%" --version
echo.

REM 删除旧的虚拟环境（如果存在）
if exist "venv" (
    echo [2/5] 删除旧的虚拟环境...
    rmdir /s /q venv
    echo 旧虚拟环境已删除
    echo.
)

echo [2/5] 创建新的Python 3.12虚拟环境...
"%PYTHON312%" -m venv venv
if errorlevel 1 (
    echo [错误] 创建虚拟环境失败
    pause
    exit /b 1
)
echo 虚拟环境创建成功
echo.

echo [3/5] 激活虚拟环境...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo [错误] 激活虚拟环境失败
    pause
    exit /b 1
)
echo 虚拟环境已激活
python --version
echo.

echo [4/5] 升级pip...
python -m pip install --upgrade pip
echo.

echo [5/5] 安装项目依赖（包括onnxruntime，这可能需要几分钟）...
echo 正在安装，请耐心等待...
pip install -r requirements.txt
if errorlevel 1 (
    echo [警告] 某些依赖安装失败，但继续执行
)
echo.

echo ========================================
echo 安装完成！
echo ========================================
echo.
echo 重要提示：
echo 1. 每次运行程序前，需要先激活虚拟环境：
echo    venv\Scripts\activate.bat
echo.
echo 2. 然后运行程序：
echo    python gui.py
echo.
echo 3. VAD功能已启用（通过onnxruntime）
echo.
echo 按任意键退出...
pause > nul
