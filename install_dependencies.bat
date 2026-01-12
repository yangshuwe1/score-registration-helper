@echo off
echo 正在安装依赖包...
echo.

echo [1/3] 安装基础依赖...
pip install openpyxl>=3.1.2
pip install pandas>=2.0.0
pip install xlrd>=2.0.1
pip install edge-tts>=6.1.0
pip install numpy>=1.24.0
pip install pyinstaller>=5.13.0

echo.
echo [2/3] 安装PyTorch...
pip install torch>=2.0.0

echo.
echo [3/3] 尝试安装faster-whisper和tokenizers...
echo 注意: 如果tokenizers编译失败，将尝试安装预编译版本

REM 先尝试安装tokenizers的预编译版本
pip install tokenizers --only-binary :all: 2>nul
if %errorlevel% neq 0 (
    echo tokenizers预编译版本不可用，尝试从源码安装...
    pip install tokenizers
)

REM 安装faster-whisper
pip install faster-whisper>=0.10.0

echo.
echo [4/4] 安装PyAudio...
echo 注意: PyAudio可能需要Visual C++编译器
echo 如果安装失败，可以尝试: pip install pipwin && pipwin install pyaudio

pip install pyaudio>=0.2.14
if %errorlevel% neq 0 (
    echo.
    echo PyAudio安装失败！
    echo 请尝试以下方法之一:
    echo 1. 安装 pipwin: pip install pipwin
    echo 2. 然后运行: pipwin install pyaudio
    echo 3. 或者从 https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio 下载预编译版本
)

echo.
echo 安装完成！
pause
