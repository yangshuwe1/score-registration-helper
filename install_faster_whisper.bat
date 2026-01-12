@echo off
echo 正在安装 faster-whisper...
echo.

echo [1/2] 安装 av (PyAV)...
echo 注意: av 可能需要编译，请耐心等待
pip install av

echo.
echo [2/2] 安装 faster-whisper (最新版本)...
pip install faster-whisper --no-deps
pip install ctranslate2

echo.
echo 安装完成！
echo 验证安装: python -c "from faster_whisper import WhisperModel; print('安装成功')"
pause
