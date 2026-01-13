"""
语音识别模块
使用faster-whisper实现本地语音识别，支持实时识别（VAD自动检测）
"""
import wave
import numpy as np
from faster_whisper import WhisperModel
from typing import Optional, Callable
import threading
import os
import time
from scipy.io.wavfile import write as wav_write
from config import (
    WHISPER_MODEL, WHISPER_DEVICE, WHISPER_COMPUTE_TYPE,
    RECORD_DURATION, SAMPLE_RATE, CHUNK_SIZE
)

# 优先使用 sounddevice（更现代，安装更容易，支持实时录音）
try:
    import sounddevice as sd
    SOUNDDEVICE_AVAILABLE = True
    PYAUDIO_AVAILABLE = False
except ImportError:
    SOUNDDEVICE_AVAILABLE = False
    # 备用方案：使用 pyaudio
    try:
        import pyaudio
        PYAUDIO_AVAILABLE = True
    except ImportError:
        PYAUDIO_AVAILABLE = False
        print("警告: 未找到录音库。请安装: pip install sounddevice")


class SpeechRecognition:
    def __init__(self):
        self.model: Optional[WhisperModel] = None
        self.is_recording = False
        self.audio_frames = []
        self.audio_data = []  # 用于sounddevice
        self.callback = None  # 实时识别回调函数
        self._vad_warning_shown = False  # 标记VAD警告是否已显示
        self._load_model()
    
    def _load_model(self):
        """加载Whisper模型（带进度提示和资源管理）"""
        try:
            import os
            import time
            from pathlib import Path
            
            print("=" * 60)
            print("正在加载语音识别模型...")
            print(f"模型: {WHISPER_MODEL}, 设备: {WHISPER_DEVICE}, 计算类型: {WHISPER_COMPUTE_TYPE}")
            print("=" * 60)
            
            # 检查模型是否已下载
            cache_dir = Path.home() / ".cache" / "huggingface" / "hub"
            model_path = cache_dir / f"models--guillaumekln--faster-whisper-{WHISPER_MODEL}"
            
            if model_path.exists():
                print("✓ 检测到已下载的模型")
                print("  正在加载到内存（可能需要10-30秒）...")
            else:
                print("⚠ 首次运行，需要下载模型")
                print("  模型大小: 约150MB")
                print("  下载时间: 取决于网络速度（通常1-5分钟）")
                print("  请保持网络连接，不要关闭程序...")
                print("  正在下载中...")
            
            # 记录开始时间
            start_time = time.time()
            
            # 加载模型（这里可能会卡住，但会在后台线程中执行）
            print("  正在初始化模型（请稍候）...")
            self.model = WhisperModel(
                WHISPER_MODEL,
                device=WHISPER_DEVICE,
                compute_type=WHISPER_COMPUTE_TYPE,
                download_root=None  # 使用默认缓存目录
            )
            
            elapsed_time = time.time() - start_time
            print(f"✓ 模型加载完成！耗时: {elapsed_time:.1f}秒")
            print("=" * 60)
            
            self.model = WhisperModel(
                WHISPER_MODEL,
                device=WHISPER_DEVICE,
                compute_type=WHISPER_COMPUTE_TYPE,
                download_root=None  # 使用默认缓存目录
            )
            print("=" * 50)
            print("✓ 模型加载完成！")
            print("=" * 50)
        except Exception as e:
            print("=" * 50)
            print(f"✗ 加载模型失败: {e}")
            print("提示: 首次运行需要下载模型，请确保网络连接正常")
            print("     如果下载失败，请检查网络连接后重试")
            print("=" * 50)
            import traceback
            traceback.print_exc()
            self.model = None
    
    def record_audio_realtime(self, on_speech_end: Callable[[str], None],
                              silence_duration: float = 1.5,
                              min_speech_duration: float = 0.5) -> bool:
        """
        实时录音，使用VAD自动检测说话结束
        on_speech_end: 检测到说话结束后的回调函数，参数为识别的文本
        silence_duration: 静音持续时间（秒），超过此时间认为说话结束
        min_speech_duration: 最小说话时长（秒），短于此时间忽略
        返回: True表示成功识别一次语音，False表示失败
        """
        if self.model is None:
            print("模型未加载")
            return False

        if not SOUNDDEVICE_AVAILABLE and not PYAUDIO_AVAILABLE:
            print("错误: 未找到录音库")
            print("请运行: pip install sounddevice")
            return False

        self.audio_data = []
        speech_started = False
        last_speech_time = None
        recording_stopped = False  # 标记录音是否已停止

        def audio_callback(indata, frames, time_info, status):
            nonlocal speech_started, last_speech_time, recording_stopped

            if status:
                print(f"录音状态: {status}")

            if recording_stopped:
                raise sd.CallbackStop

            # 计算音频能量（简单的VAD）
            audio_level = np.abs(indata).mean()
            threshold = 0.01  # 音量阈值，可根据环境调整

            if audio_level > threshold:
                # 检测到声音
                self.audio_data.append(indata.copy())
                if not speech_started:
                    speech_started = True
                    last_speech_time = time.time()
                else:
                    last_speech_time = time.time()
            else:
                # 静音
                if speech_started:
                    # 记录静音片段
                    self.audio_data.append(indata.copy())
                    # 检查是否静音时间过长
                    if last_speech_time and (time.time() - last_speech_time) > silence_duration:
                        # 说话结束，停止录音
                        recording_stopped = True
                        raise sd.CallbackStop

        try:
            print("开始实时录音...（说完后自动识别）")
            if SOUNDDEVICE_AVAILABLE:
                # 使用sounddevice录音
                with sd.InputStream(samplerate=SAMPLE_RATE, channels=1,
                                   callback=audio_callback, blocksize=CHUNK_SIZE):
                    # 等待录音完成（由callback中的sd.CallbackStop触发）
                    while not recording_stopped:
                        time.sleep(0.1)
            else:
                # 使用pyaudio（旧方法，不支持实时VAD）
                return self._record_with_pyaudio_vad(on_speech_end, silence_duration)

            # 录音结束，处理音频
            if len(self.audio_data) > 0:
                # 合并音频数据
                audio_array = np.concatenate(self.audio_data, axis=0)
                audio_duration = len(audio_array) / SAMPLE_RATE

                if audio_duration < min_speech_duration:
                    print(f"录音太短（{audio_duration:.2f}秒），忽略")
                    return False

                # 保存为临时文件
                temp_file = "temp_recording.wav"
                wav_write(temp_file, SAMPLE_RATE, (audio_array * 32767).astype(np.int16))

                # 识别
                text = self.transcribe(temp_file)
                if text:
                    on_speech_end(text)
                    return True
                return False
            else:
                print("未录制到音频")
                return False

        except sd.CallbackStop:
            # 正常停止（VAD检测到说话结束）
            if len(self.audio_data) > 0:
                audio_array = np.concatenate(self.audio_data, axis=0)
                audio_duration = len(audio_array) / SAMPLE_RATE

                if audio_duration >= min_speech_duration:
                    temp_file = "temp_recording.wav"
                    wav_write(temp_file, SAMPLE_RATE, (audio_array * 32767).astype(np.int16))
                    text = self.transcribe(temp_file)
                    if text:
                        on_speech_end(text)
                        return True
            return False
        except Exception as e:
            print(f"实时录音失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def record_audio(self, duration: float = RECORD_DURATION) -> Optional[str]:
        """
        录制音频并保存为临时文件（传统方法）
        返回临时文件路径
        """
        if self.model is None:
            return None
        
        if SOUNDDEVICE_AVAILABLE:
            # 使用sounddevice
            try:
                print("开始录音...")
                self.is_recording = True
                recording = sd.rec(int(duration * SAMPLE_RATE), 
                                  samplerate=SAMPLE_RATE, 
                                  channels=1)
                sd.wait()  # 等待录音完成
                print("录音结束")
                
                # 保存为WAV文件
                temp_file = "temp_recording.wav"
                wav_write(temp_file, SAMPLE_RATE, (recording * 32767).astype(np.int16))
                return temp_file
            except Exception as e:
                print(f"录音失败: {e}")
                return None
        
        if not PYAUDIO_AVAILABLE:
            print("错误: 未找到录音库")
            print("请运行: pip install sounddevice")
            return None
        
        audio = None
        stream = None
        try:
            audio = pyaudio.PyAudio()
            
            # 检查可用输入设备
            input_devices = []
            for i in range(audio.get_device_count()):
                info = audio.get_device_info_by_index(i)
                if info['maxInputChannels'] > 0:
                    input_devices.append(i)
            
            if not input_devices:
                print("未找到可用的输入设备")
                return None
            
            # 打开音频流
            stream = audio.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=SAMPLE_RATE,
                input=True,
                frames_per_buffer=CHUNK_SIZE
            )
            
            print("开始录音...")
            self.is_recording = True
            self.audio_frames = []
            
            # 录制音频
            max_frames = int(SAMPLE_RATE / CHUNK_SIZE * duration)
            for _ in range(max_frames):
                if not self.is_recording:
                    break
                try:
                    data = stream.read(CHUNK_SIZE, exception_on_overflow=False)
                    self.audio_frames.append(data)
                except Exception as e:
                    print(f"读取音频数据失败: {e}")
                    break
            
            print("录音结束")
            
            if not self.audio_frames:
                print("未录制到音频数据")
                return None
            
            # 保存为WAV文件
            temp_file = "temp_recording.wav"
            wf = wave.open(temp_file, 'wb')
            wf.setnchannels(1)
            wf.setsampwidth(audio.get_sample_size(pyaudio.paInt16))
            wf.setframerate(SAMPLE_RATE)
            wf.writeframes(b''.join(self.audio_frames))
            wf.close()
            
            return temp_file
        except OSError as e:
            print(f"音频设备错误: {e}")
            print("请检查麦克风是否已连接并启用")
            return None
        except Exception as e:
            print(f"录音失败: {e}")
            return None
        finally:
            # 确保资源被释放
            if stream:
                try:
                    stream.stop_stream()
                    stream.close()
                except:
                    pass
            if audio:
                try:
                    audio.terminate()
                except:
                    pass
    
    def stop_recording(self):
        """停止录音"""
        self.is_recording = False
    
    def transcribe(self, audio_file: str) -> Optional[str]:
        """
        将音频文件转换为文字
        """
        if self.model is None:
            return None
        
        if not os.path.exists(audio_file):
            print(f"音频文件不存在: {audio_file}")
            return None
        
        try:
            print("正在识别语音...")

            # 添加 initial_prompt 来引导 whisper 识别特定格式
            # 这个 prompt 会告诉 whisper 我们期望的输入格式，提高识别准确率
            # 注意：不要在 prompt 中包含具体例子，否则 whisper 会学习这些内容
            initial_prompt = (
                "这是一个成绩登记系统。"
                "用户只会输入数字序号和数字分数，或者中文学生姓名和数字分数。"
                "不会说其他无关内容。"
            )

            # 尝试使用VAD过滤器（需要onnxruntime）
            try:
                segments, info = self.model.transcribe(
                    audio_file,
                    beam_size=5,
                    language="zh",
                    vad_filter=True,  # 启用VAD过滤，提高准确率
                    vad_parameters=dict(min_silence_duration_ms=500),
                    initial_prompt=initial_prompt  # 添加引导性 prompt
                )
            except RuntimeError as e:
                # 如果VAD不可用（缺少onnxruntime），禁用VAD重试
                if "onnxruntime" in str(e):
                    # 只在第一次显示警告
                    if not self._vad_warning_shown:
                        print("提示: VAD过滤器不可用（缺少onnxruntime），使用标准模式")
                        print("      建议安装: pip install onnxruntime")
                        self._vad_warning_shown = True
                    segments, info = self.model.transcribe(
                        audio_file,
                        beam_size=5,
                        language="zh",
                        vad_filter=False,  # 禁用VAD
                        initial_prompt=initial_prompt  # 仍然使用引导性 prompt
                    )
                else:
                    raise

            # 获取识别结果
            text = ""
            for segment in segments:
                text += segment.text

            text = text.strip()
            print(f"识别结果: {text}")

            # 清理临时文件（延迟清理，确保识别完成）
            try:
                if os.path.exists(audio_file):
                    os.remove(audio_file)
            except:
                pass

            return text if text else None
        except Exception as e:
            print(f"语音识别失败: {e}")
            import traceback
            traceback.print_exc()
            # 清理临时文件
            try:
                if os.path.exists(audio_file):
                    os.remove(audio_file)
            except:
                pass
            return None
    
    def record_and_transcribe(self, duration: float = RECORD_DURATION) -> Optional[str]:
        """
        录制并识别，一步完成
        """
        audio_file = self.record_audio(duration)
        if audio_file:
            return self.transcribe(audio_file)
        return None
