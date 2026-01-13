"""
语音识别模块 - 极简版
目标：轻量化、稳定、准确
参考：Whisper官方示例 + faster-whisper最佳实践
"""
import numpy as np
from faster_whisper import WhisperModel
from typing import Optional, Callable
import os
import time
from scipy.io.wavfile import write as wav_write
from config import WHISPER_MODEL, WHISPER_DEVICE, WHISPER_COMPUTE_TYPE, SAMPLE_RATE, CHUNK_SIZE

try:
    import sounddevice as sd
    SOUNDDEVICE_AVAILABLE = True
except ImportError:
    SOUNDDEVICE_AVAILABLE = False
    print("警告: 未找到sounddevice。请安装: pip install sounddevice")


class SpeechRecognition:
    """简单可靠的语音识别系统"""

    def __init__(self):
        self.model: Optional[WhisperModel] = None
        self.is_recording = False
        self._vad_warning_shown = False
        self._load_model()

    def _load_model(self):
        """加载Whisper模型"""
        try:
            print("=" * 50)
            print("正在加载Whisper模型...")
            print(f"模型: {WHISPER_MODEL}")
            print("=" * 50)

            start = time.time()
            self.model = WhisperModel(
                WHISPER_MODEL,
                device=WHISPER_DEVICE,
                compute_type=WHISPER_COMPUTE_TYPE
            )
            elapsed = time.time() - start

            print(f"✓ 模型加载完成！耗时: {elapsed:.1f}秒")
            print("=" * 50)

        except Exception as e:
            print(f"✗ 模型加载失败: {e}")
            self.model = None

    def record_audio_realtime(self, on_speech_end: Callable[[str], None],
                              silence_duration: float = 1.5,
                              min_speech_duration: float = 0.8) -> bool:
        """
        实时录音识别（单次）
        - 检测到语音后自动开始录音
        - 静音silence_duration秒后自动停止
        - 识别完成后调用on_speech_end回调
        - 返回True表示成功识别一次，False表示失败
        """
        if self.model is None:
            print("模型未加载")
            return False

        if not SOUNDDEVICE_AVAILABLE:
            print("错误: 需要sounddevice库")
            return False

        # 状态变量
        audio_buffer = []
        speech_started = False
        last_speech_time = None
        energy_buffer = []

        def audio_callback(indata, frames, time_info, status):
            nonlocal speech_started, last_speech_time

            if status:
                print(f"录音状态: {status}")

            if not self.is_recording:
                raise sd.CallbackStop

            # 计算音频能量（RMS）
            audio_level = np.sqrt(np.mean(indata**2))

            # 平滑能量值（5帧平均）
            energy_buffer.append(audio_level)
            if len(energy_buffer) > 5:
                energy_buffer.pop(0)
            smoothed_level = np.mean(energy_buffer)

            # VAD阈值
            threshold = 0.02

            current_time = time.time()

            if smoothed_level > threshold:
                # 检测到语音
                audio_buffer.append(indata.copy())
                if not speech_started:
                    speech_started = True
                    last_speech_time = current_time
                else:
                    last_speech_time = current_time
            else:
                # 静音
                if speech_started:
                    audio_buffer.append(indata.copy())
                    # 检查静音时长
                    if (current_time - last_speech_time) > silence_duration:
                        # 说话结束
                        raise sd.CallbackStop

        try:
            self.is_recording = True

            # 开始录音
            with sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=1,
                callback=audio_callback,
                blocksize=CHUNK_SIZE
            ):
                while self.is_recording:
                    time.sleep(0.1)

        except sd.CallbackStop:
            pass
        except Exception as e:
            print(f"录音失败: {e}")
            return False
        finally:
            self.is_recording = False

        # 处理录音
        if len(audio_buffer) == 0:
            return False

        audio_array = np.concatenate(audio_buffer, axis=0)
        duration = len(audio_array) / SAMPLE_RATE

        if duration < min_speech_duration:
            print(f"录音太短（{duration:.2f}秒），忽略")
            return False

        # 音频预处理
        audio_array = self._preprocess_audio(audio_array)

        # 保存临时文件
        temp_file = f"temp_{int(time.time() * 1000)}.wav"
        wav_write(temp_file, SAMPLE_RATE, (audio_array * 32767).astype(np.int16))

        # 识别
        text = self._transcribe(temp_file)

        # 清理临时文件
        try:
            os.remove(temp_file)
        except:
            pass

        # 回调
        if text:
            on_speech_end(text)
            return True

        return False

    def _preprocess_audio(self, audio: np.ndarray) -> np.ndarray:
        """音频预处理"""
        # 去除DC偏移
        audio = audio - np.mean(audio)
        # 归一化
        max_val = np.max(np.abs(audio))
        if max_val > 0:
            audio = audio / max_val
        # 限幅
        audio = np.clip(audio, -1.0, 1.0)
        return audio

    def _transcribe(self, audio_file: str) -> Optional[str]:
        """转录音频文件"""
        if not os.path.exists(audio_file):
            return None

        try:
            print("正在识别语音...")

            # 使用简单的prompt，引导识别序号+分数格式
            # 参考Whisper官方建议：简短、相关、不要太复杂
            prompt = "一号，二号，三号，四号，五号，100分，95分，90分，85分"

            # 转录参数（参考faster-whisper官方示例）
            try:
                segments, info = self.model.transcribe(
                    audio_file,
                    language="zh",
                    beam_size=5,
                    vad_filter=True,
                    vad_parameters=dict(
                        min_silence_duration_ms=500,
                        threshold=0.5
                    ),
                    initial_prompt=prompt,
                    temperature=0.0,
                    condition_on_previous_text=False,
                    no_speech_threshold=0.6
                )
            except RuntimeError as e:
                if "onnxruntime" in str(e).lower():
                    if not self._vad_warning_shown:
                        print("提示: VAD不可用，使用标准模式")
                        self._vad_warning_shown = True
                    segments, info = self.model.transcribe(
                        audio_file,
                        language="zh",
                        beam_size=5,
                        vad_filter=False,
                        initial_prompt=prompt,
                        temperature=0.0,
                        condition_on_previous_text=False,
                        no_speech_threshold=0.6
                    )
                else:
                    raise

            # 获取识别结果
            text = "".join(segment.text for segment in segments).strip()

            if text:
                print(f"识别结果: {text}")
                # 简单后处理
                text = self._postprocess(text)
                return text

            return None

        except Exception as e:
            print(f"识别失败: {e}")
            return None

    def _postprocess(self, text: str) -> str:
        """简单后处理"""
        import re

        # 规范化标点
        text = text.replace('，', ',').replace('。', '.')

        # 中文数字转阿拉伯数字（只在"号"和"分"前面）
        replacements = {
            '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
            '六': '6', '七': '7', '八': '8', '九': '9', '十': '10'
        }
        for cn, num in replacements.items():
            text = re.sub(f'{cn}(?=号)', num, text)
            text = re.sub(f'{cn}(?=分)', num, text)

        # 清理多余符号
        text = re.sub(r',+', ',', text)
        text = re.sub(r'[,\s]+$', '', text)

        return text.strip()

    def stop_recording(self):
        """停止录音"""
        self.is_recording = False
