"""
语音识别模块 - 重构版
实现ChatGPT式的连续对话体验：
- 持续录音，无需反复开关
- 智能VAD自动检测语音段
- 异步识别队列，避免阻塞
- 高准确率的序号/姓名+分数识别
"""
import wave
import numpy as np
from faster_whisper import WhisperModel
from typing import Optional, Callable
import threading
import queue
import os
import time
import re
from scipy.io.wavfile import write as wav_write
from config import (
    WHISPER_MODEL, WHISPER_DEVICE, WHISPER_COMPUTE_TYPE,
    RECORD_DURATION, SAMPLE_RATE, CHUNK_SIZE
)

# 优先使用 sounddevice
try:
    import sounddevice as sd
    SOUNDDEVICE_AVAILABLE = True
    PYAUDIO_AVAILABLE = False
except ImportError:
    SOUNDDEVICE_AVAILABLE = False
    try:
        import pyaudio
        PYAUDIO_AVAILABLE = True
    except ImportError:
        PYAUDIO_AVAILABLE = False
        print("警告: 未找到录音库。请安装: pip install sounddevice")


class ContinuousSpeechRecognition:
    """
    持续语音识别系统
    特点：
    1. 一次启动，持续运行
    2. 智能VAD自动分段
    3. 异步识别，不阻塞录音
    4. 高准确率识别
    """

    def __init__(self):
        self.model: Optional[WhisperModel] = None
        self._vad_warning_shown = False

        # 录音控制
        self.is_running = False  # 整个系统是否运行
        self.recording_thread = None

        # 音频缓冲
        self.audio_buffer = []
        self.buffer_lock = threading.Lock()

        # 识别队列
        self.recognition_queue = queue.Queue(maxsize=5)
        self.recognition_thread = None

        # 回调函数
        self.on_recognition_callback = None

        # VAD状态
        self.speech_started = False
        self.last_speech_time = None
        self.silence_duration = 1.5  # 静音持续时间（秒）
        self.min_speech_duration = 0.6  # 最小语音时长（秒）

        # 能量平滑缓冲
        self.energy_buffer = []
        self.energy_buffer_size = 5

        # 去重机制
        self.last_recognition = ""
        self.last_recognition_time = 0

        # 计数器
        self._file_count = 0

        # 加载模型
        self._load_model()

    def _load_model(self):
        """加载Whisper模型"""
        try:
            from pathlib import Path

            print("=" * 60)
            print("正在加载语音识别模型...")
            print(f"模型: {WHISPER_MODEL}, 设备: {WHISPER_DEVICE}")
            print("=" * 60)

            cache_dir = Path.home() / ".cache" / "huggingface" / "hub"
            model_path = cache_dir / f"models--guillaumekln--faster-whisper-{WHISPER_MODEL}"

            if model_path.exists():
                print("✓ 使用已缓存的模型")
            else:
                print("⚠ 首次运行，正在下载模型（约150MB）...")

            start_time = time.time()
            self.model = WhisperModel(
                WHISPER_MODEL,
                device=WHISPER_DEVICE,
                compute_type=WHISPER_COMPUTE_TYPE,
                download_root=None
            )

            elapsed = time.time() - start_time
            print(f"✓ 模型加载完成！耗时: {elapsed:.1f}秒")
            print("=" * 60)

        except Exception as e:
            print(f"✗ 模型加载失败: {e}")
            import traceback
            traceback.print_exc()
            self.model = None

    def start(self, on_recognition: Callable[[str], None]):
        """
        启动持续识别系统
        on_recognition: 识别结果回调函数
        """
        if self.model is None:
            print("错误: 模型未加载")
            return False

        if not SOUNDDEVICE_AVAILABLE and not PYAUDIO_AVAILABLE:
            print("错误: 未找到录音库")
            return False

        if self.is_running:
            print("系统已在运行中")
            return True

        self.on_recognition_callback = on_recognition
        self.is_running = True

        # 启动识别线程
        self.recognition_thread = threading.Thread(
            target=self._recognition_worker,
            daemon=True
        )
        self.recognition_thread.start()

        # 启动录音线程
        self.recording_thread = threading.Thread(
            target=self._recording_worker,
            daemon=True
        )
        self.recording_thread.start()

        print("=" * 60)
        print("✓ 持续识别系统已启动")
        print("  说话会自动识别，无需手动操作")
        print("  支持连续识别多个学生成绩")
        print("=" * 60)

        return True

    def stop(self):
        """停止持续识别系统"""
        if not self.is_running:
            return

        print("正在停止识别系统...")
        self.is_running = False

        # 等待线程结束
        if self.recording_thread:
            self.recording_thread.join(timeout=2)
        if self.recognition_thread:
            self.recognition_thread.join(timeout=2)

        print("识别系统已停止")

    def _recording_worker(self):
        """录音工作线程（持续运行）"""
        try:
            if SOUNDDEVICE_AVAILABLE:
                self._recording_with_sounddevice()
            else:
                self._recording_with_pyaudio()
        except Exception as e:
            print(f"录音线程异常: {e}")
            import traceback
            traceback.print_exc()

    def _recording_with_sounddevice(self):
        """使用sounddevice持续录音"""
        def audio_callback(indata, frames, time_info, status):
            if status:
                print(f"录音状态: {status}")

            if not self.is_running:
                raise sd.CallbackStop

            # 计算音频能量（RMS）
            audio_level = np.sqrt(np.mean(indata**2))

            # 能量平滑
            self.energy_buffer.append(audio_level)
            if len(self.energy_buffer) > self.energy_buffer_size:
                self.energy_buffer.pop(0)

            smoothed_level = np.mean(self.energy_buffer)

            # VAD阈值（动态调整）
            threshold = 0.015  # 降低阈值，更容易触发

            current_time = time.time()

            if smoothed_level > threshold:
                # 检测到语音
                with self.buffer_lock:
                    self.audio_buffer.append(indata.copy())

                if not self.speech_started:
                    self.speech_started = True
                    self.last_speech_time = current_time
                    print("🎤 检测到语音...")
                else:
                    self.last_speech_time = current_time
            else:
                # 静音
                if self.speech_started:
                    # 继续记录静音段（保持连续性）
                    with self.buffer_lock:
                        self.audio_buffer.append(indata.copy())

                    # 检查是否静音时间过长
                    if (current_time - self.last_speech_time) > self.silence_duration:
                        # 语音段结束
                        self._process_audio_segment()
                        self.speech_started = False

        try:
            with sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=1,
                callback=audio_callback,
                blocksize=CHUNK_SIZE
            ):
                print("🎙️  录音系统已就绪，等待语音输入...")
                while self.is_running:
                    time.sleep(0.1)
        except sd.CallbackStop:
            pass

    def _recording_with_pyaudio(self):
        """使用pyaudio持续录音（备用方案）"""
        # 简化实现，主要使用sounddevice
        print("警告: pyaudio模式不支持持续识别，请使用sounddevice")

    def _process_audio_segment(self):
        """处理一个完整的语音段"""
        with self.buffer_lock:
            if len(self.audio_buffer) == 0:
                return

            # 复制缓冲区
            audio_array = np.concatenate(self.audio_buffer, axis=0)
            self.audio_buffer = []  # 清空缓冲区

        # 检查时长
        duration = len(audio_array) / SAMPLE_RATE
        if duration < self.min_speech_duration:
            print(f"  语音太短（{duration:.2f}秒），忽略")
            return

        print(f"  语音段结束（{duration:.2f}秒），加入识别队列...")

        # 放入识别队列
        try:
            self.recognition_queue.put_nowait((audio_array, time.time()))
        except queue.Full:
            print("  警告: 识别队列已满，丢弃此语音段")

    def _recognition_worker(self):
        """识别工作线程（异步处理）"""
        while self.is_running:
            try:
                # 从队列获取音频
                audio_array, timestamp = self.recognition_queue.get(timeout=0.5)

                # 识别
                text = self._recognize_audio(audio_array)

                # 去重检查
                if self._is_duplicate(text, timestamp):
                    print("  检测到重复，跳过")
                    continue

                # 回调用户
                if text and self.on_recognition_callback:
                    self.last_recognition = text
                    self.last_recognition_time = timestamp
                    self.on_recognition_callback(text)

            except queue.Empty:
                continue
            except Exception as e:
                print(f"识别线程异常: {e}")
                import traceback
                traceback.print_exc()

    def _recognize_audio(self, audio_array: np.ndarray) -> Optional[str]:
        """识别音频数组"""
        try:
            # 音频预处理
            audio_array = self._preprocess_audio(audio_array)

            # 保存临时文件
            self._file_count += 1
            temp_file = f"temp_audio_{self._file_count}.wav"
            wav_write(temp_file, SAMPLE_RATE, (audio_array * 32767).astype(np.int16))

            # 调用Whisper识别
            text = self._transcribe(temp_file)

            # 清理临时文件
            try:
                os.remove(temp_file)
            except:
                pass

            return text

        except Exception as e:
            print(f"  识别失败: {e}")
            return None

    def _transcribe(self, audio_file: str) -> Optional[str]:
        """Whisper转录（优化版）"""
        if not os.path.exists(audio_file):
            return None

        try:
            print("  🔍 正在识别...")

            # 精简的prompt，避免污染
            # 只提供数字和格式信息，不包含容易识别的完整句子
            prompt = "1 2 3 4 5 6 7 8 9 10, 85 90 95 100"

            # 尝试使用VAD
            try:
                segments, info = self.model.transcribe(
                    audio_file,
                    beam_size=5,
                    language="zh",
                    vad_filter=True,
                    vad_parameters=dict(
                        min_silence_duration_ms=700,
                        speech_pad_ms=200,
                        threshold=0.5
                    ),
                    initial_prompt=prompt,
                    temperature=0.0,
                    condition_on_previous_text=False,
                    compression_ratio_threshold=2.4,
                    no_speech_threshold=0.5,
                    log_prob_threshold=-1.0,
                )
            except RuntimeError as e:
                if "onnxruntime" in str(e).lower():
                    if not self._vad_warning_shown:
                        print("  提示: VAD不可用，使用标准模式")
                        self._vad_warning_shown = True
                    segments, info = self.model.transcribe(
                        audio_file,
                        beam_size=5,
                        language="zh",
                        vad_filter=False,
                        initial_prompt=prompt,
                        temperature=0.0,
                        condition_on_previous_text=False,
                        compression_ratio_threshold=2.4,
                        no_speech_threshold=0.5,
                        log_prob_threshold=-1.0,
                    )
                else:
                    raise

            # 获取识别结果
            text = "".join(segment.text for segment in segments).strip()

            if not text:
                return None

            print(f"  原始: {text}")

            # 后处理
            text = self._postprocess_text(text)

            print(f"  ✓ 结果: {text}")

            return text if text else None

        except Exception as e:
            print(f"  转录失败: {e}")
            return None

    def _preprocess_audio(self, audio_array: np.ndarray) -> np.ndarray:
        """音频预处理"""
        # 去除DC偏移
        audio_array = audio_array - np.mean(audio_array)

        # 归一化
        max_val = np.max(np.abs(audio_array))
        if max_val > 0:
            audio_array = audio_array / max_val

        # 限幅
        audio_array = np.clip(audio_array, -1.0, 1.0)

        return audio_array

    def _postprocess_text(self, text: str) -> str:
        """文本后处理（简化版）"""
        if not text:
            return ""

        # 1. 基本清理
        text = text.strip()

        # 2. 去除prompt污染（如果包含明显的prompt内容）
        # 检测并移除"格式"、"序号"等prompt关键词
        pollution_keywords = ['格式', '序号', '例如', '学生', '成绩', '登记']
        for keyword in pollution_keywords:
            if keyword in text and '号' in text and '分' in text:
                # 尝试提取有效部分（号和分之间的内容）
                # 例如："格式：序号90分" -> "90分"（但这已经无法恢复序号）
                # 更好的方式是直接过滤掉
                parts = text.split(keyword)
                if len(parts) > 1:
                    # 取最后一部分（更可能是实际内容）
                    text = parts[-1].strip()
                    # 如果开头是冒号或标点，去除
                    text = re.sub(r'^[：:，,。.、]+', '', text)

        # 3. 规范化标点
        text = text.replace('，', ',').replace('。', '.').replace('、', ',')

        # 4. 去除英文噪音
        text = re.sub(r'\b[A-Za-z]+\s+[A-Za-z]+\b', '', text)
        text = re.sub(r'(?<!\d)[４](?!\d)', '', text)

        # 5. 中文数字转换
        chinese_num_map = {
            '零': '0', '一': '1', '二': '2', '三': '3', '四': '4',
            '五': '5', '六': '6', '七': '7', '八': '8', '九': '9',
            '十': '10'
        }

        for cn, num in chinese_num_map.items():
            text = re.sub(f'{cn}(?=号)', num, text)
            text = re.sub(f'{cn}(?=分)', num, text)

        # 6. 去重（处理识别中的重复）
        parts = [p.strip() for p in text.split(',')]
        unique_parts = []
        seen = set()

        for part in parts:
            if part and part not in seen:
                unique_parts.append(part)
                seen.add(part)

        text = ','.join(unique_parts)

        # 7. 清理多余符号
        text = re.sub(r',+', ',', text)
        text = re.sub(r'[,\s]+$', '', text)
        text = re.sub(r'^[,\s]+', '', text)

        return text.strip()

    def _is_duplicate(self, text: str, timestamp: float) -> bool:
        """检测重复识别"""
        if not text or not self.last_recognition:
            return False

        # 时间间隔太短（2秒内）
        if timestamp - self.last_recognition_time < 2.0:
            # 文本相似度检查
            from difflib import SequenceMatcher
            similarity = SequenceMatcher(None, text, self.last_recognition).ratio()
            if similarity > 0.7:
                return True

        return False


# 兼容旧接口的包装类
class SpeechRecognition:
    """
    兼容旧接口的包装类
    保持向后兼容，同时使用新的持续识别引擎
    """

    def __init__(self):
        self.engine = ContinuousSpeechRecognition()
        self.model = self.engine.model
        self._callback_handler = None

    def record_audio_realtime(self, on_speech_end: Callable[[str], None],
                              silence_duration: float = 1.5,
                              min_speech_duration: float = 0.6) -> bool:
        """
        兼容接口：实时录音识别一次
        实际上启动持续识别，识别一次后停止
        """
        self._callback_handler = on_speech_end

        # 启动持续识别
        def callback_wrapper(text):
            # 识别一次后停止
            self.engine.stop()
            on_speech_end(text)

        self.engine.silence_duration = silence_duration
        self.engine.min_speech_duration = min_speech_duration

        return self.engine.start(callback_wrapper)

    def stop_recording(self):
        """停止录音"""
        self.engine.stop()

    def transcribe(self, audio_file: str, use_prompt: bool = True) -> Optional[str]:
        """直接转录音频文件"""
        return self.engine._transcribe(audio_file)
