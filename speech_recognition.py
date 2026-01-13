"""
语音识别模块
使用faster-whisper实现本地语音识别，支持实时识别（VAD自动检测）
参考OpenAI Whisper最佳实践进行优化
"""
import wave
import numpy as np
from faster_whisper import WhisperModel
from typing import Optional, Callable
import threading
import os
import time
import re
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
        self._last_recognition = ""  # 上次识别结果，用于去重
        self._recognition_count = 0  # 识别计数器，用于临时文件命名
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
                              silence_duration: float = 2.0,
                              min_speech_duration: float = 0.8) -> bool:
        """
        实时录音，使用VAD自动检测说话结束
        on_speech_end: 检测到说话结束后的回调函数，参数为识别的文本
        silence_duration: 静音持续时间（秒），超过此时间认为说话结束（增加到2.0秒）
        min_speech_duration: 最小说话时长（秒），短于此时间忽略（增加到0.8秒）
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

        # 用于平滑VAD检测的缓冲区
        energy_buffer = []
        buffer_size = 5  # 使用最近5帧的平均值

        def audio_callback(indata, frames, time_info, status):
            nonlocal speech_started, last_speech_time, recording_stopped, energy_buffer

            if status:
                print(f"录音状态: {status}")

            if recording_stopped:
                raise sd.CallbackStop

            # 计算音频能量（改进的VAD）
            # 使用RMS（均方根）而不是简单平均值，更准确
            audio_level = np.sqrt(np.mean(indata**2))

            # 添加到缓冲区进行平滑
            energy_buffer.append(audio_level)
            if len(energy_buffer) > buffer_size:
                energy_buffer.pop(0)

            # 使用平滑后的能量值
            smoothed_level = np.mean(energy_buffer)

            # 动态阈值：提高到0.02，减少噪音误触发
            threshold = 0.02

            if smoothed_level > threshold:
                # 检测到声音
                self.audio_data.append(indata.copy())
                if not speech_started:
                    speech_started = True
                    last_speech_time = time.time()
                    print("检测到语音开始")
                else:
                    last_speech_time = time.time()
            else:
                # 静音
                if speech_started:
                    # 记录静音片段（保持音频连续性）
                    self.audio_data.append(indata.copy())
                    # 检查是否静音时间过长
                    if last_speech_time and (time.time() - last_speech_time) > silence_duration:
                        # 说话结束，停止录音
                        print("检测到语音结束")
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

                # 音频预处理：降噪和归一化
                audio_array = self._preprocess_audio(audio_array)

                # 使用唯一的临时文件名，避免并发冲突
                self._recognition_count += 1
                temp_file = f"temp_recording_{self._recognition_count}.wav"

                # 保存为临时文件
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
                    # 音频预处理
                    audio_array = self._preprocess_audio(audio_array)

                    # 使用唯一的临时文件名
                    self._recognition_count += 1
                    temp_file = f"temp_recording_{self._recognition_count}.wav"

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

                # 音频预处理
                recording = self._preprocess_audio(recording)

                # 保存为WAV文件（使用唯一文件名）
                self._recognition_count += 1
                temp_file = f"temp_recording_{self._recognition_count}.wav"
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

            # 保存为WAV文件（使用唯一文件名）
            self._recognition_count += 1
            temp_file = f"temp_recording_{self._recognition_count}.wav"
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

    def _preprocess_audio(self, audio_array: np.ndarray) -> np.ndarray:
        """
        音频预处理：降噪和归一化
        参考：SpeechBrain、Kaldi等开源项目
        """
        # 1. 去除直流分量（DC offset）
        audio_array = audio_array - np.mean(audio_array)

        # 2. 归一化到[-1, 1]范围
        max_val = np.max(np.abs(audio_array))
        if max_val > 0:
            audio_array = audio_array / max_val

        # 3. 简单的高通滤波，去除低频噪音
        # 使用一阶差分近似高通滤波器
        # audio_filtered = np.diff(audio_array, prepend=audio_array[0])
        # audio_array = audio_filtered * 0.97 + audio_array  # 混合原始信号

        # 4. 限幅，避免削波
        audio_array = np.clip(audio_array, -1.0, 1.0)

        return audio_array

    def _postprocess_text(self, text: str) -> str:
        """
        后处理识别结果，提高准确率
        参考：OpenAI Whisper最佳实践
        """
        if not text:
            return ""

        # 1. 去除前后空格
        text = text.strip()

        # 2. 规范化标点符号（将中文标点统一处理）
        text = text.replace('，', ',').replace('。', '.').replace('、', ',')

        # 3. 去除明显的噪音（单个特殊字符、英文字母等）
        # 保留中文、数字、基本标点
        # 去除常见的识别噪音："４"、"Welner Rogers"等英文
        noise_patterns = [
            r'\b[A-Za-z]+\s+[A-Za-z]+\b',  # 英文单词组合（如"Welner Rogers"）
            r'(?<!\d)[４](?!\d)',  # 单独的全角数字４
            r'\s+$',  # 末尾空格
            r'^\s+',  # 开头空格
        ]

        for pattern in noise_patterns:
            text = re.sub(pattern, '', text)

        # 4. 规范化数字：将中文数字转换为阿拉伯数字
        chinese_num_map = {
            '零': '0', '一': '1', '二': '2', '三': '3', '四': '4',
            '五': '5', '六': '6', '七': '7', '八': '8', '九': '9',
            '十': '10', '百': '100'
        }

        # 简单替换（更复杂的转换需要专门的库）
        for cn, num in chinese_num_map.items():
            # 只替换在"号"和"分"上下文中的数字
            text = re.sub(f'({cn})(?=号)', num, text)
            text = re.sub(f'(?<=号.{{0,20}})({cn})(?=分)', num, text)

        # 5. 去除重复片段（如"2号90分，2号90分"）
        # 检测并去除重复的模式
        parts = text.split(',')
        unique_parts = []
        seen = set()

        for part in parts:
            part_clean = part.strip()
            if part_clean and part_clean not in seen:
                unique_parts.append(part_clean)
                seen.add(part_clean)

        text = ','.join(unique_parts)

        # 6. 清理多余的逗号和空格
        text = re.sub(r',+', ',', text)  # 多个逗号合并为一个
        text = re.sub(r',\s*$', '', text)  # 去除末尾逗号
        text = re.sub(r'^\s*,', '', text)  # 去除开头逗号
        text = re.sub(r'\s+', ' ', text)  # 多个空格合并为一个

        return text.strip()

    def _is_duplicate(self, text: str) -> bool:
        """检查是否是重复识别"""
        if not text or not self._last_recognition:
            return False

        # 简单的字符串相似度比较
        # 如果新识别结果与上次结果高度相似（编辑距离小）则认为是重复
        from difflib import SequenceMatcher
        similarity = SequenceMatcher(None, text, self._last_recognition).ratio()

        return similarity > 0.8  # 相似度超过80%认为是重复
    
    def transcribe(self, audio_file: str, use_prompt: bool = True) -> Optional[str]:
        """
        将音频文件转换为文字
        use_prompt: 是否使用prompt强化识别（优先识别序号+分数格式）
        参考OpenAI Whisper最佳实践优化
        """
        if self.model is None:
            return None

        if not os.path.exists(audio_file):
            print(f"音频文件不存在: {audio_file}")
            return None

        try:
            print("正在识别语音...")

            # 构建更丰富的prompt，引导模型识别特定格式
            # 参考：https://platform.openai.com/docs/guides/speech-to-text/prompting
            prompt = None
            if use_prompt:
                # 详细的提示，包含：
                # 1. 格式示例（序号+分数）
                # 2. 常见数字（1-10, 90-100等高频分数）
                # 3. 明确的中文上下文
                prompt = (
                    "学生成绩登记：一号100分，二号95分，三号92分，四号88分，五号85分。"
                    "序号从1到10，分数从0到100分。"
                    "格式：序号加分数，例如一号100分，二号95分。"
                )

            # 尝试使用VAD过滤器（需要onnxruntime）
            try:
                segments, info = self.model.transcribe(
                    audio_file,
                    beam_size=5,
                    language="zh",
                    vad_filter=True,  # 启用VAD过滤，提高准确率
                    vad_parameters=dict(
                        min_silence_duration_ms=800,  # 增加到800ms，减少误触发
                        speech_pad_ms=300,  # 降低到300ms，减少前后噪音
                        threshold=0.6  # 提高到0.6，更保守的VAD检测
                    ),
                    initial_prompt=prompt,  # 添加prompt引导识别
                    temperature=0.0,  # 设置为0，降低随机性，提高一致性
                    condition_on_previous_text=False,  # 禁用上下文依赖，避免干扰
                    compression_ratio_threshold=2.4,  # 默认值，过滤重复内容
                    no_speech_threshold=0.6,  # 提高无语音阈值，过滤噪音
                    log_prob_threshold=-1.0,  # 默认值，过滤低概率结果
                )
            except RuntimeError as e:
                # 如果VAD不可用（缺少onnxruntime），禁用VAD重试
                if "onnxruntime" in str(e).lower():
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
                        initial_prompt=prompt,  # 添加prompt引导识别
                        temperature=0.0,  # 设置为0，降低随机性
                        condition_on_previous_text=False,  # 禁用上下文依赖
                        compression_ratio_threshold=2.4,
                        no_speech_threshold=0.6,
                        log_prob_threshold=-1.0,
                    )
                else:
                    raise

            # 获取识别结果
            text = ""
            for segment in segments:
                text += segment.text

            # 原始识别结果
            text = text.strip()
            print(f"识别结果（原始）: {text}")

            # 后处理：去噪、规范化、去重
            text = self._postprocess_text(text)
            print(f"识别结果（处理后）: {text}")

            # 检查是否是重复识别
            if self._is_duplicate(text):
                print("检测到重复识别，忽略")
                # 清理临时文件
                try:
                    if os.path.exists(audio_file):
                        os.remove(audio_file)
                except:
                    pass
                return None

            # 更新上次识别结果
            if text:
                self._last_recognition = text

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
