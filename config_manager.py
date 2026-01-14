"""
配置管理模块
负责用户配置的保存、加载和管理
"""
import json
import os
from typing import Dict, Any, Optional
from config import (
    EXCEL_COLUMNS,
    HEADER_ROWS,
    WHISPER_MODEL,
    WHISPER_DEVICE,
    WHISPER_COMPUTE_TYPE,
    RECORD_DURATION,
    SAMPLE_RATE,
    CHUNK_SIZE,
    TTS_VOICE
)


class ConfigManager:
    """配置管理器"""

    # 配置文件路径
    CONFIG_FILE = "user_config.json"

    # 默认配置
    DEFAULT_CONFIG = {
        'excel': {
            'header_rows': HEADER_ROWS,
            'columns': EXCEL_COLUMNS.copy()
        },
        'speech_recognition': {
            'model': WHISPER_MODEL,
            'device': WHISPER_DEVICE,
            'compute_type': WHISPER_COMPUTE_TYPE
        },
        'recording': {
            'duration': RECORD_DURATION,
            'sample_rate': SAMPLE_RATE,
            'chunk_size': CHUNK_SIZE
        },
        'tts': {
            'voice': TTS_VOICE
        }
    }

    # 可用的Whisper模型列表
    AVAILABLE_MODELS = ["tiny", "small", "base", "medium", "large"]

    # 模型说明
    MODEL_DESCRIPTIONS = {
        "tiny": "最快，最小（~75MB），适合低配置电脑",
        "small": "快速，小巧（~150MB），推荐一般使用",
        "base": "平衡（~150MB），速度和准确率适中",
        "medium": "慢，准确（~1.5GB），需要较好配置",
        "large": "最慢，最准确（~3GB），推荐高配置电脑"
    }

    def __init__(self):
        self.config: Dict[str, Any] = {}
        self.load_config()

    def load_config(self) -> bool:
        """
        从文件加载配置
        如果文件不存在，使用默认配置
        """
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)

                # 合并配置（保留用户配置，补充默认配置）
                self.config = self._merge_config(self.DEFAULT_CONFIG.copy(), loaded_config)
                print(f"已加载配置文件: {self.CONFIG_FILE}")
                return True
            else:
                # 使用默认配置
                self.config = self.DEFAULT_CONFIG.copy()
                print("使用默认配置")
                return True
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            # 使用默认配置
            self.config = self.DEFAULT_CONFIG.copy()
            return False

    def save_config(self) -> bool:
        """保存配置到文件"""
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            print(f"配置已保存: {self.CONFIG_FILE}")
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False

    def _merge_config(self, default: Dict, user: Dict) -> Dict:
        """
        递归合并配置
        user配置优先，但保留default中user没有的配置项
        """
        result = default.copy()
        for key, value in user.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_config(result[key], value)
            else:
                result[key] = value
        return result

    # Excel配置相关方法

    def get_header_rows(self) -> int:
        """获取表头行数"""
        return self.config['excel']['header_rows']

    def set_header_rows(self, rows: int):
        """设置表头行数"""
        if rows < 0:
            raise ValueError("表头行数不能为负数")
        self.config['excel']['header_rows'] = rows

    def get_excel_columns(self) -> Dict[str, int]:
        """获取Excel列配置"""
        return self.config['excel']['columns'].copy()

    def set_excel_column(self, field: str, column_index: int):
        """设置单个字段的列索引"""
        self.config['excel']['columns'][field] = column_index

    def set_excel_columns(self, columns: Dict[str, int]):
        """设置所有Excel列配置"""
        self.config['excel']['columns'] = columns.copy()

    # 语音识别配置相关方法

    def get_whisper_model(self) -> str:
        """获取Whisper模型名称"""
        return self.config['speech_recognition']['model']

    def set_whisper_model(self, model: str):
        """设置Whisper模型"""
        if model not in self.AVAILABLE_MODELS:
            raise ValueError(f"不支持的模型: {model}")
        self.config['speech_recognition']['model'] = model

    def get_whisper_device(self) -> str:
        """获取Whisper设备"""
        return self.config['speech_recognition']['device']

    def set_whisper_device(self, device: str):
        """设置Whisper设备"""
        self.config['speech_recognition']['device'] = device

    def get_whisper_compute_type(self) -> str:
        """获取Whisper计算类型"""
        return self.config['speech_recognition']['compute_type']

    def set_whisper_compute_type(self, compute_type: str):
        """设置Whisper计算类型"""
        self.config['speech_recognition']['compute_type'] = compute_type

    # 录音配置相关方法

    def get_record_duration(self) -> int:
        """获取录音时长"""
        return self.config['recording']['duration']

    def set_record_duration(self, duration: int):
        """设置录音时长"""
        if duration <= 0:
            raise ValueError("录音时长必须大于0")
        self.config['recording']['duration'] = duration

    def get_sample_rate(self) -> int:
        """获取采样率"""
        return self.config['recording']['sample_rate']

    def set_sample_rate(self, rate: int):
        """设置采样率"""
        if rate <= 0:
            raise ValueError("采样率必须大于0")
        self.config['recording']['sample_rate'] = rate

    # TTS配置相关方法

    def get_tts_voice(self) -> str:
        """获取TTS语音"""
        return self.config['tts']['voice']

    def set_tts_voice(self, voice: str):
        """设置TTS语音"""
        self.config['tts']['voice'] = voice

    # 重置配置

    def reset_to_default(self):
        """重置为默认配置"""
        self.config = self.DEFAULT_CONFIG.copy()

    def reset_excel_config(self):
        """重置Excel配置"""
        self.config['excel'] = self.DEFAULT_CONFIG['excel'].copy()

    def reset_speech_config(self):
        """重置语音识别配置"""
        self.config['speech_recognition'] = self.DEFAULT_CONFIG['speech_recognition'].copy()

    # 获取完整配置

    def get_full_config(self) -> Dict[str, Any]:
        """获取完整配置（副本）"""
        return self.config.copy()
