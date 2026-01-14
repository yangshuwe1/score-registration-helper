"""
配置管理模块
负责保存和加载用户自定义配置
"""
import json
import os
from pathlib import Path
from typing import Dict, Any


class ConfigManager:
    """配置管理器，支持配置的保存和加载"""

    DEFAULT_CONFIG = {
        # 模型配置
        'whisper_model': 'medium',  # tiny, small, base, medium, large
        'whisper_device': 'cpu',     # cpu, cuda
        'whisper_compute_type': 'int8',  # int8, float16, float32

        # 录音配置
        'record_duration': 5,        # 最大录音时长（秒）
        'silence_duration': 1.5,     # 静音检测时长（秒）
        'min_speech_duration': 0.5,  # 最小说话时长（秒）

        # TTS配置
        'tts_voice': 'zh-CN-XiaoxiaoNeural',  # 中文语音
        'tts_rate': 0,  # 语速调整 (-100 到 100)

        # Excel列配置
        'current_column': 'final_score',  # 默认输入列

        # 界面配置
        'window_width': 900,
        'window_height': 700,
    }

    def __init__(self, config_file: str = None):
        """
        初始化配置管理器
        config_file: 配置文件路径，默认为用户目录下的 .score_helper_config.json
        """
        if config_file is None:
            # 使用用户目录下的配置文件
            config_file = Path.home() / '.score_helper_config.json'

        self.config_file = Path(config_file)
        self.config = self.DEFAULT_CONFIG.copy()
        self.load_config()

    def load_config(self) -> bool:
        """
        从文件加载配置
        返回: True表示成功加载，False表示使用默认配置
        """
        if not self.config_file.exists():
            return False

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)

            # 合并配置（保留默认值中没有的新配置项）
            self.config.update(loaded_config)
            print(f"已加载配置: {self.config_file}")
            return True
        except Exception as e:
            print(f"加载配置失败: {e}")
            return False

    def save_config(self) -> bool:
        """
        保存配置到文件
        返回: True表示成功，False表示失败
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            print(f"配置已保存: {self.config_file}")
            return True
        except Exception as e:
            print(f"保存配置失败: {e}")
            return False

    def get(self, key: str, default: Any = None) -> Any:
        """获取配置项"""
        return self.config.get(key, default)

    def set(self, key: str, value: Any):
        """设置配置项"""
        self.config[key] = value

    def update(self, config_dict: Dict[str, Any]):
        """批量更新配置"""
        self.config.update(config_dict)

    def reset_to_default(self):
        """重置为默认配置"""
        self.config = self.DEFAULT_CONFIG.copy()

    def get_all(self) -> Dict[str, Any]:
        """获取所有配置"""
        return self.config.copy()
