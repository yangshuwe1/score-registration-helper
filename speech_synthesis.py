"""
语音合成模块
使用edge-tts实现中文语音播报
"""
import edge_tts
import asyncio
import os
import threading
import subprocess
from typing import Optional
from config import TTS_VOICE

try:
    import pygame
    PYGAME_AVAILABLE = True
except ImportError:
    PYGAME_AVAILABLE = False


class SpeechSynthesis:
    def __init__(self):
        self.voice = TTS_VOICE
        self._loop = None
        self.temp_audio_file = "temp_audio.mp3"
        if PYGAME_AVAILABLE:
            try:
                pygame.mixer.init()
            except:
                pass
    
    def _get_loop(self):
        """获取或创建事件循环"""
        try:
            loop = asyncio.get_event_loop()
            if loop.is_closed():
                raise RuntimeError("Event loop is closed")
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        return loop
    
    async def _generate_audio(self, text: str) -> str:
        """生成音频文件"""
        try:
            communicate = edge_tts.Communicate(text, self.voice)
            await communicate.save(self.temp_audio_file)
            return self.temp_audio_file
        except Exception as e:
            print(f"语音合成失败: {e}")
            return None
    
    def _play_audio(self, audio_file: str):
        """播放音频文件"""
        try:
            audio_path = os.path.abspath(audio_file)
            
            # 方案1: 使用pygame（如果可用）
            if PYGAME_AVAILABLE:
                try:
                    pygame.mixer.music.load(audio_path)
                    pygame.mixer.music.play()
                    # 等待播放完成
                    while pygame.mixer.music.get_busy():
                        pygame.time.Clock().tick(10)
                    return
                except Exception as e:
                    print(f"pygame播放失败，尝试备用方案: {e}")
            
            # 方案2: Windows系统默认播放器（主要方案）
            if os.name == 'nt':  # Windows
                try:
                    # 直接打开文件，使用系统默认播放器
                    os.startfile(audio_path)
                    # 等待一段时间让文件播放（估算音频长度）
                    import time
                    # 简单估算：假设每字符0.1秒，最少2秒，最多10秒
                    estimated_duration = min(max(len(text) * 0.1, 2), 10)
                    time.sleep(estimated_duration)
                except Exception as e:
                    print(f"Windows播放失败: {e}")
            else:
                # Linux/Mac备用方案
                try:
                    subprocess.run(['mpg123', audio_path], check=True, timeout=30)
                except:
                    try:
                        subprocess.run(['ffplay', '-autoexit', '-nodisp', audio_path], 
                                     check=True, timeout=30)
                    except:
                        print("无法播放音频，请安装mpg123或ffmpeg")
        except Exception as e:
            print(f"播放音频失败: {e}")
        finally:
            # 清理临时文件（延迟清理，确保播放完成）
            def cleanup():
                import time
                time.sleep(1)  # 等待1秒确保播放开始
                try:
                    if os.path.exists(audio_file):
                        os.remove(audio_file)
                except:
                    pass
            
            # 在后台线程中清理
            cleanup_thread = threading.Thread(target=cleanup, daemon=True)
            cleanup_thread.start()
    
    def speak(self, text: str):
        """
        同步语音播报
        """
        if not text:
            return
        
        try:
            loop = self._get_loop()
            audio_file = loop.run_until_complete(self._generate_audio(text))
            if audio_file:
                self._play_audio(audio_file)
        except Exception as e:
            print(f"语音播报失败: {e}")
    
    def speak_async(self, text: str):
        """
        异步语音播报（不阻塞）
        """
        if not text:
            return
        
        def _async_speak():
            try:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                audio_file = loop.run_until_complete(self._generate_audio(text))
                if audio_file:
                    self._play_audio(audio_file)
                loop.close()
            except Exception as e:
                print(f"异步语音播报失败: {e}")
        
        thread = threading.Thread(target=_async_speak, daemon=True)
        thread.start()
