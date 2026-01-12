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
    
    def _play_audio(self, audio_file: str, text: str = ""):
        """
        播放音频文件
        audio_file: 音频文件路径
        text: 原始文本（用于估算播放时长）
        """
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

            # 方案2: Windows系统 - 使用winsound（静默播放，不弹窗）
            if os.name == 'nt':  # Windows
                try:
                    # 首先尝试使用winsound（Python自带，不会弹窗）
                    # 但winsound只支持WAV格式，我们的是MP3，所以需要转换或使用其他方法
                    # 使用PowerShell静默播放（不会弹出窗口）
                    import time
                    # 估算播放时长
                    estimated_duration = min(max(len(text) * 0.15, 2), 10) if text else 3

                    # 使用PowerShell MediaPlayer播放（静默，不弹窗）
                    ps_command = f'''
                    Add-Type -AssemblyName presentationCore
                    $player = New-Object System.Windows.Media.MediaPlayer
                    $player.Open([uri]"{audio_path}")
                    $player.Play()
                    Start-Sleep -Seconds {int(estimated_duration)}
                    $player.Stop()
                    $player.Close()
                    '''

                    try:
                        subprocess.run(
                            ['powershell', '-Command', ps_command],
                            capture_output=True,
                            timeout=estimated_duration + 2,
                            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                        )
                        return
                    except Exception as ps_error:
                        print(f"PowerShell播放失败: {ps_error}")
                        # 如果PowerShell失败，使用startfile（会弹窗，但至少能播放）
                        os.startfile(audio_path)
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
                self._play_audio(audio_file, text)
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
                    self._play_audio(audio_file, text)
                loop.close()
            except Exception as e:
                print(f"异步语音播报失败: {e}")

        thread = threading.Thread(target=_async_speak, daemon=True)
        thread.start()
