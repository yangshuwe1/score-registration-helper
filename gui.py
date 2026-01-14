"""
GUI界面模块
使用tkinter构建图形界面
"""
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from typing import Optional
from excel_handler import ExcelHandler
from speech_recognition import SpeechRecognition
from speech_synthesis import SpeechSynthesis
from student_parser import StudentParser
from config import RECORD_DURATION
from config_manager import ConfigManager


class GradeEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("登分助手")

        # 初始化配置管理器
        self.config_manager = ConfigManager()

        # 从配置中读取窗口大小
        window_width = self.config_manager.get('window_width', 900)
        window_height = self.config_manager.get('window_height', 700)
        self.root.geometry(f"{window_width}x{window_height}")
        self.root.resizable(True, True)

        # 创建界面（先创建，再加载模型）
        self._create_widgets()

        # 显示加载提示
        self.status_label.config(
            text="正在初始化...（首次运行需要下载模型，请耐心等待）",
            foreground="blue"
        )
        self.log("=" * 50)
        self.log("程序启动中...")
        self.log("提示: 首次运行需要下载语音识别模型（约150MB）")
        self.log("     如果网络较慢，可能需要几分钟时间")
        self.log("     程序正在后台加载，请勿关闭窗口")
        self.log("=" * 50)

        # 在后台线程中初始化模块（避免阻塞界面）
        init_thread = threading.Thread(target=self._initialize_modules, daemon=True)
        init_thread.start()

        # 状态变量
        self.current_column = self.config_manager.get('current_column', 'final_score')
        self.is_recording = False
        self.last_operation = None  # 上一步操作缓存：{'row': int, 'column': str, 'old_score': float, 'new_score': float, 'name': str}
    
    def _initialize_modules(self):
        """在后台线程中初始化模块（带详细进度提示）"""
        import time
        
        try:
            # 步骤1: 初始化Excel处理器（快速）
            self.root.after(0, lambda: self.status_label.config(
                text="正在初始化Excel处理器...", foreground="blue"
            ))
            self.root.after(0, lambda: self.log("步骤 1/4: 初始化Excel处理器..."))
            self.excel_handler = ExcelHandler()
            time.sleep(0.1)  # 短暂延迟，让界面更新
            
            # 步骤2: 加载语音识别模型（可能较慢）
            self.root.after(0, lambda: self.status_label.config(
                text="正在加载语音识别模型...（首次运行需要下载，请耐心等待）", foreground="blue"
            ))
            self.root.after(0, lambda: self.log("步骤 2/4: 正在加载语音识别模型..."))
            self.root.after(0, lambda: self.log("       提示: 首次运行需要下载模型（约150MB）"))
            self.root.after(0, lambda: self.log("       如果网络较慢，可能需要几分钟"))
            self.root.after(0, lambda: self.log("       程序正在运行，请勿关闭..."))
            
            # 初始化语音识别（这里可能会卡住，但会在后台线程中）
            self.speech_recognition = SpeechRecognition()
            
            # 检查模型是否加载成功
            if self.speech_recognition.model is None:
                self.root.after(0, lambda: self.status_label.config(
                    text="模型加载失败，请检查网络连接", foreground="red"
                ))
                self.root.after(0, lambda: self.log("[错误] 语音识别模型加载失败"))
                self.root.after(0, lambda: self.log("   可能原因: 网络连接问题或磁盘空间不足"))
                return
            
            self.root.after(0, lambda: self.log("[成功] 语音识别模型加载成功"))
            time.sleep(0.1)
            
            # 步骤3: 初始化语音合成
            self.root.after(0, lambda: self.status_label.config(
                text="正在初始化语音合成...", foreground="blue"
            ))
            self.root.after(0, lambda: self.log("步骤 3/4: 正在初始化语音合成..."))
            self.speech_synthesis = SpeechSynthesis()
            time.sleep(0.1)
            
            # 步骤4: 初始化学生解析器
            self.root.after(0, lambda: self.status_label.config(
                text="正在初始化解析器...", foreground="blue"
            ))
            self.root.after(0, lambda: self.log("步骤 4/4: 正在初始化解析器..."))
            self.student_parser = StudentParser()
            time.sleep(0.1)
            
            # 初始化完成
            self.root.after(0, lambda: self.status_label.config(
                text="[就绪] 准备就绪，可以开始使用", foreground="green"
            ))
            self.root.after(0, lambda: self.record_button.config(state="normal"))
            self.root.after(0, lambda: self.log("=" * 50))
            self.root.after(0, lambda: self.log("[成功] 程序初始化完成！"))
            self.root.after(0, lambda: self.log("提示: 请先选择Excel文件"))
            self.root.after(0, lambda: self.log("=" * 50))
            
        except MemoryError as e:
            error_msg = "内存不足，模型加载失败"
            self.root.after(0, lambda: self.status_label.config(
                text=error_msg, foreground="red"
            ))
            self.root.after(0, lambda: self.log(f"[错误] {error_msg}"))
            self.root.after(0, lambda: self.log("   建议: 关闭其他程序，释放内存后重试"))
        except Exception as e:
            import traceback
            error_msg = str(e)
            self.root.after(0, lambda: self.status_label.config(
                text=f"初始化失败: {error_msg[:50]}...", foreground="red"
            ))
            self.root.after(0, lambda: self.log(f"[错误] {error_msg}"))
            self.root.after(0, lambda: self.log(f"详细错误信息已记录"))
            # 只在调试时显示详细错误
            if __debug__:
                self.root.after(0, lambda: self.log(f"详细: {traceback.format_exc()}"))
    
    def _create_widgets(self):
        """创建界面组件（使用标签页）"""
        # 创建Notebook（标签页）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # 创建主界面标签页
        main_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(main_frame, text="主界面")
        main_frame.columnconfigure(1, weight=1)
        
        # 1. 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="Excel文件", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Button(file_frame, text="选择Excel文件", command=self._select_file).grid(row=0, column=0, padx=5)
        self.file_label = ttk.Label(file_frame, text="未选择文件", foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # 2. 列选择区域
        column_frame = ttk.LabelFrame(main_frame, text="选择输入列", padding="10")
        column_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.column_var = tk.StringVar(value="final_score")
        ttk.Radiobutton(
            column_frame, text="期末成绩", variable=self.column_var,
            value="final_score", command=self._on_column_change
        ).grid(row=0, column=0, padx=10)
        ttk.Radiobutton(
            column_frame, text="平时成绩", variable=self.column_var,
            value="regular_score", command=self._on_column_change
        ).grid(row=0, column=1, padx=10)
        
        # 3. 语音输入区域
        speech_frame = ttk.LabelFrame(main_frame, text="语音输入", padding="10")
        speech_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.record_button = ttk.Button(
            speech_frame, text="开始录音", command=self._toggle_recording,
            state="disabled"
        )
        self.record_button.grid(row=0, column=0, padx=5)
        
        self.status_label = ttk.Label(speech_frame, text="请先选择Excel文件", foreground="gray")
        self.status_label.grid(row=0, column=1, padx=10)
        
        # 4. 识别结果显示
        result_frame = ttk.LabelFrame(main_frame, text="识别结果", padding="10")
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        result_frame.columnconfigure(0, weight=1)
        
        self.result_text = tk.Text(result_frame, height=3, wrap=tk.WORD)
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        result_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_text.yview)
        result_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.result_text.configure(yscrollcommand=result_scrollbar.set)
        
        # 5. 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="操作日志", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

        # 创建设置标签页
        self._create_settings_tab()

    def _create_settings_tab(self):
        """创建设置标签页"""
        settings_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(settings_frame, text="设置")

        # 1. 模型配置区域
        model_frame = ttk.LabelFrame(settings_frame, text="语音识别模型配置", padding="10")
        model_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)
        settings_frame.columnconfigure(0, weight=1)

        # 模型版本选择
        ttk.Label(model_frame, text="模型版本:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.model_var = tk.StringVar(value=self.config_manager.get('whisper_model', 'medium'))
        model_combo = ttk.Combobox(
            model_frame,
            textvariable=self.model_var,
            values=['tiny', 'small', 'base', 'medium', 'large'],
            state='readonly',
            width=15
        )
        model_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(
            model_frame,
            text="(tiny最快但不太准确, large最准确但较慢)",
            foreground="gray"
        ).grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)

        # 设备选择
        ttk.Label(model_frame, text="计算设备:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.device_var = tk.StringVar(value=self.config_manager.get('whisper_device', 'cpu'))
        device_combo = ttk.Combobox(
            model_frame,
            textvariable=self.device_var,
            values=['cpu', 'cuda'],
            state='readonly',
            width=15
        )
        device_combo.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(
            model_frame,
            text="(cuda需要NVIDIA显卡支持)",
            foreground="gray"
        ).grid(row=1, column=2, sticky=tk.W, pady=5, padx=5)

        # 计算类型
        ttk.Label(model_frame, text="计算精度:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.compute_type_var = tk.StringVar(value=self.config_manager.get('whisper_compute_type', 'int8'))
        compute_type_combo = ttk.Combobox(
            model_frame,
            textvariable=self.compute_type_var,
            values=['int8', 'float16', 'float32'],
            state='readonly',
            width=15
        )
        compute_type_combo.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Label(
            model_frame,
            text="(int8速度快内存少, float32精度高但慢)",
            foreground="gray"
        ).grid(row=2, column=2, sticky=tk.W, pady=5, padx=5)

        # 2. 录音配置区域
        record_frame = ttk.LabelFrame(settings_frame, text="录音配置", padding="10")
        record_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)

        # 最大录音时长
        ttk.Label(record_frame, text="最大录音时长(秒):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.duration_var = tk.DoubleVar(value=self.config_manager.get('record_duration', 5))
        duration_scale = ttk.Scale(
            record_frame,
            from_=1,
            to=10,
            variable=self.duration_var,
            orient=tk.HORIZONTAL,
            length=200
        )
        duration_scale.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        self.duration_label = ttk.Label(record_frame, text=f"{self.duration_var.get():.1f}秒")
        self.duration_label.grid(row=0, column=2, sticky=tk.W, pady=5, padx=5)
        duration_scale.config(command=lambda v: self.duration_label.config(text=f"{float(v):.1f}秒"))

        # 静音检测时长
        ttk.Label(record_frame, text="静音检测时长(秒):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.silence_var = tk.DoubleVar(value=self.config_manager.get('silence_duration', 1.5))
        silence_scale = ttk.Scale(
            record_frame,
            from_=0.5,
            to=3.0,
            variable=self.silence_var,
            orient=tk.HORIZONTAL,
            length=200
        )
        silence_scale.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        self.silence_label = ttk.Label(record_frame, text=f"{self.silence_var.get():.1f}秒")
        self.silence_label.grid(row=1, column=2, sticky=tk.W, pady=5, padx=5)
        silence_scale.config(command=lambda v: self.silence_label.config(text=f"{float(v):.1f}秒"))

        # 最小说话时长
        ttk.Label(record_frame, text="最小说话时长(秒):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.min_speech_var = tk.DoubleVar(value=self.config_manager.get('min_speech_duration', 0.5))
        min_speech_scale = ttk.Scale(
            record_frame,
            from_=0.1,
            to=2.0,
            variable=self.min_speech_var,
            orient=tk.HORIZONTAL,
            length=200
        )
        min_speech_scale.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        self.min_speech_label = ttk.Label(record_frame, text=f"{self.min_speech_var.get():.1f}秒")
        self.min_speech_label.grid(row=2, column=2, sticky=tk.W, pady=5, padx=5)
        min_speech_scale.config(command=lambda v: self.min_speech_label.config(text=f"{float(v):.1f}秒"))

        # 3. TTS配置区域
        tts_frame = ttk.LabelFrame(settings_frame, text="语音合成配置", padding="10")
        tts_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)

        ttk.Label(tts_frame, text="语音选择:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.tts_voice_var = tk.StringVar(value=self.config_manager.get('tts_voice', 'zh-CN-XiaoxiaoNeural'))
        tts_voice_combo = ttk.Combobox(
            tts_frame,
            textvariable=self.tts_voice_var,
            values=[
                'zh-CN-XiaoxiaoNeural',  # 晓晓（女声，普通话）
                'zh-CN-YunxiNeural',     # 云希（男声，普通话）
                'zh-CN-YunyangNeural',   # 云扬（男声，普通话）
                'zh-CN-XiaoyiNeural',    # 晓伊（女声，普通话）
                'zh-CN-YunjianNeural',   # 云健（男声，普通话）
            ],
            state='readonly',
            width=25
        )
        tts_voice_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)

        ttk.Label(tts_frame, text="语速调整:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.tts_rate_var = tk.IntVar(value=self.config_manager.get('tts_rate', 0))
        tts_rate_scale = ttk.Scale(
            tts_frame,
            from_=-50,
            to=50,
            variable=self.tts_rate_var,
            orient=tk.HORIZONTAL,
            length=200
        )
        tts_rate_scale.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        self.tts_rate_label = ttk.Label(tts_frame, text=f"{self.tts_rate_var.get():+d}%")
        self.tts_rate_label.grid(row=1, column=2, sticky=tk.W, pady=5, padx=5)
        tts_rate_scale.config(command=lambda v: self.tts_rate_label.config(text=f"{int(float(v)):+d}%"))

        # 4. 按钮区域
        button_frame = ttk.Frame(settings_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=20, padx=5)

        ttk.Button(
            button_frame,
            text="保存配置",
            command=self._save_settings
        ).grid(row=0, column=0, padx=5)

        ttk.Button(
            button_frame,
            text="重置为默认",
            command=self._reset_settings
        ).grid(row=0, column=1, padx=5)

        ttk.Button(
            button_frame,
            text="应用并重新加载模型",
            command=self._apply_and_reload_model
        ).grid(row=0, column=2, padx=5)

        # 5. 说明区域
        info_frame = ttk.LabelFrame(settings_frame, text="说明", padding="10")
        info_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5, padx=5)
        settings_frame.rowconfigure(4, weight=1)

        info_text = scrolledtext.ScrolledText(info_frame, height=8, wrap=tk.WORD)
        info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)

        info_content = """配置说明：

1. 模型版本：
   - tiny: 最快，约39MB，适合配置较低的电脑
   - small: 较快，约244MB，准确度一般
   - base: 平衡，约147MB，速度和准确度平衡
   - medium: 推荐，约1.5GB，准确度高（默认）
   - large: 最准确，约3GB，速度较慢

2. 计算设备：
   - cpu: 使用CPU计算，兼容性好（推荐）
   - cuda: 使用NVIDIA显卡加速，需要安装CUDA

3. 录音配置：
   - 最大录音时长: 单次录音的最长时间
   - 静音检测时长: 静音多久后认为说话结束
   - 最小说话时长: 低于此时长的录音将被忽略

4. 注意：更改模型配置后，需要点击"应用并重新加载模型"才能生效。
"""
        info_text.insert(tk.END, info_content)
        info_text.config(state=tk.DISABLED)

    def _save_settings(self):
        """保存设置"""
        try:
            # 更新配置
            self.config_manager.set('whisper_model', self.model_var.get())
            self.config_manager.set('whisper_device', self.device_var.get())
            self.config_manager.set('whisper_compute_type', self.compute_type_var.get())
            self.config_manager.set('record_duration', self.duration_var.get())
            self.config_manager.set('silence_duration', self.silence_var.get())
            self.config_manager.set('min_speech_duration', self.min_speech_var.get())
            self.config_manager.set('tts_voice', self.tts_voice_var.get())
            self.config_manager.set('tts_rate', self.tts_rate_var.get())
            self.config_manager.set('current_column', self.current_column)

            # 保存到文件
            if self.config_manager.save_config():
                messagebox.showinfo("成功", "配置已保存！\n\n注意：模型配置需要重新加载模型才能生效。")
                self.log("配置已保存")
            else:
                messagebox.showerror("错误", "保存配置失败")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")
            self.log(f"保存配置失败: {e}")

    def _reset_settings(self):
        """重置为默认设置"""
        if messagebox.askyesno("确认", "确定要重置为默认设置吗？"):
            self.config_manager.reset_to_default()
            # 更新界面显示
            self.model_var.set(self.config_manager.get('whisper_model'))
            self.device_var.set(self.config_manager.get('whisper_device'))
            self.compute_type_var.set(self.config_manager.get('whisper_compute_type'))
            self.duration_var.set(self.config_manager.get('record_duration'))
            self.silence_var.set(self.config_manager.get('silence_duration'))
            self.min_speech_var.set(self.config_manager.get('min_speech_duration'))
            self.tts_voice_var.set(self.config_manager.get('tts_voice'))
            self.tts_rate_var.set(self.config_manager.get('tts_rate'))
            messagebox.showinfo("成功", "已重置为默认设置")
            self.log("配置已重置为默认")

    def _apply_and_reload_model(self):
        """应用设置并重新加载模型"""
        if messagebox.askyesno(
            "确认",
            "重新加载模型可能需要几分钟时间。\n确定要继续吗？"
        ):
            # 先保存配置
            self._save_settings()

            # 禁用录音按钮
            self.record_button.config(state="disabled")
            self.status_label.config(text="正在重新加载模型，请稍候...", foreground="blue")
            self.log("开始重新加载模型...")

            # 在后台线程中重新加载
            reload_thread = threading.Thread(target=self._reload_model, daemon=True)
            reload_thread.start()

    def _reload_model(self):
        """重新加载语音识别模型"""
        try:
            # 重新初始化SpeechRecognition
            self.root.after(0, lambda: self.log("正在加载新的语音识别模型..."))
            self.speech_recognition = SpeechRecognition(
                model=self.config_manager.get('whisper_model'),
                device=self.config_manager.get('whisper_device'),
                compute_type=self.config_manager.get('whisper_compute_type')
            )

            # 重新初始化SpeechSynthesis（应用新的语音设置）
            self.root.after(0, lambda: self.log("正在更新语音合成设置..."))
            self.speech_synthesis = SpeechSynthesis(
                voice=self.config_manager.get('tts_voice'),
                rate=self.config_manager.get('tts_rate')
            )

            # 完成
            self.root.after(0, lambda: self.status_label.config(
                text="[就绪] 模型已重新加载", foreground="green"
            ))
            self.root.after(0, lambda: self.record_button.config(state="normal"))
            self.root.after(0, lambda: self.log("模型重新加载完成！"))
        except Exception as e:
            import traceback
            error_msg = str(e)
            self.root.after(0, lambda: self.status_label.config(
                text=f"加载失败: {error_msg[:50]}...", foreground="red"
            ))
            self.root.after(0, lambda: self.log(f"重新加载模型失败: {error_msg}"))
            self.root.after(0, lambda: self.log(f"详细: {traceback.format_exc()}"))

    def _select_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            if self.excel_handler.load_excel(file_path):
                self.file_label.config(text=file_path, foreground="black")
                self.record_button.config(state="normal")
                self.status_label.config(text="准备就绪，可以开始录音", foreground="green")
                self.log(f"已加载文件: {file_path}")
                self.log(f"学生总数: {self.excel_handler.get_total_students()}")
            else:
                messagebox.showerror("错误", "无法加载Excel文件，请检查文件格式")
                self.log("加载文件失败")
    
    def _on_column_change(self):
        """列选择改变"""
        self.current_column = self.column_var.get()
        column_name = "期末成绩" if self.current_column == "final_score" else "平时成绩"
        self.log(f"已切换到: {column_name}")
    
    def _toggle_recording(self):
        """切换录音状态"""
        if not self.is_recording:
            self._start_recording()
        else:
            self._stop_recording()
    
    def _start_recording(self):
        """开始实时录音（说完自动识别）"""
        if self.excel_handler.file_path is None:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
        
        self.is_recording = True
        self.record_button.config(text="停止录音", state="normal")
        self.status_label.config(text="正在监听...（说完后自动识别）", foreground="green")
        self.result_text.delete(1.0, tk.END)
        
        # 在后台线程中启动实时录音
        thread = threading.Thread(target=self._realtime_record_and_process, daemon=True)
        thread.start()
    
    def _stop_recording(self):
        """停止录音"""
        self.is_recording = False
        self.speech_recognition.stop_recording()
        self.record_button.config(text="开始录音")
        self.status_label.config(text="已停止录音", foreground="gray")
    
    def _realtime_record_and_process(self):
        """实时录音并处理（说完自动识别）"""
        def on_speech_end(text: str):
            """说话结束后的回调"""
            if not text:
                self.root.after(0, lambda: self.status_label.config(
                    text="识别失败，继续监听...", foreground="orange"
                ))
                self.root.after(0, lambda: self.log("语音识别失败，请检查录音质量"))
                # 语音提醒
                self.speech_synthesis.speak_async("识别失败，请重试")
                # 不要设置 is_recording = False，继续监听
                return
            
            # 显示识别结果
            self.root.after(0, lambda: self.result_text.insert(tk.END, text))
            self.root.after(0, lambda: self.log(f"识别结果: {text}"))
            self.root.after(0, lambda: self.status_label.config(
                text="正在处理...", foreground="blue"
            ))

            # 检查是否是撤回命令
            if self.student_parser.is_undo_command(text):
                self._handle_undo()
                return

            # 解析结果（支持多个学生）
            parsed_list = self.student_parser.parse_multiple(text)

            # 调试：显示解析后的结果
            if parsed_list:
                for p in parsed_list:
                    parse_info = f"解析: {p['type']}={p['identifier']}, 分数={p['score']}"
                    self.root.after(0, lambda info=parse_info: self.log(info))

            if not parsed_list:
                self.root.after(0, lambda: self.status_label.config(
                    text="解析失败，继续监听...请说：姓名/序号，分数", foreground="orange"
                ))
                self.root.after(0, lambda: self.log(f"解析失败，识别文本: {text}"))
                # 语音提醒
                self.speech_synthesis.speak_async("解析失败，请说序号或姓名加分数")
                # 不要设置 is_recording = False，继续监听
                return

            # 处理所有解析出的学生信息
            success_count = 0
            confirmations = []

            for parsed in parsed_list:
                # 查找学生
                if parsed['type'] == 'sequence':
                    identifier = parsed['identifier']
                    self.root.after(0, lambda i=identifier: self.log(
                        f"查找序号: {i}"
                    ))
                    row = self.excel_handler.find_student_by_sequence(parsed['identifier'])
                else:
                    identifier = parsed['identifier']
                    self.root.after(0, lambda i=identifier: self.log(
                        f"查找姓名: {i}"
                    ))
                    row = self.excel_handler.find_student_by_name(parsed['identifier'])

                if not row:
                    identifier = parsed['identifier']
                    parse_type = parsed['type']
                    self.root.after(0, lambda i=identifier, t=parse_type: self.log(
                        f"未找到学生: {i} (类型: {t})，跳过"
                    ))
                    continue

                # 获取学生信息
                student_info = self.excel_handler.get_student_info(row)
                if not student_info:
                    self.root.after(0, lambda: self.log("获取学生信息失败，跳过"))
                    continue

                # 获取旧成绩（用于撤回）
                old_score = self.excel_handler.get_score(row, self.current_column)

                # 更新成绩
                if not self.excel_handler.update_score(row, self.current_column, parsed['score']):
                    score = parsed['score']
                    self.root.after(0, lambda s=score: self.log(f"更新成绩失败，分数: {s}，跳过"))
                    continue

                # 保存上一步操作（用于撤回）
                self.last_operation = {
                    'row': row,
                    'column': self.current_column,
                    'old_score': old_score,
                    'new_score': parsed['score'],
                    'name': student_info['name']
                }

                # 生成确认文本
                confirmation = self.student_parser.format_confirmation(
                    row, student_info['name'], parsed['score']
                )
                confirmations.append(confirmation)
                success_count += 1

            # 保存文件（一次性保存所有更改）
            if success_count > 0:
                if not self.excel_handler.save_excel():
                    self.root.after(0, lambda: self.status_label.config(
                        text="保存失败，文件可能被占用", foreground="red"
                    ))
                    self.root.after(0, lambda: messagebox.showwarning(
                        "保存失败", "请关闭Excel文件后重试"
                    ))
                    self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                    self.is_recording = False
                    return

                # 显示成功信息
                self.root.after(0, lambda: self.status_label.config(
                    text=f"成功录入{success_count}条，继续监听...", foreground="green"
                ))

                # 记录日志
                for confirmation in confirmations:
                    c = confirmation  # 避免闭包问题
                    self.root.after(0, lambda conf=c: self.log(f"已更新: {conf}"))

                # 语音播报（播报所有成功的记录）
                full_confirmation = "，".join(confirmations)
                self.speech_synthesis.speak_async(full_confirmation)
            else:
                self.root.after(0, lambda: self.status_label.config(
                    text="未成功录入任何成绩", foreground="red"
                ))
                self.root.after(0, lambda: self.log("未成功录入任何成绩，请检查输入"))
                # 语音提醒
                self.speech_synthesis.speak_async("未找到学生，请重试")
            
            # 继续监听（不重置按钮，保持录音状态）
            # 用户可以继续说下一个，或点击停止

        try:
            # 循环录音，直到用户点击停止
            while self.is_recording:
                # 启动实时录音（使用配置中的参数）
                silence_duration = self.config_manager.get('silence_duration', 1.5)
                min_speech_duration = self.config_manager.get('min_speech_duration', 0.5)

                success = self.speech_recognition.record_audio_realtime(
                    on_speech_end=on_speech_end,
                    silence_duration=silence_duration,
                    min_speech_duration=min_speech_duration
                )

                if not success and self.is_recording:
                    self.root.after(0, lambda: self.status_label.config(
                        text="录音失败，请检查麦克风", foreground="red"
                    ))
                    self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                    self.is_recording = False
                    break

                # 如果还在录音状态，继续下一次循环（自动开始下一次录音）
                if self.is_recording:
                    self.root.after(0, lambda: self.status_label.config(
                        text="继续监听...（说完后自动识别）", foreground="green"
                    ))

            # 循环结束，重置按钮
            self.root.after(0, lambda: self.record_button.config(text="开始录音"))
        except Exception as e:
            import traceback
            error_msg = str(e)
            self.root.after(0, lambda: self.status_label.config(
                text=f"处理失败: {error_msg}", foreground="red"
            ))
            self.root.after(0, lambda: self.log(f"错误: {error_msg}"))
            self.root.after(0, lambda: self.log(f"详细错误: {traceback.format_exc()}"))
            self.root.after(0, lambda: self.record_button.config(text="开始录音"))
            self.is_recording = False
    
    def _record_and_process(self):
        """录音并处理（传统方法，保留作为备用）"""
        try:
            # 录音
            audio_file = self.speech_recognition.record_audio(RECORD_DURATION)
            if not audio_file:
                self.root.after(0, lambda: self.status_label.config(
                    text="录音失败，请检查麦克风", foreground="red"
                ))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return
            
            # 识别
            self.root.after(0, lambda: self.status_label.config(
                text="正在识别...", foreground="blue"
            ))
            
            text = self.speech_recognition.transcribe(audio_file)
            
            if not text:
                self.root.after(0, lambda: self.status_label.config(
                    text="识别失败，请重试", foreground="red"
                ))
                self.root.after(0, lambda: self.log("语音识别失败，请检查录音质量"))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return
            
            # 显示识别结果
            self.root.after(0, lambda: self.result_text.insert(tk.END, text))
            self.root.after(0, lambda: self.log(f"识别结果: {text}"))
            
            # 解析结果
            parsed = self.student_parser.parse(text)
            if not parsed:
                self.root.after(0, lambda: self.status_label.config(
                    text="解析失败，请说：姓名/序号，分数", foreground="red"
                ))
                self.root.after(0, lambda: self.log(f"解析失败，识别文本: {text}"))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return

            # 查找学生
            if parsed['type'] == 'sequence':
                row = self.excel_handler.find_student_by_sequence(parsed['identifier'])
            else:
                row = self.excel_handler.find_student_by_name(parsed['identifier'])
            
            if not row:
                self.root.after(0, lambda: self.status_label.config(
                    text=f"未找到学生: {parsed['identifier']}", foreground="red"
                ))
                self.root.after(0, lambda: self.log(
                    f"未找到学生: {parsed['identifier']}，请检查Excel文件"
                ))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return
            
            # 获取学生信息
            student_info = self.excel_handler.get_student_info(row)
            if not student_info:
                self.root.after(0, lambda: self.status_label.config(
                    text="获取学生信息失败", foreground="red"
                ))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return
            
            # 更新成绩
            if not self.excel_handler.update_score(row, self.current_column, parsed['score']):
                self.root.after(0, lambda: self.status_label.config(
                    text=f"更新成绩失败，分数: {parsed['score']}", foreground="red"
                ))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return
            
            # 保存文件
            if not self.excel_handler.save_excel():
                self.root.after(0, lambda: self.status_label.config(
                    text="保存失败，文件可能被占用", foreground="red"
                ))
                self.root.after(0, lambda: messagebox.showwarning(
                    "保存失败", "请关闭Excel文件后重试"
                ))
                self.root.after(0, lambda: self.record_button.config(text="开始录音"))
                self.is_recording = False
                return
            
            # 生成确认文本并播报
            confirmation = self.student_parser.format_confirmation(
                row, student_info['name'], parsed['score']
            )
            
            self.root.after(0, lambda: self.status_label.config(
                text="操作成功", foreground="green"
            ))
            self.root.after(0, lambda: self.log(
                f"已更新: {confirmation}"
            ))
            
            # 语音播报（异步，不阻塞）
            self.speech_synthesis.speak_async(confirmation)
            
            # 重置状态
            self.root.after(0, lambda: self.record_button.config(text="开始录音"))
            self.is_recording = False
            
        except Exception as e:
            import traceback
            error_msg = str(e)
            self.root.after(0, lambda: self.status_label.config(
                text=f"处理失败: {error_msg}", foreground="red"
            ))
            self.root.after(0, lambda: self.log(f"错误: {error_msg}"))
            self.root.after(0, lambda: self.log(f"详细错误: {traceback.format_exc()}"))
            self.root.after(0, lambda: self.record_button.config(text="开始录音"))
            self.is_recording = False

    def _handle_undo(self):
        """处理撤回命令"""
        if self.last_operation is None:
            self.root.after(0, lambda: self.status_label.config(
                text="没有可撤回的操作", foreground="orange"
            ))
            self.root.after(0, lambda: self.log("没有可撤回的操作"))
            self.speech_synthesis.speak_async("没有可撤回的操作")
            return

        # 获取上一步操作信息
        row = self.last_operation['row']
        column = self.last_operation['column']
        old_score = self.last_operation['old_score']
        new_score = self.last_operation['new_score']
        name = self.last_operation['name']

        self.root.after(0, lambda: self.log(
            f"撤回操作: {name}，{new_score}分 -> {old_score if old_score is not None else '空'}分"
        ))

        # 恢复旧分数
        if old_score is None:
            # 如果旧分数是None，清空单元格（设置为空字符串或0）
            # 这里我们设置为None，Excel会显示为空
            success = self.excel_handler.update_score(row, column, 0)  # 暂时设为0
            self.root.after(0, lambda: self.log("警告: 原分数为空，已设置为0"))
        else:
            success = self.excel_handler.update_score(row, column, old_score)

        if not success:
            self.root.after(0, lambda: self.status_label.config(
                text="撤回失败", foreground="red"
            ))
            self.root.after(0, lambda: self.log("撤回失败：无法更新成绩"))
            self.speech_synthesis.speak_async("撤回失败")
            return

        # 保存文件
        if not self.excel_handler.save_excel():
            self.root.after(0, lambda: self.status_label.config(
                text="保存失败，文件可能被占用", foreground="red"
            ))
            self.root.after(0, lambda: messagebox.showwarning(
                "保存失败", "请关闭Excel文件后重试"
            ))
            self.speech_synthesis.speak_async("保存失败")
            return

        # 撤回成功
        self.root.after(0, lambda: self.status_label.config(
            text=f"已撤回: {name}，恢复为{old_score if old_score is not None else '空'}分", foreground="green"
        ))
        self.root.after(0, lambda: self.log(f"撤回成功: {name}"))

        # 语音播报
        score_text = f"{old_score}分" if old_score is not None else "空"
        self.speech_synthesis.speak_async(f"已撤回，{name}，恢复为{score_text}")

        # 清空last_operation，防止重复撤回
        self.last_operation = None

    def log(self, message: str):
        """添加日志"""
        self.log_text.config(state=tk.NORMAL)
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    app = GradeEntryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
