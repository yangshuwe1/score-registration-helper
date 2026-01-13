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


class GradeEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("登分助手")
        self.root.geometry("800x600")
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
        self.current_column = 'final_score'  # 默认总成绩
        self.is_recording = False
    
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
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
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
            column_frame, text="总成绩", variable=self.column_var,
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
        column_name = "总成绩" if self.current_column == "final_score" else "平时成绩"
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

                # 更新成绩
                if not self.excel_handler.update_score(row, self.current_column, parsed['score']):
                    score = parsed['score']
                    self.root.after(0, lambda s=score: self.log(f"更新成绩失败，分数: {s}，跳过"))
                    continue

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
                # 启动实时录音
                success = self.speech_recognition.record_audio_realtime(
                    on_speech_end=on_speech_end,
                    silence_duration=1.5,  # 静音1.5秒后认为说话结束
                    min_speech_duration=0.5  # 最少0.5秒才识别
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
