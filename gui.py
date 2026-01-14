"""
GUI界面模块
使用tkinter构建图形界面
"""
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from typing import Optional, Dict, List
from excel_handler import ExcelHandler
from speech_recognition import SpeechRecognition
from speech_synthesis import SpeechSynthesis
from student_parser import StudentParser
from config_manager import ConfigManager


class GradeEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("登分助手 - 智能配置版")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # 初始化配置管理器
        self.config_manager = ConfigManager()

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
        self.current_column = 'final_score'  # 默认期末成绩
        self.is_recording = False
        self.last_operation = None  # 上一步操作缓存：{'row': int, 'column': str, 'old_score': float, 'new_score': float, 'name': str}

        # 列名缓存（用于配置界面）
        self.cached_column_names: List[str] = []
    
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
        # 创建Notebook（标签页容器）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建主界面标签页
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="  主界面  ")

        # 创建配置标签页
        self.config_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.config_tab, text="  配置  ")

        # 构建主界面
        self._create_main_tab()

        # 构建配置界面
        self._create_config_tab()

    def _create_main_tab(self):
        """创建主界面标签页"""
        # 主框架
        main_frame = ttk.Frame(self.main_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)

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

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

    def _create_config_tab(self):
        """创建配置标签页"""
        # 主框架（带滚动条）
        canvas = tk.Canvas(self.config_tab)
        scrollbar = ttk.Scrollbar(self.config_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 主配置区域
        config_frame = ttk.Frame(scrollable_frame, padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True)

        # ===== 1. Excel配置区域 =====
        excel_frame = ttk.LabelFrame(config_frame, text="Excel配置", padding="10")
        excel_frame.pack(fill=tk.X, pady=5)

        # 表头行数
        row_num = 0
        ttk.Label(excel_frame, text="表头行数:").grid(row=row_num, column=0, sticky=tk.W, padx=5, pady=5)
        self.header_rows_var = tk.IntVar(value=self.config_manager.get_header_rows())
        header_rows_spinbox = ttk.Spinbox(excel_frame, from_=0, to=10, textvariable=self.header_rows_var, width=10)
        header_rows_spinbox.grid(row=row_num, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Label(excel_frame, text="（数据从第几行开始的前一行，通常为1-3）", foreground="gray").grid(
            row=row_num, column=2, sticky=tk.W, padx=5, pady=5
        )

        # 列名行号
        row_num += 1
        ttk.Label(excel_frame, text="列名所在行:").grid(row=row_num, column=0, sticky=tk.W, padx=5, pady=5)
        self.column_name_row_var = tk.IntVar(value=self.config_manager.get_header_rows())
        column_name_row_spinbox = ttk.Spinbox(excel_frame, from_=1, to=10, textvariable=self.column_name_row_var, width=10)
        column_name_row_spinbox.grid(row=row_num, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Button(excel_frame, text="读取列名", command=self._read_column_names).grid(
            row=row_num, column=2, sticky=tk.W, padx=5, pady=5
        )

        # 分隔线
        row_num += 1
        ttk.Separator(excel_frame, orient='horizontal').grid(
            row=row_num, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10
        )

        # 列映射配置
        row_num += 1
        ttk.Label(excel_frame, text="列映射配置", font=('TkDefaultFont', 10, 'bold')).grid(
            row=row_num, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5
        )

        # 学号列
        row_num += 1
        ttk.Label(excel_frame, text="学号列:").grid(row=row_num, column=0, sticky=tk.W, padx=5, pady=5)
        self.student_id_col_var = tk.StringVar()
        self.student_id_col_combo = ttk.Combobox(
            excel_frame, textvariable=self.student_id_col_var, width=30, state="readonly"
        )
        self.student_id_col_combo.grid(row=row_num, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)

        # 姓名列
        row_num += 1
        ttk.Label(excel_frame, text="姓名列:").grid(row=row_num, column=0, sticky=tk.W, padx=5, pady=5)
        self.name_col_var = tk.StringVar()
        self.name_col_combo = ttk.Combobox(
            excel_frame, textvariable=self.name_col_var, width=30, state="readonly"
        )
        self.name_col_combo.grid(row=row_num, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)

        # 平时成绩列
        row_num += 1
        ttk.Label(excel_frame, text="平时成绩列:").grid(row=row_num, column=0, sticky=tk.W, padx=5, pady=5)
        self.regular_score_col_var = tk.StringVar()
        self.regular_score_col_combo = ttk.Combobox(
            excel_frame, textvariable=self.regular_score_col_var, width=30, state="readonly"
        )
        self.regular_score_col_combo.grid(row=row_num, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)

        # 期末成绩列
        row_num += 1
        ttk.Label(excel_frame, text="期末成绩列:").grid(row=row_num, column=0, sticky=tk.W, padx=5, pady=5)
        self.final_score_col_var = tk.StringVar()
        self.final_score_col_combo = ttk.Combobox(
            excel_frame, textvariable=self.final_score_col_var, width=30, state="readonly"
        )
        self.final_score_col_combo.grid(row=row_num, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5, pady=5)

        excel_frame.columnconfigure(2, weight=1)

        # ===== 2. 语音识别配置区域 =====
        speech_config_frame = ttk.LabelFrame(config_frame, text="语音识别配置", padding="10")
        speech_config_frame.pack(fill=tk.X, pady=5)

        # 模型选择
        ttk.Label(speech_config_frame, text="识别模型:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.model_var = tk.StringVar(value=self.config_manager.get_whisper_model())
        self.model_combo = ttk.Combobox(
            speech_config_frame, textvariable=self.model_var,
            values=ConfigManager.AVAILABLE_MODELS, width=15, state="readonly"
        )
        self.model_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_change)

        # 模型说明
        self.model_desc_label = ttk.Label(
            speech_config_frame,
            text=ConfigManager.MODEL_DESCRIPTIONS.get(self.model_var.get(), ""),
            foreground="gray"
        )
        self.model_desc_label.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

        speech_config_frame.columnconfigure(2, weight=1)

        # ===== 3. 按钮区域 =====
        button_frame = ttk.Frame(config_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="保存配置", command=self._save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="重置为默认", command=self._reset_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="应用配置", command=self._apply_config).pack(side=tk.LEFT, padx=5)

        # 加载当前配置
        self._load_current_config()

    # ===== 配置相关方法 =====

    def _load_current_config(self):
        """加载当前配置到界面"""
        # 加载Excel列配置
        columns = self.config_manager.get_excel_columns()

        # 如果还没有列名，先加载默认值
        if not self.cached_column_names:
            self.cached_column_names = [
                f"列{i}  (索引{i})" for i in range(26)
            ]

        # 更新下拉框选项
        self._update_column_combos()

        # 设置当前选中的列
        if 'student_id' in columns:
            idx = columns['student_id']
            if idx < len(self.cached_column_names):
                self.student_id_col_var.set(self.cached_column_names[idx])

        if 'name' in columns:
            idx = columns['name']
            if idx < len(self.cached_column_names):
                self.name_col_var.set(self.cached_column_names[idx])

        if 'regular_score' in columns:
            idx = columns['regular_score']
            if idx < len(self.cached_column_names):
                self.regular_score_col_var.set(self.cached_column_names[idx])

        if 'final_score' in columns:
            idx = columns['final_score']
            if idx < len(self.cached_column_names):
                self.final_score_col_var.set(self.cached_column_names[idx])

    def _update_column_combos(self):
        """更新列下拉框的选项"""
        self.student_id_col_combo['values'] = self.cached_column_names
        self.name_col_combo['values'] = self.cached_column_names
        self.regular_score_col_combo['values'] = self.cached_column_names
        self.final_score_col_combo['values'] = self.cached_column_names

    def _read_column_names(self):
        """从Excel文件读取列名"""
        if self.excel_handler.file_path is None:
            messagebox.showwarning("提示", "请先在主界面选择Excel文件")
            return

        # 获取列名所在行号
        row_num = self.column_name_row_var.get()

        # 读取列名
        column_names = self.excel_handler.get_column_names(row_num)

        if not column_names:
            messagebox.showerror("错误", "读取列名失败，请检查行号设置")
            return

        # 更新缓存
        self.cached_column_names = [
            f"{name}  (索引{i})" for i, name in enumerate(column_names)
        ]

        # 更新下拉框
        self._update_column_combos()

        # 智能匹配列名
        self._auto_match_columns(column_names)

        messagebox.showinfo("成功", f"已读取 {len(column_names)} 个列名")

    def _auto_match_columns(self, column_names: List[str]):
        """智能匹配列名到字段"""
        # 匹配规则
        rules = {
            'student_id': ['学号', '学生学号', 'ID', 'id', '编号'],
            'name': ['姓名', '学生姓名', '名字', 'name', 'Name'],
            'regular_score': ['平时', '平时成绩', '平时分', '日常', '日常成绩'],
            'final_score': ['期末', '期末成绩', '期末分', '考试', '考试成绩']
        }

        for field, keywords in rules.items():
            for i, col_name in enumerate(column_names):
                if any(keyword in col_name for keyword in keywords):
                    # 找到匹配，设置对应的下拉框
                    display_value = f"{col_name}  (索引{i})"
                    if field == 'student_id':
                        self.student_id_col_var.set(display_value)
                    elif field == 'name':
                        self.name_col_var.set(display_value)
                    elif field == 'regular_score':
                        self.regular_score_col_var.set(display_value)
                    elif field == 'final_score':
                        self.final_score_col_var.set(display_value)
                    break

    def _on_model_change(self, event=None):
        """模型选择改变时更新说明"""
        model = self.model_var.get()
        desc = ConfigManager.MODEL_DESCRIPTIONS.get(model, "")
        self.model_desc_label.config(text=desc)

    def _save_config(self):
        """保存配置到文件"""
        try:
            # 保存表头行数
            self.config_manager.set_header_rows(self.header_rows_var.get())

            # 保存列映射
            def get_column_index(var_value: str) -> int:
                """从显示值中提取列索引"""
                try:
                    # 格式: "列名  (索引X)"
                    if "(索引" in var_value:
                        idx_str = var_value.split("(索引")[1].rstrip(")")
                        return int(idx_str)
                    return 0
                except:
                    return 0

            columns = {
                'student_id': get_column_index(self.student_id_col_var.get()),
                'name': get_column_index(self.name_col_var.get()),
                'regular_score': get_column_index(self.regular_score_col_var.get()),
                'final_score': get_column_index(self.final_score_col_var.get())
            }
            self.config_manager.set_excel_columns(columns)

            # 保存语音识别配置
            self.config_manager.set_whisper_model(self.model_var.get())

            # 保存到文件
            if self.config_manager.save_config():
                messagebox.showinfo("成功", "配置已保存！\n注意：部分配置（如模型）需要重启程序才能生效。")
                self.log("配置已保存")
            else:
                messagebox.showerror("错误", "保存配置失败")

        except Exception as e:
            messagebox.showerror("错误", f"保存配置时出错: {str(e)}")

    def _reset_config(self):
        """重置为默认配置"""
        if messagebox.askyesno("确认", "确定要重置为默认配置吗？"):
            self.config_manager.reset_to_default()
            self.header_rows_var.set(self.config_manager.get_header_rows())
            self.column_name_row_var.set(self.config_manager.get_header_rows())
            self.model_var.set(self.config_manager.get_whisper_model())
            self._on_model_change()
            self._load_current_config()
            messagebox.showinfo("成功", "已重置为默认配置")
            self.log("配置已重置为默认")

    def _apply_config(self):
        """应用配置（不保存到文件）"""
        try:
            # 应用表头行数
            header_rows = self.header_rows_var.get()
            if self.excel_handler:
                self.excel_handler.update_header_rows(header_rows)

            # 应用列映射（更新config模块的全局变量）
            import config

            def get_column_index(var_value: str) -> int:
                try:
                    if "(索引" in var_value:
                        idx_str = var_value.split("(索引")[1].rstrip(")")
                        return int(idx_str)
                    return 0
                except:
                    return 0

            config.EXCEL_COLUMNS = {
                'student_id': get_column_index(self.student_id_col_var.get()),
                'name': get_column_index(self.name_col_var.get()),
                'regular_score': get_column_index(self.regular_score_col_var.get()),
                'final_score': get_column_index(self.final_score_col_var.get())
            }
            config.HEADER_ROWS = header_rows

            messagebox.showinfo("成功", "配置已应用！\n注意：语音识别模型需要重启程序才能更改。")
            self.log("配置已应用到当前会话")

        except Exception as e:
            messagebox.showerror("错误", f"应用配置时出错: {str(e)}")

    # ===== 主界面方法 =====

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
