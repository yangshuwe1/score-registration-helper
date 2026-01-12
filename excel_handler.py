"""
Excel文件处理模块
处理Excel文件的读取、查找学生、更新成绩等功能
"""
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from typing import Optional, Tuple
import os
from config import EXCEL_COLUMNS


class ExcelHandler:
    def __init__(self):
        self.file_path: Optional[str] = None
        self.df: Optional[pd.DataFrame] = None
        self.wb: Optional[openpyxl.Workbook] = None
        self.ws: Optional[openpyxl.worksheet.worksheet.Worksheet] = None
        self.header_row = 0  # 默认第一行是表头
    
    def load_excel(self, file_path: str) -> bool:
        """
        加载Excel文件
        返回True表示成功，False表示失败
        """
        try:
            if not os.path.exists(file_path):
                print(f"文件不存在: {file_path}")
                return False
            
            if not os.access(file_path, os.R_OK):
                print(f"文件无法读取: {file_path}")
                return False
            
            self.file_path = file_path
            
            # 使用pandas读取用于查找
            try:
                self.df = pd.read_excel(file_path, header=None, engine='openpyxl')
            except Exception as e:
                # 尝试使用xlrd引擎（用于.xls文件）
                try:
                    self.df = pd.read_excel(file_path, header=None, engine='xlrd')
                except Exception as e2:
                    print(f"无法读取Excel文件: {e}, {e2}")
                    return False
            
            if self.df.empty:
                print("Excel文件为空")
                return False
            
            # 使用openpyxl打开用于写入
            try:
                self.wb = load_workbook(file_path)
                self.ws = self.wb.active
            except Exception as e:
                print(f"无法打开Excel文件用于写入: {e}")
                print("提示: 请确保文件未被其他程序打开")
                return False
            
            # 检测表头：如果第一行是文本标题，则从第二行开始
            first_row = self.df.iloc[0]
            if any(isinstance(val, str) and ('学号' in str(val) or '姓名' in str(val)) 
                   for val in first_row):
                self.header_row = 1
                # 重新读取，跳过表头
                try:
                    self.df = pd.read_excel(file_path, header=0, engine='openpyxl')
                except:
                    self.df = pd.read_excel(file_path, header=0, engine='xlrd')
            else:
                self.header_row = 0
                # 如果没有表头，使用默认列名
                if len(self.df.columns) >= 4:
                    self.df.columns = ['学号', '姓名', '平时成绩', '总成绩'] + \
                                     [f'列{i}' for i in range(4, len(self.df.columns))]
                else:
                    print("Excel文件列数不足，至少需要4列（学号、姓名、平时成绩、总成绩）")
                    return False
            
            return True
        except PermissionError:
            print("文件被其他程序占用，无法打开")
            return False
        except Exception as e:
            print(f"加载Excel文件失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def find_student_by_id(self, student_id: str) -> Optional[int]:
        """
        根据学号查找学生，返回Excel中的行号（1-based，包含表头）
        """
        if self.df is None or self.df.empty:
            return None
        
        try:
            # 清理学号字符串
            student_id = str(student_id).strip()
            if not student_id:
                return None
            
            # 在学号列中查找
            id_col = self.df.iloc[:, EXCEL_COLUMNS['student_id']]
            
            # 尝试直接匹配
            matches = id_col.astype(str).str.strip() == student_id
            
            # 如果没找到，尝试匹配数字部分
            if not matches.any():
                # 提取学号中的数字
                id_num = ''.join(filter(str.isdigit, student_id))
                if id_num:
                    # 将学号列转换为字符串并提取数字
                    id_col_str = id_col.astype(str).str.replace(r'\D', '', regex=True)
                    matches = id_col_str == id_num
            
            # 如果还是没找到，尝试模糊匹配（包含关系）
            if not matches.any():
                id_num = ''.join(filter(str.isdigit, student_id))
                if id_num:
                    id_col_str = id_col.astype(str).str.replace(r'\D', '', regex=True)
                    matches = id_col_str.str.contains(id_num, na=False)
            
            if matches.any():
                # 返回第一个匹配的行号（1-based，包含表头）
                excel_row = matches.idxmax() + 1 + self.header_row
                return excel_row
            
            return None
        except Exception as e:
            print(f"查找学号失败: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def find_student_by_name(self, name: str) -> Optional[int]:
        """
        根据姓名查找学生，返回Excel中的行号（1-based，包含表头）
        """
        if self.df is None or self.df.empty:
            return None
        
        try:
            name = str(name).strip()
            if not name:
                return None
            
            # 在姓名列中查找
            name_col = self.df.iloc[:, EXCEL_COLUMNS['name']]
            
            # 精确匹配
            matches = name_col.astype(str).str.strip() == name
            
            # 如果没找到，尝试模糊匹配（包含关系）
            if not matches.any():
                matches = name_col.astype(str).str.contains(name, na=False, regex=False)
            
            if matches.any():
                # 返回第一个匹配的行号（1-based，包含表头）
                excel_row = matches.idxmax() + 1 + self.header_row
                return excel_row
            
            return None
        except Exception as e:
            print(f"查找姓名失败: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def get_student_info(self, row: int) -> Optional[dict]:
        """
        获取指定行的学生信息
        row: Excel行号（1-based）
        """
        if self.ws is None:
            return None
        
        try:
            # 转换为pandas索引（0-based，不含表头）
            pandas_row = row - 1 - self.header_row
            
            if pandas_row < 0 or pandas_row >= len(self.df):
                return None
            
            student_id = self.ws.cell(row=row, column=EXCEL_COLUMNS['student_id'] + 1).value
            name = self.ws.cell(row=row, column=EXCEL_COLUMNS['name'] + 1).value
            
            return {
                'row': row,
                'student_id': str(student_id) if student_id else '',
                'name': str(name) if name else ''
            }
        except Exception as e:
            print(f"获取学生信息失败: {e}")
            return None
    
    def update_score(self, row: int, column_type: str, score: float) -> bool:
        """
        更新指定行的成绩
        row: Excel行号（1-based）
        column_type: 'regular_score' 或 'final_score'
        score: 分数（0-100）
        """
        if self.ws is None:
            return False
        
        try:
            # 验证分数范围
            score = float(score)
            if score < 0 or score > 100:
                return False
            
            # 确定列
            if column_type == 'regular_score':
                col = EXCEL_COLUMNS['regular_score'] + 1
            elif column_type == 'final_score':
                col = EXCEL_COLUMNS['final_score'] + 1
            else:
                return False
            
            # 更新单元格
            self.ws.cell(row=row, column=col, value=score)
            
            return True
        except Exception as e:
            print(f"更新成绩失败: {e}")
            return False
    
    def save_excel(self) -> bool:
        """
        保存Excel文件
        """
        try:
            if self.wb is None or self.file_path is None:
                print("未加载Excel文件")
                return False
            
            # 检查文件是否可写
            if not os.access(self.file_path, os.W_OK):
                print(f"文件不可写: {self.file_path}")
                return False
            
            # 检查文件是否被其他程序打开
            try:
                self.wb.save(self.file_path)
                return True
            except PermissionError:
                print("文件被其他程序打开，无法保存")
                print("提示: 请关闭Excel文件后重试")
                return False
            except Exception as e:
                print(f"保存Excel文件失败: {e}")
                import traceback
                traceback.print_exc()
                return False
        except Exception as e:
            print(f"保存Excel文件失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_total_students(self) -> int:
        """
        获取学生总数
        """
        if self.df is None:
            return 0
        return len(self.df)
