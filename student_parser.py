"""
学生信息解析模块
从语音识别结果中提取学号/姓名和分数
"""
import re
from typing import Optional, Tuple, Dict


class StudentParser:
    def __init__(self):
        # 匹配学号的模式：数字+号，或第+数字+号
        self.id_pattern = re.compile(r'(?:第?)(\d+)(?:号|号学生)')
        # 匹配姓名的模式：2-4个中文字符
        self.name_pattern = re.compile(r'([\u4e00-\u9fa5]{2,4})')
        # 匹配分数的模式：数字+分
        self.score_pattern = re.compile(r'(\d+(?:\.\d+)?)(?:分)?')
    
    def parse(self, text: str) -> Optional[Dict[str, any]]:
        """
        解析语音识别结果
        返回: {
            'type': 'id' 或 'name',
            'identifier': 学号或姓名,
            'score': 分数
        }
        """
        if not text:
            return None
        
        text = text.strip()
        
        # 提取分数
        score_match = self.score_pattern.search(text)
        if not score_match:
            # 尝试提取纯数字作为分数（备用方案）
            numbers = re.findall(r'\d+(?:\.\d+)?', text)
            if not numbers:
                return None
            try:
                score = float(numbers[-1])  # 取最后一个数字作为分数
            except:
                return None
        else:
            try:
                score = float(score_match.group(1))
            except:
                return None
        
        # 验证分数范围
        if score < 0 or score > 100:
            return None
        
        # 提取学号或姓名
        # 优先匹配学号（数字+号）
        id_match = self.id_pattern.search(text)
        if id_match:
            student_id = id_match.group(1)
            return {
                'type': 'id',
                'identifier': student_id,
                'score': score
            }
        
        # 匹配姓名（2-4个中文字符）
        name_match = self.name_pattern.search(text)
        if name_match:
            name = name_match.group(1)
            # 过滤掉常见的干扰词
            if name not in ['分', '号', '行', '第', '个', '学生', '成绩']:
                return {
                    'type': 'name',
                    'identifier': name,
                    'score': score
                }
        
        return None
    
    def format_confirmation(self, row: int, name: str, score: float) -> str:
        """
        格式化确认播报文本
        """
        # 将行号转换为中文数字（简化版，只处理常见情况）
        row_str = str(row)
        return f"第{row_str}行，{name}，{score}分"
