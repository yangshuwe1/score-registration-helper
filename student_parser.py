"""
学生信息解析模块
从语音识别结果中提取学号/姓名和分数
"""
import re
from typing import Optional, Tuple, Dict, List


class StudentParser:
    def __init__(self):
        # 匹配学号的模式：数字+号，或第+数字+号
        self.id_pattern = re.compile(r'(?:第?)(\d+)(?:号|号学生)')
        # 匹配姓名的模式：2-4个中文字符
        self.name_pattern = re.compile(r'([\u4e00-\u9fa5]{2,4})')
        # 匹配分数的模式：优先匹配"数字+分"，再匹配纯数字
        self.score_with_unit_pattern = re.compile(r'(\d+(?:\.\d+)?)分')
        self.score_pattern = re.compile(r'(\d+(?:\.\d+)?)')
    
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

        # 提取分数：优先匹配"数字+分"格式
        score_match = self.score_with_unit_pattern.search(text)
        if score_match:
            # 找到了"数字+分"格式
            try:
                score = float(score_match.group(1))
            except:
                return None
        else:
            # 没有找到"数字+分"，尝试找纯数字（排除学号）
            # 先找出所有数字
            all_numbers = re.findall(r'\d+(?:\.\d+)?', text)
            if not all_numbers:
                return None

            # 如果有学号，排除学号位置的数字，取最后一个数字作为分数
            id_match = self.id_pattern.search(text)
            if id_match and len(all_numbers) > 1:
                # 有学号，取最后一个数字（假设分数在最后）
                try:
                    score = float(all_numbers[-1])
                except:
                    return None
            else:
                # 没有学号或只有一个数字，取最后一个
                try:
                    score = float(all_numbers[-1])
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
    
    def format_confirmation(self, row: int, name: str, score: float, student_id: str = None) -> str:
        """
        格式化确认播报文本
        row: 行号
        name: 姓名
        score: 分数
        student_id: 学号（可选）
        """
        # 如果提供了学号，播报学号而不是行号
        if student_id:
            return f"{student_id}号，{name}，{score}分"
        else:
            # 没有学号，播报行号
            return f"第{row}行，{name}，{score}分"

    def parse_multiple(self, text: str) -> List[Dict[str, any]]:
        """
        解析包含多个学生信息的语音识别结果
        支持格式如："1号10分，2号20分" 或 "1号,10分,2号,20分"
        返回: [{
            'type': 'id' 或 'name',
            'identifier': 学号或姓名,
            'score': 分数
        }, ...]
        """
        if not text:
            return []

        text = text.strip()
        results = []

        # 策略：使用正则表达式找出所有的"学号/姓名+分数"对
        # 模式1: "数字号...数字分" (例如："1号10分", "1号,10分")
        # 先分割成可能的条目（用逗号、句号等分隔）
        entries = re.split(r'[,，。、\s]+', text)

        # 尝试匹配"学号+分数"对
        # 更精确的模式：找所有的"数字号...数字分"
        pattern = r'(?:第?)(\d+)(?:号|号学生)[^0-9]*?(\d+(?:\.\d+)?)分?'
        matches = re.findall(pattern, text)

        if matches:
            # 找到了多个学号+分数对
            for student_id, score_str in matches:
                try:
                    score = float(score_str)
                    if 0 <= score <= 100:
                        results.append({
                            'type': 'id',
                            'identifier': student_id,
                            'score': score
                        })
                except:
                    continue

        # 如果没有找到任何匹配，尝试单个解析
        if not results:
            single_result = self.parse(text)
            if single_result:
                results.append(single_result)

        return results
