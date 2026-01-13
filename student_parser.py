"""
学生信息解析模块 - 极简版
从语音识别结果中提取序号/姓名和分数
"""
import re
from typing import Optional, Dict, List


class StudentParser:
    """简单可靠的学生信息解析器"""

    def __init__(self):
        # 匹配序号：数字+号
        self.sequence_pattern = re.compile(r'(\d+)号')
        # 匹配分数：数字+分
        self.score_pattern = re.compile(r'(\d+(?:\.\d+)?)分')
        # 匹配姓名：2-4个中文字符
        self.name_pattern = re.compile(r'[\u4e00-\u9fa5]{2,4}')

    def parse(self, text: str) -> Optional[Dict]:
        """
        解析单个学生信息
        返回: {'type': 'sequence' 或 'name', 'identifier': 序号或姓名, 'score': 分数}
        """
        if not text:
            return None

        text = text.strip()

        # 提取分数（必须）
        score_match = self.score_pattern.search(text)
        if not score_match:
            return None

        try:
            score = float(score_match.group(1))
            if score < 0 or score > 100:
                return None
        except:
            return None

        # 提取序号（优先）
        seq_match = self.sequence_pattern.search(text)
        if seq_match:
            return {
                'type': 'sequence',
                'identifier': seq_match.group(1),
                'score': score
            }

        # 提取姓名
        name_match = self.name_pattern.search(text)
        if name_match:
            name = name_match.group(0)
            # 过滤干扰词
            if name not in ['分', '号', '第', '个', '学生', '成绩', '分数', '序号']:
                return {
                    'type': 'name',
                    'identifier': name,
                    'score': score
                }

        return None

    def parse_multiple(self, text: str) -> List[Dict]:
        """
        解析多个学生信息（简单版：分割后逐个解析）
        """
        if not text:
            return []

        # 尝试分割（按逗号）
        parts = re.split(r'[,，]', text)

        results = []
        for part in parts:
            parsed = self.parse(part.strip())
            if parsed:
                results.append(parsed)

        # 如果分割后没有结果，尝试整体解析
        if not results:
            parsed = self.parse(text)
            if parsed:
                results.append(parsed)

        return results

    def format_confirmation(self, row: int, name: str, score: float) -> str:
        """格式化确认播报文本"""
        from config import HEADER_ROWS
        sequence = row - HEADER_ROWS
        return f"{sequence}号，{name}，{score}分"
