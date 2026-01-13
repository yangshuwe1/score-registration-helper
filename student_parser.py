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

        # 中文数字映射表
        self.cn_num_map = {
            '零': '0', '一': '1', '二': '2', '三': '3', '四': '4',
            '五': '5', '六': '6', '七': '7', '八': '8', '九': '9',
            '十': '10', '百': '100',
            # 大写
            '壹': '1', '贰': '2', '叁': '3', '肆': '4', '伍': '5',
            '陆': '6', '柒': '7', '捌': '8', '玖': '9', '拾': '10'
        }

        # 繁体字转简体字映射表（常用字）
        self.traditional_to_simplified = {
            '號': '号', '個': '个', '學': '学',
            '畫': '画', '書': '书', '長': '长', '門': '门',
            '開': '开', '關': '关', '來': '来', '時': '时',
            '認': '认', '證': '证', '華': '华', '國': '国',
            '點': '点', '現': '现', '實': '实', '際': '际',
            '這': '这', '發': '发', '員': '员', '會': '会',
            '經': '经', '過': '过', '還': '还', '進': '进'
        }

    def _normalize_text(self, text: str) -> str:
        """
        规范化文本：繁体转简体，中文数字转阿拉伯数字，去除干扰词
        """
        if not text:
            return text

        # 1. 繁体字转简体字
        for trad, simp in self.traditional_to_simplified.items():
            text = text.replace(trad, simp)

        # 2. 去除常见干扰词和语音识别错误
        # "四个号" -> "四号"，"个"通常是识别错误
        text = text.replace('个号', '号')
        text = text.replace('个学生', '学生')
        # 去除多余的空格、逗号前后的"号"
        text = re.sub(r'号\s*,', '号,', text)
        text = re.sub(r',\s*号', ',', text)

        # 3. 处理中文数字（支持"三十七"、"一百"等复杂组合）
        # 先处理百位数："一百五十" → "150"，"二百" → "200"
        def convert_hundreds(match):
            hundreds_digit = match.group(1) if match.group(1) else '一'
            tens_digit = match.group(2) if match.group(2) else ''
            ones_digit = match.group(3) if match.group(3) else ''

            result = int(self.cn_num_map.get(hundreds_digit, '1')) * 100
            if tens_digit:
                result += int(self.cn_num_map.get(tens_digit, '0')) * 10
            if ones_digit:
                result += int(self.cn_num_map.get(ones_digit, '0'))
            return str(result)

        # 匹配：[一-九]?百[零一-九]?[十]?[一-九]?
        text = re.sub(r'([一二三四五六七八九])?百([零一二三四五六七八九])?十?([一二三四五六七八九])?',
                     convert_hundreds, text)

        # 处理十位数："三十七" → "37"，"二十" → "20"
        def convert_tens(match):
            tens_digit = match.group(1)
            ones_digit = match.group(2) if match.group(2) else ''

            result = int(self.cn_num_map[tens_digit]) * 10
            if ones_digit:
                result += int(self.cn_num_map[ones_digit])
            return str(result)

        # 匹配：[一-九]十[一-九]? （如"三十七"、"二十"）
        text = re.sub(r'([一二三四五六七八九])十([一二三四五六七八九])?', convert_tens, text)

        # 处理"十X"（如"十一"、"十二"）
        text = re.sub(r'十([一二三四五六七八九])', lambda m: str(10 + int(self.cn_num_map[m.group(1)])), text)
        # 处理单独的"十"
        text = text.replace('十', '10')

        # 4. 替换单个中文数字（最后处理，避免干扰组合数字）
        for cn, num in self.cn_num_map.items():
            if cn not in ['十', '百']:  # 十和百已经处理过了
                text = text.replace(cn, num)

        return text
    
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
        # 规范化文本（繁体转简体，中文数字转阿拉伯数字）
        text = self._normalize_text(text)

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
    
    def format_confirmation(self, row: int, name: str, score: float) -> str:
        """
        格式化确认播报文本
        row: Excel行号（1-based，包含表头）
        name: 姓名
        score: 分数
        返回: "序号号，姓名，分数分"（例如："1号，许凯旋，30分"）
        """
        # 计算序号：row是Excel行号（1-based），减去表头行数（2行）
        # 例如：row=3 -> 序号=1（第1个学生）
        from config import HEADER_ROWS
        sequence_number = row - HEADER_ROWS
        return f"{sequence_number}号，{name}，{score}分"

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
        # 规范化文本（繁体转简体，中文数字转阿拉伯数字）
        text = self._normalize_text(text)
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
