# 配置文件

# Excel列配置（0-based索引）
# 根据学校成绩登记表格式
EXCEL_COLUMNS = {
    'student_id': 0,      # A列：学号*
    'name': 1,            # B列：姓名*
    'regular_score': 4,   # E列：平时成绩
    'final_score': 7      # H列：期末成绩
}

# 表头行数配置
HEADER_ROWS = 2  # 前2行是表头（说明行和列名行）

# 语音识别配置
# 模型选择: "tiny"(最快,最小), "small"(快,小), "base"(平衡), "medium"(慢,准确), "large"(最慢,最准确)
# 如果遇到内存不足或加载慢，可以改为 "small" 或 "tiny"
WHISPER_MODEL = "large"  # 使用base模型，平衡速度和准确率
WHISPER_DEVICE = "cpu"  # 使用CPU，兼容性更好
WHISPER_COMPUTE_TYPE = "int8"  # 量化模型，降低延迟和内存占用

# 录音配置
RECORD_DURATION = 5  # 最大录音时长（秒）
SAMPLE_RATE = 16000  # 采样率
CHUNK_SIZE = 1024    # 音频块大小

# TTS配置
TTS_VOICE = "zh-CN-XiaoxiaoNeural"  # 中文语音
