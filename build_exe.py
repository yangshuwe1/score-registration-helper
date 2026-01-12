"""
PyInstaller打包脚本
使用方法: python build_exe.py
"""
import PyInstaller.__main__
import os
import sys

def build():
    """构建exe文件"""
    print("开始打包...")
    print("注意: 首次打包可能需要较长时间，因为需要下载和包含模型文件")
    
    # 打包配置
    args = [
        'main.py',
        '--onefile',  # 单文件模式
        '--windowed',  # 无控制台窗口
        '--name=登分助手',
        '--hidden-import=edge_tts',
        '--hidden-import=faster_whisper',
        '--hidden-import=openpyxl',
        '--hidden-import=pandas',
        '--hidden-import=pyaudio',
        '--hidden-import=pygame',
        '--hidden-import=numpy',
        '--hidden-import=torch',
        '--collect-all=edge_tts',
        '--collect-all=faster_whisper',
        '--collect-all=torch',
        '--noconsole',
        '--clean',  # 清理临时文件
    ]
    
    # 添加图标（如果存在）
    icon_path = 'icon.ico'
    if os.path.exists(icon_path):
        args.append(f'--icon={icon_path}')
    
    try:
        PyInstaller.__main__.run(args)
        print("\n打包完成！")
        print("exe文件位置: dist/登分助手.exe")
        print("\n注意:")
        print("1. 首次运行exe时，模型文件会自动下载到用户目录")
        print("2. 如果exe文件较大（200-300MB），这是正常的，因为包含了模型文件")
        print("3. 确保目标机器有足够的磁盘空间")
    except Exception as e:
        print(f"打包失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    build()
