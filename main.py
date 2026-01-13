"""
登分助手主程序
"""
# -*- coding: utf-8 -*-
import sys
import os
import traceback
import tkinter as tk
from tkinter import messagebox

# 确保在 Windows 系统上使用 UTF-8 编码
if sys.platform == 'win32':
    try:
        import io
        if hasattr(sys.stdout, 'buffer'):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        if hasattr(sys.stderr, 'buffer'):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass  # 如果无法设置编码，忽略错误

def main():
    """主函数，带错误处理"""
    try:
        print("=" * 60)
        print("登分助手 - 启动中...")
        print("=" * 60)
        
        # 创建根窗口
        root = tk.Tk()
        
        # 设置窗口关闭事件
        def on_closing():
            if messagebox.askokcancel("退出", "确定要退出程序吗？"):
                try:
                    root.quit()
                    root.destroy()
                except:
                    pass
                sys.exit(0)
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # 导入GUI（在创建窗口后）
        try:
            from gui import GradeEntryApp
            app = GradeEntryApp(root)
            print("界面初始化完成")
        except Exception as e:
            print(f"界面初始化失败: {e}")
            traceback.print_exc()
            messagebox.showerror(
                "启动失败", 
                f"程序启动失败:\n{str(e)}\n\n请检查依赖是否已正确安装。"
            )
            sys.exit(1)
        
        # 运行主循环
        print("进入主循环...")
        root.mainloop()
        
    except KeyboardInterrupt:
        print("\n用户中断程序")
        sys.exit(0)
    except Exception as e:
        print(f"\n程序异常退出: {e}")
        traceback.print_exc()
        try:
            messagebox.showerror(
                "程序错误", 
                f"程序遇到错误:\n{str(e)}\n\n详细信息请查看控制台输出。"
            )
        except:
            pass
        sys.exit(1)

if __name__ == "__main__":
    main()
