#!/usr/bin/env python3
"""运行主程序并捕获错误"""
import sys
import traceback

try:
    print("正在启动3MF文件预览工具...")
    print("Python版本:", sys.version)
    
    # 导入主程序
    from qt_3mf_previewer import ThreeMfPreviewer, QApplication
    
    print("主程序导入成功")
    
    # 创建应用
    app = QApplication(sys.argv)
    print("QApplication创建成功")
    
    # 创建窗口
    window = ThreeMfPreviewer()
    print("主窗口创建成功")
    
    # 显示窗口
    window.show()
    print("主窗口显示成功")
    
    # 运行应用
    print("开始运行应用...")
    exit_code = app.exec_()
    
    print(f"应用退出，退出码: {exit_code}")
    sys.exit(exit_code)
    
except ImportError as e:
    print(f"导入错误: {str(e)}")
    traceback.print_exc()
    sys.exit(1)
except Exception as e:
    print(f"运行错误: {str(e)}")
    traceback.print_exc()
    sys.exit(1)