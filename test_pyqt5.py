#!/usr/bin/env python3
"""测试PyQt5是否正常工作"""
import sys
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout, QPushButton
from PyQt5.QtCore import Qt

def test_pyqt5():
    """测试PyQt5功能"""
    app = QApplication(sys.argv)
    
    # 创建窗口
    window = QWidget()
    window.setWindowTitle("PyQt5测试")
    window.setGeometry(100, 100, 300, 200)
    
    # 创建布局
    layout = QVBoxLayout()
    
    # 添加标签
    label = QLabel("PyQt5测试成功！")
    label.setAlignment(Qt.AlignCenter)
    layout.addWidget(label)
    
    # 添加按钮
    button = QPushButton("关闭")
    button.clicked.connect(window.close)
    layout.addWidget(button)
    
    window.setLayout(layout)
    
    # 显示窗口
    window.show()
    
    # 运行应用
    print("PyQt5测试窗口已显示，请检查是否能看到窗口...")
    print("如果能看到窗口，点击'关闭'按钮退出测试。")
    
    return app.exec_()

if __name__ == "__main__":
    try:
        print("开始PyQt5测试...")
        exit_code = test_pyqt5()
        print(f"PyQt5测试完成，退出码: {exit_code}")
        sys.exit(exit_code)
    except Exception as e:
        print(f"PyQt5测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)