import os
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow  # 假设 UI 文件位于 src/UI2.py 中
import sys

class MainWindow(QMainWindow):

    def r2bom(self, rw):
        try:
            print(f"调用 R2BOM 函数，参数为 {rw}")
            # 在这里实现 R2BOM 函数的具体逻辑
        except Exception as e:
            print(f"R2BOM 函数执行时发生错误: {e}")

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # 动态为每个按钮绑定函数
        for i in range(1, 7):
            button_name = f"pushButton_{i}"
            button = getattr(self.ui, button_name, None)
            if button:
                button.clicked.connect(lambda checked, rw=i: self.r2bom(rw))

def create_ui():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    create_ui()