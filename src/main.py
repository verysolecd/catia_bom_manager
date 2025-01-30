import os
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow  # 假设 UI 文件位于 src/UI2.py 中
import sys
import win32com

class MainWindow(QMainWindow):

    def r2bom(self, rw):
        pass


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