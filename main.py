## 主程序入口
import os
print(os.getcwd()) 
from PyQt5.QtWidgets import QApplication, QMainWindow
from  GUI.UI2 import Ui_MainWindow
import sys

class MainWindow(QMainWindow):

    def Ufunc1(self,rw):
        # 调用函数 R2Bom，并将参数 rw 传递给它
        R2Bom(self,rw)

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # 为每个按钮绑定函数
        self.ui.pushButton_1.clicked.connect(lambda: self.Ufunc1(1))
        self.ui.pushButton_2.clicked.connect(lambda: self.Ufunc1(2))
        self.ui.pushButton_3.clicked.connect(lambda: self.Ufunc1(3))
        self.ui.pushButton_4.clicked.connect(lambda: self.Ufunc1(4))
        self.ui.pushButton_5.clicked.connect(lambda: self.Ufunc1(5))
        self.ui.pushButton_6.clicked.connect(lambda: self.Ufunc1(6))
        

def create_ui():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    create_ui()
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()
def R2bom(rw):
    # 在这里实现 R2BOM 函数的具体逻辑
    print(f"调用 R2BOM 函数，参数为 {rw}")