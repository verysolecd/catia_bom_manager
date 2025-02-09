from PyQt5 import  QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow  # 假设 UI 文件位于 src/UI2.py 中
import sys
#import src.catia_Processor as PDM
from src.data_processor import ClassTDM

import win32com.client
# from pycatia import CATIA
class APPUI(QMainWindow):          
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui._setup_ui(self)
        # 动态为每个按钮绑定函数
        for i in range(1, 6):
            button_name = f"pushButton_{i}"
            button = getattr(self.ui, button_name, None)
            if button:
                button.clicked.connect(lambda checked, rw=i: self.BTNF(rw))
    def BTNF(self, rw):
            my_array = [1, 2, 3, 4, 5]
            orow=2
            start_col = 4
            TDM.inject_data(self.ui.tableWidget, orow, start_col, my_array)
         
def create_ui():
        Prog = QApplication(sys.argv)
        progwindow= APPUI()
        progwindow.show()
        sys.exit(Prog.exec_())
if __name__ == "__main__":
        create_ui()

        APPUI.BTNF(self, 0)