from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow
import sys
# import src.catia_Processor as PDM
from src.data_processor import TDM

import win32com.client
# from pycatia import CATIA



class APPUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # 动态为每个按钮绑定函数
        self.buttons = []  # 用于存储按钮对象的列表
        for i in range(6):
            button_name = f"pushButton_{i+1}"
            button = getattr(self.ui, button_name, None)
            if button:
                self.buttons.append(button)
        for idx, button in enumerate(self.buttons, start=1):
            button.clicked.connect(lambda checked, ind=idx: self.BTNF(ind))

        print("Loaded buttons:", [btn.objectName() for btn in self.buttons])

    def BTNF(self, rw):

        print(f"Button {rw} clicked")
        if rw == 1:
            # TDM.init_template()
            pass
        elif rw == 2:
            # TDM.read_selected()
            pass
        elif rw == 3:
            # TDM.modify_selected()
            pass
        elif rw == 4:
            my_array = [1, 2, 3, 4, 5]
            orow = 1
            start_col = 1
            TDM1 = TDM(self.ui.tableWidget)
            TDM1.inject_data(self.ui.tableWidget, orow, start_col, my_array)
            # pass
        elif rw == 5:
            # TDM.select_to_modify()
            pass
        elif rw == 6:
            # TDM.select_to_modify()
            pass


def create_ui():
    Prog = QApplication(sys.argv)
    progwindow = APPUI()
    progwindow.show()
    sys.exit(Prog.exec_())


if __name__ == "__main__":
    create_ui()
