from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow
import sys
from src.catia_processor import ClassPDM
from src.data_processor import ClassTDM

import win32com.client
# from pycatia import CATIA


class APP(QMainWindow):
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
        # print("Loaded buttons:", [btn.objectName() for btn in self.buttons])

    def BTNF(self, rw):
        print(f"Button {rw} clicked")
        if rw == 1:
            # TDM.init_template()
            b_dict = {}
            root_prd = None
            ini_prd(root_prd, b_dict)
            pass
        elif rw == 2:
            # TDM.read_selected()
            pass
        elif rw == 3:
            # TDM.modify_selected()
            pass
        elif rw == 4:
            PDM = ClassPDM()
            oPrd = PDM.rootPrd
            my_array = PDM.attDefault(oPrd)
            orow = 1
            start_col = 2
            print({start_col})
            TDM = ClassTDM(self.ui.tableWidget)
            TDM.inject_data(orow, start_col, my_array)  # self.ui.tableWidget,

            # pass
        elif rw == 5:
            # TDM.select_to_modify()
            pass
        elif rw == 6:
            # TDM.select_to_modify()
            pass


def create_ui():
    Prog = QApplication(sys.argv)
    progwindow = APP()
    progwindow.show()
    sys.exit(Prog.exec_())


if __name__ == "__main__":
    create_ui()
