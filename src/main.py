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

            #     On Error Resume Next
            # Set prd2wt = whois2rv() '弹窗说明读取对象
            # If Err.Number <> 0 Then
            # Err.Clear
            #  aMsgBox "没有要读取的产品"
            # Exit Sub
            # End If
            # On Error GoTo 0
            # Dim Prd2Read: Set Prd2Read = prd2wt
            # Set rng = xlsht.Range(xlsht.Cells(3, 1), xlsht.Cells(50, 14)): rng.ClearContents
            #     currRow = startrow
            #     Arry2sht infoPrd(Prd2Read), xlsht, currRow
            # Set children = Prd2Read.Products
            #     For i = 1 To children.Count
            #      currRow = i + startrow
            #      Arry2sht infoPrd(children.Item(i)), xlsht, currRow
            #     Next
            # Set Prd2Read = Nothing












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
