from pycatia import catia
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow
import sys
from itertools import count
from functools import partial
from src.data_processor import ClassTDM
from src.catia_processor import ClassPDM

import win32com.client



class APP(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui._setup_ui(self)
        self._setup_buttons()

    def _setup_buttons(self):
        self.buttons = self._get_Buttons()  # 1. 动态发现按钮
        self._connect_button_handlers()  # 2. 绑定智能事件处理

    def _get_Buttons(self):
        return [
            btn for i in count(0)
            if (btn := getattr(self.ui, f"pushButton_{i}", None)) is not None
        ][:6]  # 安全截断最多6个按钮

    def _connect_button_handlers(self):
        for idx, btn in enumerate(self.buttons):
            # 使用partial避免闭包问题
            btn.clicked.connect(partial(self._handle_button_action, idx))

    def _handle_button_action(self, button_id):
        handler = getattr(self, f"_handle_button_{button_id}", None)
        if handler and callable(handler):
            try:
                handler()
            except Exception as e:
                self._log_error(f"按钮 {button_id} 操作失败: {str(e)}")
        else:
            self._log_error(f"未定义按钮 {button_id} 的处理方法")

    def _log_error(self, message):
        print(f"错误: {message}")

    def handle_button_0(self):
        # b_dict = {}
        # root_prd = None
       # ini_prd(root_prd, b_dict)
       pass

    def handle_button_1(self):
        pass  # 暂时不做任何操作

    def handle_button_2(self):
        pass

    def handle_button_3(self):
        PDM = ClassPDM()
        oPrd = PDM.rootPrd
        my_array = PDM.attDefault(oPrd)
        orow = 0
        start_col = 0
        TDM = ClassTDM(self.ui.tableWidget)
        TDM.inject_data(orow, start_col, my_array)

    def handle_button_4(self):
        pass

    def handle_button_5(self):
        pass

def create_ui():
    Prog = QApplication(sys.argv)
    progwindow = APP()
    progwindow.show()
    sys.exit(Prog.exec_())

if __name__ == "__main__":
    create_ui()
