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
        self.setup_buttons()

    def setup_buttons(self):
        self.buttons = self.get_Buttons()  # 1. 动态发现按钮
        self.connect_button_handlers()  # 2. 绑定智能事件处理

    def get_Buttons(self):
        # return [self.ui.pushButton_0,
        #         self.ui.pushButton_1,
        #         self.ui.pushButton_2,
        #         self.ui.pushButton_3,
        #         self.ui.pushButton_4,
        #         self.ui.pushButton_5]


        return [
            btn for i in range(10)  # 使用有限范围代替无限count
            if (btn := getattr(self.ui, f"pushButton_{i}", None)) is not None
        ][:6]

    def connect_button_handlers(self):
        for idx, btn in enumerate(self.buttons):
            try:
                # 使用partial避免闭包问题
                btn.clicked.connect(partial(self.handle_button_action, idx))
                print(f"成功连接按钮 {idx} 的点击事件")
            except Exception as e:
                print(f"连接按钮 {idx} 的点击事件时出错: {e}")

    def handle_button_action(self, button_id):
        handler = getattr(self, f"handle_button_{button_id}", None)
        if handler and callable(handler):
            try:
                handler()
            except Exception as e:
                self.log_error(f"按钮 {button_id} 操作失败: {str(e)}")
        else:
            self.log_error(f"未定义按钮 {button_id} 的处理方法")

    def log_error(self, message):
        print(f"错误: {message}")

    def handle_button_0(self):
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


def StartAPP():
    Prog = QApplication(sys.argv)
    progwindow = APP()
    progwindow.show()
    sys.exit(Prog.exec_())

if __name__ == "__main__":
    StartAPP()
