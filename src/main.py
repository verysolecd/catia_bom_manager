#  Python库类
from pycatia import catia
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from itertools import count
from functools import partial
import tkinter as tk
from tkinter import messagebox

from PyQt5.QtWidgets import QMessageBox
# COM类
import win32com.client
import sys


#  自建类
from src.UI2 import Ui_MainWindow
from src.data_processor import ClassTDM
from src.catia_processor import ClassPDM
from src.UI3 import ClassUIM

# 全局变量
from Vars import global_var

from PyQt5.QtWidgets import QMessageBox


class APP(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow(self)
        self.ui._setup_ui(self)
        self.tableWidget = self.ui.tableWidget
        self.setup_buttons()
        self.TDM = ClassTDM(self.ui.tableWidget)
        self.PDM = ClassPDM()
        self.UIM = ClassUIM(self.ui)

    def setup_buttons(self):
        self.buttons = self.get_Buttons()  # 1. 动态发现按钮
        self.connect_button_handlers()  # 2. 绑定智能事件处理

    def get_Buttons(self):
        return [
            btn for i in range(10)  # 使用有限范围代替无限count
            if (btn := getattr(self.ui, f"pushButton_{i}", None)) is not None
        ][:6]

    def connect_button_handlers(self):
        for idx, btn in enumerate(self.buttons):
            try:
                # 使用partial避免闭包问题
                btn.clicked.connect(partial(self.handle_clicks, idx))
                print(f"成功连接按钮 {idx} 的点击事件")
            except Exception as e:
                print(f"连接按钮 {idx} 的点击事件时出错: {e}")

    def handle_clicks(self, button_id):
        if not self.PDM.catia:
            self.catia = self.PDM.connect_to_catia()  # 尝试连接到 CATIA
            if self.catia:
                QMessageBox.information(self, "成功", "CATIA 连接成功，继续执行操作。")
            else:
                QMessageBox.critical(self, "错误", "CATIA 连接失败，请检查相关设置。")
                return
        self.button_run(button_id)


    def button_run(self, button_id):
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

    def handle_button_0(self):  # 选择产品
        self.select_this_Prd()

    def handle_button_1(self):  # 释放产品
        pass

    def handle_button_2(self):  # 读取产品
       

        my_array = PDM.attDefault(oPrd)
        orow = 0
        start_col = 0
        TDM = ClassTDM(self.ui.tableWidget)
        TDM.inject_data(orow, start_col, my_array)

    def handle_button_3(self):     # 修改产品
        pass

    def handle_button_4(self):  # 初始化产品
        oPrd = self.catia.select_this_Prd()
        self.PDM.catia

    def handle_button_5(self):    # 生成BOM
        pass

    def select_this_Prd(self):
        reply = QMessageBox.question(self, '选择操作产品及其子产品', '是: 选择要修改的产品\n否: 修改根产品\n取消: 退出选择',
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Cancel)
        try:
            if reply == QMessageBox.Yes:
                messagebox.showinfo(
                    "提示", "在catia中选择产品")
                self.PDM.selprd()
            elif reply == QMessageBox.No:
                global_var.Prd2rw = self.PDM.catia.activedocument.rootPrd
                print("获取到 rootPrd:", global_var.Prd2rw)
            else:
                return
        except Exception as e:
            self.log_error(f"产品选择出错，请检查: {str(e)}")



    def infoPrd(self, oPrd):  # 将函数正确缩进到类内部
        try:
            oArry = [88, self.PDM.attDefault(oPrd), 0, self.PDM.attUsp(oPrd)]
            return oArry
        except Exception as e:
            self.log_error(f"获取产品信息时出错: {str(e)}")

def StartAPP():
    Prog = QApplication(sys.argv)
    progwindow = APP()
    progwindow.show()
    sys.exit(Prog.exec_())

if __name__ == "__main__":
    StartAPP()
