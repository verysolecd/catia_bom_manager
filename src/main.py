#  Python库类
# from pycatia import catia

from PyQt5.QtWidgets import QApplication, QMainWindow
# from itertools import count
from functools import partial
from tkinter import messagebox

from PyQt5.QtWidgets import QMessageBox
# COM类

import sys

#  自建类
from src.UI2 import ClassUI
from src.data_processor import ClassTDM
from src.catia_processor import ClassPDM
from src.catia_processor import CATIAConnectionError
from src.UI3 import ClassUIM

# 全局变量
from Vars import global_var


class ClassApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = ClassUI(self)
        self.tableWidget = self.ui.tableWidget
        self.statusbar = self.ui.statusbar
        self.setup_buttons()
        self.UIM = ClassUIM()
        self.TDM = ClassTDM(self.ui.tableWidget)
        self.PDM = ClassPDM()
        self.catia = None
        self.ui.pushButton_0.setEnabled(False)
        # if global_var.Prd2Rw is None:
        #     self.ui.pushButton_1.setEnabled(False)
        self.ui.pushButton_4.setEnabled(False)
        self.ui.pushButton_5.setEnabled(False)

    def setup_buttons(self):
        self.buttons = self.get_Buttons()  # 1. 动态发现按钮
        self.connect_button_handlers()  # 2. 绑定智能事件处理

    def get_Buttons(self):
        return [
            btn for i in range(10)  # 使用有限范围代替无限count
            if (btn := getattr(self.ui, f"pushButton_{i}", None)) is not None
        ][:7]

    def connect_button_handlers(self):
        for idx, btn in enumerate(self.buttons):
            try:
                # 使用partial避免闭包问题
                btn.clicked.connect(partial(self.handle_clicks, idx))
                # print(f"成功连接按钮 {idx} 的点击事件")
            except Exception as e:
                print(f"连接按钮 {idx} 的点击事件时出错: {e}")

    def handle_clicks(self, button_id):
        self.catia = None
        if self.catia is None:
            try:
                self.catia = self.PDM.connect_to_catia()
                if self.catia is not None:
                    pass
                    # QMessageBox.information(self, "成功", "CATIA 连接成功，请继续执行操作。")
            except CATIAConnectionError as e:
                # 处理连接失败：记录日志、提示用户或尝试启动CATIA
                msg = f"错误: {e}，请打开catia和你的文档，再点击按钮继续操作。"
                # print(f"错误: {e}")
                self.catia = None
                QMessageBox.critical(self, "错误", msg)  # 处理其他未预期的异常
                return  # 停止运行
            except Exception as e:
                self.catia = None
                msg = f"未知错误: {e}，请检查后重新运行"
                # print(msg)
                QMessageBox.critical(self, "错误", msg)  # 处理其他未预期的异常
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
        self.ui.pushButton_0.setEnabled(False)
        # self.root_or_select()

    def handle_button_1(self):  # 释放产品
        if global_var.Prd2Rw is None:
            msg = "当前未选择待修改产品"
        else:
            msg = f"释放产品成功: {global_var.Prd2Rw.partnumber}"
            global_var.Prd2Rw = None
            self.TDM.clear_table()
            self.UIM.adjust_tab_width(self.tableWidget)
            self.UIM.set_table_readonly(self.tableWidget)  # 设置只读
        QMessageBox.information(self, "提示", msg)

    def handle_button_2(self):  # 读取产品
        global_var.Prd2Rw = self.PDM.catia.activedocument.product
        oprd = global_var.Prd2Rw
        data = self.PDM.attDefault(oprd)
        oRow = 0
        self.TDM.inject_data(oRow, data)
        if oprd.Products.Count > 0:
            for product in oprd.Products:
                oRow += 1
                data = self.PDM.attDefault(product)
                self.TDM.inject_data(oRow, data)
        self.UIM.set_table_readonly(self.tableWidget)  # 设置只读

    def handle_button_3(self):     # 修改产品
        oRow = 0
        data = self.TDM.extract_data(oRow)
        self.PDM.attModify(global_var.Prd2Rw, data)

        for product in global_var.Prd2Rw.Products:
            oRow += 1
            data = self.TDM.extract_data(oRow)
            self.PDM.attModify(product, data)

    def handle_button_4(self):  # 初始化产品
        pass
        # oPrd = self.catia.root_or_select()
        # self.PDM.catia

    def handle_button_5(self):
        pass
        # 如果你暂时没有具体的实现，可以使用 pass 作为占位符

    def handle_button_6(self):
        self.TDM.clear_table()
        self.UIM.adjust_tab_width(self.tableWidget)
        self.UIM.set_table_readonly(self.tableWidget)  # 设置只读





    def root_or_select(self):
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
    APP = ClassApp()
    APP.setWindowTitle("catia bom manager--作者：键盘造车手")
    APP.show()
    sys.exit(Prog.exec_())


if __name__ == "__main__":
    StartAPP()
