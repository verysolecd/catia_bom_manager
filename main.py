#  Python库类
import logging
from PyQt5.QtWidgets import QApplication, QMainWindow
# from itertools import count
from functools import partial
from PyQt5.QtWidgets import QMessageBox
# COM类
import sys

# 将 src 目录添加到 sys.path
src_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
sys.path.append(src_path)


#  自建类
from src.UI2 import ClassUI
from src.data_processor import ClassTDM
from src.catia_processor import ClassPDM
from src.catia_processor import CATerror
from src.UI3 import ClassUIM

# 全局变量
from src.Vars import gVar


class ClassApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.UI = ClassUI(self)
        self.tableWidget = self.UI.tableWidget
        self.statusbar = self.UI.statusbar
        self.setup_buttons()
        self.UIM = ClassUIM()
        self.TDM = ClassTDM(self.UI.tableWidget)
        self.PDM = ClassPDM()
        self.catia = None
    def setup_buttons(self):
        self.buttons = self.get_Buttons()
        self.connect_button_handlers()

    def get_Buttons(self):
        return [
            btn for i in range(10)
            if (btn := getattr(self.UI, f"pushButton_{i}", None)) is not None
        ][:7]

    def connect_button_handlers(self):
        for idx, btn in enumerate(self.buttons):
            try:
                btn.clicked.connect(partial(self.handle_clicks, idx))
            except Exception as e:
                print(f"连接按钮 {idx} 的点击事件时出错: {e}")

    def handle_clicks(self, button_id):
        if self.catia is None:
            try:
                self.catia = self.PDM.connect_to_catia()
                if self.catia is not None:
                    pass
            except CATerror as e:
                self.catia = None
                QMessageBox.information(self, "错误", "CATIA连接失败，请打开CATIA文档")
                return
            except Exception as e:
                self.catia = None
                self.log_error(f"未知错误: {e}，请检查后重新运行")
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
        QMessageBox.information(self, "错误", message)

    def handle_button_0(self):  # 选择产品
        self.root_or_select()
        imsg = f"你选择的是{gVar.Prd2Rw.PartNumber}" if gVar.Prd2Rw is not None else "未选择产品"
        QMessageBox.information(self, "提示", imsg)

    def handle_button_1(self):  # 释放产品
        msg = "当前未选择待修改产品" if gVar.Prd2Rw is None else f"释放产品成功: {gVar.Prd2Rw.PartNumber}"
        if gVar.Prd2Rw is not None:
            gVar.Prd2Rw = None
            self.handle_button_6()
            # self.TDM.clear_table()
            # self.UIM.adjust_tab_width(self.tableWidget)
            # self.UIM.set_table_readonly(self.tableWidget)
            QMessageBox.information(self, "提示", msg)

    def handle_button_2(self):  # 读取产品
        if gVar.Prd2Rw is None:
            QMessageBox.information(self, "提示", "当前未选择待修改产品")
            return
        oprd = gVar.Prd2Rw
        self.TDM.inject_data(0, self.PDM.info_Prd(oprd))
        for i, product in enumerate(oprd.Products, start=1):
            self.TDM.inject_data(i, self.PDM.info_Prd(product))
        self.UIM.set_table_readonly(self.tableWidget)  # 设置只读

    def handle_button_3(self):     # 修改产品
        if gVar.Prd2Rw is None:
            QMessageBox.information(self, "提示", "当前未选择待修改产品")
            return
        oRow = 0
        data = self.TDM.extract_data(oRow)
        self.PDM.attModify(gVar.Prd2Rw, data)
        for product in gVar.Prd2Rw.Products:
            oRow += 1
            data = self.TDM.extract_data(oRow)
            self.PDM.attModify(product, data)

    def handle_button_4(self):  # 初始化产品
        if gVar.Prd2Rw is None:
            QMessageBox.information(self, "提示", "当前未选择待修改产品")
            return
        try:
            oprd = gVar.Prd2Rw
            self.PDM.init_Product(oprd)
        except Exception as e:
            self.log_error(f"产品初始化失败: {str(e)}")
            imsg = f"产品初始化失败: {str(e)}"
            QMessageBox.information(self, "提示", imsg)

    def handle_button_5(self):  # 生成Bom
        try:
            oprd = gVar.Prd2Rw
            LV = 1
            data = self.PDM.recurPrd(oprd, LV)
            print(data)
            self.TDM.generate_bom(data)

        except Exception as e:
            self.log_error(f"产品Bom生成失败: {str(e)}")
            imsg = f"产品Bom生成失败: {str(e)}"
            QMessageBox.information(self, "提示", imsg)
        pass

    def handle_button_6(self):  # 清空表格
        self.TDM.clear_table()
        self.UI.set_head(self.tableWidget)
        self.UIM.adjust_tab_width(self.tableWidget)
        self.UIM.set_table_readonly(self.tableWidget)  # 设置只读

    def root_or_select(self):
        reply = QMessageBox.question(self, '选择操作产品及其子产品', '是: 选择要修改的产品\n否: 修改根产品\n取消: 退出选择',
                                     QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Cancel)
        try:
            if reply == QMessageBox.Yes:
                self.catia.visible = True
                gVar.Prd2Rw = self.PDM.selPrd()
                gVar.Prd2Rw.ApplyWorkMode(2)
            elif reply == QMessageBox.No:
                gVar.Prd2Rw = self.PDM.catia.activedocument.product
                gVar.Prd2Rw.ApplyWorkMode(2)
            else:
                gVar.Prd2Rw = None
                return
        except Exception as e:
            self.log_error(f"产品选择出错，请检查: {str(e)}")


def StartAPP():
    Prog = QApplication(sys.argv)
    APP = ClassApp()
    APP.setWindowTitle("catia bom manager--作者：键盘造车手")
    APP.show()
    sys.exit(Prog.exec_())


if __name__ == "__main__":
    StartAPP()
