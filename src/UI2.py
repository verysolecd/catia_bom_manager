# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'mainWindow.ui'
# Created by: PyQt5 UI code generator 5.15.11

from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PyQt5.QtGui import QPixmap

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        try:
            MainWindow.setObjectName("MainWindow")
            MainWindow.resize(1920, 1080)
            self.centralwidget = QtWidgets.QWidget(MainWindow)
            self.centralwidget.setObjectName("centralwidget")
            main_layout = QHBoxLayout(self.centralwidget)
            # 创建表格部件并使用布局管理器
            self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
            self.tableWidget.setRowCount(24)  # 设置表格的行数
            self.tableWidget.setColumnCount(14)  # 设置表格的列数
            self.tableWidget.setObjectName("tableWidget")
            main_layout.addWidget(self.tableWidget, stretch=8)
            right_layout = QVBoxLayout()
            header_labels = [
            "质量\nMass",
            "厚度\nThickness",
            "零件号\nPartnumber",
            "更改\n件号",
            "英文名称\nNomenclature",
            "更改\n英文名",
            "中文名称\nDefinition",
            "更改\n中文名",
            "实例名\nInstanceName",
            "更改\n实例名",
            "材料\nmaterial",
            "定义\n材料",
            "密度\nMaterial",
            "更改\n密度"
            ]
            self.tableWidget.setHorizontalHeaderLabels(header_labels)
            # 设置表格头部背景颜色为灰色
            header = self.tableWidget.horizontalHeader()
            header.setStyleSheet("QHeaderView::section { background-color: #808080; color: white; }")
            for row in range(self.tableWidget.rowCount()):
                for col in [0, 1, 2, 4, 6, 8, 10, 12]:  # 第1、3、5列不可编辑
                    item = QtWidgets.QTableWidgetItem()
                    item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    item.setBackground(QtGui.QColor(200, 200, 200))  # 浅灰色
                    self.tableWidget.setItem(row, col, item)
            self.tableWidget.resizeColumnsToContents()
            # 设置列的调整模式为自动拉伸
            for col in range(self.tableWidget.columnCount()):
                header.setSectionResizeMode(col, QtWidgets.QHeaderView.Stretch)
            # 创建按钮并使用布局管理器
            button_layout = QVBoxLayout()
            for i, text in enumerate([
               "1 选择待修改产品", "2 释放待修改产品", "3 初始化产品模板",
            "4 读取选择的产品", "5 修改选择的产品", "6 遍历生成产品BOM"
            ]):
                button = self.create_button(i, text,)                
                setattr(self, f"pushButton_{i+1}", button)  # 关键修复
                button_layout.addWidget(button)
            right_layout.addLayout(button_layout, stretch=8)           
            # 第三部分：图片
            imgpath = 'resources/icons/IDcard.jpg'
            pixmap = QPixmap(imgpath)
            if not pixmap.isNull():
                scaled_pixmap = pixmap.scaled(250, 250, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
                wxpic = QLabel()
                wxpic.setPixmap(scaled_pixmap)
                wxpic.setAlignment(QtCore.Qt.AlignCenter)
                right_layout.addWidget(wxpic, stretch=2)
            else:
                print(f"Failed to load image: {imgpath}")

            main_layout.addLayout(right_layout, stretch=2)
            MainWindow.setCentralWidget(self.centralwidget)
            self.statusbar = QtWidgets.QStatusBar(MainWindow)
            self.statusbar.setObjectName("statusbar")
            MainWindow.setStatusBar(self.statusbar)
            self.menubar = QtWidgets.QMenuBar(MainWindow)
            self.menubar.setGeometry(QtCore.QRect(0, 0, 1387, 32))
            self.menubar.setObjectName("menubar")
            self.menucatia = QtWidgets.QMenu(self.menubar)
            self.menucatia.setObjectName("menucatia")
            self.menu = QtWidgets.QMenu(self.menubar)
            self.menu.setObjectName("menu")
            self.menu_2 = QtWidgets.QMenu(self.menubar)
            self.menu_2.setObjectName("menu_2")
            self.menu_3 = QtWidgets.QMenu(self.menubar)
            self.menu_3.setObjectName("menu_3")
            MainWindow.setMenuBar(self.menubar)
            self.menubar.addAction(self.menucatia.menuAction())
            self.menubar.addAction(self.menu.menuAction())
            self.menubar.addAction(self.menu_2.menuAction())
            self.menubar.addAction(self.menu_3.menuAction())           
            QtCore.QMetaObject.connectSlotsByName(MainWindow)
        except Exception as e:
            print(f"Error setting up UI: {e}")

    def create_button(self, index, text):
        button = QPushButton(self.centralwidget)
        button.setText(text)
        button.setObjectName(f"pushButton_{index + 1}")
        font = QtGui.QFont()
        font.setPointSize(14)  # 设置字体大小为12
        button.setFont(font)
        return button

    # def retranslateUi(self, MainWindow):
    #     _translate = QtCore.QCoreApplication.translate
    #     MainWindow.setWindowTitle(_translate("MainWindow", "程序窗口"))
    #     for i, text in enumerate([
    #         "选择待修改产品", "释放待修改产品", "初始化产品模板",
    #         "读取选择的产品", "修改选择的产品", "遍历生成产品BOM"
    #     ]):
    #         getattr(self, f"pushButton_{i + 1}").setText(_translate("MainWindow", text))
    #     self.menucatia.setTitle(_translate("MainWindow", "菜单"))
    #     self.menu.setTitle(_translate("MainWindow", "设置 "))
    #     self.menu_2.setTitle(_translate("MainWindow", "关于"))
    #     self.menu_3.setTitle(_translate("MainWindow", "作者"))


    from PyQt5 import  QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow  # 假设 UI 文件位于 src/UI2.py 中
import sys
#import src.catia_Processor as PDM
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
        for i in range(0, 5):
            button_name = f"pushButton_{i+1}"
            button = getattr(self.ui, button_name, None)
            if button:
                self.buttons.append(button)
        for i, button in enumerate(self.buttons):
            button.clicked.connect(lambda checked, idx=i+1:self.BTNF(idx))
    def BTNF(self, rw):
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
            orow=1
            start_col = 1
            TDM.inject_data(self.ui.tableWidget, orow, start_col, my_array)
            #pass
        elif rw == 5:
            # TDM.select_to_modify()
            pass
        elif rw == 6:
            # TDM.select_to_modify()
            pass
def create_ui():
        Prog = QApplication(sys.argv)
        progwindow= APPUI()
        progwindow.show()
        sys.exit(Prog.exec_())
if __name__ == "__main__":
        create_ui()


