# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'mainWindow.ui'
# Created by: PyQt5 UI code generator 5.15.11

from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QGridLayout, QDockWidget
from PyQt5.QtGui import QPixmap
from src.UI3 import ClassUIM
# 全局变量
from src.Vars import global_var
# 自用全局变量
header_style = """
            QHeaderView::section {
                font-size: 18px;
                font-family: Dengxian;
                font-weight: bold;
                color: blue;
                background-color: #808080;
            }
            """
menu_names = ["菜单", "设置", "关于", "作者"]
header_labels = [
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
    "密度\nDensity",
    "更改\n密度",
    "质量\nMass",
    "厚度\nThickness"
]
menu_names = ["菜单", "设置", "关于", "作者"]
btnames = [
    "选择\n产品", "释放\n修改产品", "读取\n产品",
    "修改\n产品", "生成\n产品BOM", "初始化\n产品"
]
class Ui_MainWindow(object):
    def __init__(self, mainwindow):
        super().__init__()
        self.mainwindow = mainwindow
        self.UIM = ClassUIM(self.mainwindow)
        self._setup_ui()

    def _setup_ui(self):
        try:
            self.mainwindow.setObjectName("mainwindow")
            self.mainwindow.resize(1920, 1080)
            self.centralwidget = QtWidgets.QWidget(mainwindow)
            self.centralwidget.setObjectName("centralwidget")
            self.mainwindow.setCentralWidget(self.centralwidget)
            # 使用布局管理器
            main_layout, tab_layout, button_layout, pic_layout = self.init_layout()
            self.add_table(tab_layout)  # 1 创建表格部件并使用布局管理器
            self.add_buttons(button_layout)  # 2 创建按钮并使用布局管理器
            self.add_wxpic(pic_layout)    # 3 第三部分：图片
            self.add_statusbar()  # 3 设置状态栏
            self.add_menubar()  # 4 设置菜单栏
            self.centralwidget.setLayout(main_layout)
            self.mainwindow.resizeEvent = self.on_resize_event
            # self.retranslateUi(MainWindow)
        except Exception as e:
            print(f"Error setting up UI: {e}")

    def init_layout(self):
        main_layout = QHBoxLayout(self.centralwidget)
        tab_layout = QVBoxLayout()
        right_layout = QVBoxLayout()
        button_layout = QGridLayout()
        pic_layout = QHBoxLayout()
        main_layout.addLayout(tab_layout, 10)
        main_layout.addLayout(right_layout, 1)
        right_layout.addLayout(button_layout, 2)
        right_layout.addStretch(80)
        right_layout.addLayout(pic_layout, 2)
        return (main_layout,
                tab_layout,
                button_layout,
                pic_layout
                )

    def add_menubar(self):
        self.menubar = QtWidgets.QMenuBar(self.mainwindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1387, 32))
        self.menubar.setObjectName("menubar")
        self.mainwindow.setMenuBar(self.menubar)
        for i, name in enumerate(menu_names):
            menu = QtWidgets.QMenu(self.menubar)
            menu.setObjectName(f"oMenu_{i}")
            menu.setTitle(name)
            setattr(self, f"oMenu_{i}", menu)
            self.menubar.addAction(menu.menuAction())
        QtCore.QMetaObject.connectSlotsByName(mainwindow)

    def add_buttons(self, button_layout):
        for i, desc in enumerate(btnames):
            button = self._create_button(i, desc)
            setattr(self, f"pushButton_{i}", button)
            button.setFixedSize(180, 180)
            font = QtGui.QFont()
            font.setPointSize(12)  # 设置字体大小为12
            button.setFont(font)
            button_layout.addWidget(button, i // 2, i % 2)

    def _create_button(self, index, desc):
        button = QPushButton(self.centralwidget)
        button.setText(desc)
        button.setObjectName(f"pushButton_{index}")
        return button

    def add_table(self, tab_layout):
        self.tableWidget = QtWidgets.QTableWidget()
        crow = 30
        ccol = 14
        self.tableWidget.setRowCount(crow)  # 设置表格的行数
        self.tableWidget.setColumnCount(ccol)  # 设置表格的列数
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setHorizontalHeaderLabels(header_labels)
        self.tableWidget.horizontalHeader().setStyleSheet(header_style)
        self.tableWidget.itemChanged.connect(
            lambda: self.adjust_tab_width(self.tableWidget))
        self.adjust_tab_width(self.tableWidget)  # 调整列宽
        tab_layout.addWidget(self.tableWidget)
        

    def adjust_tab_width(self, table_widget):
        header = table_widget.horizontalHeader()
        col_count = table_widget.columnCount()
        header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        max_widths = [max(table_widget.horizontalHeader().sectionSize(
            col), table_widget.columnWidth(col)) for col in range(col_count)]
        total_max_width = sum(max_widths)
        avail_width = table_widget.viewport().width()

        if total_max_width > avail_width * 1.0:  # 允许20%溢出
            header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
            header.setCascadingSectionResizes(True)
            for col, width in enumerate(max_widths):
                table_widget.setColumnWidth(col, width)
        else:
            header.setSectionResizeMode(QtWidgets.QHeaderView.Fixed)
            if total_max_width and avail_width:
                allocated = 0
                for col in range(col_count - 1):
                    table_widget.setColumnWidth(
                        col, int(avail_width * (max_widths[col] / total_max_width)))
                    allocated += table_widget.columnWidth(col)
                table_widget.setColumnWidth(
                    col_count - 1, max(avail_width - allocated, max_widths[-1]))
            else:
                default_w = max(avail_width // col_count, 50)
                header.setSectionSizes([default_w] * col_count)
        self.tableWidget.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive)

    def add_statusbar(self):
        self.statusbar = QtWidgets.QStatusBar(
            self.mainwindow)  # 修改为 self.mainwindow
        self.statusbar.setObjectName("statusbar")
        self.mainwindow.setStatusBar(self.statusbar)  # 修改为 self.mainwindow
        self.statusbar.setFixedHeight(40)
        self.statusbar.setStyleSheet(
            "QStatusBar::item { font-size: 50px; font-family: Arial; }")
        self.statusbar.showMessage("ready for use")

    def add_wxpic(self, pic_layout):
        imgpath = 'resources/icons/wxpic.png'
        pixmap = QPixmap(imgpath)
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(
                pixmap.width() // 5, pixmap.height() // 5, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
            wxpic = QLabel()
            wxpic.setPixmap(scaled_pixmap)
            wxpic.setAlignment(QtCore.Qt.AlignCenter)
            pic_layout.addWidget(wxpic)
        else:
            print(f"Failed to load image: {imgpath}")


    # def on_dock_location_changed(self, mainwindow):
    #     dock_widget = self.sender()
    #     if self.mainwindow:
    #         location = self.mainwindow.dockWidgetArea(dock_widget)
    #         if location in [QtCore.Qt.LeftDockWidgetArea, QtCore.Qt.RightDockWidgetArea, QtCore.Qt.TopDockWidgetArea, QtCore.Qt.BottomDockWidgetArea]:
    #             dock_widget.setFloating(False)
    #             dock_widget.showMaximized()
    def on_resize_event(self, event):
        # 窗口尺寸变化时调用 _adjust_tab_width 方法
        self.adjust_tab_width(self.tableWidget)
        # 调用原始的 resizeEvent 方法
        return QtWidgets.QMainWindow.resizeEvent(self.mainwindow, event)
