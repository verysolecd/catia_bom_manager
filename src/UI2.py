# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'mainWindow.ui'
# Created by: PyQt5 UI code generator 5.15.11

from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QGridLayout, QDockWidget
from PyQt5.QtGui import QPixmap

# 全局变量
prd_2rw = None
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

class Ui_MainWindow(object):
    def _setup_ui(self, main_window):
        try:
            main_window.setObjectName("MainWindow")
            main_window.resize(1920, 1080)
            self.centralwidget = QtWidgets.QWidget(main_window)
            self.centralwidget.setObjectName("centralwidget")
            main_window.setCentralWidget(self.centralwidget)
            # 使用布局管理器
            (main_layout,
             tab_layout,
             button_layout,
             pic_layout) = self._init_layout()
            self._add_table(tab_layout)  # 1 创建表格部件并使用布局管理器
            self._add_buttons(button_layout)  # 2 创建按钮并使用布局管理器
            self._add_wxpic(pic_layout)    # 3 第三部分：图片
            self._add_statusbar(main_window)  # 3 设置状态栏
            self._add_menubar(main_window)  # 4 设置菜单栏
            self.centralwidget.setLayout(main_layout)
            # self.retranslateUi(MainWindow)
        except Exception as e:
            print(f"Error setting up UI: {e}")

    def _init_layout(self):
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

    def _add_menubar(self, main_window):
        self.menubar = QtWidgets.QMenuBar(main_window)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1387, 32))
        self.menubar.setObjectName("menubar")
        main_window.setMenuBar(self.menubar)
        for i, name in enumerate(menu_names):
            menu = QtWidgets.QMenu(self.menubar)
            menu.setObjectName(f"oMenu_{i}")
            menu.setTitle(name)
            setattr(self, f"oMenu_{i}", menu)
            self.menubar.addAction(menu.menuAction())
        QtCore.QMetaObject.connectSlotsByName(main_window)

    def _add_buttons(self, button_layout):
        for i, desc in enumerate([
            "选择\n产品", "释放\n修改产品", "初始化\n产品",
            "读取\n产品", "修改\n产品", "生成\n产品BOM"
        ]):
            button = self._create_button(i, desc)
            button.setFixedSize(180, 180)
            setattr(self, f"pushButton_{i}", button)
            font = QtGui.QFont()
            font.setPointSize(12)  # 设置字体大小为12
            button.setFont(font)
            button_layout.addWidget(button, i // 2, i % 2)

    def _create_button(self, index, desc):
        button = QPushButton(self.centralwidget)
        button.setText(desc)
        button.setObjectName(f"pushButton_{index}")
        return button

    def _add_table(self, tab_layout):
        self.tableWidget = QtWidgets.QTableWidget()
        crow = 30
        ccol = 14
        self.tableWidget.setRowCount(crow)  # 设置表格的行数
        self.tableWidget.setColumnCount(ccol)  # 设置表格的列数
        self.tableWidget.setObjectName("tableWidget")

        self.tableWidget.setHorizontalHeaderLabels(header_labels)
        # 设置表头字体为蓝色加粗体
        # 设置表格头部背景颜色为灰色
        header_style = """
            QHeaderView::section {
                font-size: 18px;
                font-family: Dengxian;
                font-weight: bold;
                color: blue;
                background-color: #808080;
            }
            """
        self.tableWidget.horizontalHeader().setStyleSheet(header_style)
        self._set_table_readonly(self.tableWidget)  # 设置只读
        self.tableWidget.itemChanged.connect(
            lambda: self._adjust_table_column_width(self.tableWidget))
        self._adjust_table_column_width(self.tableWidget)  # 根据表头内容调整列宽

        tab_layout.addWidget(self.tableWidget)

    def _adjust_table_column_width(self, table_widget):
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

    def _set_table_readonly(self, table_widget):
        readonly_cols = [0, 2, 4, 6, 8, 10, 12, 13]
        for col in readonly_cols:
            for row in range(table_widget.rowCount()):
                item = table_widget.item(row, col)
                if item is None:
                    item = QtWidgets.QTableWidgetItem()
                    table_widget.setItem(row, col, item)
                color = 192
                item.setBackground(QtGui.QColor(color, color, color))
                # 设置单元格为只读
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)

    def _add_statusbar(self, main_window):
        self.statusbar = QtWidgets.QStatusBar(main_window)
        self.statusbar.setObjectName("statusbar")
        main_window.setStatusBar(self.statusbar)
        self.statusbar.setFixedHeight(40)
        self.statusbar.setStyleSheet(
            "QStatusBar::item { font-size: 50px; font-family: Arial; }")
        self.statusbar.showMessage("ready for use")
        self._update_statusbar()

    def _add_wxpic(self, pic_layout):
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

    def _update_statusbar(self):
        global prd_2rw
        if prd_2rw is None:
            self.statusbar.showMessage("未选择产品")
        else:
            msg = prd_2rw.name
            self.statusbar.showMessage(msg)

    def on_dock_location_changed(self):
        # 获取当前停靠窗口的位置
        dock_widget = self.sender()
        location = self.main_window.dockWidgetArea(dock_widget)

        # 如果停靠窗口停靠在主窗口的边缘，将其最大化
        if location in [QtCore.Qt.LeftDockWidgetArea, QtCore.Qt.RightDockWidgetArea, QtCore.Qt.TopDockWidgetArea, QtCore.Qt.BottomDockWidgetArea]:
            dock_widget.setFloating(False)
            dock_widget.showMaximized()

    # def retranslateUi(self, MainWindow):
    #     _translate = QtCore.QCoreApplication.translate
    #     MainWindow.setWindowTitle(_translate("MainWindow", "程序窗口"))
    #     for menu_name, title in menu_names.items():
    #         menu = getattr(self, menu_name, None)
    #         if menu:
    #             menu.setTitle(_translate("MainWindow", title))
    # def retranslateUi(self, MainWindow):
    #     _translate = QtCore.QCoreApplication.translate
    #     MainWindow.setWindowTitle(_translate("MainWindow", "程序窗口"))
    #     for menu_name, title in menu_names.items():
    #         menu = getattr(self, menu_name, None)
    #         if menu:
    #             menu.setTitle(_translate("MainWindow", title))
