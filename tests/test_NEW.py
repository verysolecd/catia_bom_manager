# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'mainWindow.ui'
# Created by: PyQt5 UI code generator 5.15.11

from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QGridLayout, QDockWidget
from PyQt5.QtGui import QPixmap

# 全局变量
prd_2rw = None
menu_names = ["菜单", "设置", "关于", "作者"]


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        try:
            MainWindow.setObjectName("MainWindow")
            MainWindow.resize(1920, 1080)
            self.centralwidget = QtWidgets.QWidget(MainWindow)
            self.centralwidget.setObjectName("centralwidget")
            MainWindow.setCentralWidget(self.centralwidget)
            main_layout, tab_layout, button_layout, pic_layout = self.iniLayout()
            self.centralwidget.setLayout(main_layout)
            # 1 创建表格部件并使用布局管理器
            self.add_table(tab_layout)
            # 2 创建按钮并使用布局管理器
            self.add_buttons(button_layout)
            # 3 第三部分：图片
            self.add_wxpic(pic_layout)
            # 3 设置状态栏
            self.add_statusbar(MainWindow)
            # 4 设置菜单栏
            self.add_menubar(MainWindow)
            # self.retranslateUi(MainWindow)
        except Exception as e:
            print(f"Error setting up UI: {e}")

    def iniLayout(self):
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
        return main_layout, tab_layout, button_layout, pic_layout

    def add_menubar(self, MainWindow):
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1387, 32))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        for i, name in enumerate(menu_names):
            menu = QtWidgets.QMenu(self.menubar)
            menu.setObjectName(f"oMenu_{i}")
            menu.setTitle(name)
            setattr(self, f"oMenu_{i}", menu)
            self.menubar.addAction(menu.menuAction())
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def add_buttons(self, button_layout):
        for i, text in enumerate([
            "选择\n产品", "释放\n修改产品", "初始化\n产品",
            "读取\n产品", "修改\n产品", "生成\n产品BOM"
        ]):
            button = self.create_button(i, text)
            font = QtGui.QFont()
            font.setPointSize(12)  # 设置字体大小为12
            button.setFont(font)
            button.setFixedSize(180, 180)
            setattr(self, f"pushButton_{i + 1}", button)
            button_layout.addWidget(button, i // 2, i % 2)

    def create_button(self, index, text):
        button = QPushButton(self.centralwidget)
        button.setText(text)
        button.setObjectName(f"pushButton_{index + 1}")

        return button

    def add_table(self, tab_layout):
        self.tableWidget = QtWidgets.QTableWidget()
        crow = 30
        ccol = 14
        self.tableWidget.setRowCount(crow)
        self.tableWidget.setColumnCount(ccol)
        self.tableWidget.setObjectName("tableWidget")
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
            "密度\nMaterial",
            "更改\n密度",
            "质量\nMass",
            "厚度\nThickness"
        ]
        self.tableWidget.setHorizontalHeaderLabels(header_labels)

        header = self.tableWidget.horizontalHeader()
        self.TabReadOnly(self.tableWidget)  # 设置只读

        # 设置表头样式
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

        # 初始调整列宽并延迟布局计算
        self.tableWidget.resizeColumnsToContents()
        QtCore.QTimer.singleShot(0, self.adjust_table_columns)

        # 安装事件过滤器以响应视图大小变化
        self.tableWidget.viewport().installEventFilter(self)

        # 内容变化时触发调整
        self.tableWidget.itemChanged.connect(self.on_item_changed)

        # 将表格添加到布局
        tab_layout.addWidget(self.tableWidget)

    def adjust_table_columns(self):
        header = self.tableWidget.horizontalHeader()
        # 临时设置为Interactive以正确计算列宽
        header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)
        self.tableWidget.resizeColumnsToContents()

        total_width = sum(self.tableWidget.columnWidth(col) for col in range(self.tableWidget.columnCount()))
        available_width = self.tableWidget.viewport().width()

        if total_width < available_width:
            header.setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        else:
            header.setSectionResizeMode(QtWidgets.QHeaderView.Interactive)

    def on_item_changed(self, item):
        self.adjust_table_columns()

    def eventFilter(self, source, event):
        if source is self.tableWidget.viewport() and event.type() == QtCore.QEvent.Resize:
            self.adjust_table_columns()
        return super().eventFilter(source, event)

        tab_layout.addWidget(self.tableWidget)
    def TabReadOnly(self, tableWidget):
        self.tableWidget = tableWidget
        readonly_cols = [0, 2, 4, 6, 8, 10, 12, 13]
        for col in readonly_cols:
            for row in range(tableWidget.rowCount()):
                item = tableWidget.item(row, col)
                if item is None:
                    item = QtWidgets.QTableWidgetItem()
                    tableWidget.setItem(row, col, item)
                color = 192
                item.setBackground(QtGui.QColor(color, color, color))
                # 设置单元格为只读
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)

    def add_statusbar(self, MainWindow):
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.statusbar.setFixedHeight(40)
        self.statusbar.setStyleSheet(
            "QStatusBar::item { font-size: 50px; font-family: Arial; }")
        self.statusbar.showMessage("ready for use")
        self.update_statusbar()

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



    def update_statusbar(self):
        global prd_2rw
        if prd_2rw is None:
            self.statusbar.showMessage("未选择产品")
        else:
            msg = prd_2rw.name
            self.statusbar.showMessage(msg)

    def on_dock_location_changed(self):
        # 获取当前停靠窗口的位置
        dock_widget = self.sender()
        location = MainWindow.dockWidgetArea(dock_widget)

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
