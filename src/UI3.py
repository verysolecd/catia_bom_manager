# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'mainWindow.ui'
# Created by: PyQt5 UI code generator 5.15.11

from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QGridLayout, QDockWidget, QTableWidget, QTableWidgetItem, QStatusBar
from PyQt5.QtGui import QPixmap
# 全局变量
from src.Vars import global_var

class ClassUIM(object):
    def __init__(self, mainwindow):
        self.mainWindow = mainwindow
        self.tablewidget = self.mainWindow.tableWidget
        self.statusbar = mainwindow.statusbar
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
        table_widget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Interactive)  # 修改为 table_widget
    def set_table_readonly(self, table_widget):
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

    def update_statusbar(self):
        if global_var.prd_2rw is None:
            self.statusbar.showMessage("当前未选择产品")
        else:
            msg = global_var.prd_2rw.name
            self.statusbar.showMessage(msg)

    def on_resize_event(self, event):
        # 窗口尺寸变化时调用 _adjust_tab_width 方法
        self.adjust_tab_width(self.tablewidget)
        # 调用原始的 resizeEvent 方法
        return QtWidgets.QMainWindow.resizeEvent(self.mainWindow, event)
