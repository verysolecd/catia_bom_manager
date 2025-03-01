


# 获取当前文件所在目录

import sys
import os

from PyQt5 import QtCore, QtWidgets, QtGui
from src.Vars import gVar

# # 全局变量
# from src.Vars import gVar

class ClassUIM(object):
    def __init__(self):
        super().__init__()


    def adjust_tab_width(self, table_widget):
        header = table_widget.horizontalHeader()
        col_count = table_widget.columnCount()
        header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        max_widths = [max(table_widget.horizontalHeader().sectionSize(
            col), table_widget.columnWidth(col)) for col in range(col_count)]
        total_max_width = sum(max_widths)
        avail_width = table_widget.viewport().width()

        if total_max_width > avail_width * 1.0:
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
        table_widget.horizontalHeader().setSectionResizeMode(
            QtWidgets.QHeaderView.Interactive)

    def set_table_readonly(self, tableWidget):
        for col in gVar.read_only_cols:
            for row in range(tableWidget.rowCount()):
                item = tableWidget.item(row, col)
                if item is None:
                    item = QtWidgets.QTableWidgetItem()
                    tableWidget.setItem(row, col, item)
                color = 201
                item.setBackground(QtGui.QColor(color, color, color))
                # 设置单元格为只读
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)

    
