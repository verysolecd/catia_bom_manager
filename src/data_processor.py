from PyQt5.QtWidgets import QTableWidgetItem
import logging
# 全局变量
cols_to_extract = [1, 3, 5, 7]  # ，9，11]
cols_to_inject = [0, 2, 4, 6, 12, 12, 8, 13, 11]
bom_cols = [0, 1, 2, 3, 4, 6, 7, 10, 11]


class ClassTDM():
    def __init__(self, tableWidget):
        self.tableWidget = tableWidget

    def inject_data(self, row, data):
        # if len(data) == len(cols_to_inject):
        # for col_index, value in zip(cols_to_inject, data):
        #     oItem = QTableWidgetItem(str(value))
        #     self.tableWidget.setItem(row, col_index, oItem)

        for value, col in zip(data, cols_to_inject):
            oItem = QTableWidgetItem(str(value))
            self.tableWidget.setItem(row, col, oItem)


    def extract_data(self, row):
        try:
            return [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) is not None else None for col in cols_to_extract]
        except Exception as e:
            return [None] * len(cols_to_extract)  # 错误的话全部返回为none

    def clear_table(self):
        # 先将行数和列数设置为 0
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)
        # 重新设置行数和列数
        crow = 30
        ccol = 14
        self.tableWidget.setRowCount(crow)
        self.tableWidget.setColumnCount(ccol)
