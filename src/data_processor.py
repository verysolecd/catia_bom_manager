from PyQt5.QtWidgets import QTableWidgetItem
import logging
# 全局变量
eCols = [1, 3, 5, 7]  # ，9，11]
iCols = [0, 2, 4, 6, 12, 12, 10, 13]
bom_cols = [0, 1, 2, 3, 4, 6, 7, 10, 11]


class ClassTDM():
    def __init__(self, tableWidget):
        self.tableWidget = tableWidget

    def inject_data(self, row, data):
        # if len(data) == len(iCols):
        # for col_index, value in zip(iCols, data):
        #     oItem = QTableWidgetItem(str(value))
        #     self.tableWidget.setItem(row, col_index, oItem)

        for value, col in zip(data, iCols):
            oItem = QTableWidgetItem(str(value))
            self.tableWidget.setItem(row, col, oItem)


    def extract_data(self, row):
        try:
            return [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) is not None else "" for col in eCols]
        except Exception as e:
            print(f"提取数据时出错: {str(e)}")
            return [""] * len(eCols)  # 返回默认值列表

    def clear_table(self):
        # 先将行数和列数设置为 0
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)
        # 重新设置行数和列数
        crow = 30
        ccol = 14
        self.tableWidget.setRowCount(crow)
        self.tableWidget.setColumnCount(ccol)
