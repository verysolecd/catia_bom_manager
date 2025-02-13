from PyQt5.QtWidgets import QTableWidgetItem
from UI3 import ClassUIM

# 全局变量
extract_cols = [1, 3, 5, 7]  # ，9，11]
bom_cols = [0, 1, 2, 3, 4, 6, 7, 10, 11]
inject_cols = [0, 2, 4, 6, 8, 10, 12, 13,]


class ClassTDM():
    def __init__(self, tableWidget):
        self.tableWidget = tableWidget

    def inject_data(self, row, data):
        # if len(data) == len(inject_cols):
        for col_index, value in zip(inject_cols, data):
            oItem = QTableWidgetItem(str(value))
            self.tableWidget.setItem(row, col_index, oItem)


    def extract_data(self, row):
        try:
            return [self.tableWidget.item(row, col).text() if self.tableWidget.item(row, col) is not None else "" for col in extract_cols]
        except Exception as e:
            print(f"提取数据时出错: {str(e)}")
            return [""] * len(extract_cols)  # 返回默认值列表

    def clear_table(self):
        # 先将行数和列数设置为 0
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(0)
        # 重新设置行数和列数
        crow = 30
        ccol = 14
        self.tableWidget.setRowCount(crow)
        self.tableWidget.setColumnCount(ccol)
