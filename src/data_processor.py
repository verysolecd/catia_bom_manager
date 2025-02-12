from PyQt5.QtWidgets import QTableWidgetItem

# 全局变量
extract_cols = [3, 4, 5, 6, 10, 11]
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
        return [self.tableWidget.item(row, col).text() for col in extract_cols]
