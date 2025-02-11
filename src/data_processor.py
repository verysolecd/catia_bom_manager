from PyQt5.QtWidgets import QTableWidgetItem

# 全局变量
extract_cols = [3, 4, 5, 6, 10, 11]
bom_cols = [0, 1, 2, 3, 4, 6, 7, 10, 11]
inject_cols = [0, 2, 4, 6, 8, 10, 12, 13,]
class ClassTDM():
    def __init__(self, tableWidget):
        self.tableWidget = tableWidget

    def read_table(self):
        """
        从 tableWidget 中读取数据并返回一个二维数组
        """
        # data = []
        # row_count = self.tableWidget.rowCount()
        # col_count = self.tableWidget.columnCount()

        # for row in range(row_count):
        #     row_data = []
        #     for col in range(col_count):
        #         item = self.tableWidget.item(row, col)
        #         if item is not None:
        #             row_data.append(item.text())
        #         else:
        #             row_data.append('')
        #     data.append(row_data)

        # return data
    def inject_data(self, row, data):
        # if len(data) == len(inject_cols):
        for col_index, value in zip(inject_cols, data):
            oItem = QTableWidgetItem(str(value))
            self.tableWidget.setItem(row, col_index, oItem)
