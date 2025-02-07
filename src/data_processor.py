from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem


class TDM():
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


    def inject_data(self, tableWidget, row, start_col, array):
        for col, cell_data in enumerate(array):
            item = QTableWidgetItem(str(cell_data))
            tableWidget.setItem(row, start_col + col, item)
            print(f"Set item at row {row}, col {start_col + col} with value {cell_data}")
