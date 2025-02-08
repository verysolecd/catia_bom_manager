from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
from src.UI2 import Ui_MainWindow


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

    def inject_data(self, row, start_col, array):
        for ind, aItem in enumerate(array):
            item = QTableWidgetItem(str(aItem))
            self.tableWidget.setItem(row, start_col + ind * 2, item)
            toUIm = Ui_MainWindow()
            toUIm.TabReadOnly(self.tableWidget)
            print(f"已经写入第{row}行数据")
