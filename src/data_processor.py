from PyQt5.QtWidgets import QTableWidgetItem
from src.UI2 import Ui_MainWindow
# 全局变量
inject_cols = [1, 2, 7, 8, 9, 12]
extract_cols = [3, 4, 5, 6, 10, 11]
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
    def inject_data(self, row, start_col, data):
        for ind, aItem in enumerate(array):
            item = QTableWidgetItem(str(aItem))
            self.tableWidget.setItem(row, start_col + ind * 2, item)
            toUIm = Ui_MainWindow()
            toUIm.set_table_readonly(self.tableWidget)
            print(f"已经写入第{row}行数据")
        """
        将数组数据写入到 QTableWidget 中指定的行和列
        :param table_widget: QTableWidget 实例
        :param oRow: 要写入数据的行索引
        :param data: 包含 6 个元素的数组
        """

        if len(data) != len(columns):
            raise ValueError("数据长度必须为 6，以匹配指定的列数")
        for col_index, value in zip(columns, data):
            oItem = QTableWidgetItem(str(value))
            table_widget.setItem(row, col_index, oItem)
