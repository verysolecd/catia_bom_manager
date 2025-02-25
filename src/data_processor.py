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

        # 批量创建 QTableWidgetItem 对象
        items = [QTableWidgetItem(str(value)) for value in data]

        # 批量设置表格项
        for col_index, item in zip(cols_to_inject, items):
            self.tableWidget.setItem(row, col_index, item)



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

    def generate_bom(self, data_bom):

        from openpyxl import Workbook
        if not hasattr(self, '_workbook'):
            self._workbook = Workbook()
            self._worksheet = self._workbook.active
            self._worksheet.append(bom_head)
        for row_data in data_bom:
            self._worksheet.append(row_data)
        temp_file_path = "temp_bom.xlsx"
        self._workbook.save(temp_file_path)
        try:
            if os.name == 'nt':  # Windows 系统
                os.startfile(temp_file_path)
            elif os.name == 'posix':  # Linux 系统
                subprocess.call(['xdg-open', temp_file_path])
            elif os.name == 'darwin':  # macOS 系统
                subprocess.call(['open', temp_file_path])
        except Exception as e:
            print(f"无法打开文件: {e}")

        input("请手动保存 Excel 文件，保存完成后按回车键继续...")

        try:        # 删除临时文件
            os.remove(temp_file_path)
        except Exception as e:
            print(f"无法删除临时文件: {e}")
        return self._workbook
