from PyQt5.QtWidgets import QTableWidgetItem
import logging
# 全局变量
cols_to_extract = [1, 3, 5, 7, 9, 11]
cols_to_inject = [0, 2, 4, 6, 12, 12, 8, 13, 10]
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
            data = [self.tableWidget.item(row, col).text() if self.tableWidget.item(
                row, col) is not None else None for col in cols_to_extract]
            print(data)
            return data
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
        import os
        import subprocess
        from tempfile import NamedTemporaryFile

        try:
            # 创建新工作簿
            wb = Workbook()
            ws = wb.active

            # 添加标题（假设第一行已有标题）
            # ws.append(["序号", "零件号", "名称", "数量", "材料", ...])  # 根据实际列名添加

            # 写入数据从第二行开始
            for row_data in data_bom:
                ws.append(row_data)

            # 创建临时文件
            with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                temp_path = tmp.name
                wb.save(temp_path)

            # 自动打开文件
            if os.name == 'nt':
                os.startfile(temp_path)
            elif os.name == 'posix':
                subprocess.call(['xdg-open', temp_path])
            elif os.name == 'darwin':
                subprocess.call(['open', temp_path])

            # 等待用户确认
            QMessageBox.information(None, "提示",
                                    "请保存生成的BOM文件后，点击确定继续操作。\n"
                                    "文件路径：" + temp_path,
                                    QMessageBox.Ok)

        except Exception as e:
            logging.error(f"BOM生成失败: {str(e)}")
            raise
        finally:
            # 清理临时文件
            if 'temp_path' in locals():
                try:
                    os.remove(temp_path)
                except Exception as e:
                    logging.warning(f"临时文件清理失败: {str(e)}")

        return wb
