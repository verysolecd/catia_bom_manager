from PyQt5.QtWidgets import QTableWidgetItem
import logging
from contextlib import contextmanager
import logging
import pandas as pd
import os
import subprocess
from tempfile import NamedTemporaryFile
from PyQt5.QtWidgets import QMessageBox
# 全局变量
cols_to_extract = [1, 3, 5, 7, 9, 11]
cols_to_inject = [0, 2, 4, 6, 12, 12, 8, 13, 10]
bom_cols = [0, 1, 2, 3, 4, 6, 7, 10, 11]
# bom_head = [
#     "No./n编号",
#     "Layout/n层级",
#     "PN/n零件号",
#     "Nomenclature/n英文名称",
#     "Definition/n中文名称",
#     "Picture/n图像",
#     "Quantity/n数量(PCS)",
#     "Weight/n单质量",
#     "Total Weight/n总质量",
#     "",
#     "Material/n材料",
#     "Thickness/n厚度(mm)",
#     "TS/n抗拉",
#     "YS/n屈服",
#     "EL/n延伸率",]

bom_head = [
    "No./n编号",
    "Layout/n层级",
    "PN/n零件号",
    "Nomenclature/n英文名称",
    "Definition/n中文名称",
    "Picture/n图像",
    "Quantity/n数量(PCS)",
    "Weight/n单质量",
    # "Total Weight/n总质量",
    # "",
    "Material/n材料",
    "Thickness/n厚度(mm)",
    "Density/n密度",]
# "TS/n抗拉",
# "YS/n屈服",
# "EL/n延伸率",

bom_cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]

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

    def generate_bom(self, data):
        try:
            # 使用传入的 data 创建 DataFrame
            df = pd.DataFrame(data, columns=bom_head)

            # 使用 NamedTemporaryFile 创建临时文件
            with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                temp_path = tmp.name
                # 使用 pandas 导出 Excel
                df.to_excel(temp_path, index=False, engine='openpyxl')

                # 自动打开文件
                if os.name == 'nt':
                    os.startfile(temp_path)
                elif os.name == 'posix':
                    subprocess.call(['xdg-open', temp_path])
                elif os.name == 'darwin':
                    subprocess.call(['open', temp_path])
            result = QMessageBox.information(None, "提示",
                                             "请完成对BOM的操作后保存，然后继续操作。\n"
                                             "文件路径：" + temp_path,
                                             QMessageBox.Ok)

            if result == QMessageBox.Ok:
                # 用户点击确定后，尝试删除临时文件
                try:
                    os.remove(temp_path)
                    logging.info(f"临时文件 {temp_path} 已成功删除。")
                except Exception as e:
                    logging.warning(f"临时文件清理失败: {str(e)}")

        except Exception as e:
            # 导入 logging 模块
            logging.error(f"BOM生成失败: {str(e)}")
            raise
        finally:
            # 清理临时文件
            if 'temp_path' in locals():
                try:
                    # os.remove(temp_path)
                    print("临时文件已删除")
                except Exception as e:
                    logging.warning(f"临时文件清理失败: {str(e)}")
        return temp_path
