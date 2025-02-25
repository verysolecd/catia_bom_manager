# atest.py

import win32com.client
import win32api
import win32con
import os
import subprocess
import pandas as pd  # 添加pandas导入

import numpy as np


from src.data_processor import ClassTDM
from src.catia_processor import ClassPDM
from src.catia_processor import CATerror



class test1(ClassPDM):
    
    def __init__(self):
        super().__init__()  # 必须调用父类初始化
        # 这里可以添加子类特有的初始化
        self.test_data = []



    def test_1(self):  # 添加缺失的self参数
        try:
            # 使用父类连接方法替代直接获取CATIA对象
            self.connect_to_catia()  # 调用继承自ClassPDM的方法
            rootprd = self.catia.ActiveDocument.Product  # 使用父类已有属性
        except CATerror as e:
            raise RuntimeError(f"初始化失败: {str(e)}")

        data = self.recurPrd(rootprd, 1, 1)  # 修正调用方式
        # print(data)
        self.inject_sht(data)

    def recurPrd(self, oPrd, LV, counter, all_rows_data=None):  # 修正参数顺序
        if all_rows_data is None:
            all_rows_data = []
        row_data = [counter, LV, *self.info_Prd(oPrd)]
        all_rows_data.append(row_data)
        if oPrd.Products.Count > 0:  # 修正COM属性大小写
            for i in range(1, oPrd.Products.Count + 1):
                child = oPrd.Products.Item(i)  # 修正方法名大小写
                all_rows_data = self.recurPrd(
                    child, LV + 1, counter + 1, all_rows_data)  # 修正参数传递顺序
        return all_rows_data

    def inject_sht(self, data):

        # 将二维数组转换为DataFrame（假设bomhead在父类中定义）
        df = pd.DataFrame(data, columns=["序号", "层级", *self.bomhead])

        temp_file_path = "temp_bom.xlsx"
        df.to_excel(temp_file_path, index=False)  # 一次性写入整个DataFrame

        try:
            if os.name == 'nt':
                os.startfile(temp_file_path)
            elif os.name == 'posix':
                subprocess.call(['xdg-open', temp_file_path])
            elif os.name == 'darwin':
                subprocess.call(['open', temp_file_path])
        except Exception as e:
            print(f"无法打开文件: {e}")

        input("按回车键删除临时文件...")

        try:
            os.remove(temp_file_path)
        except Exception as e:
            print(f"无法删除临时文件: {e}")

        return self._workbook


if __name__ == '__main__':
    test1().test_1()  # 创建实例后调用
