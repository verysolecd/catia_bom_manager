import os
import subprocess
from openpyxl import Workbook


class YourClass:  # 假设代码在类中
    def recurPrd(self, oPrd, oRowNb, LV, inject_cols):
        if not hasattr(self, '_workbook'):
            self._workbook = Workbook()
            self._worksheet = self._workbook.active
            self._worksheet.append(bom_head)

        all_bom_rows = []  # 用于存储所有 BOM 行数据的数组
        self._collect_bom_rows(oPrd, LV, all_bom_rows, inject_cols)

        # 一次性将所有 BOM 行数据写入 Excel
        for row_data in all_bom_rows:
            self._worksheet.append(row_data)

        # 临时保存文件
        temp_file_path = "temp_bom.xlsx"
        self._workbook.save(temp_file_path)

        # 打开临时文件
        try:
            if os.name == 'nt':  # Windows 系统
                os.startfile(temp_file_path)
            elif os.name == 'posix':  # Linux 系统
                subprocess.call(['xdg-open', temp_file_path])
            elif os.name == 'darwin':  # macOS 系统
                subprocess.call(['open', temp_file_path])
        except Exception as e:
            print(f"无法打开文件: {e}")

        # 提示用户保存文件
        input("请手动保存 Excel 文件，保存完成后按回车键继续...")

        # 删除临时文件
        try:
            os.remove(temp_file_path)
        except Exception as e:
            print(f"无法删除临时文件: {e}")

        return self._workbook

    def _collect_bom_rows(self, oPrd, LV, all_bom_rows, inject_cols):

        prd_info = self.info_Prd(oPrd)
        row_data = [
            LV,  # 层级
            prd_info[0],  # 零件号
            prd_info[1],  # 名称
            prd_info[2],  # 定义
            prd_info[3],  # 实例名
            prd_info[4],  # 数量
            prd_info[5],  # 质量
            prd_info[6],  # 材料
            prd_info[7],  # 厚度
            prd_info[8]  # 密度
        ]
        # 根据 inject_cols 筛选需要的列数据
        filtered_row_data = [row_data[i]
                             for i in inject_cols if i < len(row_data)]
        all_bom_rows.append(filtered_row_data)

        if oPrd.products.count > 0:
            for i in range(1, oPrd.products.count + 1):
                child = oPrd.products.item(i)
                self._collect_bom_rows(
                    child, LV + 1, all_bom_rows, inject_cols)

