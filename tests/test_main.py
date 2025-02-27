import logging
import pandas as pd
import os
import subprocess
from tempfile import NamedTemporaryFile

def generate_bom(self, data):
    try:
        # 使用传入的 data 创建 DataFrame
        df = pd.DataFrame(data, columns=[
            "序号", "层级", "零件号", "英文名称", "中文名称",
            "图像", "数量", "单质量", "总质量", "材料",
            "厚度", "抗拉", "屈服", "延伸率", "备注"
        ])

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

    except Exception as e:
        # 导入 logging 模块
        logging.error(f"BOM生成失败: {str(e)}")
        raise
    finally:
        # 清理临时文件
        if 'temp_path' in locals():
            try:
                os.remove(temp_path)
            except Exception as e:
                logging.warning(f"临时文件清理失败: {str(e)}")
    return temp_path