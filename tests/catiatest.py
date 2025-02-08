from PyQt5 import  QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from src.UI2 import Ui_MainWindow  # 假设 UI 文件位于 src/UI2.py 中
import sys
from src.catia_processor import ClassPDM
from src.data_processor import ClassTDM 

import win32com.client
from pycatia import catia

def test1():
    PDM=ClassPDM()
    oprd=PDM.rootPrd
    print(f"Button {oprd.name} clicked")
    print({oprd.part_number})
    print("123")
    
if __name__ == '__main__':
    test1()
