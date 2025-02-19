# atest.py

import win32com.client
import win32api
import win32con


def test_1():

    catia = win32com.client.GetActiveObject("catia.Application")
    if catia:
        oprd = catia.ActiveDocument.Product
        catia.visible = True
        win32api.MessageBox(0, "请选择产品", "提示", win32con.MB_OK)

if __name__ == '__main__':
    test_1()
