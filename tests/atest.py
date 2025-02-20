# atest.py

import win32com.client
import win32api
import win32con


def test_1():

    catia = win32com.client.GetActiveObject("catia.Application")
    if catia:
        rootprd = catia.ActiveDocument.Product
        catia.visible = True
    oprd = rootprd.products.item(1)
    qty = _countPrd(oprd)
    print(qty)


def _countPrd(oPrd):
    count = 0
    parent = getattr(oPrd.parent, 'parent', None)
    if parent and hasattr(parent, 'products'):
        for i in range(1, parent.products.count + 1):
            bros = parent.products.item(i)
            if bros.PartNumber == oPrd.PartNumber:
                count += 1
    return count or 1  # 默认返回1表示没有父级的情况


if __name__ == '__main__':
    test_1()
