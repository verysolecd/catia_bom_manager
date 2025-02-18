# atest.py

import win32com.client
import pywintypes


att = [None] * 6
attNames = [
    "cm",
    "iBodys",
    "iMaterial",
    "iDensity",
    "iMass",
    "iThickness"
]

allPN = {}

iType = {
    attNames[0]: "String",
    attNames[1]: "list",
    attNames[2]: "String",
    attNames[3]: "Density",
    attNames[4]: "Mass",
    attNames[5]: "Length"
}


class CATIAProcessor:
    def __init__(self):
        self.catia = None

    def connect_to_catia(self):
        try:
            self.catia = win32com.client.GetActiveObject("catia.Application")
            return self.catia
        except pywintypes.com_error as e:
            raise Exception("无法连接到CATIA，请确保CATIA已启动")

    def init_refPrd(self, oPrd):
        refprd = oPrd.ReferenceProduct
        colls = refprd.UserRefProperties
        for i in range(2, 6):  # 修改循环条件以包含所有需要处理的属性
            if self._att_Obj_Value(colls, attNames[i])[0] is None:
                if i == 2:
                    att[i] = colls.CreateString(attNames[i], "")
                elif 4 <= i <= 5:
                    att[i] = colls.CreateDimension(
                        attNames[i], iType[attNames[i]], 0)

    def _att_Obj_Value(self, collection, itemName):
        try:
            att = collection.Item(itemName)
            att_value = att.Value
            return [att, att_value]
        except Exception:
            return [None, None]


def test_1():
    processor = CATIAProcessor()
    catia = processor.connect_to_catia()
    if catia:
        oprd = catia.ActiveDocument.Product
        processor.init_refPrd(oprd)


if __name__ == '__main__':
    test_1()
