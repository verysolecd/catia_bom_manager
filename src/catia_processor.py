# -*- coding: utf-8 -*-
# this module is to manage catia data to get or define the product attributes


import win32com.client
import pywintypes
import win32api
import win32con
import win32gui


# 全局参数
from src.Vars import gVar

allPN = {}
att = [None] * 6
attNames = ["cm",
            "iBodys",
            "iMaterial",
            "iDensity",
            "iMass",
            "iThickness"]

iType = {}
iType = {
    "cm": "String",
    "iBodys": "list",
    "iMaterial": "String",
    "iDensity": "Density",
    "iMass": "Mass",
    "iThickness": "Length"
}


class CATIAConnectionError(Exception):
    """连接CATIA失败的专用异常"""


class CATerror(Exception):
    """catia错误异常"""


class ClassPDM():
    def __init__(self):
        self.catia = None
        self.active_document = None
        self.rootPrd = None

    def connect_to_catia(self):
        try:
            # 连接CATIA的具体实现
            self.catia = win32com.client.GetActiveObject("catia.Application")
            return self.catia
        except pywintypes.com_error as e:
            self.catia = None
            raise CATIAConnectionError("你没开catia")

        return self.catia

    def selPrd(self) -> 'ProductReference':
        if not (self.catia and self.catia.ActiveDocument):
            raise CATIAConnectionError("CATIA未连接")
        try:
            self.catia.Visible = True
            self.catia.ActiveWindow.WindowState = 0
            self.catia.Top = True
        except AttributeError:
            pass
        try:
            selection = self.catia.ActiveDocument.Selection
            selection.Clear()
            if (status := selection.SelectElement2(
                    ("Product",), "请选择产品", False)) == "cancel":
                win32api.MessageBox(0, "用户取消", "提示", win32con.MB_OK)
                return None

            return selection.Item(1).LeafProduct.ReferenceProduct \
                if selection.Count2 == 1 else \
                (win32api.MessageBox(0, "请选择单个产品",
                 "错误", win32con.MB_OK), None)[1]

        except Exception as e:
            win32api.MessageBox(0, f"选择失败: {str(e)}", "错误", win32con.MB_OK)
            raise CATerror("选择操作异常") from e


    def init_refPrd(self, refPrd):
        colls = refPrd.UserRefProperties  # 第一部分，初始化产品的属性
        for i in range(2, 6):  # 创建属性
            if self._att_Obj_Value(colls, attNames[i])[0] is None:
                if i == 2:
                    att[i] = colls.CreateString(attNames[i], "")
                if 4 <= i <= 5:
                    att[i] = colls.CreateDimension(
                        attNames[i], iType[attNames[i]], 0)
        try:  # 第二部分，若有子part初始化part
            self._init_Prt(refPrd, attNames)
        except Exception:  # 若没有part则单独发布imass
            colls = refPrd.Publications
            i = 4
            if self._att_Obj_Value(colls, attNames[i])[0] is None:
                oRef = refPrd.CreateReferenceFromName(att[i].Name)
                colls.Add(attNames[i])
                colls.SetDirect(attNames[i], oRef)

    def _init_Prt(self, refPrd, attNames):
        self._init_Parameters(refPrd)
        self._init_Formula(refPrd)

    def _init_Parameters(self, refPrd):
        oPrt = refPrd.Parent.Part
        mbd = oPrt.MainBody
        oSet = oPrt.Parameters.RootParameterSet.ParameterSets
        try:
            colls = oSet.Item("cm").DirectParameters
        except Exception:
            oSet.CreateSet("cm")
            colls = oSet.Item("cm").DirectParameters

        for i in range(1, 5):  # 不存在则创建
            if self._att_Obj_Value(colls, attNames[i])[0] is None:
                if i == 1:
                    att[i] = colls.Createlist(attNames[i], "")
                if 4 <= i <= 5:
                    att[i] = colls.CreateDimension(
                        attNames[i], iType[attNames[i]], 0)
            else:
                att[i] = colls.Item(attNames[i])

        if not self._att_Obj_Value(att[1], mbd.Name)[0]:
            att[1].ValueList.Add(mbd)

        for i in range[3, 4, 5]:
            pubs = refPrd.Publications
            if self._att_Obj_Value(pubs, attNames[i])[0] is None:
                if i in [3, 4, 5]:
                    oref = refPrd.CreateReferenceFromName(att[i].Name)
                    oPub = pubs.Add(attNames[i])  # 添加发布
                    pubs.SetDirect(attNames[i], oref)  # 设置发布元素
                if i == 2:
                    att[i] = refPrd.UserRefProperties.Item(attNames[i])
                    oref = refPrd.CreateReferenceFromName(att[i].Name)
                    oPub = pubs.Add(attNames[i])  # 添加发布
                    pubs.SetDirect(attNames[i], oref)  # 设置发布元素

    def _init_Formula(self, refPrd):
        oPrt = refPrd.Parent.Part
        eStr = ["CalM",
                "let lst = cm\\iBodys; let V = 0; let Vol = 0; for i = 1 to lst.Size() do { V = smartVolume(lst.GetItem(i)); Vol = Vol + V; }; cm\\iMass = Vol * cm\\iDensity"]
        colls = oPrt.Relations
        oRule = colls.Item(eStr[0]) if self._att_Obj_Value(colls, eStr[0])[
            0] else colls.CreateProgram(eStr[0], "cal of mass", eStr[1])
        if oRule.Value != eStr[1]:
            oRule.Modify(eStr[1])
        gName = [
            ("CMAS", attNames[4], "cm\\iMass"),
            ("CTK", attNames[5], "cm\\iThickness")]
        for rName, att_index, rSource in gName:
            rTarget = refPrd.UserRefProperties.Item(att_index)
            oFml = colls.Item(rName) if self._att_Obj_Value(colls, rName)[
                0] else colls.CreateFormula(rName, " ", rTarget, rSource)
            if oFml.Value != rSource:
                oFml.Modify(rSource)


    def _att_Obj_Value(self, collection, itemName):
        try:
            att = collection.Item(itemName)
            att_value = att.Value
            return [att, att_value]
        except Exception:
            return [None, None]

    def init_Product(self, oPrd, allPN):
        refPrd = oPrd.ReferenceProduct
        if refPrd.PartNumber not in allPN:
            allPN[refPrd.PartNumber] = 1
            self.init_refPrd(refPrd)
        if refPrd.Products.count > 0:
            for product in refPrd.Products:
                self.init_Product(product, allPN)
        allPN.clear()


    def attDefault(self, oPrd):
        refPrd = oPrd.referenceproduct
        att_default = [refPrd.partnumber,
                       refPrd.nomenclature, refPrd.definition, oPrd.name]
        # 0   1   2   3
        # pn nom def name
        return att_default

    def info_Prd(self, oPrd):
        refPrd = oPrd.referenceproduct
        infoPrd = [None]*6
        infoPrd[0] = refPrd.PartNumber
        infoPrd[1] = refPrd.Nomenclature
        infoPrd[2] = refPrd.Defintion
        infoPrd[3] = oPrd.Name

        usrAtt = ["iMaterial",
                      "iDensity",
                      "iMass",
                      "iThickness"
                  ]


        infoPrd[4] = refPrd.userrefproperties.item("iMaterial").value

        infoPrd[5] = refPrd.userrefproperties.item("iDensity").value

        att_usp_Names = ["iMass", "iThickness", "iMass", "iThickness"]

        return infoPrd

    def _askValue(self, colls, myname):
        try:
            askValue = colls.Item(myname)
            return askValue
        except Exception as e:
            return "N\\A"

    def attModify(self, oPrd, data):
        refPrd = oPrd.referenceproduct
        att_names = ["PartNumber", "Nomenclature", "Definition", "Name"]
        for att_name, new_value in zip(att_names, data):
            try:
                att_value = getattr(refPrd, att_name)
                if new_value and new_value != att_value:
                    setattr(oPrd, att_name, new_value)
            # except AttributeError:
            #     raise CATerror("缺少属性")
            except IndexError:
                break  # 若 data 元素不足，直接结束循环



#
#
#     def recurPrd(oPrd, xlsht, oRowNb, LV):
#         bDict = {}
#         bomRowPrd(oPrd, LV, xlsht, oRowNb)
#         if oPrd.products.count > 0:
#             for i in range(1, oPrd.products.count + 1):
#                 if oPrd.products.item(i).part_number not in bDict:
#                     bDict[oPrd.products.item(i).part_number] = 1
#                     oRowNb += 1
#                     recurPrd(oPrd.products.item(i), xlsht, oRowNb, LV + 1)
#     def Assmass(oPrd):
#         total = 0
#         children = oPrd.products
#         if oPrd.products.count > 0:
#             for i in range(1, children.count + 1):
#                 total += Assmass(children.item(i)) + children.item(
#                     i).reference_product.user_ref_properties.item("iMass").value
#             oPrd.reference_product.user_ref_properties.item(
#                 "iMass").value = total
#         else:
#             total = oPrd.reference_product.user_ref_properties.item(
#                 "iMass").value
#         return total
#     def Dictbros(oPrd):
#         oDict = {}
#         for i in range(1, oPrd.products.count + 1):
#             pn = oPrd.products.item(i).part_number
#             if pn in oDict:
#                 oDict[pn] += 1
#             else:
#                 oDict[pn] = 1
#         return oDict

#
