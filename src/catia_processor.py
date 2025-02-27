# -*- coding: utf-8 -*-
# this module is to manage catia data to get or define the product attributes


import win32com.client
import pywintypes
import win32api
import win32con
from openpyxl import Workbook
import os
import subprocess
# 全局参数

all_rows_data = None
counter = 0
bom_cols = [0, 1, 2, 3, 4, 6, 7, 10, 11]
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

bom_head = [
    "No./n编号",
    "Layout/n层级",
    "PN/n零件号",
    "Nomenclature/n英文名称",
    "Definition/n中文名称",
    "Picture/n图像",
    "Quantity/n数量(PCS)",
    "Weight/n单质量",
    "Total Weight/n总质量",
    "",
    "Material/n材料",
    "Thickness/n厚度(mm)",
    "TS/n抗拉",
    "YS/n屈服",
    "EL/n延伸率",]
bom_cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]




class CATerror(Exception):
    """catia错误异常"""
    pass

class ClassPDM():
    def __init__(self):
        self.catia = None
        self.active_document = None
        self.rootPrd = None
        self.data_bom = []
        global all_rows_data, counter
        all_rows_data = []
        counter = [0]  # 使用列表保持计数器状态

    def connect_to_catia(self):
        try:
            self.catia = win32com.client.GetActiveObject("catia.Application")
            return self.catia
        except pywintypes.com_error as e:
            raise CATerror("请打开catia和产品")
        except Exception as e:
            raise CATerror(f"错误: {str(e)}，请检查")


    def amsg(imsg):
        win32api.MessageBox(0, imsg, "Error", win32con.MB_OK)

    def selPrd(self):
        if not (self.catia and self.catia.ActiveDocument):
            raise CATerror("请打开catia和产品")
        try:
            self.catia.Visible = True
            self.catia.ActiveWindow.WindowState = 0
            self.catia.Top = True
        except AttributeError:
            pass
        try:
            oSel = self.catia.ActiveDocument.Selection
            oSel.Clear()
            if (status := oSel.SelectElement2(("Product",), "请选择产品", False)) == "cancel":
                self.amsg("用户取消")
                return None
            return oSel.Item(1).LeafProduct.ReferenceProduct \
                if oSel.Count2 == 1 else \
                (self.amsg("请只选择一个产品"), None)[1]
        except Exception as e:
            self.amsg(f"选择失败: {str(e)}")
            raise CATerror("选择操作异常") from e

    def init_refPrd(self, refPrd):
        colls = refPrd.UserRefProperties  # 第一部分，初始化产品的属性
        for i in range(2, 6):  # 创建属性
            if self._att_Value(colls, attNames[i])[0] is None:
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
            if self._att_Value(colls, attNames[i])[0] is None:
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
            if self._att_Value(colls, attNames[i])[0] is None:
                if i == 1:
                    att[i] = colls.Createlist(attNames[i], "")
                if 4 <= i <= 5:
                    att[i] = colls.CreateDimension(
                        attNames[i], iType[attNames[i]], 0)
            else:
                att[i] = colls.Item(attNames[i])

        if not self._att_Value(att[1], mbd.Name)[0]:
            att[1].ValueList.Add(mbd)

        # 修改为正确的语法
        for i in range[2, 3, 4, 5]:
            pubs = refPrd.Publications
            if self._att_Value(pubs, attNames[i])[0] is None:
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
        oRule = colls.Item(eStr[0]) if self._att_Value(colls, eStr[0])[
            0] else colls.CreateProgram(eStr[0], "cal of mass", eStr[1])
        if oRule.Value != eStr[1]:
            oRule.Modify(eStr[1])
        gName = [
            ("CMAS", attNames[4], "cm\\iMass"),
            ("CTK", attNames[5], "cm\\iThickness")]
        for rName, att_index, rSource in gName:
            rTarget = refPrd.UserRefProperties.Item(att_index)
            oFml = colls.Item(rName) if self._att_Value(colls, rName)[
                0] else colls.CreateFormula(rName, " ", rTarget, rSource)
            if oFml.Value != rSource:
                oFml.Modify(rSource)

    def _att_Value(self, collection, itemName):
        try:
            att = collection.Item(itemName)
            att_value = att.Value
            return [att, att_value]
        except pywintypes.com_error as e:
            return [None, "N/A"]
            # raise CATerror("属性错误")
        except Exception as e:
        except pywintypes.com_error as e:
            return [None, "N/A"]
            # raise CATerror("属性错误")
        except Exception as e:
            return [None, "N/A"]

    def init_Product(self, oPrd):
        refPrd = oPrd.ReferenceProduct
        if refPrd.PartNumber not in allPN:
            allPN[refPrd.PartNumber] = 1
            self.init_refPrd(refPrd)
        if refPrd.Products.count > 0:
            for product in refPrd.Products:
                self.init_Product(product)
        allPN.clear()


    def attDefault(self, oPrd):
        # 修改为正确的大小写
        refPrd = oPrd.ReferenceProduct
        att_default = [refPrd.PartNumber,
                       refPrd.Nomenclature, refPrd.Definition, oPrd.Name]
        # 0   1   2   3
        # pn nom def name
        return att_default

    def info_Prd(self, oPrd):
        # 修改为正确的大小写
        refPrd = oPrd.ReferenceProduct
        infoPrd = [None]*9
        infoPrd[0] = refPrd.PartNumber
        infoPrd[1] = refPrd.Nomenclature
        infoPrd[2] = refPrd.Definition
        infoPrd[3] = oPrd.Name
        infoPrd[4] = self._countPrd(oPrd)

        usrp = refPrd.UserRefProperties


        infoPrd[5] = self._att_Value(usrp, "iMass")[1]
        # try:
        #     mass_value = round(float(self._att_Value(usrp, "iMass")[1]), 3)
        #     infoPrd[5] = f"{mass_value:.3f}"
        # except Exception:
        #     infoPrd[5] = "N/A"

        infoPrd[6] = self._att_Value(usrp, "iMaterial")[1]
        infoPrd[7] = self._att_Value(usrp, "iThickness")[1]
        try:
            colls = refPrd.Parent.Part.Parameters.RootParameterSet.ParameterSets.Item(
                "cm").DirectParameters
            infoPrd[8] = self._att_Value(colls, "iDensity")[1]
        except Exception:
            infoPrd[8] = "N/A"
        return infoPrd

    def _countPrd(self, oPrd):
        count = 0
        parent = getattr(oPrd.parent, 'parent', None)
        if parent and hasattr(parent, 'products'):
            for i in range(1, parent.products.count + 1):
                bros = parent.products.item(i)
                if bros.PartNumber == oPrd.PartNumber:
                    count += 1
        return count or 1  # 默认返回1表示没有父级的情况

    def attModify(self, oPrd, data):
        refPrd = oPrd.referenceproduct
        att_names = ["PartNumber", "Nomenclature", "Definition", "Name"]
        for att_name, new_value in zip(att_names, data):
            try:
                att_value = getattr(refPrd, att_name)
                if new_value and new_value != "" and new_value != att_value:
                    setattr(oPrd, att_name, new_value)
            except IndexError:
                pass
        try:
            # data[4]  # iMateiral
            if data[4] and data[4] != "" and data[4] != refPrd.UserRefProperties.Item("iMaterial").Value:
                refPrd.UserRefProperties.Item("iMaterial").Value = data[4]
        except Exception as e:
            pass
        try:
            # data[5]  # iDensity
            if data[5] and data[5] != "" and data[5] != refPrd.parent.part.parameters.rootparameterset.parametersets.Item(
                    "cm").DirectParameters.Item("iDensity").Value:
                refPrd.parent.part.parameters.rootparameterset.parametersets.Item(
                    "cm").DirectParameters.Item("iDensity").Value = data[5]
        except Exception as e:
            pass


    def Assmass(oPrd):
        total = 0
        children = oPrd.products
        if oPrd.products.count > 0:
            for i in range(1, children.count + 1):
                total += Assmass(children.item(i)) + children.item(
                    i).reference_product.user_ref_properties.item("iMass").value
            oPrd.reference_product.user_ref_properties.item(
                "iMass").value = total
        else:
            total = oPrd.reference_product.user_ref_properties.item(
                "iMass").value
        return total

    def recurPrd(self, oPrd, LV):
        # 初始化参数
        global all_rows_data, counter
        if all_rows_data is None:
            all_rows_data = []
            counter = [0]  # 使用列表保持计数器状态

        counter[0] += 1  # 递增全局计数器
        row_data = [counter[0], LV, *self.info_Prd(oPrd)]
        all_rows_data.append(row_data)

        # 递归处理子产品（修正属性和方法名大小写）
        if oPrd.Products.Count > 0:
            bdict = {}
            for idx in range(1, oPrd.Products.Count + 1):
                child = oPrd.Products.Item(idx)
                part_number = child.PartNumber
                if part_number not in bdict:
                    bdict[part_number] = 1
                    self.recurPrd(child, LV + 1)
        return all_rows_data


if __name__ == "__main__":

    PDM = ClassPDM()
    catia = PDM.connect_to_catia()
    rootPrd = catia.ActiveDocument.Product
    data = PDM.recurPrd(rootPrd, 1)
    print(data)
