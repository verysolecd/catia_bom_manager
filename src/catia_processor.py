# -*- coding: utf-8 -*-
# this module is to manage catia data to get or define the product attributes
import win32com.client
import pywintypes
# 全局参数
from Vars import gVar
att = [None] * 6
attNames = ["cm",
            "iBodys",
            "iMaterial",
            "iDensity",
            "iMass",
            "iThickness"]


allPN = {}

iType = {}
iType[attNames[0]] = "String"
iType[attNames[1]] = "list"
iType[attNames[2]] = "String"
iType[attNames[3]] = "Density"
iType[attNames[4]] = "Mass"
iType[attNames[5]] = "Length"




class CATIAConnectionError(Exception):
    """连接CATIA失败的专用异常"""
    pass
class CATerror(Exception):
    """catia错误异常"""
    pass

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
            # 捕获COM接口错误后抛出自定义异常
            raise CATIAConnectionError("你没开catia")
            self.catia = None
        return self.

    # 假设 allPN 是一个全局字典

    def init_refPrd(self, oPrd):
        refprd = oPrd.ReferenceProduct
        colls = refprd.UserRefProperties
        for i in range(2, 4, 5):
            if self._att_Obj_Value(colls, attNames[i])[0] is None:
                if i == 2:
                    att[i] = colls.CreateString(attNames[i], "")
                elif 4 <= i <= 5:
                    att[i] = colls.CreateDimension(
                        attNames[i], iType[attNames[i]], 0)

            # def init_my_ref_prd(o_prd):
            #     """
            #     初始化参考产品
            #     """
            #     ref_prd = o_prd.ReferenceProduct
            #     att_type = {}
            #     att_type[attNames[0]] = "String"
            #     att_type[attNames[1]] = "list"
            #     att_type[attNames[2]] = "String"
            #     att_type[attNames[3]] = "Density"
            #     att_type[attNames[4]] = "Mass"
            #     att_type[attNames[5]] = "Length"

            #     colls = ref_prd.UserRefProperties
            #     parameter_obj = [None] * 6

            #     for i in range(2, 6):
            #         if not has_att(colls, attNames[i]):
            #             if i == 2:
            #                 parameter_obj[i] = colls.CreateString(attNames[i], "TBD")
            #             elif 4 <= i <= 5:
            #                 parameter_obj[i] = colls.CreateDimension(
            #                     attNames[i], att_type[attNames[i]], 0)
            #         else:
            #             parameter_obj[i] = colls.Item(attNames[i])

            #     try:
            #         o_prt = ref_prd.Parent.Part
            #         ini_prt(o_prd, attNames)
            #     except Exception:
            #         colls = ref_prd.Publications
            #         i = 4
            #         if not has_att(colls, attNames[i]):
            #             o_ref = ref_prd.CreateReferenceFromName(parameter_obj[i].Name)
            #             o_pub = colls.Add(attNames[i])
            #             colls.SetDirect(attNames[i], o_ref)

            #     o_prd.Update()


     # 对于PRD创建参数，"iMaterial" "iMass" "iThickness"

    def _att_Obj_Value(self, collection, itemName):
        try:
            # 尝试获取指定名称的属性
            att = collection.item(itemName)
            att_value = att.value
        except Exception:
            # 若属性不存在
            attribute = None
            attribute_value = None
        result = [attribute, attribute_value]

        
        
        
        
        pass

    def init_Prd(self, oPrd, oDict):
        if oPrd.PartNumber not in allPN:
            allPN[oPrd.PartNumber] = 1
            init_refPrd(oPrd)
        if oPrd.Products.count > 0:
            for product in oPrd.Products:
                init_Prd(product, oDict)
        allPN.clear()



    # def selprd(self):
    #     self.catia.visible = True
    #     self.catia.message_box("请选择产品", 16)
    #     oSel = self.active_document.selection
    #     oSel.clear()
    #     filter_type = ("Product",)
    #     # 执行选择操作
    #     status = oSel.select_element2(filter_type, "请选择要读取的产品", False)
    #     if status == "cancel":
    #         self.catia.message_box("操作已取消", 16)
    #         oSel.clear()
    #         return None
    #     elif status == "Normal":
    #         if oSel.count2 == 1:
    #             selected_item = oSel.item(1)  # 获取选中的项目
    #             if hasattr(selected_item, 'leaf_product'):
    #                 leaf_product = selected_item.leaf_product
    #                 if hasattr(leaf_product, 'reference_product'):
    #                     reference_product = leaf_product.reference_product
    #                     oSel.clear()
    #                     return reference_product
    #                 else:
    #                     self.catia.message_box("选中的产品没有reference_product属性", 16)
    #             else:
    #                 self.catia.message_box("选中的对象不是有效的产品", 16)
    #         else:
    #             self.catia.message_box("请仅选择一个产品", 16)
    #     oSel.clear()
    #     return None

#     def whois2rv():
#         """
#         选择产品的函数
#         """
#         root = tk.Tk()
#         root.withdraw()
#         infobox = messagebox.askyesnocancel("请选择产品", "“是”选择产品，“否”读取根产品，“取消退出”")
#         if infobox is None:
#             return None
#         elif infobox:
#             try:
#                 return selprd()
#             except Exception:
#                 return None
#         elif not infobox:
#             caa = catia()
#             return caa.active_document.product
#         return None

#     def ReadAllPrd(self, oPrd):
#         self.refprd = oPrd.reference_product

    def attDefault(self, oPrd):
        refprd = oPrd.ReferenceProduct
        att_default = [refprd.partnumber,
                       refprd.nomenclature, refprd.definition, oPrd.name]
        # 0   1   2   3
        # pn nom def name
        return att_default



    def attUsp(self, oPrd):
        # attNames = ["cm", "iBodys", "iMaterial",
        #         "iDensity", "iMass", "iThickness"]
        # 0    1       2          3       4      5
        # cm iBodys iMaterial iDensity iMass
        attNames = ["iMaterial", "iDensity", "iMass", "iThickness"]

        refprd = oPrd.reference_product
        att_usp = [None]*6
        att_usp[0] = ""
        for i, att_name in enumerate(attNames):
            colls = refprd.user_ref_properties
            try:
                att_value = getattr(refprd, att_name)
                att_usp[i+1] = att_value
            except AttributeError:
                att_usp[i+1] = "N/A"
        for i in range(2, 5):
            if i in [2, 4, 5]:
                colls = refprd.user_ref_properties
                att_usp[i] = self._askValue(colls, attNames[i])
            elif i == 3:
                try:
                    oPrt = refprd.parent.part
                    colls = oPrt.parameters.root_parameter_set.parameter_sets.item(
                        attNames[1]).direct_parameters
                    att_usp[i] = self._askValue(
                        colls, attNames[i])
                except Exception as e:
                    att_usp[i] = "N\\A"
        return att_usp
    def _askValue(colls, myname):
        try:
            askValue = colls.Item(myname)
            return askValue
        except Exception as e:
            return "N/A"

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

    def init_prd(self, oPrd):
        refPrd = oPrd.reference_product
        att_names = ["PartNumber", "Nomenclature", "Definition", "Name"]


    # def infoPrd(self, oPrd):
    #     oArry = [88, self.attDefault(oPrd), 0, self.attUsp(oPrd)]
    #     return oArry


#   def bomRowPrd(oPrd, LV):
#         oDict = {}
#         QTy = 1
#         if isinstance(oPrd.parent, pycatia.product_structure_interfaces.products):
#             oDict = Dictbros(oPrd.parent.parent)
#             QTy = oDict[oPrd.part_number]
#         return [LV, attDefault(oPrd), QTy, attUsp(oPrd)]

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

#     from pycatia import catia


# # 假设这是全局变量
# allPN = {}
# attNames = ["iMaterial", "iMass", "iThickness", "iDensity", "iBodys"]

    # def has_att(colls, name):
    #     """
    #     检查集合中是否存在指定名称的元素
    #     """
    #     try:
    #         if isinstance(colls, (list, tuple)):
    #             for item in colls:
    #                 if hasattr(item, 'Name') and item.Name == name:
    #                     return True
    #         elif hasattr(colls, 'Item'):
    #             colls.Item(name)
    #             return True
    #     except Exception:
    #         return False
    #     return False


# def ini_prd(o_prd, o_dict):
#     """
#     初始化产品
#     """
#     if o_prd.PartNumber not in allPN:
#         allPN[o_prd.PartNumber] = 1
#         init_my_ref_prd(o_prd)
#     if o_prd.Products.Count > 0:
#         for product in o_prd.Products:
#             ini_prd(product, o_dict)
#     allPN.clear()


# def init_my_ref_prd(o_prd):
#     """
#     初始化参考产品
#     """
#     ref_prd = o_prd.ReferenceProduct
#     att_type = {}
#     att_type[attNames[0]] = "String"
#     att_type[attNames[1]] = "list"
#     att_type[attNames[2]] = "String"
#     att_type[attNames[3]] = "Density"
#     att_type[attNames[4]] = "Mass"
#     att_type[attNames[5]] = "Length"

#     colls = ref_prd.UserRefProperties
#     parameter_obj = [None] * 6

#     for i in range(2, 6):
#         if not has_att(colls, attNames[i]):
#             if i == 2:
#                 parameter_obj[i] = colls.CreateString(attNames[i], "TBD")
#             elif 4 <= i <= 5:
#                 parameter_obj[i] = colls.CreateDimension(
#                     attNames[i], att_type[attNames[i]], 0)
#         else:
#             parameter_obj[i] = colls.Item(attNames[i])

#     try:
#         o_prt = ref_prd.Parent.Part
#         ini_prt(o_prd, attNames)
#     except Exception:
#         colls = ref_prd.Publications
#         i = 4
#         if not has_att(colls, attNames[i]):
#             o_ref = ref_prd.CreateReferenceFromName(parameter_obj[i].Name)
#             o_pub = colls.Add(attNames[i])
#             colls.SetDirect(attNames[i], o_ref)

#     o_prd.Update()


# def ini_prt(o_prd, attNames):
#     """
#     初始化零件
#     """
#     ref_prd = o_prd.ReferenceProduct
#     o_prt = ref_prd.Parent.Part
#     mbd = o_prt.MainBody

#     try:
#         colls = o_prt.Parameters.RootParameterSet.ParameterSets.Item(
#             "cm").DirectParameters
#     except Exception:
#         colls = o_prt.Parameters.RootParameterSet.ParameterSets.CreateSet("cm")
#         colls = o_prt.Parameters.RootParameterSet.ParameterSets.Item(
#             "cm").DirectParameters

#     att_type = {}
#     att_type[attNames[0]] = "String"
#     att_type[attNames[1]] = "list"
#     att_type[attNames[2]] = "String"
#     att_type[attNames[3]] = "Density"
#     att_type[attNames[4]] = "Mass"
#     att_type[attNames[5]] = "Length"

#     parameter_obj = [None] * 6

#     for i in range(1, 6):
#         if not has_att(colls, attNames[i]):
#             if i == 1:
#                 parameter_obj[i] = colls.CreateList(attNames[i])
#             elif i == 2:
#                 parameter_obj[i] = colls.CreateString(attNames[i], "TBD")
#             elif 3 <= i <= 5:
#                 parameter_obj[i] = colls.CreateDimension(
#                     attNames[i], att_type[attNames[i]], 0)
#         else:
#             parameter_obj[i] = colls.Item(attNames[i])

#     lst = parameter_obj[1]
#     if not has_att(lst.ValueList, mbd.Name):
#         lst.ValueList.Add(mbd)

#     o_prt.Update()
#     o_prd.Update()

#     pubs = ref_prd.Publications
#     for i in range(3, 6):
#         if not has_att(pubs, attNames[i]):
#             o_ref = ref_prd.CreateReferenceFromName(parameter_obj[i].Name)
#             o_pub = pubs.Add(attNames[i])
#             pubs.SetDirect(attNames[i], o_ref)

#     o_prt.Update()
#     o_prd.Update()

#     str_ekl = ["CalM",
#                "let lst(list) set lst=cm\\iBodys let V(Volume) V=0 let Vol(Volume) Vol=0 let i(integer) i=1 for i while i<=lst.Size() {V=smartVolume(lst.GetItem(i)) Vol=Vol+V i=i+1} cm\\iMass=Vol*cm\\iDensity"]
#     colls = o_prt.Relations
#     if not has_att(colls, str_ekl[0]):
#         o_rule = colls.CreateProgram(str_ekl[0], "cal of mass", str_ekl[1])
#     else:
#         o_rule = colls.Item(str_ekl[0])
#         if o_rule.Value != str_ekl[1]:
#             o_rule.Modify(str_ekl[1])

#     rl_name = "CMAS"
#     rl_target = ref_prd.UserRefProperties.Item(attNames[4])
#     rl_source = "cm\\iMass"
#     if not has_att(colls, rl_name):
#         o_formula = colls.CreateFormula(rl_name, " ", rl_target, rl_source)
#     else:
#         o_formula = colls.Item(rl_name)
#         if o_formula.Value != rl_source:
#             o_formula.Modify(rl_source)
#     print(o_formula.Value)

#     rl_name = "CTK"
#     rl_target = ref_prd.UserRefProperties.Item(attNames[5])
#     rl_source = "cm\\iThickness"
#     if not has_att(colls, rl_name):
#         o_formula = colls.CreateFormula(rl_name, " ", rl_target, rl_source)
#     else:
#         o_formula = colls.Item(rl_name)
#         if o_formula.Value != rl_source:
#             o_formula.Modify(rl_source)

    # def no_prd():
    #     """
    #     释放待修改的产品
    #     """

    #     prd2rw = None
    #     messagebox.showinfo("提示", "已释放待修改的产品")
