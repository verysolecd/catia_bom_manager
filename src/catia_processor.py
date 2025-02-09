# -*- coding: utf-8 -*-
# this module is to manage catia data to get or define the product attributes
from pycatia import catia
import pycatia

# 全局参数
attNames = []


class ClassPDM():
    def __init__(self):
        self.catia = catia()
        self.documents = self.catia.documents
        self.active_document = self.catia.active_document
        self.rootPrd = self.active_document.product
        self.initialize_array()  # 确保没有传递任何参数

#     def selprd():
#         """
#         选择产品的具体操作
#         """
#         messagebox.showinfo("提示", "请选择要读取的产品")
#         o_sel = self.active_document.Selection
#         o_sel.Clear()
#         i_type = ["Product"]
#         if o_sel.Count2 == 0:
#             status = o_sel.SelectElement2(i_type, "请选择要读取的产品", False)
#         if status == "Cancel":
#             return None
#         if status == "Normal" and o_sel.Count2 == 1:
#             result = o_sel.Item(1).LeafProduct.ReferenceProduct
#             o_sel.Clear()
#             return result
#         messagebox.showinfo("提示", "请只选择一个产品")
#         return None

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





    def initialize_array(self):
        global attNames
        attNames = ["cm", "iBodys", "iMaterial",
                    "iDensity", "iMass", "iThickness"]
        return attNames

    def attDefault(self, oPrd):
        refprd = oPrd.reference_product
        att_default = [refprd.part_number,
                       refprd.nomenclature, refprd.definition, oPrd.name]
        # 0   1   2   3
        # pn nom def name
        return att_default

#     def attUsp(oPrd):
#         refprd = oPrd.reference_product
#         att_usp = [None]*6
#         att_usp[0] = ""
#         for i in range(1, 5):
#             if i in [1, 3, 4]:
#                 colls = refprd.user_ref_properties
#                 att_usp[i] = thisParameterValue(colls, attNames[i])
#             elif i == 2:
#                 try:
#                     oPrt = refprd.parent.part
#                     colls = oPrt.parameters.root_parameter_set.parameter_sets.item(
#                         "cm").direct_parameters
#                     att_usp[i] = thisParameterValue(colls, attNames[i])
#                 except:
#                     att_usp[i] = "N\A"
#         return att_usp

#     def bomRowPrd(oPrd, LV):
#         oDict = {}
#         QTy = 1
#         if isinstance(oPrd.parent, pycatia.product_structure_interfaces.products):
#             oDict = Dictbros(oPrd.parent.parent)
#             QTy = oDict[oPrd.part_number]
#         return [LV, attDefault(oPrd), QTy, attUsp(oPrd)]

#     def infoPrd(oPrd):
#         return [0, attDefault(oPrd), 0, attUsp(oPrd)]

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
# att_usp_Names = ["iMaterial", "iMass", "iThickness", "iDensity", "iBodys"]


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
#     att_type[att_usp_Names[0]] = "String"
#     att_type[att_usp_Names[1]] = "list"
#     att_type[att_usp_Names[2]] = "String"
#     att_type[att_usp_Names[3]] = "Density"
#     att_type[att_usp_Names[4]] = "Mass"
#     att_type[att_usp_Names[5]] = "Length"

#     colls = ref_prd.UserRefProperties
#     parameter_obj = [None] * 6

#     for i in range(2, 6):
#         if not has_att(colls, att_usp_Names[i]):
#             if i == 2:
#                 parameter_obj[i] = colls.CreateString(att_usp_Names[i], "TBD")
#             elif 4 <= i <= 5:
#                 parameter_obj[i] = colls.CreateDimension(
#                     att_usp_Names[i], att_type[att_usp_Names[i]], 0)
#         else:
#             parameter_obj[i] = colls.Item(att_usp_Names[i])

#     try:
#         o_prt = ref_prd.Parent.Part
#         ini_prt(o_prd, att_usp_Names)
#     except Exception:
#         colls = ref_prd.Publications
#         i = 4
#         if not has_att(colls, att_usp_Names[i]):
#             o_ref = ref_prd.CreateReferenceFromName(parameter_obj[i].Name)
#             o_pub = colls.Add(att_usp_Names[i])
#             colls.SetDirect(att_usp_Names[i], o_ref)

#     o_prd.Update()


# def ini_prt(o_prd, att_usp_Names):
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
#     att_type[att_usp_Names[0]] = "String"
#     att_type[att_usp_Names[1]] = "list"
#     att_type[att_usp_Names[2]] = "String"
#     att_type[att_usp_Names[3]] = "Density"
#     att_type[att_usp_Names[4]] = "Mass"
#     att_type[att_usp_Names[5]] = "Length"

#     parameter_obj = [None] * 6

#     for i in range(1, 6):
#         if not has_att(colls, att_usp_Names[i]):
#             if i == 1:
#                 parameter_obj[i] = colls.CreateList(att_usp_Names[i])
#             elif i == 2:
#                 parameter_obj[i] = colls.CreateString(att_usp_Names[i], "TBD")
#             elif 3 <= i <= 5:
#                 parameter_obj[i] = colls.CreateDimension(
#                     att_usp_Names[i], att_type[att_usp_Names[i]], 0)
#         else:
#             parameter_obj[i] = colls.Item(att_usp_Names[i])

#     lst = parameter_obj[1]
#     if not has_att(lst.ValueList, mbd.Name):
#         lst.ValueList.Add(mbd)

#     o_prt.Update()
#     o_prd.Update()

#     pubs = ref_prd.Publications
#     for i in range(3, 6):
#         if not has_att(pubs, att_usp_Names[i]):
#             o_ref = ref_prd.CreateReferenceFromName(parameter_obj[i].Name)
#             o_pub = pubs.Add(att_usp_Names[i])
#             pubs.SetDirect(att_usp_Names[i], o_ref)

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
#     rl_target = ref_prd.UserRefProperties.Item(att_usp_Names[4])
#     rl_source = "cm\\iMass"
#     if not has_att(colls, rl_name):
#         o_formula = colls.CreateFormula(rl_name, " ", rl_target, rl_source)
#     else:
#         o_formula = colls.Item(rl_name)
#         if o_formula.Value != rl_source:
#             o_formula.Modify(rl_source)
#     print(o_formula.Value)

#     rl_name = "CTK"
#     rl_target = ref_prd.UserRefProperties.Item(att_usp_Names[5])
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
#     global prd2wt
#     prd2wt = None
#     messagebox.showinfo("提示", "已释放待修改的产品")
