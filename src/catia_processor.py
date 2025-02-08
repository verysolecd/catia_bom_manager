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

    def attUsp(oPrd):
        refprd = oPrd.reference_product
        att_usp = [None]*6
        att_usp[0] = ""
        for i in range(1, 5):
            if i in [1, 3, 4]:
                colls = refprd.user_ref_properties
                att_usp[i] = thisParameterValue(colls, attNames[i])
            elif i == 2:
                try:
                    oPrt = refprd.parent.part
                    colls = oPrt.parameters.root_parameter_set.parameter_sets.item(
                        "cm").direct_parameters
                    att_usp[i] = thisParameterValue(colls, attNames[i])
                except:
                    att_usp[i] = "N\A"
        return att_usp

    def bomRowPrd(oPrd, LV):
        oDict = {}
        QTy = 1
        if isinstance(oPrd.parent, pycatia.product_structure_interfaces.products):
            oDict = Dictbros(oPrd.parent.parent)
            QTy = oDict[oPrd.part_number]
        return [LV, attDefault(oPrd), QTy, attUsp(oPrd)]

    def infoPrd(oPrd):
        return [0, attDefault(oPrd), 0, attUsp(oPrd)]

    def recurPrd(oPrd, xlsht, oRowNb, LV):
        bDict = {}
        bomRowPrd(oPrd, LV, xlsht, oRowNb)
        if oPrd.products.count > 0:
            for i in range(1, oPrd.products.count + 1):
                if oPrd.products.item(i).part_number not in bDict:
                    bDict[oPrd.products.item(i).part_number] = 1
                    oRowNb += 1
                    recurPrd(oPrd.products.item(i), xlsht, oRowNb, LV + 1)

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

    def Dictbros(oPrd):
        oDict = {}
        for i in range(1, oPrd.products.count + 1):
            pn = oPrd.products.item(i).part_number
            if pn in oDict:
                oDict[pn] += 1
            else:
                oDict[pn] = 1
        return oDict

    def hasAtt(colls, parameterName):
        try:
            colls.item(parameterName)
            return True
        except:
            return False

    def thisParameterValue(colls, parameterName):
        try:
            return colls.item(parameterName).value
        except:
            return "N\A"

    import pycatia

# Assume that the following functions and variables are defined elsewhere
# att_usp_Names
# allPN
# xlApp
# CATIA


def AttModify(oPrd, oArry):
    """Modify product attributes from an array to CATIA"""
    refprd = oPrd.ReferenceProduct
    pArry = infoPrd(oPrd)
    ind = 1
    for i in range(1, 5):
        if oArry[ind][i] != "" and oArry[ind][i] != pArry[ind][i]:
            if i == 1:
                refprd.PartNumber = oArry[1][i]
            elif i == 2:
                refprd.Nomenclature = oArry[1][i]
            elif i == 3:
                refprd.definition = oArry[1][i]
            elif i == 4:
                oPrd.Name = oArry[1][i]

    ind = 3
    i = 2
    if oArry[ind][i] != "" and pArry[ind][i] != oArry[ind][i]:
        colls = refprd.UserRefProperties
        colls.Item(att_usp_Names[i]).Value = oArry[ind][i]

    i = 3
    try:
        oPrt = refprd.Parent.Part
        colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item(
            "cm").DirectParameters
        if oArry[ind][i] != "" and oArry[ind][i] != pArry[ind][i]:
            colls.Item(att_usp_Names[i]).Value = oArry[ind][i]
    except Exception as e:
        pass


def iniPrd(oPrd, oDict):
    """Initialize product"""
    if oPrd.PartNumber not in allPN:
        allPN[oPrd.PartNumber] = 1
        initMyRefPrd(oPrd)

    if oPrd.Products.Count > 0:
        for product in oPrd.Products:
            iniPrd(product, oDict)

    allPN.clear()


def initMyRefPrd(oPrd):
    """Initialize reference product"""
    refprd = oPrd.ReferenceProduct
    attType = {
        att_usp_Names[0]: "String",
        att_usp_Names[1]: "list",
        att_usp_Names[2]: "String",
        att_usp_Names[3]: "Density",
        att_usp_Names[4]: "Mass",
        att_usp_Names[5]: "Length"
    }
    colls = refprd.UserRefProperties
    parameterObj = [None] * 6
    for i in range(2, 6):
        if not hasAtt(colls, att_usp_Names[i]):
            if i == 2:
                parameterObj[i] = colls.CreateString(att_usp_Names[i], "TBD")
            elif 4 <= i <= 5:
                parameterObj[i] = colls.CreateDimension(
                    att_usp_Names[i], attType[att_usp_Names[i]], 0)
        else:
            parameterObj[i] = colls.Item(att_usp_Names[i])

    try:
        oPrt = refprd.Parent.Part
        iniPrt(oPrt, att_usp_Names)
    except Exception as e:
        colls = refprd.Publications
        i = 4
        if not hasAtt(colls, att_usp_Names[i]):
            oref = refprd.CreateReferenceFromName(parameterObj[i].Name)
            oPub = colls.Add(att_usp_Names[i])
            colls.SetDirect(att_usp_Names[i], oref)

    oPrd.Update()


def iniPrt(oPrd, att_usp_Names):
    """Initialize part"""
    refprd = oPrd.ReferenceProduct
    oPrt = refprd.Parent.Part
    MBD = oPrt.MainBody

    try:
        colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item(
            "cm").DirectParameters
    except Exception as e:
        colls = oPrt.Parameters.RootParameterSet.ParameterSets.CreateSet("cm")
        colls = oPrt.Parameters.RootParameterSet.ParameterSets.Item(
            "cm").DirectParameters

    attType = {
        att_usp_Names[0]: "String",
        att_usp_Names[1]: "list",
        att_usp_Names[2]: "String",
        att_usp_Names[3]: "Density",
        att_usp_Names[4]: "Mass",
        att_usp_Names[5]: "Length"
    }

    parameterObj = [None] * 6
    for i in range(1, 6):
        if not hasAtt(colls, att_usp_Names[i]):
            if i == 1:
                parameterObj[i] = colls.CreateList(att_usp_Names[i])
            elif i == 2:
                parameterObj[i] = colls.CreateString(att_usp_Names[i], "TBD")
            elif 3 <= i <= 5:
                parameterObj[i] = colls.CreateDimension(
                    att_usp_Names[i], attType[att_usp_Names[i]], 0)
        else:
            parameterObj[i] = colls.Item(att_usp_Names[i])

    lst = parameterObj[1]
    if not hasAtt(lst.ValueList, MBD.Name):
        lst.ValueList.Add(MBD)

    oPrt.Update()
    oPrd.Update()

    for i in range(3, 6):
        pubs = refprd.Publications
        if not hasAtt(pubs, att_usp_Names[i]):
            if i in [3, 4, 5]:
                oref = refprd.CreateReferenceFromName(parameterObj[i].Name)
                oPub = pubs.Add(att_usp_Names[i])
                pubs.SetDirect(att_usp_Names[i], oref)
            elif i == 2:
                parameterObj[i] = refprd.UserRefProperties.Item(
                    att_usp_Names[i])
                oref = refprd.CreateReferenceFromName(parameterObj[i].Name)
                oPub = pubs.Add(att_usp_Names[i])
                pubs.SetDirect(att_usp_Names[i], oref)

    oPrt.Update()
    oPrd.Update()

    strEKL = [
        "CalM",
        "let lst(list) set lst=cm\\iBodys let V(Volume) V=0 let Vol(Volume) Vol=0 let i(integer) i=1 for i while i<=lst.Size() {V=smartVolume(lst.GetItem(i)) Vol=Vol+V i=i+1} cm\\iMass=Vol*cm\\iDensity"
    ]
    colls = oPrt.Relations
    if not hasAtt(colls, strEKL[0]):
        oRule = colls.CreateProgram(strEKL[0], "cal of mass", strEKL[1])
    else:
        oRule = colls.Item(strEKL[0])
        if oRule.Value != strEKL[1]:
            oRule.Modify(strEKL[1])

    RLname = "CMAS"
    RLtarget = refprd.UserRefProperties.Item(att_usp_Names[4])
    RLsource = "cm\\iMass"
    colls = oPrt.Relations
    if not hasAtt(colls, RLname):
        oFormula = colls.CreateFormula(RLname, " ", RLtarget, RLsource)
    else:
        oFormula = colls.Item(RLname)
        if oFormula.Value != RLsource:
            oFormula.Modify(RLsource)
    print(oFormula.Value)

    RLname = "CTK"
    RLtarget = refprd.UserRefProperties.Item(att_usp_Names[5])
    RLsource = "cm\\iThickness"
    if not hasAtt(colls, RLname):
        oFormula = colls.CreateFormula(RLname, "    ", RLtarget, RLsource)
