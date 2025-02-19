# atest.py

import win32com.client
import pywintypes

att = [None] * 6
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

    def info_Prd(self, oPrd):
        refPrd = oPrd.referenceproduct
        infoPrd = [None]*6
        infoPrd[0] = refPrd.PartNumber
        infoPrd[1] = refPrd.Nomenclature
        infoPrd[2] = refPrd.Defintion
        infoPrd[3] = oPrd.Name

        infoPrd[4] = refPrd.userrefproperties.item("iMaterial").value

        infoPrd[5] =
        try:
            att_usp[i] = self.thisParameterValue(
                refprd.parent.part.parameters.root_parameter_set.parameter_sets.item(
                    "cm").direct_parameters,
                att_usp_Names[i]
            )


except Exception:
    att_usp[i] = "N\\A"

        infoPrd[6] = refPrd.userrefproperties.item("iThickness").value
        return infoPrd

    def attUsp(self, oPrd):
        attNames = ["iMaterial", "iDensity", "iMass", "iThickness"]

        refPrd = oPrd.referenceproduct
        att_usp = [None]*6
        att_usp[0] = ""
        for i, att_name in enumerate(attNames):
                colls = refPrd.user_ref_properties
                try:
                    att_value = getattr(refPrd, att_name)
                    att_usp[i+1] = att_value
                except AttributeError:
                    att_usp[i+1] = "N/A"
            for i in range(2, 5):
                if i in [2, 4, 5]:
                    colls = refPrd.user_ref_properties
                    att_usp[i] = self._askValue(colls, attNames[i])
                elif i == 3:
                    try:
                        oPrt = refPrd.parent.part
                        colls = oPrt.parameters.root_parameter_set.parameter_sets.item(
                            attNames[1]).direct_parameters
                        att_usp[i] = self._askValue(
                            colls, attNames[i])
                    except Exception as e:
                        att_usp[i] = "N\\A"
            return att_usp


def test_1():
    processor = CATIAProcessor()
    catia = processor.connect_to_catia()
    if catia:
        oprd = catia.ActiveDocument.Product
        processor.init_refPrd(oprd)

if __name__ == '__main__':
    test_1()
