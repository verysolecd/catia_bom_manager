# -*- coding: utf-8 -*-
# this module is to manage catia data to get or define the product attributes
from pycatia import catia
import pycatia


class cadmanager():
    def __init__(self):
        self.catia = catia()
        self.documents = self.catia.documents
        self.active_document = self.catia.active_document
        self.product = self.active_document.is_product

    def iniPrd(self)

    def infoPrd(self，oPrd):
        refPrd = oPrd.reference_product
        self.refPrd = refPrd
        return refPrd

    def att_default(self, oPrd):
        refPrd = oPrd.reference_product
        oArry = [refprd.part_number, refprd.nomenclature,
                 refprd.definition, oPrd.name]
        return oArry

    def att_usp(self, oPrd):
        refPrd = oPrd.reference_product
        oArry = [refprd.part_number, refprd.nomenclature,
                 refprd.definition, oPrd.name]
        return oArry







    def get_product(self):
        return self.product

    def get_product_name(self):
        return self.product.name

    def get_product_type(self):
        return self.product.product_type

    def get_product_description(self):
        return self.product.description

    def get_product_attributes(self):
        return self.product.attributes

    def get_product_attribute(self, attribute_name):
        for attribute in self.product.attributes:
            if attribute.name == attribute_name:
                return attribute.value
