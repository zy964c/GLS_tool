import win32com.client
import os
import json
import codecs
from random import randint
from functions import sta_value, inch_to_mm


class CenterBin(object):

    CATIA = win32com.client.Dispatch('catia.application')
    oFileSys = CATIA.FileSystem
    work_path = os.getcwd()
    work_path_lib = work_path + '\LIBRARY'
    extention = '.CATPart'

    def __init__(self, name, sta, plug_value, component_name=None):
        self.sta = sta
        self.name = name
        self.plug_value = plug_value
        self.component_name = component_name

    def set_name(self, new_name):
        self.component_name = new_name

    def parse_ss(self, layout):

        j = json.load(codecs.open(layout, 'r', 'utf-8-sig'))
        content = j['Layout']['Children']
        rand = str(randint(0, 10000))
        incl_subtype = ''
        for i in content:
            if i["ObjectType"] == "Stowbin" and i["Column"] == "Center" and i["STA"] == float(self.sta):
                size = int(i["Width"])
                bin_type = i["StowbinType"]
                subtype = i["BinSubtype"]
                known_subtypes = ["OFCR", "OFAR", "Horseshoe", "Hybrid"]
                made_subtypes = []
                if subtype in known_subtypes:
                    if subtype is not None and subtype in made_subtypes:
                        incl_subtype = subtype
                else:
                    print "new subtype: " + str(subtype)
                bin_name = str(size) + 'IN_Center_' + bin_type + '_' + incl_subtype
                PartDocPath = self.work_path_lib + '\\' + bin_name
                PartDocPath1 = PartDocPath + rand + self.extention
                try:
                    CenterBin.oFileSys.CopyFile(PartDocPath + self.extention, PartDocPath1, False)
                except:
                    break
                ICM_Product = CenterBin.CATIA.ActiveDocument.Product.Products.Item(self.name)
                ICM_Products = ICM_Product.Products
                PartDoc = CenterBin.CATIA.Documents.Open(PartDocPath1)
                NewComponent = ICM_Products.AddExternalComponent(PartDoc)
                PartDoc.Close()
                CenterBin.oFileSys.DeleteFile(PartDocPath1)
                NewComponent.PartNumber = str(size) + 'IN STA' + sta_value(inch_to_mm(float(self.sta)), self.plug_value)
                self.set_name(NewComponent.Name)
                position = [1, 0, 0, 0, 1, 0, 0, 0, 1, inch_to_mm(float(self.sta)), 0, 0]
                NewComponent.Move.Apply(position)
                return self.component_name
        print 'Such reference geometry is not found: ' + PartDocPath
        return None

    def parse_ss1(self, layout):

        j = json.load(codecs.open(layout, 'r', 'utf-8-sig'))
        content = j['Layout']['Children']
        with open('Bin_subtypes.txt', 'w') as output:
            for i in content:
                #if i["ObjectType"] == "Stowbin" and i["Column"] == "Center":
                try:
                    if i["Column"] == "Center":
                        content = 'STA' + str(i["STA"]) + ': ' + str(i["BinSubtype"])
                        output.write(content + '\n')
                except KeyError:
                    continue


if __name__ == "__main__":

    irm = 'Product1.1'
    dash_nine = 240
    c = CenterBin(irm, '0489', dash_nine)
    c1 = CenterBin(irm, '0513', dash_nine)
    c2 = CenterBin(irm, '0561', dash_nine)
    c3 = CenterBin(irm, '0609', dash_nine)
    c4 = CenterBin(irm, '0657', dash_nine)
    c5 = CenterBin(irm, '0801', dash_nine)
    c6 = CenterBin(irm, '0825', dash_nine)
    c7 = CenterBin(irm, '0873', dash_nine)
    c8 = CenterBin(irm, '0921', dash_nine)
    c9 = CenterBin(irm, '0969', dash_nine)
    c10 = CenterBin(irm, '1017', dash_nine)
    m = [c, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10]
    for n in m:
        print n.parse_ss(n.work_path + '\\json_cus_data\\787_9_NEO_ZB874.json')
    #c.parse_ss1(c.work_path + '\\json_cus_data\\787_9_NEO_ZB874.json')

