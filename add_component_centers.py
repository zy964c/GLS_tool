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
    extention = '.CATProduct'

    def __init__(self, name, sta, plug_value, bin_order, component_name=None):
        self.sta = sta
        self.name = name
        self.plug_value = plug_value
        self.bin_order = bin_order
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
                known_subtypes = ["OFCR", "OFAR", "Horseshoe", "Hybrid", "PCP"]
                made_subtypes = ["PCP", "OFCR", "OFAR"]
                if subtype in known_subtypes:
                    if subtype is not None and subtype in made_subtypes:
                        incl_subtype = subtype
                else:
                    print "new subtype: " + str(subtype)
                if incl_subtype == 'OFCR' and int(self.sta) < 465 and self.bin_order == 1:
                    incl_subtype += '_FWD'
                print int(self.sta), incl_subtype
                bin_name = str(size) + 'IN_Center_' + bin_type + '_prod_' + incl_subtype
                PartDocPath = self.work_path_lib + '\\' + bin_name
                PartDocPath1 = PartDocPath + rand + self.extention
                try:
                    CenterBin.oFileSys.CopyFile(PartDocPath + self.extention, PartDocPath1, False)
                except:
                    print 'Such reference geometry is not found: ' + PartDocPath
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

        return None

# helper to get info from json export

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

    layout = os.getcwd() + '\\json_cus_data\\787_9_GUN_ZB910.json'
    k = json.load(codecs.open(layout, 'r', 'utf-8-sig'))
    content = k['Layout']['Children']
    sta = [i["STA"] for i in content if i["ObjectType"] == "Stowbin" and i["Column"] == "Center"]
    print sta

    irm = 'Product1.1'
    dash_nine = 240
    m = []

    def add_zero(j):
        if j < 1000:
            return '0'+str(j)
        return str(j)
    for l in map(add_zero, sta):
        sbin = CenterBin(irm, l, 2, dash_nine)
        m.append(sbin)
    print len(m)
    for n in m:
        print n.parse_ss(layout)
    #c.parse_ss1(c.work_path + '\\json_cus_data\\787_9_NEO_ZB874.json')

