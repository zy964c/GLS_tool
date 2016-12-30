import math
import win32com.client
import os
import pdb
from functions import inch_to_mm, mm_to_inch, sta_value
from ringposts import redirect


class Ref(object):

    bin_breaker = []
    sta_value_pairs = []
    sta_values_fake = []
    angle = 0
    order_of_templete_product = 4
    make_carms = False
    make_carms_UI = False
    dash_number = 1000
    enovia_connected = True
    create_irms = True
    parameters_set = False
    CATIA = win32com.client.Dispatch('catia.application')
    oFileSys = CATIA.FileSystem
    work_path = os.getcwd()
    work_path_lib = work_path + '\LIBRARY'
    print work_path_lib

 #  @staticmethod

    def __init__(self, customer, sta, side, plug, bin_order, irm_ln, all_irm_parts, path=work_path_lib, name=None,
                 component_name=None):
        self.plug = plug
        self.path = path
        self.customer = customer
        self.sta_to_find = sta
        self.side_to_find = side
        self.name = name
        self.component_name = component_name
        self.bin_order = bin_order
        self.irm_ln = irm_ln
        self.all_irm_parts = all_irm_parts

    def set_plug(self, new_plug):
        self.plug = new_plug

    def set_path(self, new_path):
        self.path = new_path

    def set_customer(self, new_customer):
        self.path = new_customer

    def set_sta_to_find(self, new_sta_to_find):
        self.sta_to_find = new_sta_to_find

    def set_side_to_find(self, new_side_to_find):
        self.side_to_find = new_side_to_find

    def set_name(self, new_name):
        self.component_name = new_name

    def converter(self):

        f = open(str(self.customer) + '.txt')
        s_raw = f.readlines()
        #print s_raw
        s_all = []
        state1 = True
        for element in s_raw:
            if 'CTR' in element:
                state1 = False
                continue
            elif '#' in element and 'CTR' not in element:
                state1 = True
                continue
            elif '#' not in element and state1 is True:
                s_all.append(element.replace(' ', '').replace('fairing', '1').replace('premium', '2')
                             .replace('prem', '2').replace('EXT', '3').replace('\r\n', '').split(","))
            else:
                continue

        #print s_all

        s1 = s_all[0]
        s2 = s_all[1]
        s3 = s_all[2]
        s4 = s_all[3]
        s5 = s_all[4]
        s6 = s_all[5]
        s1 = s1[::-1]
        s2 = s2[::-1]

        return s1, s2, s3, s4, s5, s6

    def instantiate_nonconstant_components(self):
        """
        Instantiates 4 products for non-constant sections
        :return: no return
        """

        try:
            ICM = Ref.CATIA.ActiveDocument
        except:
            ICM = Ref.CATIA.Documents.Add('Product')

#        ICM_1 = ICM.Product
#        ICM_Products = ICM_1.Products

        ICM_1 = ICM.Product
        ICM_Products_irms = ICM_1.Products
        ICM_Product = ICM_Products_irms.Item(self.name)
        ICM_Products = ICM_Product.Products
        ICM_Products_ref = ICM_Product.ReferenceProduct.Products

        global new_component1
        try:
            new_component1 = ICM_Products.Item(self.name + '_' + 'non-constant_41_LH.1')
        except:
            new_component1 = ICM_Products.AddNewComponent("Product", self.name + '_' + 'non-constant_41_LH')
#            name_ref = new_component1.Name
#            new_component1_ref = new_component1.RefefrenceProduct.Name = new_component1.Name[:-2]

        global ICM_Sec41_LH_Products
        ICM_Sec41_LH_Products = new_component1.Products

        global new_component2
        try:
            new_component2 = ICM_Products.Item(self.name + '_' + 'non-constant_41_RH.1')
        except:
            new_component2 = ICM_Products.AddNewComponent("Product", self.name + '_' + 'non-constant_41_RH')

        global ICM_Sec41_RH_Products
        ICM_Sec41_RH_Products = new_component2.Products

        global new_component3
        try:
            new_component3 = ICM_Products.Item(self.name + '_' + 'non-constant_47_LH.1')
        except:
            new_component3 = ICM_Products.AddNewComponent("Product", self.name + '_' + 'non-constant_47_LH')

        global ICM_Sec47_LH_Products
        ICM_Sec47_LH_Products = new_component3.Products

        global new_component4
        try:
            new_component4 = ICM_Products.Item(self.name + '_' + 'non-constant_47_RH.1')
        except:
            new_component4 = ICM_Products.AddNewComponent("Product", self.name + '_' + 'non-constant_47_RH')

        global ICM_Sec47_RH_Products
        ICM_Sec47_RH_Products = new_component4.Products

    def add_component(self, s, side, section, location, plug_value, name=None):
        """
        :param s: a list containing information about bin run
        :param side: 'LH' or 'RH' side
        :param section: 'constant' or 'nonconstant'
        :param location: 'nose', 'middle' or 'tail'
        :param plug_value: insert plug variable
        :return: Doesn't return anything, builds ECS layout using CATIA objects library and sets instance IDs. Modifies
        globals: sta_value_pairs, sta_values_fake
        """

        ICM = Ref.CATIA.ActiveDocument
        ICM_1 = ICM.Product
        ICM_Products_irms = ICM_1.Products
        #debugging
        #print name
        #for i in xrange(1, ICM_Products_irms.Count+1):
        #    print ICM_Products_irms.Item(i).Name
        ICM_Product = ICM_Products_irms.Item(name)
        ICM_Products = ICM_Product.Products
        #bin_breaker = []
        extention = '.CATProduct'
        global sta_value_pairs
        global sta_values_fake
        global angle
        x_coord = inch_to_mm(465)
        x_coord_nonconstant = inch_to_mm(0)
        fake_coord_nonconstant_41 = inch_to_mm(459)
        if plug_value == 240:
            fake_coord_nonconstant_47 = inch_to_mm(1863)
        elif plug_value == 456:
            fake_coord_nonconstant_47 = inch_to_mm(2079)
        elif plug_value == 0:
            fake_coord_nonconstant_47 = inch_to_mm(1623)

        if plug_value == 0:
            door2_coord = 0
        elif plug_value == 456:
            door2_coord = 240
        else:
            door2_coord = 120

        if side == 'LH' and location == 'middle':
            iteration = 0
        elif side == 'RH' and location == 'middle':
            iteration = 100
        elif side == 'LH' and location == 'nose':
            iteration = 200
        elif side == 'RH' and location == 'nose':
            iteration = 300
        elif side == 'LH' and location == 'tail':
            iteration = 400
        elif side == 'RH' and location == 'tail':
            iteration = 500

        if location == 'tail':
            Ref.angle = 3.125
        else:
            Ref.angle = 5

        rad = math.radians(Ref.angle)
        #print Ref.angle

        index = 0

        for number in s:

            nozzl_type = 'ECO'
            dow_type = 'DWNR_STD-STRT'
            ligval_ammount = 1
            Arch = ''

            bins = ['36', '42', '48', '362', '422', '482']
            bin_twenty_four = ['24', '242', '2432', '243']

            if number in bins:
                stowbin = True
                btype = 'BIN'
            elif number in bin_twenty_four:
                stowbin = 'twenty_four'
                #dow_type = 'DWNR_JOG-STRT'
                btype = 'BIN'
            else:
                stowbin = False
                btype = 'FAIRING'

            if str(number) == 'door':
                x_coord = inch_to_mm(693 + door2_coord)
                index += 1
                continue

            else:

                Rotate5 = [0.996194698, -0.087155742, 0, 0.087155742, 0.996194698, 0, 0, 0, 1, inch_to_mm(466.61647022),
                           inch_to_mm(0.08471639), 0]
                Rotate185 = [-0.996194698, -0.087155742, 0, 0.087155742, -0.996194698, 0, 0, 0, 1, inch_to_mm(466.61647018),
                             inch_to_mm(-0.084716377), 0]
                Rotate_5 = [0.998512978, 0.054514501, 0, -0.054514501, 0.998512978, 0, 0, 0, 1,
                            inch_to_mm(1618.61663822 + plug_value), inch_to_mm(0.17865996), 0]
                Rotate_185 = [-0.998512978, 0.054514501, 0, -0.054514501, -0.998512978, 0, 0, 0, 1,
                              inch_to_mm(1618.61663822 + plug_value), inch_to_mm(-0.17865996), 0]

                #print int(number)

                # checking area around DOOR 2:

                if index != (len(s) - 1) and (s[index + 1] == 'door' or s[index - 1] == 'door'):

                    if (side == 'LH' and s[index + 1] == 'door') or (side == 'RH' and s[index - 1] == 'door'):

                        #print 'RH door2'

                        Arch = 'ARCH'

                        if int(number) == 24:
                            dow_type = 'DWNR_JOG-RGHT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 243:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 36:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 42:

                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 18 or int(number) == 181:
                            number = '18'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 30 or int(number) == 301:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 54 or int(number) == 541:
                            number = '54'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 60 or int(number) == 601:
                            number = '60'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 241:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 361:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 421:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 481:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                        # PREMIUM:

                        elif int(number) == 242:
                            dow_type = 'DWNR_JOG-RGHT'
                            number = '24'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2432:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:
                            number = '18'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 362:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:
                            number = '54'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:
                            number = '60'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2412 or int(number) == 2421:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 3612 or int(number) == 3621:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 4212 or int(number) == 4221:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 4812 or int(number) == 4821:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                        else:
                            x_coord += inch_to_mm(int(number))
                            iteration += 1
                            index += 1
                            continue

                    elif (side == 'LH' and s[index - 1] == 'door') or (side == 'RH' and s[index + 1] == 'door'):

                        #print 'LH door2'

                        Arch = 'ARCH'

                        if int(number) == 24:
                            dow_type = 'DWNR_JOG-LEFT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 243:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 36:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 42:

                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 18 or int(number) == 181:
                            number = '18'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 30 or int(number) == 301:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 54 or int(number) == 541:
                            number = '54'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 60 or int(number) == 601:
                            number = '60'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 241:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 361:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 421:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 481:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                            # PREM:

                        elif int(number) == 242:
                            dow_type = 'DWNR_JOG-LEFT'
                            number = '24'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2432:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:
                            number = '18'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 362:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:
                            number = '54'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:
                            number = '60'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2412 or int(number) == 2421:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 3612 or int(number) == 3621:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 4212 or int(number) == 4221:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 4812 or int(number) == 4821:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                        else:
                            x_coord += inch_to_mm(int(number))
                            iteration += 1
                            index += 1
                            continue

                            #  NOT around DOOR 2:

                elif (location == 'nose' and stowbin is not True and stowbin != 'twenty_four') or location == 'middle':

                    if int(number) == 24:

                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_LH_solids'
                            dow_type = 'DWNR_JOG-LEFT'
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_RH_solids'
                            dow_type = 'DWNR_JOG-RGHT'
                        else:
                            PartDocPath = self.path + '\Twenty_four_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 243:
                        number = '30'
                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_LH_solids'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_RH_solids'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 36:

                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Thirty_six_DR3_solids'
                            dow_type = 'DWNR_JOG-STRT'
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Thirty_six_DR3_solids'
                            dow_type = 'DWNR_JOG-STRT'
                        else:
                            PartDocPath = self.path + '\Thirty_six_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 48:

                        iteration += 1
                        index += 1
                        if sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_LH'
                        elif sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_RH'
                        else:
                            PartDocPath = self.path + '\Fourty_eight_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention

                    elif int(number) == 12 or int(number) == 121:
                        number = '12'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twelve_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention

                    elif int(number) == 18 or int(number) == 181:
                        number = '18'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Eighteen_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 30 or int(number) == 301:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 54 or int(number) == 541:
                        number = '54'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fifty_four_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 60 or int(number) == 601:
                        number = '60'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Sixty_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 72 or int(number) == 721:
                        number = '72'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Seventy_two_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 241:

                        number = '24'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention

                    elif int(number) == 361:

                        number = '36'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 421:

                        number = '42'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 481:

                        number = '48'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                        # Premium plenums:

                    elif int(number) == 242:

                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_LH_solids_pr'
                            dow_type = 'DWNR_JOG-LEFT'
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_RH_solids_pr'
                            dow_type = 'DWNR_JOG-RGHT'
                        else:
                            PartDocPath = self.path + '\Twenty_four_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 2432:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        if index != (len(s) - 1) and side == 'RH' and s[index] == '72' or index != (
                                len(s) - 1) and side == 'LH' and s[index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_LH_solids_pr'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        elif index != (len(s) - 1) and side == 'RH' and s[index - 2] == '72' or index != (
                                len(s) - 1) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_RH_solids_pr'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 362:

                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 422:

                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 482:

                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        if sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_LH'
                        elif sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_RH'
                        else:
                            PartDocPath = self.path + '\Fourty_eight_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 122 or int(number) == 1212 or int(number) == 1221:

                        number = '12'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twelve_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:

                        number = '18'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Eighteen_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:

                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:

                        number = '54'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fifty_four_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:

                        number = '60'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Sixty_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 722 or int(number) == 7212 or int(number) == 7221:

                        number = '72'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Seventy_two_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 2412 or int(number) == 2421:

                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 3612 or int(number) == 3621:

                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 4212 or int(number) == 4221:

                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 4812 or int(number) == 4821:

                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    else:
                        x_coord += inch_to_mm(int(number))
                        iteration += 1
                        index += 1
                        continue

                        # SECTION 41:

                elif location == 'nose' and stowbin is True or stowbin == 'twenty_four' and location == 'nose':

                    if int(number) == 24:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 36:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 48:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 242:

                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 362:

                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 422:

                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 482:

                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                        # SECTION 47:

                else:

                    if int(number) == 24:
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 36:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 48:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 12 or int(number) == 121:
                        number = '12'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twelve_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 18 or int(number) == 181:
                        number = '18'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Eighteen_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 30 or int(number) == 301:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 241:

                        number = '24'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 361:

                        number = '36'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 421:

                        number = '42'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 481:

                        number = '48'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    else:
                        x_coord += inch_to_mm(int(number))
                        iteration += 1
                        index += 1
                        continue

                # for lower plenums:

                if stowbin is True or 'twenty_four':
                    if number == '36':
                        L_PL_size1 = '36'
                    elif number == '42':
                        L_PL_size1 = '42'
                    elif number == '48':
                        L_PL_size1 = '48'
                    else:
                        L_PL_size1 = '24'

                if section == 'constant':
                    sta_current = sta_value(x_coord, plug_value)
                elif location == 'nose':
                    sta_current = sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                        plug_value)
                else:
                    sta_current = sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant),
                        plug_value)

                if self.sta_to_find == sta_current and self.side_to_find == side:

                    if stowbin is True or 'twenty_four':
                        #pdb.set_trace()
                        PartDocPath = redirect(self, PartDocPath)

                    Ref.oFileSys.CopyFile(PartDocPath + extention, PartDocPath1, False)
                    PartDoc = Ref.CATIA.Documents.Open(PartDocPath1)

                    if section == 'constant':

                        NewComponent = ICM_Products.AddExternalComponent(PartDoc)
                        PartDoc.Close()
                        Ref.oFileSys.DeleteFile(PartDocPath1)
                        NewComponent.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                        NewComponent.Name = str(number) + 'IN STA ' + sta_value(x_coord, plug_value) + ' ' + side + Arch
                        self.set_name(NewComponent.Name)

                    elif section == 'nonconstant' and side == 'LH' and location == 'nose':

                        NewComponent = ICM_Sec41_LH_Products.AddExternalComponent(PartDoc)
                        PartDoc.Close()
                        Ref.oFileSys.DeleteFile(PartDocPath1)
                        RenamingToolProd = new_component1.ReferenceProduct
                        Pr = RenamingToolProd.Products
                        Prod = RenamingToolProd.Products.Item(Pr.count)
                        Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                        Prod.Name = str(number) + 'IN STA ' + sta_value(
                            (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                            plug_value) + ' ' + side
                        self.set_name(Prod.Name)

                        NewComponent.Move.Apply(Rotate5)

                    elif section == 'nonconstant' and side == 'RH' and location == 'nose':

                        NewComponent = ICM_Sec41_RH_Products.AddExternalComponent(PartDoc)
                        PartDoc.Close()
                        Ref.oFileSys.DeleteFile(PartDocPath1)
                        RenamingToolProd = new_component2.ReferenceProduct
                        Pr = RenamingToolProd.Products
                        Prod = RenamingToolProd.Products.Item(Pr.count)
                        Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                        Prod.Name = str(number) + 'IN STA ' + sta_value(
                            (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                            plug_value) + ' ' + side
                        self.set_name(Prod.Name)

                        NewComponent.Move.Apply(Rotate185)

                    elif section == 'nonconstant' and side == 'LH' and location == 'tail':

                        NewComponent = ICM_Sec47_LH_Products.AddExternalComponent(PartDoc)
                        PartDoc.Close()
                        Ref.oFileSys.DeleteFile(PartDocPath1)
                        RenamingToolProd = new_component3.ReferenceProduct
                        Pr = RenamingToolProd.Products
                        Prod = RenamingToolProd.Products.Item(Pr.count)
                        Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                        Prod.Name = str(number) + 'IN STA ' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                        plug_value) + ' ' + side
                        self.set_name(Prod.Name)

                        NewComponent.Move.Apply(Rotate_5)

                    elif section == 'nonconstant' and side == 'RH' and location == 'tail':

                        NewComponent = ICM_Sec47_RH_Products.AddExternalComponent(PartDoc)
                        PartDoc.Close()
                        Ref.oFileSys.DeleteFile(PartDocPath1)
                        RenamingToolProd = new_component4.ReferenceProduct
                        Pr = RenamingToolProd.Products
                        Prod = RenamingToolProd.Products.Item(Pr.count)
                        Prod.PartNumber = str(number) + '_' + nozzl_type + '_' + btype + '_' + str(iteration)
                        Prod.Name = str(number) + 'IN STA ' + sta_value((fake_coord_nonconstant_47 + x_coord_nonconstant),
                                                                        plug_value) + ' ' + side
                        self.set_name(Prod.Name)

                        NewComponent.Move.Apply(Rotate_185)


                if section == 'constant':
                    trouble2 = mm_to_inch(x_coord)
                    Ref.bin_breaker.append(int(trouble2))
                    Ref.sta_values_fake.append(sta_value(x_coord, plug_value))
                    Ref.sta_value_pairs.append(x_coord)
                    #print bin_breaker

                elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                elif section == 'nonconstant' and side == 'LH' and location == 'tail':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                elif section == 'nonconstant' and side == 'RH' and location == 'tail':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                #print Ref.sta_value_pairs
                #print Ref.sta_values_fake

                if location == 'nose':
                    x_coord_nonconstant -= inch_to_mm(int(number))

                x = x_coord_nonconstant * math.cos(rad)
                y = x_coord_nonconstant * math.sin(rad)

                position = [1, 0, 0, 0, 1, 0, 0, 0, 1, x_coord, 0, 0]
                position_non = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, -y, 0]
                position_non_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (inch_to_mm(int(number)) * math.cos(rad)),
                                   y + (inch_to_mm(int(number)) * math.sin(rad)), 0]
                position90 = [-1, 0, 0, 0, -1, 0, 0, 0, 1, x_coord + inch_to_mm(int(number)), 0, 0]  # 90 deg rotation
                position_non_47 = [1, 0, 0, 0, 1, 0, 0, 0, 1, x, y, 0]
                position_non_47_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (inch_to_mm(int(number)) * math.cos(rad)),
                                      (y + (inch_to_mm(int(number)) * math.sin(rad))) * (-1), 0]

                if side == 'LH' and section == 'constant':
                    # NewComponentRef = NewComponent.ReferenceProduct
                    if self.sta_to_find == sta_current and self.side_to_find == side:
                        NewComponent.Move.Apply(position)
                    #print side
                    #print x_coord
                elif side == 'RH' and section == 'constant':
                    if self.sta_to_find == sta_current and self.side_to_find == side:
                        NewComponent.Move.Apply(position90)
                    #print side
                    #print x_coord
                elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                    if self.sta_to_find == sta_current and self.side_to_find == side:
                        NewComponent.Move.Apply(position_non)
                    #print section
                    #print x_coord_nonconstant
                elif section == 'nonconstant' and side == 'RH' and location == 'nose':
                    if self.sta_to_find == sta_current and self.side_to_find == side:
                        NewComponent.Move.Apply(position_non_RH)
                    #print section
                    #print x_coord_nonconstant
                elif section == 'nonconstant' and side == 'LH' and location == 'tail':
                    if self.sta_to_find == sta_current and self.side_to_find == side:
                        NewComponent.Move.Apply(position_non_47)
                    x_coord_nonconstant += inch_to_mm(int(number))
                    #print section
                    #print x, y
                    #print x_coord_nonconstant
                elif section == 'nonconstant' and side == 'RH' and location == 'tail':
                    if self.sta_to_find == sta_current and self.side_to_find == side:
                        NewComponent.Move.Apply(position_non_47_RH)
                    x_coord_nonconstant += inch_to_mm(int(number))
                    #print section
                    #print x, y
                    #print x_coord_nonconstant

                x_coord += inch_to_mm(int(number))

    def add_component_coord(self, s, side, section, location, plug_value, target='coord'):
        """
        :param s: a list containing information about bin run
        :param side: 'LH' or 'RH' side
        :param section: 'constant' or 'nonconstant'
        :param location: 'nose', 'middle' or 'tail'
        :param plug_value: insert plug variable
        :return: Doesn't return anything, builds ECS layout using CATIA objects library and sets instance IDs. Modifies
        globals: sta_value_pairs, sta_values_fake
        """
        #ICM = Ref.CATIA.ActiveDocument
        #ICM_1 = ICM.Product
        #ICM_Products = ICM_1.Products
        bin_breaker = []
        extention = '.CATProduct'
        global sta_value_pairs
        global sta_values_fake
        global angle
        x_coord = inch_to_mm(465)
        x_coord_nonconstant = inch_to_mm(0)
        fake_coord_nonconstant_41 = inch_to_mm(459)
        if plug_value == 240:
            fake_coord_nonconstant_47 = inch_to_mm(1863)
        elif plug_value == 456:
            fake_coord_nonconstant_47 = inch_to_mm(2079)
        elif plug_value == 0:
            fake_coord_nonconstant_47 = inch_to_mm(1623)

        if plug_value == 0:
            door2_coord = 0
        elif plug_value == 456:
            door2_coord = 240
        else:
            door2_coord = 120

        if side == 'LH' and location == 'middle':
            iteration = 0
        elif side == 'RH' and location == 'middle':
            iteration = 100
        elif side == 'LH' and location == 'nose':
            iteration = 200
        elif side == 'RH' and location == 'nose':
            iteration = 300
        elif side == 'LH' and location == 'tail':
            iteration = 400
        elif side == 'RH' and location == 'tail':
            iteration = 500

        if location == 'tail':
            Ref.angle = 3.125
        elif location == 'nose':
            Ref.angle = 5
        else:
            Ref.angle = 0

        rad = math.radians(Ref.angle)
        #print Ref.angle

        index = 0

        for number in s:

            nozzl_type = 'ECO'
            dow_type = 'DWNR_STD-STRT'
            ligval_ammount = 1
            Arch = ''

            bins = ['36', '42', '48', '362', '422', '482']
            bin_twenty_four = ['24', '242', '2432', '243']

            if number in bins:
                stowbin = True
                btype = 'BIN'
            elif number in bin_twenty_four:
                stowbin = 'twenty_four'
                dow_type = 'DWNR_JOG-STRT'
                btype = 'BIN'
            else:
                stowbin = False
                btype = 'FAIRING'

            if str(number) == 'door':
                x_coord = inch_to_mm(693 + door2_coord)
                index += 1
                continue

            else:

                Rotate5 = [0.996194698, -0.087155742, 0, 0.087155742, 0.996194698, 0, 0, 0, 1, inch_to_mm(466.61647022),
                           inch_to_mm(0.08471639), 0]
                Rotate185 = [-0.996194698, -0.087155742, 0, 0.087155742, -0.996194698, 0, 0, 0, 1, inch_to_mm(466.61647018),
                             inch_to_mm(-0.084716377), 0]
                Rotate_5 = [0.998512978, 0.054514501, 0, -0.054514501, 0.998512978, 0, 0, 0, 1,
                            inch_to_mm(1618.61663822 + plug_value), inch_to_mm(0.17865996), 0]
                Rotate_185 = [-0.998512978, 0.054514501, 0, -0.054514501, -0.998512978, 0, 0, 0, 1,
                              inch_to_mm(1618.61663822 + plug_value), inch_to_mm(-0.17865996), 0]

                #print int(number)

                # checking area around DOOR 2:

                if index != (len(s) - 1) and (s[index + 1] == 'door' or s[index - 1] == 'door'):

                    if (side == 'LH' and s[index + 1] == 'door') or (side == 'RH' and s[index - 1] == 'door'):

                        #print 'RH door2'

                        Arch = 'ARCH'

                        if int(number) == 24:
                            dow_type = 'DWNR_JOG-RGHT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 243:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 36:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 42:

                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 18 or int(number) == 181:
                            number = '18'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 30 or int(number) == 301:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 54 or int(number) == 541:
                            number = '54'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 60 or int(number) == 601:
                            number = '60'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 241:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 361:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 421:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 481:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                        # PREMIUM:

                        elif int(number) == 242:
                            dow_type = 'DWNR_JOG-RGHT'
                            number = '24'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2432:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:
                            number = '18'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 362:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_RH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:
                            number = '54'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:
                            number = '60'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2412 or int(number) == 2421:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 3612 or int(number) == 3621:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 4212 or int(number) == 4221:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 4812 or int(number) == 4821:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_RH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                        else:
                            x_coord += inch_to_mm(int(number))
                            iteration += 1
                            index += 1
                            continue

                    elif (side == 'LH' and s[index - 1] == 'door') or (side == 'RH' and s[index + 1] == 'door'):

                        #print 'LH door2'

                        Arch = 'ARCH'

                        if int(number) == 24:
                            dow_type = 'DWNR_JOG-LEFT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 243:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 36:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 42:

                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 18 or int(number) == 181:
                            number = '18'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 30 or int(number) == 301:
                            number = '30'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 54 or int(number) == 541:
                            number = '54'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 60 or int(number) == 601:
                            number = '60'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 241:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 361:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 421:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 481:
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                            # PREM:

                        elif int(number) == 242:
                            dow_type = 'DWNR_JOG-LEFT'
                            number = '24'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2432:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_arch_EXT_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:
                            number = '18'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Eighteen_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:
                            number = '30'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 362:
                            dow_type = 'DWNR_JOG-STRT'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_arch_LH_solids'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:
                            number = '54'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fifty_four_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:
                            number = '60'
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Sixty_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention


                        elif int(number) == 2412 or int(number) == 2421:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Twenty_four_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '24'

                        elif int(number) == 3612 or int(number) == 3621:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Thirty_six_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '36'

                        elif int(number) == 4212 or int(number) == 4221:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_two_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '42'

                        elif int(number) == 4812 or int(number) == 4821:
                            nozzl_type = 'PREM'
                            iteration += 1
                            index += 1
                            PartDocPath = self.path + '\Fourty_eight_fairing_arch_LH_solids_pr'
                            PartDocPath1 = PartDocPath + str(iteration) + extention

                            number = '48'

                        else:
                            x_coord += inch_to_mm(int(number))
                            iteration += 1
                            index += 1
                            continue

                            #  NOT around DOOR 2:

                elif (location == 'nose' and stowbin is not True and stowbin != 'twenty_four') or location == 'middle':

                    if int(number) == 24:

                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_LH_solids'
                            dow_type = 'DWNR_JOG-LEFT'
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_RH_solids'
                            dow_type = 'DWNR_JOG-RGHT'
                        else:
                            PartDocPath = self.path + '\Twenty_four_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 243:
                        number = '30'
                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_LH_solids'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_RH_solids'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 36:

                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Thirty_six_DR3_solids'
                            dow_type = 'DWNR_JOG-STRT'
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Thirty_six_DR3_solids'
                            dow_type = 'DWNR_JOG-STRT'
                        else:
                            PartDocPath = self.path + '\Thirty_six_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 48:

                        iteration += 1
                        index += 1
                        if sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_LH'
                        elif sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_RH'
                        else:
                            PartDocPath = self.path + '\Fourty_eight_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention

                    elif int(number) == 12 or int(number) == 121:
                        number = '12'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twelve_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention

                    elif int(number) == 18 or int(number) == 181:
                        number = '18'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Eighteen_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 30 or int(number) == 301:
                        number = '30'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 54 or int(number) == 541:
                        number = '54'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fifty_four_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 60 or int(number) == 601:
                        number = '60'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Sixty_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 72 or int(number) == 721:
                        number = '72'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Seventy_two_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 241:

                        number = '24'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention

                    elif int(number) == 361:

                        number = '36'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 421:

                        number = '42'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 481:

                        number = '48'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_fairing_solids'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                        # Premium plenums:

                    elif int(number) == 242:

                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        if index != len(s) and side == 'RH' and s[index] == '72' or index != len(s) and side == 'LH' and s[
                                index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_LH_solids_pr'
                            dow_type = 'DWNR_JOG-LEFT'
                        elif index != len(s) and side == 'RH' and s[index - 2] == '72' or index != len(
                                s) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_DR3_RH_solids_pr'
                            dow_type = 'DWNR_JOG-RGHT'
                        else:
                            PartDocPath = self.path + '\Twenty_four_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 2432:
                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        if index != (len(s) - 1) and side == 'RH' and s[index] == '72' or index != (
                                len(s) - 1) and side == 'LH' and s[index - 2] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_LH_solids_pr'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        elif index != (len(s) - 1) and side == 'RH' and s[index - 2] == '72' or index != (
                                len(s) - 1) and side == 'LH' and s[index] == '72':
                            PartDocPath = self.path + '\Twenty_four_EXT_DR3_RH_solids_pr'
                            dow_type = 'DWNR_JOG-STRT'
                            ligval_ammount = 2
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 362:

                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 422:

                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 482:

                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        if sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'LH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'LH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_LH'
                        elif sta_value(x_coord, plug_value) == '1569' and plug_value != 456 and side == 'RH' or sta_value(
                                x_coord, plug_value) == '1618+47' and plug_value == 456 and side == 'RH':
                            PartDocPath = self.path + '\Fourty_eight_horseshoe_solids_RH'
                        else:
                            PartDocPath = self.path + '\Fourty_eight_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 122 or int(number) == 1212 or int(number) == 1221:

                        number = '12'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twelve_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 182 or int(number) == 1812 or int(number) == 1821:

                        number = '18'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Eighteen_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 302 or int(number) == 3012 or int(number) == 3021:

                        number = '30'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 542 or int(number) == 5412 or int(number) == 5421:

                        number = '54'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fifty_four_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 602 or int(number) == 6012 or int(number) == 6021:

                        number = '60'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Sixty_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 722 or int(number) == 7212 or int(number) == 7221:

                        number = '72'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Seventy_two_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 2412 or int(number) == 2421:

                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 3612 or int(number) == 3621:

                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 4212 or int(number) == 4221:

                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 4812 or int(number) == 4821:

                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_fairing_solids_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    else:
                        x_coord += inch_to_mm(int(number))
                        iteration += 1
                        index += 1
                        continue

                        # SECTION 41:

                elif location == 'nose' and stowbin is True or stowbin == 'twenty_four' and location == 'nose':

                    if int(number) == 24:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 36:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 48:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_solids_sec41'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 242:

                        number = '24'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 362:

                        number = '36'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 422:

                        number = '42'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 482:

                        number = '48'
                        nozzl_type = 'PREM'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_solids_sec41_pr'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                        # SECTION 47:

                else:

                    if int(number) == 24:
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 36:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 42:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 48:

                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 12 or int(number) == 121:
                        number = '12'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Twelve_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 18 or int(number) == 181:
                        number = '18'
                        index += 1
                        PartDocPath = self.path + '\Eighteen_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 30 or int(number) == 301:
                        number = '30'
                        index += 1
                        PartDocPath = self.path + '\Thirty_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 241:

                        number = '24'
                        index += 1
                        PartDocPath = self.path + '\Twenty_four_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 361:

                        number = '36'
                        index += 1
                        PartDocPath = self.path + '\Thirty_six_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 421:

                        number = '42'
                        iteration += 1
                        index += 1
                        PartDocPath = self.path + '\Fourty_two_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    elif int(number) == 481:

                        number = '48'
                        index += 1
                        PartDocPath = self.path + '\Fourty_eight_fairing_solids_sec47'
                        PartDocPath1 = PartDocPath + str(iteration) + extention


                    else:
                        x_coord += inch_to_mm(int(number))
                        index += 1
                        continue

                ref_name = PartDocPath.split('\\')[-1]
                #print ref_name
                # for lower plenums:
                if section == 'constant':
                    sta_current = sta_value(x_coord, plug_value)
                elif location == 'nose':
                    sta_current = sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                        plug_value)
                else:
                    sta_current = sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant),
                        plug_value)

                if self.sta_to_find == sta_current and self.side_to_find == side:

                    if section == 'constant':
                        if target == 'coord':
                            return x_coord
                        elif target == 'name':
                            return ref_name
                        elif target == 'size':
                            return number
                        else:
                            pass

                    else:
                        if target == 'coord':
                            return x_coord_nonconstant
                        elif target == 'name':
                            return ref_name
                        elif target == 'size':
                            return number
                        else:
                            pass

                if section == 'constant':
                    trouble2 = mm_to_inch(x_coord)
                    Ref.bin_breaker.append(int(trouble2))
                    Ref.sta_values_fake.append(sta_value(x_coord, plug_value))
                    Ref.sta_value_pairs.append(x_coord)

                elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_41 + x_coord_nonconstant - inch_to_mm(int(number))),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                elif section == 'nonconstant' and side == 'LH' and location == 'tail':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                elif section == 'nonconstant' and side == 'RH' and location == 'tail':
                    Ref.sta_values_fake.append(sta_value(
                        (fake_coord_nonconstant_47 + x_coord_nonconstant),
                        plug_value))
                    Ref.sta_value_pairs.append(x_coord_nonconstant)

                x = x_coord_nonconstant * math.cos(rad)
                y = x_coord_nonconstant * math.sin(rad)

                position = [1, 0, 0, 0, 1, 0, 0, 0, 1, x_coord, 0, 0]
                position_non = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + Rotate5[9], -y + Rotate5[10], 0 + Rotate5[11]]
                position_non_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (inch_to_mm(int(number)) * math.cos(rad)) + Rotate185[9],
                                   y + (inch_to_mm(int(number)) * math.sin(rad)) + Rotate185[10], 0 + Rotate185[11]]
                position90 = [-1, 0, 0, 0, -1, 0, 0, 0, 1, x_coord + inch_to_mm(int(number)), 0, 0]  # 90 deg rotation
                position_non_47 = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + Rotate_5[9], y + Rotate_5[10], 0 + Rotate_5[11]]
                position_non_47_RH = [1, 0, 0, 0, 1, 0, 0, 0, 1, x + (inch_to_mm(int(number)) * math.cos(rad)) + Rotate_185[9],
                                      (y + (inch_to_mm(int(number)) * math.sin(rad))) * (-1) + Rotate_185[10], 0 + Rotate_185[11]]

                if location == 'nose':
                    x_coord_nonconstant -= inch_to_mm(int(number))

                if side == 'LH' and section == 'constant':
                    if self.sta_to_find == sta_current and self.side_to_find == side and target == 'position':   
                        return position[-3:]

                elif side == 'RH' and section == 'constant':
                    if self.sta_to_find == sta_current and self.side_to_find == side and target == 'position':
                        return position90[-3:]

                elif section == 'nonconstant' and side == 'LH' and location == 'nose':
                    if self.sta_to_find == sta_current and self.side_to_find == side and target == 'position':
                        xyz1 = position_non[-3:]
                        xyz1[0] = xyz1[0] - math.cos(rad) * inch_to_mm(int(number))
                        xyz1[1] = xyz1[1] + math.sin(rad) * inch_to_mm(int(number))
                        return xyz1

                elif section == 'nonconstant' and side == 'RH' and location == 'nose':
                    if self.sta_to_find == sta_current and self.side_to_find == side and target == 'position':
                        xyz = position_non_RH[-3:]
                        xyz[0] = xyz[0] - math.cos(rad) * inch_to_mm(int(number))
                        xyz[1] = xyz[1] - math.sin(rad) * inch_to_mm(int(number))
                        return xyz

                elif section == 'nonconstant' and side == 'LH' and location == 'tail':
                    if self.sta_to_find == sta_current and self.side_to_find == side and target == 'position':
                        return position_non_47[-3:]
                    x_coord_nonconstant += inch_to_mm(int(number))

                elif section == 'nonconstant' and side == 'RH' and location == 'tail':
                    if self.sta_to_find == sta_current and self.side_to_find == side and target == 'position':
                        return position_non_47_RH[-3:]
                    x_coord_nonconstant += inch_to_mm(int(number))

                x_coord += inch_to_mm(int(number))

    def section(self):

        sec = 'constant'
        if '+' not in self.sta_to_find:
            if float(self.sta_to_find) < 465.0:
                sec = '41'
                self.angle = 5
            elif float(self.sta_to_find) > 1623.0:
                sec = '47'
            else:
                sec = 'constant'
        return sec

    def find_sta(self, origin, size, kind='sta', **kwargs):

        size = size.replace('\n', '')
        sec = 'constant'
        if '+' not in self.sta_to_find:
            if float(self.sta_to_find) < 465:
                sec = '41'
                self.angle = 5
            elif float(self.sta_to_find) > 1617:
                sec = '47'
                self.angle = 3.125
            else:
                sec = 'constant'
                self.angle = 0
        
        ringposts_dict = {'constant': {'24': [[1.5, 93.9905, 266.25],
                                              [22.5, 93.9905, 266.25]],
                                       '36': [[1.5, 93.9905, 266.25],
                                              [34.5, 93.9905, 266.25]],
                                       '42': [[1.5, 93.9905, 266.25],
                                              [40.5, 93.9905, 266.25]],
                                       '48': [[1.5, 93.9905, 266.25],
                                              [46.5, 93.9905, 266.25]]},
                               '41': {'24': [[-6.6975, 93.7635, 266.25],
                                              [14.2226, 95.5938, 266.25]],
                                      '36': [[-6.6975, 93.7635, 266.25],
                                              [26.1769, 96.6397, 266.25]],
                                      '42': [[-6.6975, 93.7635, 266.25],
                                              [32.1541, 97.1626, 266.25]],
                                      '48': [[-6.6975, 93.7635, 266.25],
                                              [38.1312, 97.6856, 266.25]]},
                               '47': {'24': [[6.6216, 93.7689, 266.25],
                                              [27.5904, 92.6241, 266.25]],
                                      '36': [[6.6216, 93.7689, 266.25],
                                              [39.5725, 91.9700, 266.25]],
                                      '42': [[6.6216, 93.7689, 266.25],
                                              [45.5636, 91.6429, 266.25]],
                                      '48': [[6.6216, 93.7689, 266.25],
                                              [51.5547, 91.3157, 266.25]]}}
                                              
        bushing_dict =   {'constant': {'24': [[5.6856+0.25, (93.9905-0.6519), (266.25+0.6152)],
                                              [23.5+0.25-5.6856, (93.9905-0.6519), (266.25+0.6152)]],
                                       '36': [[5.6856+0.25, (93.9905-0.6519), (266.25+0.6152)],
                                              [35.5+0.25-5.6856, (93.9905-0.6519), (266.25+0.6152)]],
                                       '42': [[5.6856+0.25, (93.9905-0.6519), (266.25+0.6152)],
                                              [41.5+0.25-5.6856, (93.9905-0.6519), (266.25+0.6152)]],
                                       '48': [[5.6856+0.25, (93.9905-0.6519), (266.25+0.6152)],
                                              [47.5+0.25-5.6856, (93.9905-0.6519), (266.25+0.6152)]]},
                               '41': {'24': [[-6.6975, 93.7635, 266.25],
                                              [14.2226, 95.5938, 266.25]],
                                      '36': [[-6.6975, 93.7635, 266.25],
                                              [26.1769, 96.6397, 266.25]],
                                      '42': [[-6.6975, 93.7635, 266.25],
                                              [32.1541, 97.1626, 266.25]],
                                      '48': [[-6.6975, 93.7635, 266.25],
                                              [38.1312, 97.6856, 266.25]]},
                               '47': {'24': [[6.6216, 93.7689, 266.25],
                                              [27.5904, 92.6241, 266.25]],
                                      '36': [[6.6216, 93.7689, 266.25],
                                              [39.5725, 91.9700, 266.25]],
                                      '42': [[6.6216, 93.7689, 266.25],
                                              [45.5636, 91.6429, 266.25]],
                                      '48': [[6.6216, 93.7689, 266.25],
                                              [51.5547, 91.3157, 266.25]]}}
                                              
        if sec == 'constant':        
            offset_STA = [0.25, 60.6643, 284.8828]
            offset_mounting = [(0.25 - 0.0787), 62.1737, 288.7698]
            offset_alignment = [(0.25 + 0.4), 61.8299, 290.500]

        elif sec == '47':

            offset_STA = [3.5567, 60.5605, 284.8828]
            offset_mounting = [3.5604, 62.0719, 288.7698]
            offset_alignment = [4.0197, 61.7025, 290.500]

        elif sec == '41':

            offset_STA = [-5.0382, 60.4552, 284.8828]
            offset_mounting = [-5.2481, 61.9520, 288.7698]
            offset_alignment = [-4.7413, 61.6513, 290.500]

        try:
            offset_ringpost1 = ringposts_dict[sec][size][0]
            offset_ringpost2 = ringposts_dict[sec][size][1]
            offset_bushing1 = bushing_dict[sec][size][0]
            offset_bushing2 = bushing_dict[sec][size][1]

        except KeyError:
            if kind == 'ringpost1' or kind == 'ringpost2' or kind == 'bushing1' or kind  == 'bushing2':
                return None
            else:
                pass
        
        if kind == 'sta':
            offset = offset_STA
        elif kind == 'mounting':
            offset = offset_mounting
        elif kind == 'alignment':
            offset = offset_alignment
        elif kind == 'ringpost1':
            offset = offset_ringpost1
        elif kind == 'ringpost2':
            offset = offset_ringpost2
        elif kind == 'bushing1':
            offset = offset_bushing1
        elif kind == 'bushing2':
            offset = offset_bushing2
        elif kind == 'camera':
            for key in kwargs:
                #if key == coord_origin:
                    offset_inch = kwargs[key]
                    offset = []
                    for i in offset_inch:
                        offset.append(i / 25.4)
                    offset[1] = abs(offset[1])
                    print offset
        
        rad = math.radians(self.angle)

        
        if sec == '41' and self.side_to_find == 'RH':
            sta_offset_LH = [offset[0] * 25.4, -1 * offset[1] * 25.4, offset[2] * 25.4]
            sta_offset_RH = [offset[0] * 25.4 - math.cos(rad) * inch_to_mm(int(size)), offset[1] * 25.4 -
                            math.sin(rad) * inch_to_mm(int(size)), offset[2] * 25.4]
        else:
            sta_offset_LH = [offset[0] * 25.4, -1 * offset[1] * 25.4, offset[2] * 25.4]
            sta_offset_RH = [offset[0] * 25.4 - math.cos(rad) * inch_to_mm(int(size)), offset[1] * 25.4 +
                            math.sin(rad) * inch_to_mm(int(size)), offset[2] * 25.4]


        sta_pos = []
        for i in xrange(3):
            if 'L' in self.side_to_find:
                sta_pos.append(origin[i] + sta_offset_LH[i])
            elif 'R' in self.side_to_find:
                sta_pos.append(origin[i] + sta_offset_RH[i])
        return sta_pos

    def build(self):

        Ref.instantiate_nonconstant_components(self)

        s1, s2, s3, s4, s5, s6 = self.converter()

        if s3 != ['']:
            Ref.add_component(self, s3, 'LH', 'constant', 'middle', self.plug, self.name)
        if s4 != ['']:
            Ref.add_component(self, s4, 'RH', 'constant', 'middle', self.plug, self.name)
        if s1 != ['']:
            Ref.add_component(self, s1, 'LH', 'nonconstant', 'nose', 0, self.name)
        if s2 != ['']:
            Ref.add_component(self, s2, 'RH', 'nonconstant', 'nose', 0, self.name)
        if s5 != ['']:
            Ref.add_component(self, s5, 'LH', 'nonconstant', 'tail', self.plug, self.name)
        if s6 != ['']:
            Ref.add_component(self, s6, 'RH', 'nonconstant', 'tail', self.plug, self.name)

    def get_x_coord(self):

        s1, s2, s3, s4, s5, s6 = self.converter()

        if s3 != ['']:
            result = Ref.add_component_coord(self, s3, 'LH', 'constant', 'middle', self.plug)
            if result is not None:
                return result
        if s4 != ['']:
            result = Ref.add_component_coord(self, s4, 'RH', 'constant', 'middle', self.plug)
            if result is not None:
                return result
        if s1 != ['']:
            result = Ref.add_component_coord(self, s1, 'LH', 'nonconstant', 'nose', 0)
            if result is not None:
                return result
        if s2 != ['']:
            result = Ref.add_component_coord(self, s2, 'RH', 'nonconstant', 'nose', 0)
            if result is not None:
                return result
        if s5 != ['']:
            result = Ref.add_component_coord(self, s5, 'LH', 'nonconstant', 'tail', self.plug)
            if result is not None:
                return result
        if s6 != ['']:
            result = Ref.add_component_coord(self, s6, 'RH', 'nonconstant', 'tail', self.plug)
            if result is not None:
                return result

    def get_ref_name(self):

        s1, s2, s3, s4, s5, s6 = self.converter()

        if s3 != ['']:
            result = Ref.add_component_coord(self, s3, 'LH', 'constant', 'middle', self.plug, 'name')
            if result is not None:
                return result
        if s4 != ['']:
            result = Ref.add_component_coord(self, s4, 'RH', 'constant', 'middle', self.plug, 'name')
            if result is not None:
                return result
        if s1 != ['']:
            result = Ref.add_component_coord(self, s1, 'LH', 'nonconstant', 'nose', 0, 'name')
            if result is not None:
                return result
        if s2 != ['']:
            result = Ref.add_component_coord(self, s2, 'RH', 'nonconstant', 'nose', 0, 'name')
            if result is not None:
                return result
        if s5 != ['']:
            result = Ref.add_component_coord(self, s5, 'LH', 'nonconstant', 'tail', self.plug, 'name')
            if result is not None:
                return result
        if s6 != ['']:
            result = Ref.add_component_coord(self, s6, 'RH', 'nonconstant', 'tail', self.plug, 'name')
            if result is not None:
                return result

    def get_position_sta(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'))

    def get_position_mounting(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='mounting')

    def get_position_alignment(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='alignment')

    def get_position_offset_ringpost1(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='ringpost1')

    def get_position_offset_ringpost2(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='ringpost2')

    def get_position_offset_bushing1(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='bushing1')

    def get_position_offset_bushing2(self):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='bushing2')

    def get_position_camera(self, origin):

        return self.find_sta(self.get_position(), self.get_position(target='size'), kind='camera', coord_origin=origin)

    def get_position(self, target='position'):

        s1, s2, s3, s4, s5, s6 = self.converter()

        if s3 != ['']:
            result = Ref.add_component_coord(self, s3, 'LH', 'constant', 'middle', self.plug, target)
            if result is not None:
                return result
        if s4 != ['']:
            result = Ref.add_component_coord(self, s4, 'RH', 'constant', 'middle', self.plug, target)
            if result is not None:
                return result
        if s1 != ['']:
            result = Ref.add_component_coord(self, s1, 'LH', 'nonconstant', 'nose', 0, target)
            if result is not None:
                return result
        if s2 != ['']:
            result = Ref.add_component_coord(self, s2, 'RH', 'nonconstant', 'nose', 0, target)
            if result is not None:
                return result
        if s5 != ['']:
            result = Ref.add_component_coord(self, s5, 'LH', 'nonconstant', 'tail', self.plug, target)
            if result is not None:
                return result
        if s6 != ['']:
            result = Ref.add_component_coord(self, s6, 'RH', 'nonconstant', 'tail', self.plug, target)
            if result is not None:
                return result

    def remove_component(self):

        ICM = Ref.CATIA.ActiveDocument
        ICM_1 = ICM.Product
        ICM_Products_irms = ICM_1.Products
        ICM_Product = ICM_Products_irms.Item(self.name)
        ICM_Products = ICM_Product.Products

        try:
            ICM_Products.Remove(self.component_name)
            return 0
        except:
            pass
        non_const_prods = [self.name + '_' + 'non-constant_41_LH.1', self.name + '_' + 'non-constant_41_RH.1',
                           self.name + '_' + 'non-constant_47_LH.1', self.name + '_' + 'non-constant_47_RH.1']
        for i in non_const_prods:
            prod = ICM_Products.Item(i)
            prods = prod.Products
            try:
                prods.Remove(self.component_name)
                break
            except:
                continue

if __name__ == "__main__":

    #ecs = Ref('787_9_KAL_ZB656', '0465', 'LH', 240, 1, 4, [], name='GLS_STA0561-0657_OB_LH_CAI')
    #ecs1 = Ref('787_9_KAL_ZB656', '0609', 'LH', 240, 2, 4, [], name='GLS_STA0561-0657_OB_LH_CAI')
    #ecs2 = Ref('787_9_KAL_ZB656', '0609+48', 'LH', 240, 3, 4, [], name='GLS_STA0561-0657_OB_LH_CAI')
    ecs3 = Ref('787_9_KAL_ZB656', '0609+96', 'LH', 240, 4, 4, ['1X5005-210000##ALT68'], name='GLS_STA0561-0657_OB_LH_CAI')
    #ecs.build()
    #ecs1.build()
    #ecs2.build()
    ecs3.build()
    #ecs.remove_component()
    #ecs1.remove_component()
    #ecs2.remove_component()
    #ecs3.remove_component()
    #name = ecs.get_ref_name()

#    ecs4 = Ref('787_9_KAL_ZB656', '0465', 'LH', 240, name='new_instance')
#    ecs4.build()
#    ecs4.remove_component()

#    bins = []
#    ecs = Ref('BRI', '0345', 'LH')
#    ecs1 = Ref('BRI', '0345', 'RH')
#    ecs2 = Ref('BRI', '0369', 'LH')
#    ecs3 = Ref('BRI', '0369', 'RH')
#    ecs4 = Ref('BRI', '0411', 'LH')
#    ecs5 = Ref('BRI', '0411', 'RH')
#    ecs6 = Ref('BRI', '1689', 'LH')
#    ecs7 = Ref('BRI', '1689', 'RH')
#    ecs8 = Ref('BRI', '1665', 'LH')
#    ecs9 = Ref('BRI', '1665', 'RH')
#    ecs10 = Ref('BRI', '0465', 'LH')
#    ecs11 = Ref('BRI', '0465', 'RH')
#    bins.extend((ecs, ecs1, ecs2, ecs3, ecs4, ecs5, ecs6, ecs7, ecs8, ecs9, ecs10, ecs11))
#    for ecs in bins:
#        ecs.build()
#        coord = ecs.get_x_coord()
#        name = ecs.get_ref_name()
#        position = ecs.get_position()
#        sta = ecs.get_position_sta()
#        mounting = ecs.get_position_mounting()
#        alignment = ecs.get_position_alignment()
#        ringpost1 = ecs.get_position_offset_ringpost1()
#        ringpost2 = ecs.get_position_offset_ringpost2()
#        print 'x coordinate in mm: ' + str(coord)
#        print 'reference surface name: ' + str(name)
#        print 'origin coordinates: ' + str(position)
#        print 'sta point coordinates: ' + str(sta)
#        print 'mounting point coordinates: ' + str(mounting)
#        print 'alignment point coordinates: ' + str(alignment)
#        print 'ringpost1 point coordinates: ' + str(ringpost1)
#        print 'ringpost2 point coordinates: ' + str(ringpost2)
#
#        
#        CATIA = win32com.client.Dispatch('catia.application')
#        productDocument1 = CATIA.ActiveDocument
#        documents = CATIA.Documents
#        Product = productDocument1.Product
#        collection = Product.Products
#        part1 = collection.AddNewComponent("Part", ecs.sta_to_find + '_' + ecs.side_to_find)
#        part12 = collection.Item(collection.Count)
#        name1 = part12.PartNumber
#        print name1
#        current_doc = documents.Item(name1 + '.CATPart')
#        current_part = current_doc.Part
#        HybridShapeFactory1 = current_part.HybridShapeFactory
#        hybrid_bodies1 = current_part.HybridBodies
#        hybrid_body1 = hybrid_bodies1.Item('Geometrical Set.1')
#        points = []
#        names = ['sta', '1X5005-412(XXX) Mounting Feature', '1X5005-412(XXX) Alignment Feature',
#                 'ringpost1', 'ringpost2']
#        points.extend((sta, mounting, alignment, ringpost1, ringpost2))
#        for i in xrange(len(points)):
#            try:
#                point_added = HybridShapeFactory1.AddNewPointCoord(points[i][0], points[i][1], points[i][2])
#            except:
#                continue    
#            if hybrid_body1 is not None:
#                hybrid_body1.AppendHybridShape(point_added)
#                point_added.Name = names[i]
#        current_part.Update()
