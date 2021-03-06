from json_lookup import json_lookup_origin
import re
import math
import win32com.client


class JD(object):

    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    Product = productDocument1.Product
    collection = Product.Products

    def __init__(self, pn, irm_name, part_name, plug_value, bin_type):
        self.pn = pn
        self.irm_name = irm_name
        self.part_name = part_name
        self.plug_value = plug_value
        self.bin_type = bin_type
        self.product = JD.collection.Item(self.irm_name)
        self.collection1 = self.product.Products

    def __str__(self):
        return '({0.pn!s}, {0.irm_type!s})'.format(self)

    def get_name(self):
        return self.find_name()

    def get_part_name(self):
        return self.part_name

    def find_name(self):

        re_ceiling_light = re.compile('1J5009-22\d{4}-\d##ALT\d+')
        re_noise_seal = re.compile('C519510-\d+')
        re_sidewall_light_bracket = re.compile('1X5005-420\d00-\d')
        re_disconnect_bracket = re.compile('836Z1510-1$')
        re_disconnect_bracket_36 = re.compile('836Z1510-2$')
        re_disconnect_bracket_24 = re.compile('836Z1510-3$')
        re_disconnect_bracket_48_42 = re.compile('836Z1510-22$')
        re_disconnect_bracket2 = re.compile('836Z1510-24$')
        re_disconnect_bracket_36_ringpost = re.compile('836Z1510-2_ringpost$')
        re_disconnect_bracket_24_ringpost = re.compile('836Z1510-3_ringpost$')
        re_disconnect_bracket_48_42_ringpost = re.compile('836Z1510-22_ringpost$')
        re_disconnect_bracket2_pclamp = re.compile('836Z1510-24_pclamp$')
        re_sight_block = re.compile('836Z15\d0-([5-9]|12|13)')
        re_ceiling_latch = re.compile('C519503-52(3|5).*')
        re_ofcr_ob_latch = re.compile('C519502-11$')
        #re_sidewall_light_bracket = re.compile('1X5005-420\d00-\d')
        re_riser_to_light = re.compile('C519503-515##ALT1$')
        re_nofar_light = re.compile('1J5009-3010\d{2}-\d##ALT\d+')
        re_power_supply = re.compile('1X5005-300000-\d##ALT\d+')
        re_ringpost = re.compile('^(IC830Z3000-1.+rp.*(FWD|AFT)*)')
        re_suddle_clamp = re.compile('^(IC830Z3000-1.+jd31)$')
        re_ringpost_rail = re.compile('^(IC830Z3000-1.+jd28)$')
        re_bushing = re.compile('^(IC830Z3000-1[^jd]*.[^rp]$)')
        re_mounting_bracket = re.compile('1X5005-411100-\d+')
        re_inboard_ceiling_lights = re.compile('1J5009-211\d{3}-\d##ALT\d+')
        #re_inboard_latches = re.compile('C519503-525$')
        re_nofar_light_mounting_bracket = re.compile('1X5005-431100-1$')
        re_nofar_light_mounting_bracket2 = re.compile('1X5005-431200-1$')
        re_power_supply_mounting_bracket = re.compile('832Z7450-1$')
        # Have not found JD for the two lights below
        re_ofcr_light = re.compile('1J5009-41[1,2,3]0\d{2}-\d##ALT\d+')
        re_ofar_light = re.compile('1J5009-31[1,2]0\d{2}-\d##ALT\d+')

        jd_dict_new = {re_ceiling_light: '01',
                       re_noise_seal: '02',
                       re_sidewall_light_bracket: '03',
                       re_disconnect_bracket: '04',
                       re_disconnect_bracket_36: '05',
                       re_disconnect_bracket_24: '06',
                       re_disconnect_bracket_48_42: '07',
                       re_disconnect_bracket2: '08',
                       re_ringpost: '09',
                       re_disconnect_bracket_36_ringpost: '11',
                       re_disconnect_bracket_24_ringpost: '12',
                       re_disconnect_bracket_48_42_ringpost: '13',
                       re_disconnect_bracket2_pclamp: '14',                      
                       re_sight_block: '15',
                       re_bushing: '17',
                       re_ceiling_latch: '19',
                       re_ofcr_ob_latch: '21',
                       re_riser_to_light: '24',
                       re_nofar_light: '26',
                       re_power_supply: '27',
                       re_ringpost_rail: '28',
                       re_suddle_clamp: '31',
                       re_mounting_bracket: '40',
                       re_inboard_ceiling_lights: '41',
                       #re_inboard_latches: '42',
                       re_nofar_light_mounting_bracket: '45',
                       re_nofar_light_mounting_bracket2: '46',
                       re_power_supply_mounting_bracket: '47'}

        re_list = jd_dict_new.keys()

        for r in re_list:
            m = r.match(self.pn)
            if m is not None:
                jd_set = jd_dict_new[r]
                if jd_set == '19' and self.bin_type == 1:
                    if self.check_slc_light():
                        jd_set = '25'
                    elif self.check_riser():
                        jd_set = '20'
                elif jd_set == '19' and self.bin_type == 2:
                    jd_set = '42'
                elif jd_set == '04':
                    if self.check_disconnect_brkt(self.plug_value):
                        jd_set = '29'
                elif jd_set == '03':
                    if not self.sidewall_light_bracket():
                        if self.check_sidewall():
                            jd_set = '22'
                        else:
                            jd_set = '23'
                elif jd_set == '27' and self.bin_type == 2:
                    jd_set = '48'
                return jd_set                      
            else:
                continue

        return None


    def get_slc_origins(self):

        x_slc = []
        for n in xrange(1, self.collection1.Count+1):
            part_pn = self.collection1.Item(n).PartNumber
            if '222348' in part_pn or '221348' in part_pn:
                x_coord = json_lookup_origin(self.collection1.Item(n).Name)[-3]
                x_slc.append(x_coord)
            else:
                continue
        if len(x_slc) > 0:
            return x_slc
        else:
            return None

    def get_riser_origins(self, detail='raisers'):
        
        x_riser = []
        d_type = {'raisers': 'C519503-515', 'sidewall': '1X5005-420100-1'}
        pn_to_find = d_type[detail]
        for n in xrange(1, self.collection1.Count+1):
            part_pn = self.collection1.Item(n).PartNumber
            if pn_to_find in part_pn:
                x_coord = json_lookup_origin(self.collection1.Item(n).Name)[-3]
                x_riser.append(x_coord)
            else:
                continue
        if len(x_riser) > 0:
            return x_riser
        else:
            return None

    def check_riser(self):

        x_coord_latch = json_lookup_origin(self.part_name)[-3]
        x_coord_latch_list = [x_coord_latch+(25.4*355.273),
                              x_coord_latch-(25.4*355.273)]
        riser_list = self.get_riser_origins()
        #print riser_list
        if riser_list is not None:
            for riser in riser_list:
                for latch in x_coord_latch_list:
                    if (math.ceil(riser*1)/1) == (math.ceil(latch*1)/1):
                        return True
        return False

    def check_sidewall(self):

        x_coord_sidewall = json_lookup_origin(self.part_name)[-3]
        #print x_coord_sidewall
        sidewall_list = self.get_riser_origins('sidewall')
        #print sidewall_list
        if sidewall_list is not None:
            sidewall_list.append(x_coord_sidewall)
            sidewall_list.sort()
            positions = [i for i,x in enumerate(sidewall_list) if x == x_coord_sidewall]
            if len(positions) > 0:
                order = positions[0] + 1
                if order % 2 == 0:
                    return False
                else:
                    return True
            else:
                return True

    def check_slc_light(self):

        x_coord_latch = json_lookup_origin(self.part_name)[-3]
        x_coord_latch_list = [x_coord_latch, x_coord_latch+(25.4*42.0),
                              x_coord_latch-(25.4*42.0)]
        slc_list = self.get_slc_origins()
        if slc_list is not None:
            for slc in slc_list:
                for latch in x_coord_latch_list:
                    #print math.ceil(slc*1)/1
                    #print math.ceil(latch*1)/1
                    if (math.ceil(slc*1)/1) == (math.ceil(latch*1)/1):
                        return True
        return False

    def check_disconnect_brkt(self, plug_value):
        
        pos = json_lookup_origin(self.part_name)
        coords = slice(9, 12)
        wl = 275.0 * 25.4
        s47_sta = (1617.0 + plug_value) * 25.4
        if pos[coords][2] > wl:
            if pos[coords][0] > s47_sta:
                return True
        return False

    def sidewall_light_bracket(self):

        pos = json_lookup_origin(self.part_name)
        coords = slice(9, 12)
        s41_sta = 465.0 * 25.4
        if pos[coords][0] < s41_sta:
            return True
        return False
           
#    def if_slc(self):
#if __name__ == "__main__":

    #irm_name = 'GLS_STA0561-0657_OB_LH_CAI'
    #catia = win32com.client.Dispatch('catia.application')
    #productDocument1 = catia.ActiveDocument
    #Product = productDocument1.Product
    #collection = Product.Products
    #product = collection.Item(irm_name)
    #collection1 = product.Products
    #for n in xrange(1, collection1.Count+1):
    #    part_pn = collection1.Item(n).PartNumber
    #    part_name = collection1.Item(n).Name
    #    j1 = JD(part_pn, irm_name, part_name)
    #    print str(j1.get_part_name()) + ' : ' + str(j1.get_name())
