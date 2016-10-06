import json
##from pprint import pprint
from collections import OrderedDict
#from functions import sta_value
import codecs


def door_sta(plug_value):

    if plug_value == 240:
        door_station = (690.0 + 120.0)
    elif plug_value == 456:
        door_station = (690.0 + 240.0)
    else:
        door_station = 690.0
    return door_station


def parse_ss(layout, plug_value):

    #with open(layout) as f:
    # decode json to utf-8
    j = json.load(codecs.open(layout, 'r', 'utf-8-sig'))

    sec41_lh = {}
    sec41_rh = {}
    constant_lh = {}
    constant_rh = {}
    sec47_lh = {}
    sec47_rh = {}
    aft_dr2_lh = False
    aft_dr2_rh = False
    allowed_types = ['Bin', 'OverMonumentFairing', 'OverDoorFairing']
    #j = json_lookup_stonesoup()
    content = j['Layout']['Children']
    #print len(content)
    #pprint(content[0])
    door_station = door_sta(plug_value)
    for i in content:
        if i["ObjectType"] == 'Stowbin':
            sta = i["STA"]
            side = i["Column"]
            bin_type = i["StowbinType"]
            size = int(i["Width"])
            print 'STA: ' + str(i["STA"]) + ' Side: ' + str(i["Column"]) + ' Type: ' + str(i["StowbinType"]) + ' Size: ' + str(i["Width"])
            if 'Fairing' in bin_type:
                value = str(size) + ' fairing'
            else:
                value = str(size)
            if bin_type in allowed_types:
                if sta < 465.0:
                    if side == 'Left':
                        sec41_lh[sta] = value
                    elif side == 'Right':
                        sec41_rh[sta] = value
                elif 465.0 <= sta < (1617.0 + plug_value):
                    if sta >= door_station:
                        if side == 'Left' and aft_dr2_lh is False:
                            constant_lh[door_station] = 'door'
                            aft_dr2_lh = True
                        elif side == 'Right' and aft_dr2_rh is False:
                            constant_rh[door_station] = 'door'
                            aft_dr2_rh = True
                    if side == 'Left':
                        constant_lh[sta] = value
                    elif side == 'Right':
                        constant_rh[sta] = value
                else:
                    if side == 'Left':
                        sec47_lh[sta] = value
                    elif side == 'Right':
                        sec47_rh[sta] = value

    sec41_lh_ordered = OrderedDict(sorted(sec41_lh.items(), key=lambda t: t[0]))
    sec41_rh_ordered = OrderedDict(sorted(sec41_rh.items(), key=lambda t: t[0]))
    constant_lh_ordered = OrderedDict(sorted(constant_lh.items(), key=lambda t: t[0]))
    constant_rh_ordered = OrderedDict(sorted(constant_rh.items(), key=lambda t: t[0]))
    sec47_lh_ordered = OrderedDict(sorted(sec47_lh.items(), key=lambda t: t[0]))
    sec47_rh_ordered = OrderedDict(sorted(sec47_rh.items(), key=lambda t: t[0]))

    print sec41_lh_ordered
    print sec41_rh_ordered
    print constant_lh_ordered
    print constant_rh_ordered
    print sec47_lh_ordered
    print sec47_rh_ordered

    #output_name = layout.replace('.json', '.txt')
    output_name = j['Layout']["MajorModel"] + '_' + j['Layout']["MinorModel"] + '_' + j['Layout']["Customer"] + '_' + j['Layout']["Effectivity"] + '.txt'
    f = open(output_name,'w')
    f.write('#Sec41 LH:\n')
    for k in range(len(sec41_lh_ordered.values())-1):
        f.write(sec41_lh_ordered.values()[k] + ', ')
    try:
        f.write((sec41_lh_ordered.values())[-1])
    except IndexError:
        print 'export issue: list index out of range'
    f.write('\n')
    f.write('#Sec41 RH:\n')
    for t in range(len(sec41_rh_ordered.values())-1):
        f.write(sec41_rh_ordered.values()[t] + ', ')
    try:
        f.write((sec41_rh_ordered.values())[-1])
    except IndexError:
        print 'export issue: list index out of range'
    f.write('\n')
    f.write('#constant LH:\n')
    for u in range(len(constant_lh_ordered.values())-1):
        f.write(constant_lh_ordered.values()[u] + ', ')
    try:
        f.write((constant_lh_ordered.values())[-1])
    except IndexError:
        print 'export issue: list index out of range'
    f.write('\n')
    f.write('#constant RH:\n')
    for p in range(len(constant_rh_ordered.values())-1):
        f.write(constant_rh_ordered.values()[p] + ', ')
    try:
        f.write((constant_rh_ordered.values())[-1])
    except IndexError:
        print 'export issue: list index out of range'
    f.write('\n')
    f.write('#Sec47 LH:\n')
    for y in range(len(sec47_lh_ordered.values())-1):
        f.write(sec47_lh_ordered.values()[y] + ', ')
    try:
        f.write((sec47_lh_ordered.values())[-1])
    except IndexError:
        print 'export issue: list index out of range'
    f.write('\n')
    f.write('#Sec47 RH:\n')
    for w in range(len(sec47_rh_ordered.values())-1):
        f.write(sec47_rh_ordered.values()[w] + ', ')
    try:
        f.write((sec47_rh_ordered.values())[-1])
    except IndexError:
        print 'export issue: list index out of range'
    f.write('\n')
    f.close()

    return output_name

    #pprint(sec41_lh)
    #pprint(sec41_rh)
    #pprint(constant_lh)
    #pprint(constant_rh)
    #pprint(sec47_lh)
    #pprint(sec47_rh)

if __name__ == "__main__":

    parse_ss('JAL_ZB424.json', 240)

            
                                                                     

