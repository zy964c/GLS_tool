import json
#import pprint
from time import gmtime, strftime


def json_lookup(pn):
    """
    returns points coordinates
    """
    with open('data\\lights.json') as f:
        data = json.load(f)
        #print data
        try:
            points_dict = data[pn]
            placeholder = points_dict.values()
            #f.close()
            print pn
            return placeholder
        except:
            current_time = strftime("%Y_%m_%d_%H_%M_%S", gmtime())
            log = open('temp\\log.txt', 'a')
            log.write(current_time + ' - missing from joint definition library: ' + pn + '\n')
            log.close()
            alt_pos_index = pn.find('#')
            if alt_pos_index == -1:
                return None
            pn_base = pn[:alt_pos_index]
            for key in data:
                if pn_base in key:
                    return json_lookup(key)
            return None


def json_lookup_name(pn):
    """
    returns name of the part
    """
    with open('data\\names.json') as f:
        data = json.load(f)
        try:
            name = data[pn]
            #f.close()
            return str(name)
        except:
            return None


def json_lookup_origin(instance_id):
    """
    returns instance rotation matrix
    """
    with open('temp\\coord.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[instance_id]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            #f.close()
            return result
        except:
            return None


def json_lookup_point(carm_pn, point_name):
    """
    returns point coordinates
    """
    with open('temp\\' + carm_pn + '.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[point_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            #f.close()
            return result
        except:
            return None


def json_lookup_fl(pn):
    """
    returns points coordinates
    """
    with open('data\\flagnote.json') as f:
        data = json.load(f)
        try:
            points_dict = data[pn]
            placeholder = points_dict.values()
            print pn
            #f.close()
            return placeholder
        except:
            current_time = strftime("%Y_%m_%d_%H_%M_%S", gmtime())
            log = open('temp\\log.txt', 'a')
            log.write(current_time + ' - missing from flagnote library: ' + pn + '\n')
            log.close()
            alt_pos_index = pn.find('#')
            if alt_pos_index == -1:
                return None
            pn_base = pn[:alt_pos_index]
            for key in data:
                if pn_base in key:
                    return json_lookup_fl(key)
            return None


def json_lookup_fl_keys(pn):
    """
    returns points coordinates
    """
    with open('data\\flagnote.json') as f:
        data = json.load(f)
        try:
            points_dict = data[pn]
            placeholder = points_dict.keys()
            print pn
            #f.close()
            return placeholder
        except:
            alt_pos_index = pn.find('#')
            if alt_pos_index == -1:
                return None
            pn_base = pn[:alt_pos_index]
            for key in data:
                if pn_base in key:
                    return json_lookup_fl_keys(key)
            return None



def json_lookup_flagnote(carm_pn, point_name):
    """
    returns point coordinates
    """
    with open('temp\\flagnote_' + carm_pn + '.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[point_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            #f.close()
            return result
        except:
            return None


def json_lookup_axis(view):
    """
    returns coord of an axis
    """
    axis_name = view.replace('Text Plane', 'Axis System')
    with open('data\\axis.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[axis_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            #f.close()
            return result
        except:
            return None


def json_lookup_camera(camera_name):
    """
    returns coord of an axis
    """
    with open('data\\cameras.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[camera_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            #f.close()
            return result
        except:
            return None


def json_lookup_stonesoup():
    """
    returns points coordinates
    """
    with open('Layout_FromXML.json') as f:
        data = json.load(f)
    #f.close()
    return data


def json_lookup_components(pn):
    """
    returns points coordinates
    """
    with open('data\\jd_vectors.json') as f:
        data = json.load(f)
        try:
            components = data[pn]
            #f.close()
            return components
        except:
            return None


if __name__ == "__main__":
    
    #print json_lookup('IC830Z3000-1.13.2_jd28')
    print json_lookup_fl('IC830Z3000-1.3.2.11')
    print json_lookup_fl_keys('IC830Z3000-1.3.2.11')
    #print json_lookup_origin('1631-10_STA0561-0657_OB-OMF_LH_CARM')
    #print json_lookup_axis('Inboard Facing Out - Top Support - Text Plane 41 RH')
    #print json_lookup_camera("JD01 Ceiling Light Typical")
    #j = json_lookup_stonesoup()
    #for key in j:
    #    if 'Layout' in key:
    #        pprint.pprint(j[key])
