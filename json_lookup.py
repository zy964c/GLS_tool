import json
#import pprint


def json_lookup(pn):
    """
    returns points coordinates
    """
    with open('lights.json') as f:
        data = json.load(f)
        try:
            points_dict = data[pn]
            placeholder = points_dict.values()
            return placeholder
        except:
            return None


def json_lookup_name(pn):
    """
    returns name of the part
    """
    with open('names.json') as f:
        data = json.load(f)
        try:
            name = data[pn]
            return str(name)
        except:
            return None


def json_lookup_origin(instance_id):
    """
    returns instance rotation matrix
    """
    with open('coord.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[instance_id]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            return result
        except:
            return None


def json_lookup_point(carm_pn, point_name):
    """
    returns point coordinates
    """
    with open(carm_pn + '.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[point_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            return result
        except:
            return None


def json_lookup_fl(pn):
    """
    returns points coordinates
    """
    with open('flagnote.json') as f:
        data = json.load(f)
        try:
            points_dict = data[pn]
            placeholder = points_dict.values()
            return placeholder
        except:
            return None


def json_lookup_fl_keys(pn):
    """
    returns points coordinates
    """
    with open('flagnote.json') as f:
        data = json.load(f)
        try:
            points_dict = data[pn]
            placeholder = points_dict.keys()
            return placeholder
        except:
            return None


def json_lookup_flagnote(carm_pn, point_name):
    """
    returns point coordinates
    """
    with open('flagnote_' + carm_pn + '.txt') as f:
        data = json.load(f)
        try:
            matrix_str = data[point_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            return result
        except:
            return None


def json_lookup_axis(view):
    """
    returns coord of an axis
    """
    axis_name = view.replace('Text Plane', 'Axis System')
    with open('axis.txt') as f:  
        data = json.load(f)
        try:
            matrix_str = data[axis_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            return result
        except:
            return None


def json_lookup_camera(camera_name):
    """
    returns coord of an axis
    """
    with open('cameras.txt') as f:  
        data = json.load(f)
        try:
            matrix_str = data[camera_name]
            matrix_list = matrix_str.replace(',', '').split()
            result = []
            for j in matrix_list:
                result.append(float(str(j)))
            return result
        except:
            return None
            
def json_lookup_stonesoup():
    """
    returns points coordinates
    """
    with open('Layout_FromXML.json') as f:
        data = json.load(f)
    return data
    
def json_lookup_components(pn):
    """
    returns points coordinates
    """
    with open('jd_vectors.json') as f:
        data = json.load(f)
        try:
            components = data[pn]
            return components
        except:
            return None


if __name__ == "__main__":
    
    print json_lookup('C519510-33')
    print json_lookup_name('1J5009-221136-0##ALT4')
    print json_lookup_name('C519510-33')
    print json_lookup_name('C519503-523##ALT2')
    print json_lookup_name('836Z1510-9')
    print json_lookup_name('1X5005-300000-0##ALT69')
    print json_lookup_name('836Z1510-1')
    print json_lookup_name('1J5009-301024-0##ALT1')
    print json_lookup_origin('1631-10_STA0561-0657_OB-OMF_LH_CARM')
    print json_lookup_point('CA836Z1631-1', 'Point.192')
    print json_lookup_axis('Inboard Facing Out - Top Support - Text Plane 41 RH')
    print json_lookup_camera("JD01 Ceiling Light Typical")
    #j = json_lookup_stonesoup()
    #for key in j:
    #    if 'Layout' in key:
    #        pprint.pprint(j[key])
