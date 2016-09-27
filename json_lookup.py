import json
#import pprint


def json_lookup(pn):
    """
    returns points coordinates
    """
    with open('C:\\temp\\GLS\\lights.json') as f:
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
    with open('C:\\temp\\GLS\\coord.txt') as f:
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
    with open('C:\\temp\\GLS\\' + carm_pn + '.txt') as f:
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
    with open('C:\\temp\\GLS\\flagnote.json') as f:
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
    with open('C:\\temp\\GLS\\flagnote.json') as f:
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
    with open('C:\\temp\\GLS\\flagnote_' + carm_pn + '.txt') as f:
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
    with open('C:\\temp\\GLS\\axis.txt') as f:
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
    with open('C:\\temp\\GLS\\cameras.txt') as f:
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
    with open('C:\\temp\\GLS\\jd_vectors.json') as f:
        data = json.load(f)
        try:
            components = data[pn]
            return components
        except:
            return None


if __name__ == "__main__":
    
    print json_lookup('C519510-33')
    print json_lookup_origin('1631-10_STA0561-0657_OB-OMF_LH_CARM')
    print json_lookup_axis('Inboard Facing Out - Top Support - Text Plane 41 RH')
    print json_lookup_camera("JD01 Ceiling Light Typical")
    #j = json_lookup_stonesoup()
    #for key in j:
    #    if 'Layout' in key:
    #        pprint.pprint(j[key])
