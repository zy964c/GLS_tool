import win32com.client
import transformations
import numpy
from external_component import add_carm_as_external_component
from json_lookup import json_lookup_point
#import math
import pdb


def add_std_ref1(carm_pn, irm_instance_id, documents, selection1, products):

    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    hybridBodies1 = carm_part.HybridBodies
    hybridBody1 = hybridBodies1.Item("Reference Geometry")
    hybridBodies2 = hybridBody1.HybridBodies
    hybridBody2 = hybridBodies2.Item("Non-Instantiated Standard Parts")
    hybridBodyJD = hybridBodies1.Item("Joint Definitions")
    hybridBodiesJD = hybridBodyJD.HybridBodies
    axisSystems1 = carm_part.AxisSystems
    axis = axisSystems1.Item('Local Axis System')
    reference1 = carm_part.CreateReferenceFromObject(axis)
    HybridShapeFactory1 = carm_part.HybridShapeFactory
    part_num = 'std_parts'
    #added_part = add_carm_as_external_component(part_num + ' ' + carm_pn, irm_instance_id, inserted=part_num)
    if hybridBodiesJD.Count > 0:
        #init_pos = [0, 0, 0]
        #axis_rotation = [1, 0, 0]
        for geoset in xrange(1, hybridBodiesJD.Count+1):
            hb = hybridBodiesJD.Item(geoset)
            points = hb.HybridShapes
            if points.Count < 1:
                continue
            added_part = add_carm_as_external_component(part_num + ' ' + carm_pn, irm_instance_id, inserted=part_num)
            init_pos = [0, 0, 0]
            axis_rotation = [1, 0, 0]
            #glob_fivd = []
            for fidv in xrange(1, points.Count+1):
                if 'FIDV' in points.Item(fidv).Name:
                    #pdb.set_trace()
                    reference2 = carm_part.CreateReferenceFromObject(points.Item(fidv))
                    line_dir = HybridShapeFactory1.AddNewDirection(reference2)
                    line_dir.RefAxisSystem = reference1
                    glob_fivd = [line_dir.GetXVal(), line_dir.GetYVal(), line_dir.GetZVal()]

            for point in xrange(1, points.Count+1):
                try:
                    if len(glob_fivd) == 0:
                        break
                except UnboundLocalError:
                    break
                if not 'FIDV' in points.Item(point).Name:
                    coordinates = json_lookup_point(carm_pn, points.Item(point).Name)
                    pos = [coordinates[0]-init_pos[0], coordinates[1]-init_pos[1], coordinates[2]-init_pos[2]]
                    if axis_rotation == glob_fivd:
                        new_position = [1, 0, 0, 0, 1, 0, 0, 0, 1] + pos
                    else:
                        new_position = add_std_ref(glob_fivd, coord=pos, glob_x_axis=axis_rotation)
                    #print new_position
                    axis_rotation = glob_fivd
                    added_part[1].Move.Apply(new_position)
                    init_pos = [coordinates[0], coordinates[1], coordinates[2]]
                    selection1.Clear()
                    selection1.Add(added_part[1])
                    selection1.Search(str('(NAME = *JD' + hb.Name[-2:] + '*), sel'))
                    if selection1.Count2 < 1:
                        print 'JD' + hb.Name[-2:] + ' std part representation not found'
                        continue
                    try:
                        selection1.Copy()
                    except:
                        'selection error'
                        continue
                    selection1.Clear()
                    #print 'JD' + hb.Name[-2:]
                    selection1.Add(hybridBody2)
                    selection1.PasteSpecial('CATPrtResultWithOutLink')
                    #selection1.Paste()
                    selection1.Clear()
                    carm_part.Update()

            products.Remove(added_part[1].name)
        std_parts_surfaces = hybridBody2.HybridShapes
        if std_parts_surfaces.Count > 0:
            print 'changing color settings of each standard part representations'
            for surface in xrange(1, std_parts_surfaces.Count+1):
                selection1.Clear()
                selection1.Add(std_parts_surfaces.Item(surface))
                red, green, blye, inh_flag = change_color(std_parts_surfaces.Item(surface).Name)
                selection1.visProperties.SetRealColor(red, green, blye, inh_flag)
                selection1.Clear()


def add_std_ref(glob_fivd, coord=[0, 0, 0], glob_x_axis=[1, 0, 0]):

    perp_v = transformations.vector_product(glob_x_axis, glob_fivd)
    ang_v = transformations.angle_between_vectors(glob_x_axis, glob_fivd)
    rotation_matrix = transformations.rotation_matrix(-ang_v, perp_v)
    return [item for sublist in [list(tup[:-1]) for tup in rotation_matrix[:-1]] for item in sublist] + coord


def change_color(name):

    parts_color_dict = {'JD02 NY-5G-53-10 REF': 'grey', 'JD02 NY-5P-53-3-51 REF': 'grey',
                        'JD11 BACN10YR3CD REF': 'grey',
                        'JD11 BACW10BP3NPK REF': 'red', 'JD12 BACN10YR3CD REF': 'grey', 'JD12 BACW10BP3NPK REF': 'blue',
                        'JD13 BACN10YR3CD REF': 'grey', 'JD13 BACW10BP3NPK REF': 'blue', 'JD14 BACN10YR3CD REF': 'grey',
                        'JD14 BACW10BP3NPK REF': 'red', 'JD19 BACN11AL1CM REF': 'grey', 'JD21 BACN11AL1CM REF': 'grey',
                        'JD24 BACN11AL1CM REF': 'grey', 'JD27 BACN10YR3CD REF': 'grey', 'JD27 BACW10BP3NPK REF': 'grey',
                        'JD28 BACN10YR3CD REF': 'grey', 'JD28 BACW10BP3NPK REF': 'red', 'JD29 BACN10YR3CD REF': 'grey',
                        'JD29 BACW10BP3NPK REF': 'red', 'JD30 BACN10YR3CD REF': 'grey', 'JD30 BACW10BP3NPK REF': 'red',
                        'JD31 BACN10YR3CD REF': 'grey', 'JD31 BACW10BP3NPK REF': 'red'}

    colors_dict = {'red': (255, 0, 0, 0), 'blue': (0, 128, 255, 0), 'grey': (196, 179, 209, 0),
                   'yellow': (255, 255, 0, 0)}
    try:
        color_name = parts_color_dict[name]
    except KeyError:
        color_name = 'yellow'
    return colors_dict[color_name]


if __name__ == "__main__":

    #global_x_axis = [1, 0, 0]
    #b = [-6.4778e-014, 0.72696, -0.68668]
    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    documents = catia.Documents
    product1 = productDocument1.Product
    collection_irms = product1.Products
    irm = collection_irms.Item('GLS_STA1618-1732_OB_LH_CAI')
    collection_irms_irm = irm.Products
    #add_std_ref('CA836Z1131-46', 'GLS_STA0561-0657_OB_LH_CAI')
    #add_std_ref('CA836Z1191-46', 'GLS_STA1618-1732_OB_LH_CAI')
    #add_std_ref('CA836Z1231-46', 'GLS_STA0561-0657_OB_RH_CAI')
    add_std_ref1('CA836Z1191-46_2017_01_27_10_40_53', 'GLS_STA1618-1732_OB_LH_CAI', documents, selection1,
                 collection_irms_irm)

    #perp_v = transformations.vector_product(global_x_axis, b)
    #print perp_v
    #ang_v = transformations.angle_between_vectors(global_x_axis, b, directed=True)
    #print math.degrees(ang_v)
    #R = transformations.rotation_matrix(-ang_v, perp_v)
    #print R
    #v0 = numpy.random.random(3)
    #print v0
    #v1 = transformations.vector_norm(v0)
    #print type(R)
    #my_list = [list(tup[:-1]) for tup in R[:-1]]
    #flatten = [item for sublist in [list(tup[:-1]) for tup in R[:-1]] for item in sublist]
    #print R
    #print flatten