import win32com.client
import transformations
import numpy
from external_component import add_carm_as_external_component
from json_lookup import json_lookup_components, json_lookup_point
import math


def add_std_ref(carm_pn, irm_instance_id, glob_x_axis=[1, 0, 0]):

    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    hybridBodies1 = carm_part.HybridBodies
    hybridBody1 = hybridBodies1.Item("Reference Geometry")
    hybridBodies2 = hybridBody1.HybridBodies
    #hybridBody2 = hybridBodies2.Item("Non-Instantiated Standard Parts")
    hybridBodyJD = hybridBodies1.Item("Joint Definitions")
    hybridBodiesJD = hybridBodyJD.HybridBodies
    axisSystems1 = carm_part.AxisSystems
    axis = axisSystems1.Item('Local Axis System')
    reference1 = carm_part.CreateReferenceFromObject(axis)
    HybridShapeFactory1 = carm_part.HybridShapeFactory
    #selection1.Add(hybridBody1)

    for geoset in xrange(1, hybridBodiesJD.Count+1):
        hb = hybridBodiesJD.Item(geoset)
        points = hb.HybridShapes
        if points.Count < 1:
            continue
        for fidv in xrange(1, points.Count+1):
            if 'FIDV' in points.Item(fidv).Name:
                reference2 = carm_part.CreateReferenceFromObject(points.Item(fidv))
                line_dir = HybridShapeFactory1.AddNewDirection(reference2)
                line_dir.RefAxisSystem = reference1
                glob_fivd = [line_dir.GetXVal(), line_dir.GetYVal(), line_dir.GetZVal()]
                print hybridBodiesJD.Item(geoset).Name + ': ' + str(glob_fivd)
                perp_v = transformations.vector_product(glob_x_axis, glob_fivd)
                ang_v = transformations.angle_between_vectors(glob_x_axis, glob_fivd, directed=True)
                rotation_matrix = transformations.rotation_matrix(ang_v, perp_v)
                print math.degrees(ang_v)
                print rotation_matrix
                rotation_matrix_flat = [item for sublist in [list(tup[:-1]) for tup in rotation_matrix[:-1]] for item in sublist]
                part_num = 'jd_std'

        for point in xrange(1, points.Count+1):
            if not 'FIDV' in points.Item(point).Name:
                coordinates = json_lookup_point(carm_pn, points.Item(point).Name)
                new_position = rotation_matrix_flat + [coordinates[0], coordinates[1], coordinates[2]]
                added_part = add_carm_as_external_component(part_num + ' ' + carm_pn, irm_instance_id, inserted=part_num)
                added_part[1].Move.Apply(new_position)

                # then we need to copy surfaces to the carm and delete the source and we are done

if __name__ == "__main__":

    #global_x_axis = [1, 0, 0]
    #b = [0, -1, 0]

    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    documents = catia.Documents

    add_std_ref('CA836Z1191-46_2017_01_20_19_11_16', 'GLS_STA1618-1732_OB_LH_CAI')

    #perp_v = transformations.vector_product(global_x_axis, b)
    #print perp_v
    #ang_v = transformations.angle_between_vectors(global_x_axis, b)
    #print ang_v
    #R = transformations.rotation_matrix(ang_v, perp_v)
    #print R
    #v0 = numpy.random.random(3)
    #print v0
    #v1 = transformations.vector_norm(v0)
    #print type(R)
    #my_list = [list(tup[:-1]) for tup in R[:-1]]
    #flatten = [item for sublist in [list(tup[:-1]) for tup in R[:-1]] for item in sublist]
    #print R
    #print flatten