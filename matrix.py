from json_lookup import json_lookup_axis, json_lookup_camera
from sympy import Matrix, linsolve, symbols
import math
from views import Annotation


def axis_coord(point, view):

    ac = json_lookup_axis(view)
    x, y, z = symbols("x, y, z")
    A = Matrix([[ac[3], ac[6], ac[9]], [ac[4], ac[7], ac[10]], [ac[5], ac[8], ac[11]]])
    b = Matrix([(point[0] - ac[0]), (point[1] - ac[1]), (point[2] - ac[2])])
    result = linsolve((A, b), [x, y, z])
    output = [j for i in result for j in i]
    return output


def mod(coord, offset):
    
    mod_coord = [coord[0] + offset[0], coord[1] + offset[1], coord[2]]
    return mod_coord


def rotate_vector(angle, vector):

    rad = math.radians(angle)
    m1 = Matrix([[math.cos(rad), -1 * math.sin(rad), 0], [math.sin(rad), math.cos(rad), 0], [0, 0, 1]])
    m2 = Matrix([vector[0], vector[1], vector[2]])
    result = m1*m2
    output = [i for i in result]
    #for i in result:
    #    output.append(i)
    return output

if __name__ == "__main__":


    #print mod(axis_coord([0.0, 0.0, (400.0*25.4)],
          #'Inboard Facing Out - Lower Support - Axis System LH'), [(3 * 25.4), (3 * 25.4)])
    sight_dir = slice(3, 6)
    sight_direction = json_lookup_camera("REF GEOMETRY")[sight_dir]
    #print rotate_vector(5.0, sight_direction)

    for j in Annotation.views:
        print axis_coord([0, 0, 0], j)
