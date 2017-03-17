import win32com.client
# from functions import inch_to_mm
# from pprint import pprint
from json_lookup import json_lookup_camera
from add_component import Ref
from matrix import rotate_vector


def map_camera_names(pn):
        """Maps camera names to numbers in order"""
        
        catia = win32com.client.Dispatch('catia.application')
        documents1 = catia.Documents
        cam_dict = {}
        partDocument1 = documents1.Item(pn + '.CATPart')
        cameras = partDocument1.Cameras
        for i in xrange(1, cameras.Count+1):
            camera = cameras.Item(i)
            cam_dict[str(camera.name)] = i
        return cam_dict


def update_camera(cam_name, omf):

        cam_coord = slice(0, 3)
        sight_dir = slice(3, 6)
        up_dir = slice(6, 9)
        coord_origin = json_lookup_camera(cam_name, omf.side_to_find)[cam_coord]
        sight_direction = json_lookup_camera(cam_name, omf.side_to_find)[sight_dir]
        up_direction = json_lookup_camera(cam_name, omf.side_to_find)[up_dir]
        coord_origin_global = omf.get_position_camera(coord_origin)
        return coord_origin_global, sight_direction, up_direction


def cameras(pn, omf):
        
        catia = win32com.client.Dispatch('catia.application')
        documents1 = catia.Documents
        partDocument1 = documents1.Item(pn + '.CATPart')
        cameras_collection = partDocument1.Cameras
        if omf.section == '41' and omf.side_to_find != 'CTR':
                angle = 5.0
        elif omf.section == '47' and omf.side_to_find != 'CTR':
                angle = -3.125
        else:
                angle = 0  
        for i in xrange(1, cameras_collection.Count+1):
            camera = cameras_collection.Item(i)
            viewpoint = camera.Viewpoint3D
            if camera.Name == 'Wire Routing Typical':
                continue
            viewpoint_data = update_camera(camera.Name, omf)
            cam_coord, sight_dir, up_dir = viewpoint_data
            sight_dir_rotated = rotate_vector(angle, sight_dir)
            up_dir_rotated = rotate_vector(angle, up_dir)
            if omf.side_to_find == 'RH':
                    sight_dir_rotated[1] = sight_dir_rotated[1] * -1
                    up_dir_rotated[1] = up_dir_rotated[1] * -1
            sight_dir_rotated = [float(buffer(str(sight_dir_rotated[0]))),
                                 float(buffer(str(sight_dir_rotated[1]))),
                                 float(buffer(str(sight_dir_rotated[2])))]
            up_dir_rotated = [float(buffer(str(up_dir_rotated[0]))),
                              float(buffer(str(up_dir_rotated[1]))),
                              float(buffer(str(up_dir_rotated[2])))]
            #print cam_coord
            viewpoint.PutOrigin(cam_coord)
            viewpoint.PutSightDirection(sight_dir_rotated)
            viewpoint.PutUpDirection(up_dir_rotated)


if __name__ == "__main__":

        #pprint(map_camera_names('seed_fairing_lh'))
        omf1 = Ref('787_9_JAL_ZB424', '1239', 'LH', 240, 0, [])
        print update_camera("JD01 Ceiling Light Typical", omf1)
        cameras('CA836Z1661-2', 'LH', omf1)

