import win32com.client
from json_lookup import json_lookup_origin


def make_axis(instance_id, carm_pn):

    catia = win32com.client.Dispatch('catia.application')

    # obtain instance origin

    pos = json_lookup_origin(instance_id)

    # obtain carm document

    documents = catia.Documents
    carm = None
    for i in range(1, documents.Count + 1):
        if carm_pn in documents.Item(i).Name:
            carm = documents.Item(i)
        else:
            continue
        if carm is None:
            return None
    part1 = carm.Part

    # make axis system

    axisSystems1 = part1.AxisSystems
    axisSystem1 = axisSystems1.Add()

    # modify axis system

    axisSystem1.OriginType = 1
    axisSystem1.PutOrigin(pos[9:12])
    axisSystem1.XAxisType = 1
    axisSystem1.PutXAxis(pos[:3])
    axisSystem1.YAxisType = 1
    axisSystem1.PutYAxis(pos[3:6])
    axisSystem1.ZAxisType = 1
    axisSystem1.PutZAxis(pos[6:9])
    part1.UpdateObject(axisSystem1)
    axisSystem1.IsCurrent = False
    part1.Update()

if __name__ == "__main__":

    make_axis('1691-4_STA1713_LT-18IN-OB-NFAR_LH', 'CA836Z1691-4')
