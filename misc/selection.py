import win32com.client


catia = win32com.client.Dispatch('catia.application')


def selection(part_name1, instance_id1):
    part_name2 = part_name1.replace('-', '*').replace('_', '*').replace('+', '*')
    instance_id2 = instance_id1.replace('-', '*').replace('_', '*').replace('+', '*')
    selection1.Search(str('Name = *' + str(instance_id2) + '*,all'))
    selection1.Search(str('Name = *' + str(part_name2) + '*, sel'))
    selection1.Search(str('Name = *' + 'JD_points' + '*, sel'))
    selected = selection1.Count2
    last_one = selection1.Item2(selected).Value
    selection1.Clear()
    selection1.Add(last_one)
    selection1.Search(str('Name =' + '*Point.*' + ', sel'))
    selection1.Copy()
    selection1.Clear()


def selection_delete(part_name1, instance_id1):
    part_name2 = part_name1.replace('-', '*').replace('_', '*').replace('+', '*')
    instance_id2 = instance_id1.replace('-', '*').replace('_', '*').replace('+', '*')
    selection1.Search(str('Name = *' + str(instance_id2) + '*,all'))
    selection1.Search(str('Name = *' + str(part_name2) + '*, sel'))
    selection1.Search(str('Name = *' + 'JD_points' + '*, sel'))
    selection1.Delete()
