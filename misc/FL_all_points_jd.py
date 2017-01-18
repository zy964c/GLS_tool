import json

import win32com.client


if __name__ == "__main__":
    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    collector = catia.ActiveDocument.Product.Products
    gls_parts_dict = {}
    line_dict = {}
    for part in range(1, collector.Count + 1):
        line_dict.clear()
        print collector.Item(part).Name
        selection1.Add(collector.Item(part))
        selection1.Search(str('(Part Design.Geometrical Set.Name=gla_jd_points), sel'))
        selection1.Search(str('(Part Design.Line.Name=*), sel'))
        for selected in xrange(1, selection1.Count2 + 1):
            line = selection1.Item2(selected).Value
            line_dir = line.dir
            line_comp = [line_dir.GetXVal(), line_dir.GetYVal(), line_dir.GetZVal()]
#            line_dict[str(line.Name)] = line_comp
#            dict_paste = line_dict.copy()
            gls_parts_dict[str(collector.Item(part).Name[:-2])] = line_comp
            #print gls_parts_dict
    

    with open('Y:\\My Documents\\Python Scripts\\jd_vectors.json', 'w') as f:
        json.dump(gls_parts_dict, f, sort_keys=True, indent=4, separators=(',', ': '))
        
    #pprint(data)

# for checking------------------------------------------------------------------------------
    #point1 = gls_parts_dict['1J5009-221142-0##ALT3']['Point.93']
    #print point1
    #points = gls_parts_dict['1J5009-221142-0##ALT3']
    #for key in points:
        #print points[key]
# ------------------------------------------------------------------------------------------


def make_point(coord):

        catia = win32com.client.Dispatch('catia.application')
        part_name = 'points'
        collector = catia.ActiveDocument.Product.Products
        new_part = collector.AddNewComponent("Part", part_name)
        documents1 = catia.Documents
        partDocument1 = documents1.Item(part_name + ".CATPart")
        part1 = partDocument1.Part
        hybridBodies1 = part1.HybridBodies
        hybridBody1 = hybridBodies1.Add()
        hybridBody1.Name = "JD_points"
        part1.InWorkObject = hybridBody1
        HybridShapeFactory1 = part1.HybridShapeFactory
        point_added = HybridShapeFactory1.AddNewPointCoord(coord[0], coord[1], coord[2])
        hybridBody1.AppendHybridShape(point_added)
        part1.Update()
        
#make_point(point1)


