import win32com.client
from json_lookup import json_lookup_point, json_lookup_flagnote
from views import AnnotationFactory
import re
from matrix import axis_coord, mod
from pprint import pprint


def add_jd_annotation(carm_pn, sta_value, jd_number1, side, irm_type):
    """Adds JOINT DEFINITION XX annotation"""
    
    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    documents = catia.Documents
    #side = instance_id[-6: -4]
    sec = 'constant'
    if '+' not in sta_value:
        if float(sta_value) < 465:
            sec = '41'
        elif float(sta_value) > 1617:
            sec = '47'
        else:
            sec = 'constant'
    for jd_number in xrange(1, jd_number1+1):
        if jd_number > 9:
            jd_number_formatted = str(jd_number)
        else:
            jd_number_formatted = '0' + str(jd_number)
        annot_text = 'JOINT DEFINITION ' + jd_number_formatted
        geoset_name = 'Joint Definition ' + jd_number_formatted
        #print geoset_name
        carm_doc = documents.Item(carm_pn + ".CATPart")
        carm_part = carm_doc.Part
        KBE = carm_part.GetCustomerFactory("BOEAnntFactory")
        annotations = AnnotationFactory(irm_type)
        view_name = annotations.get_view_number('JD ' + jd_number_formatted,
                                                side, sec)
        if view_name is None:
            continue
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        TPSViews = ann_set1.TPSViews
        view_to_activate = TPSViews.Item(view_name + 1)
        ann_set1.ActiveView = view_to_activate
        userSurfaces1 = carm_part.UserSurfaces
        geosets = carm_part.HybridBodies
        geoset1 = geosets.Item('Joint Definitions')
        geosets1 = geoset1.HybridBodies
        
        try:
            geoset2 = geosets1.Item(geoset_name)
        except:
            continue
        
        points = geoset2.HybridShapes
        fidv = 0
        for p in xrange(1, points.Count + 1):
            if 'FIDV' in points.Item(p).Name:
                fidv += 1
        wb = str(catia.GetWorkbenchId())
        if wb != 'PrtCfg':
            selection1.Clear()
            selection1.Add(carm_part)
            catia.StartWorkbench("PrtCfg")
            selection1.Clear()
        try:
            reference1 = carm_part.CreateReferenceFromObject(points.Item(1))
        except:
            continue
        userSurface1 = userSurfaces1.Generate(reference1)
        coordinates = json_lookup_point(carm_pn, points.Item(1).Name)
        coord_new = mod(axis_coord(coordinates,
                                   annotations.get_view_name('JD ' + jd_number_formatted, side, sec), irm_type),
                                                             [(2.5 * 25.4),
                                                              (2.5 * 25.4)])

        for point in xrange(2, points.Count + 1 - fidv):
            reference2 = carm_part.CreateReferenceFromObject(points.Item(point))
            userSurface1.AddReference(reference2)

        annotationFactory1 = ann_set1.AnnotationFactory
        annotation1 = annotationFactory1.CreateEvoluateText(userSurface1, coord_new[0], coord_new[1], coord_new[2],
                                                            True)
        KBE.InitializeAnntFactory(carm_doc)
        capture_dict = set_capture_dict(carm_pn)
        #annotation1 = KBE.CreateTextWithLeader(points.Item(1), view_to_activate, annot_text, 137, 68, 20)
        Captures1 = ann_set1.Captures
        capture1 = Captures1.Item(capture_dict[annotations.get_capture('JD ' + jd_number_formatted)])
        KBE.AssociateAnntCapture(annotation1, capture1)
        #capture1.Current = True
        ann_text = annotation1.Text()
        ann1text_2d = ann_text.Get2dAnnot()
        ann1text_2d.Text = annot_text
        ann1text_2d.SetFontSize(0, 0, 8)
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(anns.Count)
        selection1.Clear()
        selection1.Add(sta_annotation)
        selection1.visProperties.SetShow(1)
        carm_part.Update()


def set_capture_dict(carm_pn):
    
    capture_dict = {}
    catia = win32com.client.Dispatch('catia.application')
    documents = catia.Documents
    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    ann_sets = carm_part.AnnotationSets
    ann_set1 = ann_sets.Item(1)
    Captures1 = ann_set1.Captures
    for i in xrange(1, Captures1.Count+1):
        capture_dict[Captures1.Item(i).Name] = i
    #pprint(capture_dict)
    return capture_dict


def activate_view(carm_pn, view_name):

    catia = win32com.client.Dispatch('catia.application')
    documents = catia.Documents
    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    ann_sets = carm_part.AnnotationSets
    ann_set1 = ann_sets.Item(1)
    TPSViews = ann_set1.TPSViews
    view_to_activate = TPSViews.Item(view_name)
    ann_set1.ActiveView = view_to_activate


def activate_top_prod():

    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    selection1.Clear()
    product1 = productDocument1.Product
    selection1.Add(product1)
    catia.StartCommand('FrmActivate')
    selection1.Clear()


def access_captures(instance_id, number):

    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    Product = productDocument1.Product
    collection = Product.Products
    irm_ref = collection.Item(instance_id).ReferenceProduct
    parts = irm_ref.Products
    carm_ref = parts.Item(parts.Count).ReferenceProduct
    carm_parent = carm_ref.Parent
    carm_part = carm_parent.Part
    ann_sets = carm_part.AnnotationSets
    ann_set1 = ann_sets.Item(1)
    captures1 = ann_set1.Captures
    capture1 = captures1.Item(number)
    capture1.ManageHideShowBody = True
    capture1.DisplayCapture()


def add_annotation(carm_pn, sta_value, side, irm_type):
    """Adds annotation"""
    
    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    documents = catia.Documents
    #side = instance_id[-6: -4]
    bin_number = 0
    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    hybrid_bodies = carm_part.HybridBodies
    lib = hybrid_bodies.Item('Construction Geometry')
    hybrid_bodies1 = lib.HybridBodies
    flagnote = hybrid_bodies1.Item('flagnote')
    points = flagnote.HybridShapes
    sec = 'constant'
    if '+' not in sta_value[0][0]:
        if int(sta_value[bin_number][0]) < 465:
            sec = '41'
        elif int(sta_value[bin_number][0]) > 1617:
            sec = '47'
        else:
            sec = 'constant'

    re_ceiling_light_marker = re.compile('\d{3}Z\d{4}-4(2|3)\d')
    re_sidewall_light_marker = re.compile('\d{3}Z\d{4}-45\d')
    re_nofar_light_marker = re.compile('\d{3}Z\d{4}-44(5|6)')
    re_term_plug_marker_ceiling = re.compile('804Z3000-211.C')
    re_term_plug_marker_sidewall = re.compile('804Z3000-211$')
    re_term_plug_marker_nofar = re.compile('804Z3000-211.N')
    re_power_supply_marker = re.compile('804Z3000-210')
    re_dict = {re_ceiling_light_marker: 'light_marker',
               re_sidewall_light_marker: 'sidewall_light_marker',
               re_nofar_light_marker: 'nofar_light_marker',
               re_term_plug_marker_ceiling: 'term_plug_marker_ceiling',
               re_term_plug_marker_sidewall: 'term_plug_marker_sidewall',
               re_term_plug_marker_nofar: 'term_plug_marker_nofar',
               re_power_supply_marker: 'power_supply_marker'}
    re_list = re_dict.keys()
    
    display_ones = ['1X5005-412(XXX) Alignment Feature', '1X5005-412(XXX) Mounting Feature',
                    'Fixing Tower', '1X5005-420(X)00 Mounting Bracket', 'Locking Tabs', '4']
    not_s47 = ['1X5005-412(XXX) Alignment Feature', '1X5005-412(XXX) Mounting Feature']
    added_annots = []
    added_fl = ['FL28', 'FL29', 'FL38']
    for point in xrange(1, points.Count+1):
        point_name = points.Item(point).Name

        if 'sta' in point_name:
            annot_text_check = 'sta'
            annot_text = 'STA ' + sta_value[bin_number][0] + '\n   ' + side[0] + 'BL 61\n   WL 285\n    REF'
            bin_number += 1
        elif '804Z3000' in point_name:
            for r in re_list:
                m = r.match(point_name)
                if m is not None:
                    annot_text_check = re_dict[r]
                    annot_text = point_name
                    break
                else:
                    continue
        elif 'FL' in point_name:
            annot_text_check = point_name
            annot_text = point_name[2:]
            added_fl.append(annot_text_check)
        else:
            annot_text_check = point_name
            annot_text = point_name
        while '.' in annot_text:
            annot_text = annot_text[:-2]
        while '.' in annot_text_check:
            annot_text_check = annot_text_check[:-2]
        if annot_text in not_s47 and sec == '47':
            continue
        if annot_text in display_ones:
            if annot_text in added_annots:
                continue
            else:
                added_annots.append(annot_text)
        
        annotations = AnnotationFactory(irm_type)
        coordinates = json_lookup_flagnote(carm_part.name, point_name)
        view_name = annotations.get_view_number(annot_text_check, side, sec)
        if view_name is None:
            continue
        if isinstance(view_name, list) and coordinates[1] < 0:
            view_name = view_name[0]
        elif isinstance(view_name, list) and coordinates[1] > 0:
            view_name = view_name[1]
        print view_name
        carm_doc = documents.Item(carm_pn + ".CATPart")
        carm_part = carm_doc.Part
        KBE = carm_part.GetCustomerFactory("BOEAnntFactory")
        
        ann_sets = carm_part.AnnotationSets
        ann_set1 = ann_sets.Item(1)
        TPSViews = ann_set1.TPSViews
        view_to_activate = TPSViews.Item(view_name + 1)
        ann_set1.ActiveView = view_to_activate
        
        wb = str(catia.GetWorkbenchId())
        if wb != 'PrtCfg':
            selection1.Clear()
            selection1.Add(carm_part)
            catia.StartWorkbench("PrtCfg")
            selection1.Clear()

        userSurfaces1 = carm_part.UserSurfaces
        reference1 = carm_part.CreateReferenceFromObject(points.Item(point))
        userSurface1 = userSurfaces1.Generate(reference1)
        coordinates = json_lookup_flagnote(carm_part.name, point_name)
        #print point_name + ':'
        #print coordinates
        if 'marker' in annot_text_check:
            offset = [(2.5 * 25.4), (3.5 * 25.4)]
        elif 'sta' in annot_text_check:
            offset = [0, (25.0 * 25.4)]
        else:
            offset = [(2.5 * 25.4), (2.5 * 25.4)]
        coord_new = mod(axis_coord(coordinates, annotations.get_view_name(annot_text_check, side, sec), irm_type),
                        offset)
        annotationFactory1 = ann_set1.AnnotationFactory
        KBE.InitializeAnntFactory(carm_doc)
        
        if 'FL' in annot_text_check:
            annotation1 = annotationFactory1.CreateFlagNote(userSurface1)
            annotation1.SetXY(coord_new[0], coord_new[1])
            annotation1.AddLeader()
            fl = annotation1.FlagNote()
            KBE.SetFlagNoteText(annotation1, annot_text, 1)
            parameters_collection = carm_part.Parameters
            fl_param_text = parameters_collection.Item('Annotation Notes:\\FL' + annot_text).Value
            fl.AddURL(fl_param_text)
                                  
        else:
            annotation1 = annotationFactory1.CreateEvoluateText(userSurface1,
                                                                coord_new[0],
                                                                coord_new[1],
                                                                coord_new[2],
                                                                True)
        #annotation1 = KBE.CreateTextWithLeader(points.Item(point), view_to_activate, annot_text, 20, -10, -1*coordinates[0])
        #annotation1 = KBE.CreateFlagNoteAnnotation(annotationFactory1, annot_text, 2)
        Captures1 = ann_set1.Captures
        capture_dict = set_capture_dict(carm_pn)
        capture_name = annotations.get_capture(annot_text_check)
        if type(capture_name) is list:
            for capture in capture_name:
                capture1 = Captures1.Item(capture_dict[capture])
                KBE.AssociateAnntCapture(annotation1, capture1)
        else:
            capture1 = Captures1.Item(capture_dict[capture_name])
            KBE.AssociateAnntCapture(annotation1, capture1)
        #capture1.Current = True
        if 'FL' not in annot_text_check:
            ann_text = annotation1.Text()
            ann1text_2d = ann_text.Get2dAnnot() 
            ann1text_2d.Text = annot_text
            ann1text_2d.SetFontSize(0, 0, 8)
        if 'FL' in point_name:
            pass
#            ann1text_2d.AnchorPosition = 12
#            ann1text_2d.FrameType = 7
        elif 'sta' in point_name:
            text_leaders = ann1text_2d.Leaders
            text_leader1 = text_leaders.Item(1)
            text_leader1.HeadSymbol = 1
            ann1text_2d.SetFontSize(0, 0, 18)
            ann1text_2d.AnchorPosition = 4
            ann1text_2d.FrameType = 3
        anns = ann_set1.Annotations
        sta_annotation = anns.Item(anns.Count)
        selection1.Clear()
        selection1.Add(sta_annotation)
        selection1.visProperties.SetShow(1)
        selection1.Clear()
        carm_part.Update()
    parameters1 = carm_part.Parameters
    for param in xrange(1, 41):
        fl_added = False
        fl_param = parameters1.Item('Annotation Notes:\\FL' + str(param))
        for s in added_fl:
            if 'FL' + str(param) in s:
                fl_added = True
                break
        if not fl_added:
            selection1.Add(fl_param)
        else:
            continue
    selection1.Delete()
    selection1.Clear()
    carm_part.Update()
            

if __name__ == "__main__":

    input_config = [['1623', 42], ['1665', 24]]
    add_jd_annotation('CA836Z1191-46_2017_01_17_19_43_34', input_config[0][0], 30, 'GLS_STA0561-0657_OB_LH_CAI')
    #add_annotation('CA836Z1191-41', input_config, 'GLS_STA1618-1732_OB_LH_CAI')
