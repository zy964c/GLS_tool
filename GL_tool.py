import win32com.client
import glob
import sys
import os
import Tkinter as tk
import tkMessageBox
from external_component import add_carm_as_external_component
from json_lookup import json_lookup, json_lookup_fl, json_lookup_fl_keys, json_lookup_components, json_lookup_origin
from jd import JD
from add_component import Ref
from change_inst_id import change_inst_id
from axis import make_axis
from jd_annot import add_jd_annotation, activate_top_prod, add_annotation
from camera import cameras
from summarize import std_parts
from json_parsing import parse_ss
from functions import sta_value
from subprocess import check_call

ref_objects = []


def return_part(prdct_id, part_id):
    """
    find detail in tree
    """
    to_p = collection.Item(prdct_id)
    Product2 = to_p.ReferenceProduct
    Product2Products = Product2.Products
    product_forpaste = Product2Products.Item(part_id)
    Part3 = product_forpaste.ReferenceProduct
    PartDocument3 = Part3.Parent
    geom_elem4 = PartDocument3.Part
    return geom_elem4


def part_design(geom_elem):
    """
    switch to part design
    """
    selection1.Clear()
    selection1.Add(geom_elem)
    catia.StartWorkbench("PrtCfg")
    selection1.Clear()


def create_point(part_name1, instance_id1, carm_name2, carm_pn, part_pn, plug_value):
    """
    create point and copy it
    :param plug_value:
    """

    points_js = json_lookup(part_pn)
    if points_js is not None:
        for i in xrange(1, documents.Count + 1):
            if carm_pn in documents.Item(i).Name:
                partDocument1 = documents.Item(i)
        current_part = partDocument1.Part
        HybridShapeFactory1 = current_part.HybridShapeFactory
        hybridBody2 = jd_set(carm_name2, instance_id1, part_name1, part_pn, plug_value)
        make_axis(part_name1, carm_pn)
        axisSystems1 = current_part.AxisSystems
        axis = axisSystems1.Item(axisSystems1.Count)
        i = 0
        for coord in points_js:
            reference1 = current_part.CreateReferenceFromObject(axis)
            point_added = HybridShapeFactory1.AddNewPointCoord(coord[0],
                                                               coord[1],
                                                               coord[2])
            point_added.RefAxisSystem = reference1
            if hybridBody2 is not None:
                points = hybridBody2.HybridShapes
                if points.Count > 0 and 'FIDV' in points.Item(points.Count).Name:
                    selection1.Add(points.Item(points.Count))
                    selection1.Delete()
                    selection1.Clear()
                hybridBody2.AppendHybridShape(point_added)
                point_added.Name = hybridBody2.Name + '.' + str(points.Count)
                parameters1 = current_part.Parameters
                param_hole_qty = parameters1.Item('Joint Definitions\\' +
                                                  hybridBody2.Name +
                                                  '\\Hole Quantity')
                param_hole_qty.Value = points.Count
                i += 1
        selection1.Clear()
        selection1.Add(axis)
        selection1.visProperties.SetShow(1)
        selection1.Clear()
        current_part.Update()


def create_point_fl(part_name1, carm_pn, part_pn, minor):
    """
    create point and copy it
    """

    points_js = json_lookup_fl(part_pn)
    keys_js = json_lookup_fl_keys(part_pn)
    if points_js is not None:
        for i in xrange(1, documents.Count + 1):
            if carm_pn in documents.Item(i).Name:
                partDocument1 = documents.Item(i)
        current_part = partDocument1.Part
        HybridShapeFactory1 = current_part.HybridShapeFactory
        hybrid_bodies1 = current_part.HybridBodies
        hybrid_body1 = hybrid_bodies1.Item('Construction Geometry')
        hybrid_bodies2 = hybrid_body1.HybridBodies
        hybridBody2 = hybrid_bodies2.Item('flagnote')
        make_axis(part_name1, carm_pn)
        axisSystems1 = current_part.AxisSystems
        axis = axisSystems1.Item(axisSystems1.Count)
        for coord, key in zip(xrange(len(points_js)), xrange(len(keys_js))):
            reference1 = current_part.CreateReferenceFromObject(axis)
            point_added = HybridShapeFactory1.AddNewPointCoord((points_js[coord])[0],
                                                               (points_js[coord])[1],
                                                               (points_js[coord])[2])
            point_added.RefAxisSystem = reference1
            if hybridBody2 is not None:
                hybridBody2.AppendHybridShape(point_added)
                point_added.Name = keys_js[key]
                if point_added.Name == 'FL10' or point_added.Name == '804Z3000-211':
                    pos = json_lookup_origin(part_name1)
                    coords = slice(9, 12)
                    wl = 275.0 * 25.4
                    s47_sta = (1617.0 + minor) * 25.4
                    if pos[coords][2] > wl:
                        if pos[coords][0] > s47_sta:
                            if point_added.Name == 'FL10':
                                point_added.Name = 'FL16'
                            else:
                                point_added.Name = '804Z3000-211.N'
                        else:
                            if point_added.Name == 'FL10':
                                point_added.Name = 'FL12'
                            else:
                                point_added.Name = '804Z3000-211.C'
                
        selection1.Clear()
        selection1.Add(axis)
        selection1.visProperties.SetShow(1)
        selection1.Clear()
        current_part.Update()
        

def create_point_sta(carm_pn, omf1):

    for i in xrange(1, documents.Count + 1):
            if carm_pn in documents.Item(i).Name:
                partDocument1 = documents.Item(i)
    current_part = partDocument1.Part
    HybridShapeFactory1 = current_part.HybridShapeFactory
    hybrid_bodies1 = current_part.HybridBodies
    hybrid_body1 = hybrid_bodies1.Item('Construction Geometry')
    hybrid_bodies2 = hybrid_body1.HybridBodies
    hybridBody2 = hybrid_bodies2.Item('flagnote')
    sta = omf1.get_position_sta()
    mounting_feature = omf1.get_position_mounting()
    alignment_feature = omf1.get_position_alignment()
    points = []
    names = ['sta', '1X5005-412(XXX) Mounting Feature',
             '1X5005-412(XXX) Alignment Feature']
    points.extend((sta, mounting_feature, alignment_feature))
    for i in xrange(len(points)):
        if points[i] is not None:
            point_added = HybridShapeFactory1.AddNewPointCoord(points[i][0],
                                                               points[i][1],
                                                               points[i][2])
            if hybridBody2 is not None:
                hybridBody2.AppendHybridShape(point_added)
                point_added.Name = names[i]
        current_part.Update()


def create_jd_vectors2(part_name1, instance_id1, carm_name2, carm_pn, part_pn, plug_value):

    points_js = json_lookup_components(part_pn)
    if points_js is not None:
        for i in xrange(1, documents.Count + 1):
            if carm_pn in documents.Item(i).Name:
                partDocument1 = documents.Item(i)
        current_part = partDocument1.Part
        HybridShapeFactory1 = current_part.HybridShapeFactory
        hybridBody2 = jd_set(carm_name2, instance_id1, part_name1, part_pn, plug_value)
        if hybridBody2 is not None:
            hs = hybridBody2.HybridShapes
            try:
                point = hs.Item(hs.Count)
            except:
                print 'No points in JD'
                return None
        else:
            return None
        make_axis(part_name1, carm_pn)
        axisSystems1 = current_part.AxisSystems
        axis = axisSystems1.Item(axisSystems1.Count)
        direction = HybridShapeFactory1.AddNewDirectionByCoord(points_js[0],
                                                               points_js[1],
                                                               points_js[2])
        reference1 = current_part.CreateReferenceFromObject(axis)
        reference2 = current_part.CreateReferenceFromObject(point)                                               
        direction.RefAxisSystem = reference1
        vector = HybridShapeFactory1.AddNewLinePtDir(reference2, direction,
                                                     0.000000,
                                                     25.400000,
                                                     False)
        hybridBody2.AppendHybridShape(vector)
        vector.Name = 'FIDV_' + hybridBody2.Name[-2:]
        selection1.Clear()
        selection1.Add(axis)
        selection1.visProperties.SetShow(1)
        selection1.Clear()
        current_part.Update()
    else:
        return None
    
    
def jd_set(carm_name1, instance_id1, part_name1, part_pn, plug_value):
    """
    returns jd geoset
    :param plug_value:
    """
    part1 = return_part(instance_id1, carm_name1)
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item('Joint Definitions')
    hybridBodies2 = hybridBody1.HybridBodies
    joint = JD(part_pn, instance_id1, part_name1, plug_value)
    jd_name = joint.get_name()
    print jd_name
    if jd_name is not None:
        hybridBody2 = hybridBodies2.Item('Joint Definition ' + jd_name)
        return hybridBody2
    else:
        return None


def reference(size, instance_id1, part_name1, sta1, side1, customer, plug_value, bin_order, irm_len, all_irm_parts, pn):

    global ref_objects
    ref1 = Ref(customer, sta1, side1, plug_value, bin_order, irm_len, all_irm_parts, name=instance_id1)
    ref1.build()
    product = collection.Item(instance_id1)
    collection1 = product.Products
    for i in xrange(1, collection1.Count + 1):
            prod = collection1.Item(i)
            print ref1.component_name
            if ref1.component_name in collection1.Item(i).Name:
                ref_part = collection1.Item(i)
                break
            try:
                collection2 = prod.Products
            except:
                continue
            if collection2.Count > 0:
                for j in xrange(1, collection2.Count + 1):
                    if ref1.component_name in collection2.Item(j).Name:
                        ref_part = collection2.Item(j)
                        break
            else:
                continue
    set_visibility(pn, False)
    ref_part.ApplyWorkMode(2)
    selection1.Add(ref_part)
    selection1.Search(str('(NAME = *' + str(size) + '*IN*REF* + NAME = BACS31H1A*WMA*REF*), sel'))
    try:
            selection1.Copy()
    except:
            pass
    else:
            selection1.Clear()
            geom_elem5 = return_part(instance_id1, part_name1)
            hybridBodies1 = geom_elem5.HybridBodies
            hybridBody1 = hybridBodies1.Item("Reference Geometry")
            selection1.Add(hybridBody1)
            selection1.PasteSpecial('CATPrtResultWithOutLink')
            geom_elem5.Update()
            selection1.visProperties.SetRealColor(210, 180, 140, 0)
            selection1.visProperties.SetRealOpacity(255, 0)
            selection1.Clear()
    selection1.Add(ref_part)

    #try:
    #    if int(sta1) < 465 and bin_order == irm_len:
    #        selection1.Search(str('(Name=*Ring*Post*REF* + Name=*Lower*Bushing*REF*), sel'))
    #except:
    #    pass
    #if selection1.Count2 == 0:
    #    if '1X5005-210000' in [x[:x.find('##')] for x in all_irm_parts] and bin_order == irm_len:
    #        selection1.Search(str('(Name=*Ring*Post*REF* + Name=*Lower*Bushing*REF*), sel'))
    #    elif 'L' in side1:
    #        selection1.Search(str('(Name=*Ring*Post*REF* + Name=*Lower*Bushing*REF*FWD*), sel'))
    #    elif 'R' in side1:
    #        selection1.Search(str('(Name=*Ring*Post*REF* + Name=*Lower*Bushing*REF*AFT*), sel'))

    selection1.Search(str('(Name=*Ring*Post*REF* + Name=*Lower*Bushing*REF*), sel'))

    try:
        selection1.Copy()
    except:
        print 'SOLIDS NOT FOUND'
        ref_objects.append(ref1)
        return None
    else:
        selection1.Clear()
        carm_part = return_part(instance_id1, part_name1)
        #KBE = carm_part.GetCustomerFactory("KBEFactory")
        selection1.Add(carm_part)
        selection1.PasteSpecial('CATPrtResultWithOutLink')
        carm_part.Update()
        selection1.Clear()
    set_visibility(pn, True)
    ref_objects.append(ref1)
#    ref1.remove_component()


def deleter():

    """
    :return: no return, deletes all library data
    """

    last_prod = collection.Count
    first_delete = last_prod - 4
    for prod in xrange(first_delete, last_prod + 1):
        product_inwork = collection.Item(prod)
        selection1.Add(product_inwork)
    selection1.Delete()
    selection1.Clear()
    Product.Update()


def set_visibility(carm_pn, action):

    carm_doc = documents.Item(carm_pn + ".CATPart")
    ann_sets = carm_doc.Part.AnnotationSets
    ann_set1 = ann_sets.Item(1)
    Captures1 = ann_set1.Captures
    for capture in xrange(1, Captures1.Count+1):
        current_capture = Captures1.Item(capture)
        current_capture.ManageHideShowBody = action


def capture_del(carm_pn, instance_id):

    capture_list = ['Engineering Definition Release',
                    'Reference Geometry',
                    'All Annotations',
                    'Wire Routing Typical']
    bin_captures = ['JD17 Single Wire Bushings Typical']
    if 'OMF' not in instance_id:
        for item in bin_captures:
            capture_list.append(item)

    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    ann_sets = carm_part.AnnotationSets
    ann_set1 = ann_sets.Item(1)
    Captures1 = ann_set1.Captures
    found = 0
    selection1.Clear()
    for capture in xrange(1, Captures1.Count+1):
        current_capture = Captures1.Item(capture)
        print current_capture.Name
        annots = current_capture.Annotations
        if annots.Count > 0 or current_capture.Name in capture_list:
            continue
        else:
            selection1.Add(current_capture)
            found += 1
    if found > 0:
        selection1.Delete()
        selection1.Clear()
        Product.Update()


def jd_del(carm_pn):

    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    hybrid_bodies1 = carm_part.HybridBodies
    hybrid_body1 = hybrid_bodies1.Item('Joint Definitions')
    hybrid_bodies2 = hybrid_body1.HybridBodies
    found = 0
    selection1.Clear()
    for geoset in xrange(1, hybrid_bodies2.Count+1):
        hb = hybrid_bodies2.Item(geoset)
        print hb.Name
        points = hb.HybridShapes
        if points.Count > 0:
            continue
        else:
            selection1.Add(hb)
            found += 1
    if found > 0:
        selection1.Delete()
        selection1.Clear()
        Product.Update()


def rename_part_body(carm_pn):
        # type: (object) -> object

        """renames part body and activates it"""

        carm_doc = documents.Item(carm_pn + ".CATPart")
        carm_part = carm_doc.Part
        bodies = carm_part.Bodies
        part_body = bodies.Item(1)
        part_body.name = carm_pn
        #carm_part.InWorkObject = part_body
        hyb_bodies = carm_part.HybridBodies
        st_notes = hyb_bodies.Item('Standard Notes:')
        carm_part.InWorkObject = st_notes


def scan_parts(collection_parts, instance_id, carm_name, pn, ringposts, plug_value):
    
    for n in xrange(1, collection_parts.Count+1):
                next_part = collection_parts.Item(n)
                print collection_parts.Item(n).Name

                collection_parts2 = next_part.Products
                if collection_parts2.Count > 0:
                    scan_parts(collection_parts2, instance_id, carm_name, pn, ringposts, plug_value)
                else:
                    part_name = collection_parts.Item(n).Name
                    print part_name
                    part_pn = collection_parts.Item(n).PartNumber
                    if part_pn in ringposts:
                        for i in ringposts[part_pn]:
                            print i
                            create_point(part_name, instance_id, carm_name, pn, i, plug_value)
                            create_jd_vectors2(part_name, instance_id, carm_name, pn, i, plug_value)
                    create_point(part_name, instance_id, carm_name, pn, part_pn, plug_value)
                    create_jd_vectors2(part_name, instance_id, carm_name, pn, part_pn, plug_value)
                    create_point_fl(part_name, pn, part_pn, plug_value)


class Application(tk.Frame):
        def __init__(self, root):
            tk.Frame.__init__(self, root)
            root.geometry("300x500")
            root.title("GLS")

            self.parent_frame = tk.Frame(bd=3, relief='ridge')
            self.parent_frame_top = tk.Frame()
            self.parent_frame_minor = tk.Frame(self.parent_frame_top, bd=3, relief='ridge')
            self.parent_frame_type = tk.Frame(self.parent_frame_top, bd=3, relief='ridge')
            self.f0 = tk.Frame(self.parent_frame_minor)
            self.f01 = tk.Frame(self.parent_frame_minor)
            self.f10 = tk.Frame(self.parent_frame_type)
            self.f11 = tk.Frame(self.parent_frame_type)
            self.f1 = tk.Frame()
            self.f2 = tk.Frame(self.parent_frame)
            self.f3 = tk.Frame(self.parent_frame)
            self.f4 = tk.Frame(self.parent_frame)
            self.f5 = tk.Frame(self.parent_frame)
            self.f6 = tk.Frame(self.parent_frame)
            self.f7 = tk.Frame(self.parent_frame)
            self.f8 = tk.Frame(self.parent_frame)
            self.f9 = tk.Frame()

            self.f1.pack(fill='both', expand=1, side='top')
            self.parent_frame_top.pack(fill='both', expand=1, side='top', padx=10, pady=5)
            self.parent_frame_minor.pack(fill='both', expand=1, side='left', padx=10, pady=5)
            self.f0.pack(fill='both', expand=1, side='top')
            self.f01.pack(fill='both', expand=1, side='top')
            self.parent_frame_type.pack(fill='both', expand=1, side='left', padx=10, pady=5)
            self.f10.pack(fill='both', expand=1, side='top')
            self.f11.pack(fill='both', expand=1, side='top')
            self.parent_frame.pack(fill='both', expand=1, side='top', padx=10, pady=5)
            self.f2.pack(side='top')
            self.f3.pack(side='top')
            self.f4.pack(side='top')
            self.f5.pack(side='top')
            self.f6.pack(side='top')
            self.f7.pack(side='top')
            self.f8.pack(side='top')
            self.f9.pack(side='top', padx=10, pady=5)
         
            self.cus = tk.StringVar()
            self.sta1 = tk.StringVar()
            self.sta2 = tk.StringVar()
            self.sta3 = tk.StringVar()
            self.sta4 = tk.StringVar()
            self.sta5 = tk.StringVar()
            self.sta6 = tk.StringVar()
            self.size1 = tk.IntVar()
            self.size2 = tk.IntVar()
            self.size3 = tk.IntVar()
            self.size4 = tk.IntVar()
            self.size5 = tk.IntVar()
            self.size6 = tk.IntVar()
            self.plug = tk.IntVar()
            self.bin_type = tk.IntVar()

            self.plug.set(240)
            self.bin_type.set(1)
            self.size1.set(0)
            self.size2.set(0)
            self.size3.set(0)
            self.size4.set(0)
            self.size5.set(0)
            self.size6.set(0)
#            button_opt = {'fill': Tkconstants.BOTH, 'padx': 0, 'pady': 0}
#            listbox = tk.Listbox(self.f3)
#            listbox.pack()
#            for item in bin_sizes:
#                listbox.insert('end', item)

            customers_search_json = glob.glob('*.json')
            customers_search_txt = glob.glob('*.txt')
            customers = customers_search_json + customers_search_txt

            # if len(customers_search_json) > 0:
            #     for entry in customers_search_json:
            #         entry_upd = entry.replace('.json', '')
            #         customers.append(entry_upd)

            bin_sizes = [0, 12, 18, 24, 30, 36, 42, 48, 54, 60, 72]
            
#            bin_sizes2 = (0, 12, 18, 24, 30, 36, 42, 48, 54, 60, 72)
#            
#            cb = ttk.Combobox(self.f3, values=bin_sizes2, state='readonly')
#            cb.current(1)
#            cb.pack()

            tk.Radiobutton(self.f0, text="787-9", font=('Helvetica', 12), variable=self.plug, value=240).pack(side='left', padx=5, pady=5)
            tk.Radiobutton(self.f10, text="OUTBD", font=('Helvetica', 12), variable=self.bin_type, value=1).pack(side='left', padx=5, pady=5)
            tk.Radiobutton(self.f11, text="CTR", font=('Helvetica', 12), variable=self.bin_type, value=2).pack(side='left', padx=5, pady=5)
            tk.Label(self.f1, text="CUSTOMER:", font=('Helvetica', 14)).pack(side='left', padx=5)
            tk.Label(self.f2, text="STA", font=('Helvetica', 14)).pack(side='left', padx=50, pady=5)
            tk.Label(self.f2, text="BIN SIZE",  font=('Helvetica', 14)).pack(side='left', padx=25, pady=5)
            tk.Entry(self.f3, textvariable=self.sta1, width=10).pack(side='left',  padx=5, pady=5)
            tk.Label(self.f3, text="-", font=('Helvetica', 14)).pack(side='left', padx=15)
            tk.Entry(self.f4, textvariable=self.sta2, width=10, state='disabled').pack(side='left', padx=5, pady=5)
            tk.Label(self.f4, text="-", font=('Helvetica', 14)).pack(side='left', padx=15)
            tk.Entry(self.f5, textvariable=self.sta3, width=10, state='disabled').pack(side='left', padx=5, pady=5)
            tk.Label(self.f5, text="-", font=('Helvetica', 14)).pack(side='left', padx=15)
            tk.Entry(self.f6, textvariable=self.sta4, width=10, state='disabled').pack(side='left', padx=5, pady=5)
            tk.Label(self.f6, text="-", font=('Helvetica', 14)).pack(side='left', padx=15)
            tk.Entry(self.f7, textvariable=self.sta5, width=10, state='disabled').pack(side='left', padx=5, pady=5)
            tk.Label(self.f7, text="-", font=('Helvetica', 14)).pack(side='left', padx=15)
            tk.Entry(self.f8, textvariable=self.sta6, width=10, state='disabled').pack(side='left', padx=5, pady=5)
            tk.Label(self.f8, text="-", font=('Helvetica', 14)).pack(side='left', padx=15)
            apply(tk.OptionMenu, (self.f3, self.size1) + tuple(bin_sizes)).pack(side='left', padx=10, pady=5)
            apply(tk.OptionMenu, (self.f4, self.size2) + tuple(bin_sizes)).pack(side='left', padx=10, pady=5)
            apply(tk.OptionMenu, (self.f5, self.size3) + tuple(bin_sizes)).pack(side='left', padx=10, pady=5)
            apply(tk.OptionMenu, (self.f6, self.size4) + tuple(bin_sizes)).pack(side='left', padx=10, pady=5)
            apply(tk.OptionMenu, (self.f7, self.size5) + tuple(bin_sizes)).pack(side='left', padx=10, pady=5)
            apply(tk.OptionMenu, (self.f8, self.size6) + tuple(bin_sizes)).pack(side='left', padx=10, pady=5)
            apply(tk.OptionMenu, (self.f1, self.cus) + tuple(customers)).pack(side='left', padx=5, pady=0)
            tk.Radiobutton(self.f01, text="787-10", font=('Helvetica', 12), variable=self.plug, value=456).pack(side='left', padx=5, pady=5)
            tk.Button(self.f9, text="SUBMIT", command=self.start,  font=('Helvetica', 14)).pack()
        # need to run progressbar in another thread    
#        def progress_start(self):
#
#            self.progress()
#            self.start()
#            
#        def progress(self):
#
#            top = tk.Toplevel(self)
#            w = ttk.Progressbar(top, orient='horizontal', mode='indeterminate')
#            w.pack(expand=True, fill=tk.BOTH, side=tk.TOP)
#            w.start(10)
            
        def start(self):

            global ref_objects

            plug_value = self.plug.get()
            input_config = []
            work_path_folder = os.getcwd()
            
            size1_value = self.size1.get()
            size2_value = self.size2.get()
            size3_value = self.size3.get()
            size4_value = self.size4.get()
            size5_value = self.size5.get()
            size6_value = self.size6.get()

            sta1 = self.sta1.get()
            sta1_value_x = sum([float(x) for x in sta1.split('+')])
            sta1_value = sta_value(sta1_value_x*25.4, plug_value)
            self.sta1.set(str(sta1_value))
   
            sta2_value_x = sta1_value_x + size1_value
            sta2_value = sta_value(sta2_value_x*25.4, plug_value)
            if size2_value != 0:
                self.sta2.set(sta2_value)
    
            sta3_value_x = sta2_value_x + size2_value
            sta3_value = sta_value(sta3_value_x*25.4, plug_value)
            if size3_value != 0:
                self.sta3.set(sta3_value)

            sta4_value_x = sta3_value_x + size3_value
            sta4_value = sta_value(sta4_value_x*25.4, plug_value)
            if size4_value != 0:
                self.sta4.set(sta4_value)

            sta5_value_x = sta4_value_x + size4_value
            sta5_value = sta_value(sta5_value_x*25.4, plug_value)
            if size5_value != 0:
                self.sta5.set(sta5_value)
 
            sta6_value_x = sta5_value_x + size5_value
            sta6_value = sta_value(sta6_value_x*25.4, plug_value)
            if size6_value != 0:
                self.sta6.set(sta6_value)
                
            root.update()

            customer_txt = self.cus.get()
            if '.json' in customer_txt:
                customer_txt = parse_ss(customer_txt, plug_value)
            customer = customer_txt.replace('.txt', '')
            
            sta_list = [sta1_value, sta2_value, sta3_value, sta4_value,
                        sta5_value, sta6_value]
            size_list = [size1_value, size2_value, size3_value, size4_value,
                         size5_value, size6_value]
                         
            for sta, size in zip(sta_list, size_list):
                if sta != '' and size != 0:
                    if sta[0] != '0' and sta[0] != '1':
                        sta = '0' + sta
                    input_config.append([sta, size])
                
            print customer
            print input_config
            selection1.SelectElement2(['AnyObject'], 'CHOOSE IRM', False)
            selected1 = selection1.Item2(1).Value
            pn = 'CA' + str(selected1.PartNumber)[2:]
            print pn
            instance_id = selected1.Name
            print instance_id
            side1 = str(selected1.Name)[:-4]
            side = side1[-2:]
            print side
            ringposts = {'836Z1510-24': ['836Z1510-24_pclamp'],
                         '836Z1510-22': ['836Z1510-22_ringpost'],
                         '836Z1510-2': ['836Z1510-2_ringpost'],
                         '836Z1510-3': ['836Z1510-3_ringpost'],
                         'IC830Z3000-1.3.2': ['IC830Z3000-1.3.2_rp'],
                         'IC830Z3000-1.5': ['IC830Z3000-1.5_rp'],
                         'IC830Z3000-1.10': ['IC830Z3000-1.10_rp'],
                         'IC830Z3000-1.13': ['IC830Z3000-1.13_rp'],
                         'IC830Z3000-1.Twenty_four_arch_LH': ['IC830Z3000-1.Twenty_four_arch_LH_rp'],
                         'IC830Z3000-1.Twenty_four_arch_RH': ['IC830Z3000-1.Twenty_four_arch_RH_rp'],
                         'IC830Z3000-1.3.3': ['IC830Z3000-1.3.3_rp', 'IC830Z3000-1.3.3_jd28',
                                              'IC830Z3000-1.3.3_jd31'],
                         'IC830Z3000-1.5.1': ['IC830Z3000-1.5.1_rp', 'IC830Z3000-1.5.1_jd28',
                                              'IC830Z3000-1.5.1_jd31'],
                         'IC830Z3000-1.10.2': ['IC830Z3000-1.10.2_rp', 'IC830Z3000-1.10.2_jd28',
                                               'IC830Z3000-1.10.2_jd31'],
                         'IC830Z3000-1.13.2': ['IC830Z3000-1.13.2_rp', 'IC830Z3000-1.13.2_jd28',
                                               'IC830Z3000-1.13.2_jd31'],
                         'IC830Z3000-1.12.2': ['IC830Z3000-1.12.2_jd31'],
                         'IC830Z3000-1.2.2': ['IC830Z3000-1.2.2_jd31'],
                         'IC830Z3000-1.9.2': ['IC830Z3000-1.9.2_jd31'],
                         'IC830Z3000-1.4.2': ['IC830Z3000-1.4.2_jd31'],
                         'IC830Z3000-1.6.2': ['IC830Z3000-1.6.2_jd31'],
                         'IC830Z3000-1.11.2': ['IC830Z3000-1.11.2_jd31'],
                         'IC830Z3000-1.14.2': ['IC830Z3000-1.14.2_jd31']}

            product = collection.Item(instance_id)
            product.ApplyWorkMode(2)
            collection1 = product.Products
            all_irm_parts = []
            for j in xrange(1, collection1.Count + 1):
                all_irm_parts.append(collection1.Item(j).name)
            #collection_rename = product.ReferenceProduct.Products
            #str_to_replace = '1251-2'
            #str_new = '1251-41'
            #for m in xrange(1, collection_rename.Count+1):
            #    part1_name = collection_rename.Item(m).Name
            #    try:
            #        part1_new_name = part1_name.replace(str_to_replace, str_new)
            #        collection_rename.Item(m).Name = part1_new_name
             #   except:
             #       continue
            pn = add_carm_as_external_component(pn, instance_id)
            change_inst_id(pn, instance_id)
            carm_name = collection1.Item(collection1.Count).Name

            bin_order = 0
            for fairing in input_config:
                bin_order += 1
                reference(fairing[1], instance_id, carm_name, fairing[0], side, customer, plug_value, bin_order,
                          len(input_config), all_irm_parts, pn)
                omf = Ref(customer, fairing[0], side, plug_value, bin_order, len(input_config), all_irm_parts, name=instance_id)
                create_point_sta(pn, omf)

            # Scan parts for JD points and flagnotes:
            for prod in xrange(1, collection.Count + 1):
                if collection.Item(prod).name == instance_id:
                    collection1 = collection.Item(prod).Products
                    for prod1 in xrange(1, collection1.Count + 1):
                        parent_name = collection1.Item(prod1).Name
                        collection2 = collection1.Item(prod1).ReferenceProduct.Products
                        if collection2.Count > 0:
                            for prod2 in xrange(1, collection2.Count + 1):
                                cur_name = collection2.Item(prod2).Name
                                collection2.Item(prod2).Name = parent_name + '_' + cur_name
            try:
                check_call('cd ' + work_path_folder + ' & ' + 'Helpers.exe coord', shell=True)
            except:
                sys.exit("running external process json_export_console error")

            scan_parts(collection1, instance_id, carm_name, pn, ringposts, plug_value)

            for ref in ref_objects:
                ref.remove_component()
            
            try:
                check_call('cd ' + work_path_folder + ' & ' + 'Helpers.exe jd', shell=True)
            except:
                sys.exit("running external process GetPointCoordinates error")
            add_jd_annotation(pn, input_config[0][0], 31, instance_id)
            # access_captures(instance_id, 1)
            try:
                check_call('cd ' + work_path_folder + ' & ' + 'Helpers.exe fn', shell=True)
            except:
                sys.exit("running external process GetFlagNoteCoordinates error")
            add_annotation(pn, input_config, instance_id)
            capture_del(pn, instance_id)
            jd_del(pn)
            std_parts(pn)
            cameras(pn, side, omf)
            rename_part_body(pn)
            selection1.Clear()
            #for f in xrange(4):
            #    selection1.Clear()
            #    selection1.Add(collection.Item(collection.Count - f))
            #    selection1.visProperties.SetShow(1)
            #    selection1.Clear()
            activate_top_prod()
            Product.Update()
            root.destroy()
            
if __name__ == "__main__":

    root = tk.Tk()
    #root.resizable(0, 0)
    root.resizable(width=True, height=True)
    Application(root).pack()
    try:
        catia = win32com.client.Dispatch('catia.application')
        productDocument1 = catia.ActiveDocument
        Product = productDocument1.Product
        collection = Product.Products
        selection1 = productDocument1.Selection
        selection2 = productDocument1.Selection
        documents = catia.Documents
    except:
        tkMessageBox.showwarning("CATIA error",
                                 "Make sure to have CATProduct opened in CATIA before running an application")

        root.destroy()
    else:
        root.mainloop()