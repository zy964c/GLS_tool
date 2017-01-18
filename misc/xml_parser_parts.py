import xmltodict
import os
from xlwt import Workbook, easyxf
import sys

class Kit(object):

    def __init__(self, kit_name, part_notes, kit_type, kit_pn, kit_components, annotations):
        
        self.kit_name = kit_name
        self.part_notes = part_notes
        self.kit_type = kit_type
        self.kit_pn = kit_pn
        self.kit_components = kit_components
        self.annotations = annotations

    def find_sub_kit_pn(self):

        if self.kit_type != 'kit' or self.kit_components == []:
            return None
        sub_kit_pn = []
        try:
            for sub_kit in self.kit_components:
                part_dict = {}
                if str(sub_kit['NoteTitle']) == 'INSTANCE NUMBER':
                    part_dict['PartNumber'] = sub_kit['PartNumber']
                    part_dict['Name'] = sub_kit['Name']
                    part_dict['Qty'] = sub_kit['Qty']
                    sub_kit_pn.append(part_dict)
        except:
            if str(self.kit_components['NoteTitle']) == 'INSTANCE NUMBER':
                part_dict = {}
                part_dict['PartNumber'] = self.kit_components['PartNumber']
                part_dict['Name'] = self.kit_components['Name']
                part_dict['Qty'] = self.kit_components['Qty']
                sub_kit_pn.append(part_dict)

        return sub_kit_pn

    def find_kit_parts(self):

        if self.kit_type != 'sub_kit':
            return None
        kit_parts = []
        try:
            for u in self.kit_components:
                part_dict = {}
                part_dict['PartNumber'] = u['PartNumber']
                part_dict['Name'] = u['Name']
                part_dict['Qty'] = u['Qty']
                kit_parts.append(part_dict)
        except TypeError:
            part_dict = {}
            part_dict['PartNumber'] = self.kit_components['PartNumber']
            part_dict['Name'] = self.kit_components['Name']
            part_dict['Qty'] = self.kit_components['Qty']
            kit_parts.append(part_dict)
            
        return kit_parts
    
    def find_std(self, pn):

        std_list = ['PER']
        output = []
        output_set = set()
        std_type = ''
        found = False
        notes = self.part_notes
        annotations = self.annotations
        plces_to_look = [notes, annotations]

        for j in plces_to_look:
            for i in j:
                split_by_point = i.split('.')
                for sentence in split_by_point:
                    found = False
                    std_type = ''
                    if pn in str(sentence):
                        note_divided = sentence.split(' ')
                        for word in note_divided:
                            if found is True:
                                std_type += (' ' + word)
                                continue
                            for std in std_list:
                                if std == word:    
                                    found = True
                                    std_type += word
                        if found is True:
                            output.append(std_type)
        for t in output:
            output_set.add(t)
        return list(output_set)
        
def kit_factory(doc):
    
    kit_name = (doc['Root']['Header']['Name'])
    try:
        part_notes = (doc['Root']['PartSpecificData']['PartNotes']['Line'])
    except:
        part_notes = []
    try:
        annotations = (doc['Root']['PartSpecificData']['Annotations']['Line'])
    except:
        annotations = []
    if not 'PROVIDED' in kit_name:
        kit_type = 'kit'
    else:
        kit_type = 'sub_kit'       
    kit_pn = (doc['Root']['Header']['PartNumber'])
    try:
        kit_components = (doc['Root']['Components']['Component'])
    except KeyError:
        kit_components = []
    kit_current = Kit(kit_name, part_notes, kit_type, kit_pn, kit_components, annotations)
    return kit_current

def find_file(file_name):

    for fn in os.listdir('.'):
        if os.path.isfile(fn) and '.xml' in fn and file_name in fn:
            with open(fn) as fd:
                doc = xmltodict.parse(fd.read())
                return doc
    return None

if __name__ == "__main__":
    xml_files = []
    for fn in os.listdir('.'):
        if os.path.isfile(fn) and '.xml' in fn:
            xml_files.append(fn)
    row_count = 0
    style1 = easyxf('alignment: wrap True')
    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1',cell_overwrite_ok=True)
    sheet1.col(0).width = 10000
    sheet1.col(1).width = 20000
    sheet1.col(2).width = 2000
    sheet1.col(3).width = 10000
    sheet1.col(4).width = 20000
    sheet1.row(row_count).write(0, 'Part Number')
    sheet1.row(row_count).write(1, 'Part Name')
    sheet1.row(row_count).write(2, 'QTY*')
    sheet1.row(row_count).write(3, 'Associated Drawing')
    sheet1.row(row_count).write(4, 'Associated Manual/Specification')
    row_count += 1
    
    for xmlfile in xml_files:
        #if row_count > 100:
            #book.save('experiment.xls')
            #sys.exit("stop")
        with open(xmlfile) as fd:
            doc = xmltodict.parse(fd.read())
        try:
            kit_current = kit_factory(doc)
        except KeyError:
            continue
        if kit_current.kit_type == 'kit':       
            kit_pn = kit_current.kit_pn
            kit_name = kit_current.kit_name
            sub_kit_pn = kit_current.find_sub_kit_pn()
            print kit_pn
            print sub_kit_pn
            if sub_kit_pn is None:
                continue
            for sub_kit_part in sub_kit_pn:
                sub_kit_part_pn = sub_kit_part['PartNumber']
                sub_kit_part_name = sub_kit_part['Name']
                sub_kit_part_qty = sub_kit_part['Qty']
                print sub_kit_part_pn
                if sub_kit_part_pn is not None:
                    try:
                        subkit_current = kit_factory(find_file(sub_kit_part_pn))
                    except:
                        continue
                    parts = subkit_current.find_kit_parts()
                    #print type(parts)
                    if parts is not None:
                        for part in parts:
                            part_pn = part['PartNumber']
                            part_name = part['Name']
                            part_qty = part['Qty']
                            alt = ''
                            if '##ALT' in part_pn:
                                location = part_pn.find('##ALT')
                                part_pn_full = part_pn
                                part_pn = part_pn[:location]
                                alt = part_pn_full[location:]
                            fond_std_kit = kit_current.find_std(part_pn)
                            sheet1.row(row_count).write(0, part_pn + alt, style1)
                            sheet1.row(row_count).write(1, part_name, style1)
                            sheet1.row(row_count).write(2, part_qty, style1)
                            sheet1.row(row_count).write(3, kit_pn, style1)
                            if len(fond_std_kit) > 0:
                                reformatted_output = ''
                                for item in range(len(fond_std_kit)):
                                    if fond_std_kit[item] == fond_std_kit[-1]:
                                        reformatted_output += (fond_std_kit[item])
                                    else:
                                        reformatted_output += (fond_std_kit[item] + '\n')
                                    #reformatted_output = reformatted_output.rstrip()
                                sheet1.row(row_count).write(4, reformatted_output, style1)
                            else:
                                sheet1.row(row_count).write(4, 'N/A', style1)
                            row_count += 1
                    elif parts is None:
                        continue
                else:
                    continue
    book.save('output_table.xls')
