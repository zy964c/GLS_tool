import win32com.client
# from pprint import pprint


def summarize(carm_pn):
    
    catia = win32com.client.Dispatch('catia.application')
    documents = catia.Documents
    #for doc in xrange(1, documents.Count + 1):
        #print documents.Item(doc).Name
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    parameters = carm_part.Parameters
    hybridBodies1 = carm_part.HybridBodies
    hybridBody1 = hybridBodies1.Item("Joint Definitions")
    hybridBodies2 = hybridBody1.HybridBodies
    
    std_dict = {}
    
    for hb in xrange(1, hybridBodies2.Count + 1):
        current_folder = hybridBodies2.Item(hb)
        points = current_folder.HybridBodies
        std_parts = points.Item('Non Instantiated Standard Parts')
        hq = 'Joint Definitions\\' + current_folder.Name + '\\Hole Quantity'
        selection1.Add(std_parts)
        selection1.Search(str('Knowledgeware.Parameter, sel'))
        for n in xrange(1, selection1.Count + 1):
            std_param = selection1.Item2(n).Value
            std_param_text = str(std_param.Value)
            #print std_param_text
            bac_name = std_param_text.split('|')[0]
            hole_qty = parameters.Item(hq).Value
            if bac_name not in std_dict:
                std_dict[bac_name] = [hole_qty, std_param_text.split('|')[1]]
            else:
                std_dict[bac_name] = [std_dict[bac_name][0] + hole_qty,
                                     std_dict[bac_name][1]]
        selection1.Clear()

    return std_dict
    
def std_parts(carm_pn):
    
    std_dict = summarize(carm_pn)
    catia = win32com.client.Dispatch('catia.application')
    documents = catia.Documents
    productDocument1 = catia.ActiveDocument
    selection1 = productDocument1.Selection
    carm_doc = documents.Item(carm_pn + ".CATPart")
    carm_part = carm_doc.Part
    parameters = carm_part.Parameters
    hybridBodies1 = carm_part.HybridBodies
    try:
        std = hybridBodies1.Item("Standard Parts:")
    except:
        std = hybridBodies1.Add()
        std.Name = "Standard Parts:"
    for key in std_dict:
        key_frmtd = key
        while key_frmtd[-1] == ' ':
            key_frmtd = key_frmtd[:-1]
        param = parameters.CreateString(key_frmtd, str(std_dict[key][0]) +
                               '|' + key + '|' + std_dict[key][1])                    
        selection1.Add(param)
        selection1.Copy()
        selection1.Clear()
        selection1.Add(std)
        selection1.Paste()
        selection1.Clear()
        parameters.Remove(key_frmtd)
    
    
if __name__ == "__main__":
    
    pn = 'CA836Z1121-41'
    std_parts(pn)
