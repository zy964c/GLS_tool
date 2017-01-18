import win32com.client
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
documents = catia.Documents
selection1 = productDocument1.Selection
#selection1.Clear()
#singleselection
#selection1.SelectElement2(['Annotation'], 'CHOOSE ANNOTATION', False)
#multipal selection
TheSPAWorkbench = catia.ActiveDocument.GetWorkbench("SPAWorkbench")
print 'CHOOSE ANNOTATIONS AND CAPTURES'
selection1.SelectElement3(['AnyObject'],'CHOOSE ELEMENT', False, 1, False)
for selected in xrange(1, selection1.Count2 + 1):
    #selected1_type = selection1.Item2(selected).Type
    #print selected1_type
    selected1 = selection1.Item2(1).Value
    themeasurable = TheSPAWorkbench.GetMeasurable(selected1)
    intGeomType = themeasurable.GeometryName
    print intGeomType
    
