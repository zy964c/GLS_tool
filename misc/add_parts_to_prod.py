import win32com.client
catia = win32com.client.Dispatch('catia.application')
productDocument1 = catia.ActiveDocument
print productDocument1.name
collector = productDocument1.Product.Products
documents = catia.Documents
for i in range(1, documents.Count):
    if 'cgr' not in documents.Item(i).Name:
        collector.AddExternalComponent(documents.Item(i))
        documents.Item(i).close()
        print documents.Item(i).Name
