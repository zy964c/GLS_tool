import win32com.client


def change_inst_id(pn, instance_id):

    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    Product = productDocument1.Product
    collection = Product.Products

    irm_pn = 'IR' + pn[2:]

    for prod in xrange(1, collection.Count + 1):
        if collection.Item(prod).PartNumber == irm_pn:
            collection1 = collection.Item(prod).ReferenceProduct.Products
            collection1.Item(collection1.Count).Name = irm_pn[6:] + '_' + instance_id[4:-3] + 'CARM'

if __name__ == "__main__":

    change_inst_id('CA836Z1671-3', 'GLS_STA1293-1401+096_OB-ODF_LH_CAI')

