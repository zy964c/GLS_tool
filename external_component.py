import win32com.client
import os
from time import gmtime, strftime


def add_carm_as_external_component(pn, name, inserted='seed'):
        """Instantiates CARM from external library"""

        current_time = strftime("%Y_%m_%d_%H_%M_%S", gmtime())
        catia = win32com.client.Dispatch('catia.application')
        oFileSys = catia.FileSystem
        current_path = os.getcwd()
        productDocument1 = catia.ActiveDocument
        product1 = productDocument1.Product
        collection_irms = product1.Products
        documents = catia.Documents
        for doc in xrange(1, documents.Count+1):
                #print documents.Item(doc).Name
                if pn in documents.Item(doc).Name and 'cgr' not in documents.Item(doc).Name:
                    #print documents.Item(doc).Name
                    pn = pn + '_' + current_time
        product_to_insert_carm = collection_irms.Item(name)
        children_of_product_to_insert_carm = product_to_insert_carm.Products
        PartDocPath = current_path + '\\' + inserted + '.CATPart'
        PartDocPath1 = current_path + '\\' + pn + '.CATPart'
        oFileSys.CopyFile(PartDocPath, PartDocPath1, True)
        PartDoc = catia.Documents.NewFrom(PartDocPath1)
        PartDoc1 = PartDoc.Product
        PartDoc1.PartNumber = pn
        added_carm = children_of_product_to_insert_carm.AddExternalComponent(PartDoc)
        PartDoc.Close()
        oFileSys.DeleteFile(PartDocPath1)
        return pn, added_carm

if __name__ == "__main__":

        new_pn = add_carm_as_external_component('CA836Z1791-2', 'GLS_STA1618-1732_OB-OMF_RH_CAI')
        print new_pn[0]
