import json

from xlrd import open_workbook


def compare_to_excel(path, column_number):
    
    excel = []

    # acquiring data from excel
    #---------------------------------------------------------------
    
    book = open_workbook(path)
    sheet = book.sheet_by_index(0)

    for row_index in range(6, sheet.nrows):
        excel.append(str(sheet.cell(row_index, column_number).value))              

    result = {}
    for part in excel:
        key, value = part[:-1].split(',', 1)
        result[key] = value

    #pprint(result)

    with open('C:\\Temp\\zy964c\\names.txt', 'w') as f:
        json.dump(result, f)

if __name__ == "__main__":

    compare_to_excel('C:\\Temp\\zy964c\\gls_lib\\name1.xlsx', 0)
