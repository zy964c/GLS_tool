from xlrd import open_workbook


def compare_to_excel(path, column_number):

    excel = []

    # acquiring data from excel
    #---------------------------------------------------------------

    book = open_workbook(path)
    sheet = book.sheet_by_index(0)

    for row_index in range(0, sheet.nrows):
        excel.append(str(sheet.cell(row_index, column_number).value))

    res = []
    for k in excel:
        spl = k.split(',')
        try:
            res.append(int(spl[0]))
        except ValueError:
            continue
    return res


if __name__ == "__main__":

    print compare_to_excel('pdm.xls', 0)
