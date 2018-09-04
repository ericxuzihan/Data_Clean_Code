from xlrd import open_workbook
from xlutils.copy import copy
#import xlwt

def whole_width_test(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('Name')  # name_index is a integer that indicates the column number
    prod_class = first_row_list.index('Product Classification')
    print(name_index)  # test if the column number is expected

    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)

        for i in range(oldsheet.nrows):
            CellString = str(oldsheet.cell(i, name_index).value)
            CellString = CellString.replace("-C", "-M")
            CellString = CellString.replace("-6E", "-XXW")
            CellString = CellString.replace("-EE", "-W")
            CellString = CellString.replace("-E", "-W")
            CellString = CellString.replace("-2N", "-N")

            if str(oldsheet.cell(i, prod_class).value) == "Mens Shoes":
                CellString = CellString.replace("-2E", "-W")
                CellString = CellString.replace("-B", "-N")
                CellString = CellString.replace("-D", "-M")
                CellString = CellString.replace("-4E", "-XW")

            if str(oldsheet.cell(i, prod_class).value) == "Womens Shoes":
                CellString = CellString.replace("-4E", "-XXW")
                CellString = CellString.replace("-B", "-M")
                CellString = CellString.replace("-D", "-W")
                CellString = CellString.replace("-2E", "-XW")

            w_sheet.write(i, name_index, CellString)
    wb.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Width test(modified).xls')
    return


#######################################################################################
# ################################Driver program#######################################

whole_width_test('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Width test.xlsx')

######################################################################################
######################################################################################

# if __name__ == "__main__":