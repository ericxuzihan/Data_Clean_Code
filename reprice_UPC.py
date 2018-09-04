from xlrd import open_workbook
from xlutils.copy import copy
#import xlwt

def UPC_hash_dict(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    UPC_index = first_row_list.index('UPC')
    FBA_price_index = first_row_list.index('FBA Sale At Current BuyBox - Net Payment')
    print(UPC_index)  # test if the column number is expected
    print(FBA_price_index)
    connect_dict = {}
    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)
        for i in range(oldsheet.nrows):
            UPC_value = str(oldsheet.cell(i, UPC_index).value)
            FBA_sale_price = str(oldsheet.cell(i, FBA_price_index).value)
            connect_dict[UPC_value] = FBA_sale_price
            print(UPC_value)
            print(FBA_sale_price)
    del connect_dict['UPC']
    #print(connect_dict)
    return connect_dict, UPC_index

#######################################################################################
# ################################Driver program#######################################

connector = UPC_hash_dict('C:\\Users\\Eric\\Desktop\\New folder\\price_list_305994_file_1_of_1.xlsx')[0]
UPC_index = UPC_hash_dict('C:\\Users\\Eric\\Desktop\\New folder\\price_list_305994_file_1_of_1.xlsx')[1]
print(connector)
print(UPC_index)
######################################################################################
######################################################################################

def complete_file(complete_filename):
    rb = open_workbook(complete_filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    #w_sheet.getCells().insertColumns(8, 1)

    #w_sheet.cell_value(8, 0) == "eric is awesome"
    r_sheet.cell_value(0, 0)
    w_sheet.write(0, 8, "FBA Sale At Current BuyBox - Net Payment")

    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)
        for i in connector:
            for j in range(oldsheet.nrows):
                if i == str(oldsheet.cell(j, UPC_index).value):
                    w_sheet.write(j, 8, connector[i])
    wb.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\repricing_output.csv')

    return

complete_file('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\RepricerPricingTiersResults27.xlsx')


# if __name__ == "__main__":


"""""""""
for sheetname in rb.sheet_names():
    oldsheet = rb.sheet_by_name(sheetname)

    for i in range(oldsheet.nrows):
        CellString = str(oldsheet.cell(i, name_index).value)
        CellString = str(oldsheet.cell(i, sku_index).value)

        w_sheet.write(i, name_index, CellString)
        w_sheet.write(i, sku_index, CellString)
wb.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\anotherone.xls')
"""""""""