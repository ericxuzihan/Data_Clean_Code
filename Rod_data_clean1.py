from xlrd import open_workbook
from xlutils.copy import copy
from ast import literal_eval
#import xlwt

def UPC_hash_dict(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    for i in range(0,483):
        #print(i)
        first_row_list = r_sheet.row_values(i)
        for i in first_row_list:
        # if 'names' in i:
        # del i['names']
            i = i[51:]
            i = i[:-2]
            mylist = i.split(',')
            #python_dict = literal_eval(i)
            #target_list =list(python_dict.values())
            #print(target_list)
            print(mylist)
            #print(i)
        #first_row_list = first_row_list[50:]
        #first_row_list = first_row_list[:-1]
        #print(i)

    #if 'names' in first_row_list:
        #del first_row_list['names']
    #print(first_row_list)
    return

#######################################################################################
# ################################Driver program#######################################

UPC_hash_dict('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Rod_test.xlsx')

######################################################################################
######################################################################################



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


    w_sheet.write(0, 8, "FBA Sale At Current BuyBox - Net Payment")

    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)
        for i in connector:
            for j in range(oldsheet.nrows):
                if i == str(oldsheet.cell(j, UPC_index).value):
                    w_sheet.write(j, 8, connector[i])
    wb.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\repricing_output.csv')
    
    
def complete_file(complete_filename):
    rb = open_workbook(complete_filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    #w_sheet.getCells().insertColumns(8, 1)

    #w_sheet.cell_value(8, 0) == "eric is awesome"
    r_sheet.cell_value(0, 0)
    print(wb)



    return

complete_file('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Rod_test.xlsx')

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
    
    
    
    connector = UPC_hash_dict('C:\\Users\\Eric\\Desktop\\New folder\\price_list_305994_file_1_of_1.xlsx')[0]
UPC_index = UPC_hash_dict('C:\\Users\\Eric\\Desktop\\New folder\\price_list_305994_file_1_of_1.xlsx')[1]
print(connector)
print(UPC_index)
"""""""""