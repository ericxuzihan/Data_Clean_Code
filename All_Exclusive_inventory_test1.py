from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

def all_exclusive_inventory(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    #r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('DNA Display Name')  # name_index is a integer that indicates the column number

    print(name_index)  # test if the column number is expected

    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)
        #mens_shoe_error_list = []
        #womens_shoe_error_list = []
        list_men = []
        list_women = []
        list_uni = []
        list_none = []
        for i in range(oldsheet.nrows):
            if ' Men' in str(oldsheet.cell(i, name_index).value):
                list_men.append(i)
            elif ' Women' in str(oldsheet.cell(i, name_index).value):
                list_women.append(i)
            elif ' Unisex' in str(oldsheet.cell(i, name_index).value):
                list_uni.append(i)
            else:
                list_none.append(i)
            #CellString = str(oldsheet.cell(i, name_index).value)
        book = xlwt.Workbook(encoding="utf-8")
        sheet1 = book.add_sheet("Sheet 1")
        #sheet1.write(0, 0, "Gender")
        #sheet1.write(0, 1, "womens_row")
        for i, n in enumerate(list_men):
            sheet1.write(n, 0, "Men's")
        for i, n in enumerate(list_women):
            sheet1.write(n, 0, "Women's")
        for i, n in enumerate(list_uni):
            sheet1.write(n, 0, "Unisex")
        for i, n in enumerate(list_none):
            sheet1.write(n, 0, "None")
        #for j, k in enumerate(womens_shoe_error_list):
            #sheet1.write(j + 1, 1, k)

    book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\A1232322332test.xls')
    return

#######################################################################################
# ################################Driver program#######################################

all_exclusive_inventory('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\All Exclusive Inventory.xlsx')

######################################################################################
######################################################################################

"""""""""""""""""""""""""""""""""
                book = xlwt.Workbook(encoding="utf-8")
                sheet1 = book.add_sheet("Sheet 1")
                sheet1.write(0, 0, "mens_error_row")
                sheet1.write(0, 1, "womens_error_row")
                for i, n in enumerate(mens_shoe_error_list):
                    sheet1.write(i + 1, 0, n)
                for j, k in enumerate(womens_shoe_error_list):
                    sheet1.write(j + 1, 1, k)
"""""""""""""""



# if __name__ == "__main__":