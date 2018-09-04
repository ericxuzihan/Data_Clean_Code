from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

def Xiao_Xu_test(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    # r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    # w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    item_index = first_row_list.index('Item No.')  # name_index is a integer that indicates the column number
    print(item_index)  # test if the column number is expected
    list1 = []
    # for sheetname in rb.sheet_names():
    oldsheet = rb.sheet_by_index(0)
    for i in range(oldsheet.nrows):
        list1.append(str(oldsheet.cell(i, item_index).value))
    # print(list1, len(list1))
    list1.pop(0)
    print(list1, len(list1))
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")
    for i, n in enumerate(list1):
        sheet1.write(11 * i, 1, n)
        sheet1.write(11 * (i) + 1, 1, n)
        sheet1.write(11 * (i) + 2, 1, n)
        sheet1.write(11 * (i) + 3, 1, n)
        sheet1.write(11 * (i) + 4, 1, n)
        sheet1.write(11 * (i) + 5, 1, n)
        sheet1.write(11 * (i) + 6, 1, n)
        sheet1.write(11 * (i) + 7, 1, n)
        sheet1.write(11 * (i) + 8, 1, n)
        sheet1.write(11 * (i) + 9, 1, n)
        sheet1.write(11 * (i) + 10, 1, n)

    test_list = []
    test_list = [7, 7.5, 8, 8.5, 9, 9.5, 10, 11, 11.5, 12, 12.5] * 153
    for i, n in enumerate(test_list):
        sheet1.write(i, 2, n)


    book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Xiao_Xu_Output.xls')
    return

#######################################################################################
# ################################Driver program#######################################

Xiao_Xu_test('C:\\Users\\Eric\\Desktop\\12333333.xlsx')

######################################################################################
######################################################################################

"""""""""""""""""""""""""""""""""

"""""""""""""""



# if __name__ == "__main__":