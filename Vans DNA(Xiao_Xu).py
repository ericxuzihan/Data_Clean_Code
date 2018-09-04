from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

def Xiao_Xu_test3(filename):
    rb = open_workbook(filename)
    r_sheet1 = rb.sheet_by_index(0)
    wb = copy(rb)
    r_sheet1.cell_value(0, 0)

    first_row_list1 = r_sheet1.row_values(0)

    list1 = []
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")

    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        abc = str(oldsheet1.cell(i, 0).value)[-3:]
        sheet1.write(i, 3, abc)
        list1.append(abc)
        #list1.append(str(oldsheet1.cell(i, 2).value))


    print(list1, len(list1))

    book.save('C:\\Users\\Eric\\Desktop\\32test.xls')
    #print(list1, len(list1))
    #print(list2, len(list2))

    return
#######################################################################################
# ################################Driver program#######################################

Xiao_Xu_test3('C:\\Users\\Eric\\Desktop\\MSS Asics Sales Report & Price Change 8.3.18.xlsx')

