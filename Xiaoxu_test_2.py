from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

def Xiao_Xu_test1(filename):
    rb = open_workbook(filename)
    r_sheet1 = rb.sheet_by_index(0)
    r_sheet2 = rb.sheet_by_index(1)
    # r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    # w_sheet = wb.get_sheet(0)
    r_sheet1.cell_value(0, 0)
    r_sheet2.cell_value(0, 0)
    first_row_list1 = r_sheet1.row_values(0)
    first_row_list2 = r_sheet1.row_values(0)
    item_index1 = first_row_list1.index('Item No.')  # name_index is a integer that indicates the column number
    item_index2 = first_row_list1.index('Item No.')
    #print(item_index1)  # test if the column number is expected
    #print(item_index2)  # test if the column number is expected
    list11 = []
    list22 = []
    # for sheetname in rb.sheet_names():
    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        list11.append(str(oldsheet1.cell(i, item_index1).value))
    list11.pop(0)
    #print(list11, len(list11))
    oldsheet2 = rb.sheet_by_index(1)
    for i in range(oldsheet2.nrows):
        list22.append(str(oldsheet2.cell(i, item_index1).value))
    list22.pop(0)
    #print(list22, len(list22))

    list33 = list11+list22

    return list33


def Xiao_Xu_test2(filename):
    rb = open_workbook(filename)
    r_sheet1 = rb.sheet_by_index(0)
    r_sheet2 = rb.sheet_by_index(1)
    r_sheet3 = rb.sheet_by_index(2)
    r_sheet4 = rb.sheet_by_index(3)
    r_sheet5 = rb.sheet_by_index(4)
    r_sheet6 = rb.sheet_by_index(5)
    r_sheet7 = rb.sheet_by_index(6)
    r_sheet8 = rb.sheet_by_index(7)
    r_sheet9 = rb.sheet_by_index(8)
    r_sheet10 = rb.sheet_by_index(9)
    r_sheet11 = rb.sheet_by_index(10)
    r_sheet12 = rb.sheet_by_index(11)


    # r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    # w_sheet = wb.get_sheet(0)

    r_sheet1.cell_value(0, 0)
    r_sheet2.cell_value(0, 0)
    r_sheet3.cell_value(0, 0)
    r_sheet4.cell_value(0, 0)
    r_sheet5.cell_value(0, 0)
    r_sheet6.cell_value(0, 0)
    r_sheet7.cell_value(0, 0)
    r_sheet8.cell_value(0, 0)
    r_sheet9.cell_value(0, 0)
    r_sheet10.cell_value(0, 0)
    r_sheet11.cell_value(0, 0)
    r_sheet12.cell_value(0, 0)

    first_row_list1 = r_sheet1.row_values(0)


    item_index1 = first_row_list1.index('Item No.')  # name_index is a integer that indicates the column number
    #print(item_index1)  # test if the column number is expected

    list1 = []
    list2 = []
    list3 = []
    list4 = []
    list5 = []
    list6 = []
    list7 = []
    list8 = []
    list9 = []
    list10 = []
    list11 = []
    list12 = []


    # for sheetname in rb.sheet_names():
    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        list1.append(str(oldsheet1.cell(i, item_index1).value))
    list1.pop(0)
    #print(list1, len(list1))

    oldsheet2 = rb.sheet_by_index(1)
    for i in range(oldsheet2.nrows):
        list2.append(str(oldsheet2.cell(i, item_index1).value))
    list2.pop(0)
    #print(list2, len(list2))

    oldsheet3 = rb.sheet_by_index(2)
    for i in range(oldsheet3.nrows):
        list3.append(str(oldsheet3.cell(i, item_index1).value))
    list3.pop(0)
    #print(list3, len(list3))

    oldsheet4 = rb.sheet_by_index(3)
    for i in range(oldsheet4.nrows):
        list4.append(str(oldsheet4.cell(i, item_index1).value))
    list4.pop(0)
    #print(list4, len(list4))

    oldsheet5 = rb.sheet_by_index(4)
    for i in range(oldsheet5.nrows):
        list5.append(str(oldsheet5.cell(i, item_index1).value))
    list5.pop(0)
    #print(list5, len(list5))

    oldsheet6 = rb.sheet_by_index(5)
    for i in range(oldsheet6.nrows):
        list6.append(str(oldsheet6.cell(i, item_index1).value))
    list6.pop(0)
    #print(list6, len(list6))

    oldsheet7 = rb.sheet_by_index(6)
    for i in range(oldsheet7.nrows):
        list7.append(str(oldsheet7.cell(i, item_index1).value))
    list7.pop(0)
    #print(list7, len(list7))

    oldsheet8 = rb.sheet_by_index(7)
    for i in range(oldsheet8.nrows):
        list8.append(str(oldsheet8.cell(i, item_index1).value))
    list8.pop(0)
    #print(list8, len(list8))

    oldsheet9 = rb.sheet_by_index(8)
    for i in range(oldsheet9.nrows):
        list9.append(str(oldsheet9.cell(i, item_index1).value))
    list9.pop(0)
    #print(list9, len(list9))

    oldsheet10 = rb.sheet_by_index(9)
    for i in range(oldsheet10.nrows):
        list10.append(str(oldsheet10.cell(i, item_index1).value))
    list10.pop(0)
    #print(list10, len(list10))

    oldsheet11 = rb.sheet_by_index(10)
    for i in range(oldsheet11.nrows):
        list11.append(str(oldsheet11.cell(i, item_index1).value))
    list11.pop(0)
    #print(list11, len(list11))

    oldsheet12 = rb.sheet_by_index(11)
    for i in range(oldsheet12.nrows):
        list12.append(str(oldsheet12.cell(i, item_index1).value))
    list12.pop(0)
    #print(list12, len(list12))


    list13 = list1 + list2 + list3 + list4 + list5 + list6 + list7 + list8 + list9 + list10 + list11 + list12
    #print(list13, len(list13))

    return list13

def Xiao_Xu_test3(filename):
    rb = open_workbook(filename)
    r_sheet1 = rb.sheet_by_index(0)
    # r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    # w_sheet = wb.get_sheet(0)
    r_sheet1.cell_value(0, 0)

    first_row_list1 = r_sheet1.row_values(0)

    item_index1 = first_row_list1.index('Item No.')  # name_index is a integer that indicates the column number

    #print(item_index1)  # test if the column number is expected

    list444 = []

    # for sheetname in rb.sheet_names():
    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        list444.append(str(oldsheet1.cell(i, item_index1).value))
    list444.pop(0)
    #print(list4, len(list4))

    return list444
#######################################################################################
# ################################Driver program#######################################

new_list1 = Xiao_Xu_test1('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\Creative Recreation.xlsx')
new_list2 = Xiao_Xu_test2('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\Converse.xlsx')
new_list3 = Xiao_Xu_test3('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\Skechers1.xlsx')

######################################################################################
######################################################################################
new_list_total = new_list1 + new_list2 + new_list3
#print(new_list_total, len(new_list_total))


rb = open_workbook('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\MasterSheet-1534959684662.xlsx')
r_sheet = rb.sheet_by_index(0)
r_sheet.cell_value(0, 0)
first_row_list1 = r_sheet.row_values(0)

row_number_list = []
UPC_list = []
final_row_list = []
Brand_list = []
oldsheet123 = rb.sheet_by_index(0)

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

for rowidx in range(oldsheet123.nrows):
    for j in new_list_total:
        if j == str(oldsheet123.cell(rowidx, 0).value) and j != str(oldsheet123.cell(rowidx + 1, 0).value):
            sheet1.write(new_list_total.index(j), 0, j)
            sheet1.write(new_list_total.index(j), 1, str(oldsheet123.cell(rowidx, 3).value))
            sheet1.write(new_list_total.index(j), 2, str(oldsheet123.cell(rowidx, 4).value))



book.save('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\AABCfdfdtest.xls')



"""""""""

            UPC_list.append(str(oldsheet123.cell(rowidx, 3).value))
            Brand_list.append(str(oldsheet123.cell(rowidx, 4).value))

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

for i, n in enumerate(new_list_total):
    sheet1.write(i, 0, n)
for i, n in enumerate(UPC_list):
    sheet1.write(i, 1, n)
for i, n in enumerate(Brand_list):
    sheet1.write(i, 2, n)
"""""