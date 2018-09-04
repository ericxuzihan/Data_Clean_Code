from xlrd import open_workbook
from xlutils.copy import copy
import xlwt


def celigo_error(filename):
    rb = open_workbook(filename)
    r_sheet1 = rb.sheet_by_index(0)
    # r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    # w_sheet = wb.get_sheet(0)
    r_sheet1.cell_value(0, 0)

    first_row_list1 = r_sheet1.row_values(0)

    #item_index1 = first_row_list1.index('Item No.')  # name_index is a integer that indicates the column number

    #print(item_index1)  # test if the column number is expected

    list_size = []
    list_color = []

    new_list_1 = []
    new_list_2 = []
    # for sheetname in rb.sheet_names():
    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        abc1 = str(oldsheet1.cell(i, 8).value)
        start1 = 'size (Merchant: '
        #start2 = 'color (Merchant: '
        end = ').A'
        abcd1 = abc1[abc1.find(start1):abc1.find(end)]
        merchant_string = abcd1[6:27]
        amazon_string = abcd1[29:54]
        new_list_1.append(merchant_string)
        new_list_2.append(amazon_string)




    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")

    for i, n in enumerate(new_list_1):
        sheet1.write(i, 0, n)
    for i, n in enumerate(new_list_2):
        sheet1.write(i, 1, n)


    book.save('C:\\Users\\Eric\\Desktop\\celigo_error_output.xls')

        #list4.append(abc)
        #list4.append(str(oldsheet1.cell(i, 4).value))

    #print(list4)
    #for i in list4:
        #i[i.find(start):i.find(end)]
        #print(i)




    return


#######################################################################################
# ################################Driver program#######################################

celigo_error('C:\\Users\\Eric\\Desktop\\Original_Data_Sheet\\Celigo Errors 8-24.xlsx')


######################################################################################
######################################################################################


"""""""""

        else:
            abcd2 = abc1[abc1.find(start2):abc1.find(end)]
            merchant_string = abcd2[0:18]
            amazon_string = abcd2[24:]
            new_list_1.append(merchant_string)
            new_list_2.append(amazon_string)
            #list_color.sppend(abcd2)

        #print(abcd)

        #print(len(list4))

    while '' in list_size:
        list_size.remove('')
    #print(list4,len(list4))


    for i in list_size:
        i = i[0:18]
        new_list_1.append(i)
    print(new_list_1)
    for i in list4:
        i = i[24:]
        new_list_2.append(i)
    print(new_list_2)



    for i in list4:
        target_string = i[i.find]
    text = "Hello there @bob !"
    user = text[text.find("@") + 1:]
    print
    user
    list4.pop(0)
    print(list4)


new_list3
#print(new_list_total, len(new_list_total))

rb = open_workbook('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\MasterSheet-1534959684662.xlsx')
r_sheet = rb.sheet_by_index(0)
r_sheet.cell_value(0, 0)
first_row_list1 = r_sheet.row_values(0)

item_index1 = first_row_list1.index('SKU')  # name_index is a integer that indicates the column number

print(item_index1)  # test if the column number is expected

UPC_list = []
final_row_list = []
Brand_list = []
oldsheet123 = rb.sheet_by_index(0)

for rowidx in range(r_sheet.nrows):
    for j in new_list_total:
        if j == r_sheet.cell_value(rowidx, item_index1) and j != r_sheet.cell_value(rowidx + 1, item_index1):
            UPC_list.append(r_sheet.cell_value(rowidx, 3))
            Brand_list.append(r_sheet.cell_value(rowidx, 4))

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

for i, n in enumerate(new_list_total):
    sheet1.write(i, 0, n)
for i, n in enumerate(UPC_list):
    sheet1.write(i, 1, n)
for i, n in enumerate(Brand_list):
    sheet1.write(i, 2, n)

book.save('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\AABCfdfdtest.xls')

"""


