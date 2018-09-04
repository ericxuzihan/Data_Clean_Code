from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

def Xiao_Xu_test3(filename):
    rb = open_workbook(filename)
    r_sheet1 = rb.sheet_by_index(0)
    # r_sheet.write(25, 0, "Renew_Gender")
    wb = copy(rb)
    # w_sheet = wb.get_sheet(0)
    r_sheet1.cell_value(0, 0)

    first_row_list1 = r_sheet1.row_values(0)

    list1 = []
    list2 = []
    target_list = []
    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        list1.append(str(oldsheet1.cell(i, 2).value))

    for i in list1:
        if ':' in i:
            list2.append(i)

    for j in list2:
        s = ''.join(x for x in j if x.isdigit())
        string_length1 = (len(s) - 1) / 2
        string_length2 = (len(s) - 2) / 2
        string_length3 = (len(s) - 3) / 2
        string_length4 = (len(s) - 4) / 2
        print(int(string_length1))
        if s[0:int(string_length1)] != s[int(string_length1):(int(string_length1) * 2)] and \
            s[0:int(string_length2)] != s[int(string_length2):(int(string_length2) * 2)] and \
            s[0:int(string_length3)] != s[int(string_length3):(int(string_length3) * 2)]:
            #s[0:int(string_length4)] != s[int(string_length1):int(string_length4) * 2]:
            #print(s)
            #print(j)
            target_list.append(s)

    print(target_list, len(target_list))

    #print(list1, len(list1))
    #print(list2, len(list2))

    return
#######################################################################################
# ################################Driver program#######################################


Xiao_Xu_test3('C:\\Users\\Eric\\Desktop\\All CR8 Inventory.xlsx')


"""""""""""""""""""""""
if s[0:string_length1 + 1] != s[string_length1:string_length1 * 2 + 1] or \
        s[0:string_length2 + 1] != s[string_length1:string_length2 * 2 + 1] or \
        s[0:string_length3 + 1] != s[string_length1:string_length3 * 2 + 1] or \
        s[0:string_length4 + 1] != s[string_length1:string_length4 * 2 + 1]:
    print(s)
    target_list.append(s)

print(target_list)
"""""""""