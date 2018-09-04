from xlrd import open_workbook
from xlutils.copy import copy
import xlwt


# from ast import literal_eval

def UPC_hash_dict(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")
    for i in range(1, 482):
        first_row_list = r_sheet.row_values(i)
        new_list = first_row_list[0]
        new_list = new_list[51:]
        new_list = new_list[:-2]
        print(new_list)
        mylist = new_list.split('],')
        final_list = []
        for j in mylist:
            j = j + "]"
            if "]]" in j:
                j = j[:-1]
            # print(mylist.index(j))
            final_list.append(j)
            print(j)
            # print(final_list)
            print(final_list.index(j))
            # print(i)
            # sheet1.write(i,final_list.index(j),j)
    # book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Rod_test(output).xls')

    return


"""""""""        

            book = xlwt.Workbook(encoding="utf-8")
            sheet1 = book.add_sheet("Sheet 1")
            sheet1.write(i,mylist.index(j),j)

            #for i, n in enumerate(mens_shoe_error_list):
                #sheet1.write(i + 1, 0, n)
            #for j, k in enumerate(womens_shoe_error_list):
                #sheet1.write(j + 1, 1, k)
    book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Rod_test(output).xls')



        #sheet1.write(0, 0, "mens_error_row")
        #sheet1.write(0, 1, "womens_error_row")
        #for i in range(len(first_row_list):
            #first_row_list[i] = first_row_list[i][51:]

"""""

#######################################################################################
# ################################Driver program#######################################

UPC_hash_dict('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Rod_test.xlsx')

######################################################################################
######################################################################################