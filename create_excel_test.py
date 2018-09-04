from xlrd import open_workbook
from xlutils.copy import copy
import xlwt


a = [1,2,3,4]
b = [6,7,8,9]

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet1.write(0, 0, "man")
for i, n in enumerate(a):
    print(i)
    print(n)
    sheet1.write(i+1, 3, n)

book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\create_excel_test(1).xls')


""""""""""""""""""""""   
    for z in range(len(b)):
        for j in a:
            for k in b:
                sheet1.write(i, 1, j)
                sheet1.write(z, 2, k)
"""""""""""""""""""""""""""""""""

""""""""""""""""""""""
for z in range(len(b)):
    print(z)
    for j in a:
        print(j)
        for k in b:
            print(k)
            # sheet1.write(i, 1, j)
            # sheet1.write(z, 2, k)
"""""""""""""""""""""""