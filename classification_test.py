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
    name_index = first_row_list.index('Display Name')  # name_index is a integer that indicates the column number

    print(name_index)  # test if the column number is expected

    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)
        #mens_shoe_error_list = []
        #womens_shoe_error_list = []

        list_Tee = []
        list_T_Shirt1 = []
        list_T_Shirt2 = []
        list_Snapack = []
        list_Short = []
        list_Shorts = []
        list_Backpack = []
        list_Backpacks = []
        list_Beanie = []
        list_Belt = []
        list_Jackets = []
        list_Cap = []
        list_Bucket = []
        list_Hoodie = []
        list_Summit_Pocket = []
        list_Gully_Pullon = []
        #list_Boonie_Hat = []
        #list_Jacket = []
        #list_Jkt = []


        for i in range(oldsheet.nrows):
            if ' Tee' in str(oldsheet.cell(i, name_index).value):
                list_Tee.append(i)
            elif ' T-Shirt' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt1.append(i)
            elif ' T Shirt' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt2.append(i)
            elif ' Summit Pocket' in str(oldsheet.cell(i, name_index).value):
                list_Summit_Pocket.append(i)
            elif ' Snapback' in str(oldsheet.cell(i, name_index).value):
                list_Snapack.append(i)
            elif ' Short' in str(oldsheet.cell(i, name_index).value):
                list_Short.append(i)
            elif ' Gully Pullon' in str(oldsheet.cell(i, name_index).value):
                list_Gully_Pullon.append(i)
            elif ' Shorts' in str(oldsheet.cell(i, name_index).value):
                list_Shorts.append(i)
            elif ' Backpack' in str(oldsheet.cell(i, name_index).value):
                list_Backpack.append(i)
            elif ' Backpacks' in str(oldsheet.cell(i, name_index).value):
                list_Backpacks.append(i)
            elif ' Beanie' in str(oldsheet.cell(i, name_index).value):
                list_Beanie.append(i)
            elif ' Belt' in str(oldsheet.cell(i, name_index).value):
                list_Belt.append(i)
            elif ' Jackets' in str(oldsheet.cell(i, name_index).value):
                list_Jackets.append(i)
            elif ' Jacket' in str(oldsheet.cell(i, name_index).value):
                list_Jackets.append(i)
            elif ' Jkt' in str(oldsheet.cell(i, name_index).value):
                list_Jackets.append(i)
            elif ' Cap' in str(oldsheet.cell(i, name_index).value):
                list_Cap.append(i)
            elif ' Dad Hat' in str(oldsheet.cell(i, name_index).value):
                list_Cap.append(i)
            elif ' Boonie' in str(oldsheet.cell(i, name_index).value):
                list_Bucket.append(i)
            elif ' Bucket' in str(oldsheet.cell(i, name_index).value):
                list_Bucket.append(i)
            elif ' Hoodie' in str(oldsheet.cell(i, name_index).value):
                list_Hoodie.append(i)
            #elif:
                #list_none.append(i)
            #CellString = str(oldsheet.cell(i, name_index).value)
        book = xlwt.Workbook(encoding="utf-8")
        sheet1 = book.add_sheet("Sheet 1")
        #sheet1.write(0, 0, "Gender")
        #sheet1.write(0, 1, "womens_row")
        for i, n in enumerate(list_Tee):
            sheet1.write(n, 0, "Knits&Tees")
        for i, n in enumerate(list_T_Shirt1):
            sheet1.write(n, 0, "Knits&Tees")
        for i, n in enumerate(list_T_Shirt2):
            sheet1.write(n, 0, "Knits&Tees")
        for i, n in enumerate(list_Summit_Pocket):
            sheet1.write(n, 0, "Knits&Tees")
        for i, n in enumerate(list_Snapack):
            sheet1.write(n, 0, "Casual Hat")
        for i, n in enumerate(list_Short):
            sheet1.write(n, 0, "Shorts")
        for i, n in enumerate(list_Shorts):
            sheet1.write(n, 0, "Shorts")
        for i, n in enumerate(list_Gully_Pullon):
            sheet1.write(n, 0, "Shorts")
        for i, n in enumerate(list_Backpack):
            sheet1.write(n, 0, "Backpacks")
        for i, n in enumerate(list_Backpacks):
            sheet1.write(n, 0, "Backpacks")
        for i, n in enumerate(list_Beanie):
            sheet1.write(n, 0, "Beanies")
        for i, n in enumerate(list_Belt):
            sheet1.write(n, 0, "Belts")
        for i, n in enumerate(list_Jackets):
            sheet1.write(n, 0, "Jackets&Coats")
        for i, n in enumerate(list_Cap):
            sheet1.write(n, 0, "Baseball Caps")
        for i, n in enumerate(list_Bucket):
            sheet1.write(n, 0, "Bucket Hats")
        for i, n in enumerate(list_Hoodie):
            sheet1.write(n, 0, "Fashion Hoodies")


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