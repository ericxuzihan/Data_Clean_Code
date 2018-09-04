from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

def whole_width_test(filename):
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('Name')  # name_index is a integer that indicates the column number
    prod_class = first_row_list.index('Product Classification')
    print(name_index)  # test if the column number is expected
    print(prod_class)
    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)
        mens_shoe_error_list = []
        womens_shoe_error_list = []
        for i in range(oldsheet.nrows):
            CellString = str(oldsheet.cell(i, name_index).value)
            if str(oldsheet.cell(i, prod_class).value) == "Mens Shoes" and '-2E' in str(oldsheet.cell(i, name_index).value):
                mens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Mens Shoes" and '-6E' in str(oldsheet.cell(i, name_index).value):
                mens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Mens Shoes" and '-B' in str(oldsheet.cell(i, name_index).value):
                mens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Mens Shoes" and '-D' in str(oldsheet.cell(i, name_index).value):
                mens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Mens Shoes" and '-EE' in str(oldsheet.cell(i, name_index).value):
                mens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Mens Shoes" and '-4E' in str(oldsheet.cell(i, name_index).value):
                mens_shoe_error_list.append(i + 1)

            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-4E' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-2N' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-B' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-C' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-D' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-E' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
            elif str(oldsheet.cell(i, prod_class).value) == "Womens Shoes" and '-2E' in str(oldsheet.cell(i, name_index).value):
                womens_shoe_error_list.append(i + 1)
                print("mens_shoe_error_list is :", mens_shoe_error_list)
                print("womens_shoe_error_list is :", womens_shoe_error_list)


                book = xlwt.Workbook(encoding="utf-8")
                sheet1 = book.add_sheet("Sheet 1")
                sheet1.write(0, 0, "mens_error_row")
                sheet1.write(0, 1, "womens_error_row")
                for i, n in enumerate(mens_shoe_error_list):
                    sheet1.write(i + 1, 0, n)
                for j, k in enumerate(womens_shoe_error_list):
                    sheet1.write(j + 1, 1, k)
    book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\width_test(error_row_num_output).xls')
    return

"""""""""""""""""""""""""""""""""
            if str(oldsheet.cell(i, prod_class).value) == "Mens Shoes":
                CellString = CellString.replace("-2E", "-W")
                CellString = CellString.replace("-B", "-N")
                CellString = CellString.replace("-D", "-M")
                CellString = CellString.replace("-4E", "-XW")

            if str(oldsheet.cell(i, prod_class).value) == "Womens Shoes":
                CellString = CellString.replace("-4E", "-XXW")
                CellString = CellString.replace("-B", "-M")
                CellString = CellString.replace("-D", "-W")
                CellString = CellString.replace("-2E", "-XW")


    
    # book = open_workbook('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\random.xls')
    sheet1 = book.sheet_by_index(0)
    data1 = mens_shoe_error_list
    data2 = womens_shoe_error_list
    for i in range(sheet1.nrows):
        data1.append(sheet1.cell(i, 1).value)
        data2.append(sheet1.cell(i, 2).value)
    for i in range(oldsheet.nrows):
        CellString = str(oldsheet.cell(i, name_index).value)
        CellString = str(oldsheet.cell(i, sku_index).value)

        w_sheet.write(i, name_index, CellString)
        w_sheet.write(i, sku_index, CellString)


"""""""""""""""



#######################################################################################
# ################################Driver program#######################################

whole_width_test('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Width test.xlsx')

######################################################################################
######################################################################################

# if __name__ == "__main__":