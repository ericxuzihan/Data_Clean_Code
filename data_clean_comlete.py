from xlrd import open_workbook
from xlutils.copy import copy
import xlwt


def replace_shoe_color(filename):    #  change shoe color for example from "BLACK" to "BLK"
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('Name')  # name_index is a integer that indicates the column number
    sku_index = first_row_list.index('NetSuite Sku')  # sku_index is a integer that indicates the column number
    print(name_index)  # test if the column number is expected
    print(sku_index)

    for sheetname in rb.sheet_names():
        oldsheet = rb.sheet_by_name(sheetname)

        for i in range(oldsheet.nrows):
            CellString = str(oldsheet.cell(i, name_index).value)
            CellString = str(oldsheet.cell(i, sku_index).value)
            CellString = CellString.replace("BEIGE", "BEG")
            CellString = CellString.replace("BLACK", "BLK")
            CellString = CellString.replace("BLUE", "BLU")
            CellString = CellString.replace("BRONZE", "BNZ")
            CellString = CellString.replace("BROWN", "BRW")
            CellString = CellString.replace("GOLD", "GLD")
            CellString = CellString.replace("GREEN", "GRN")
            CellString = CellString.replace("GREY", "GRY")
            CellString = CellString.replace("METALLIC", "MET")
            CellString = CellString.replace("MULTI-COLORED", "MUL")
            CellString = CellString.replace("OFF-WHITE", "OFW")
            CellString = CellString.replace("ORANGE", "ORG")
            CellString = CellString.replace("PINK", "PNK")
            CellString = CellString.replace("PURPLE", "PUR")
            CellString = CellString.replace("SILVER", "SIL")
            CellString = CellString.replace("TRANSPARENT", "TRN")
            CellString = CellString.replace("TURQUOISE", "TRQ")
            CellString = CellString.replace("WHITE", "WHT")
            CellString = CellString.replace("YELLOW", "YEL")
            CellString = CellString.replace("ROYAL", "RYL")
            CellString = CellString.replace("DRED", "DRE")
            CellString = CellString.replace("DBLACK", "DBK")
            CellString = CellString.replace("DBLUE", "DBL")
            CellString = CellString.replace("HGREY", "HGY")
            CellString = CellString.replace("MARSALA", "MAR")
            CellString = CellString.replace("MINT", "MNT")
            CellString = CellString.replace("NAVY", "NVY")
            CellString = CellString.replace("POPPY", "POP")
            CellString = CellString.replace("FUCHSIA", "FCH")
            CellString = CellString.replace("LBLUE", "LBU")
            CellString = CellString.replace("MBLUE", "MBL")
            CellString = CellString.replace("TEAL", "TEA")
            CellString = CellString.replace("OLIVE", "OLV")
            CellString = CellString.replace("TAUPE", "TAU")
            CellString = CellString.replace("MBLACK", "MBK")
            CellString = CellString.replace("MCHARCOAL", "MCH")
            CellString = CellString.replace("OATMEAL", "OAT")
            CellString = CellString.replace("PLUM", "PLU")
            CellString = CellString.replace("TURQUOISE", "TUR")
            CellString = CellString.replace("VIOLET", "VIO")
            CellString = CellString.replace("BURGUNDY", "BUR")
            CellString = CellString.replace("MAUVE", "MAU")
            CellString = CellString.replace("TSGREEN", "TSG")
            CellString = CellString.replace("LMAUVE", "LMA")
            CellString = CellString.replace("SMINT", "SMT")
            CellString = CellString.replace("KHAKI", "KHA")
            CellString = CellString.replace("LIME", "LME")
            CellString = CellString.replace("SKY-BLUE", "SKY")
            CellString = CellString.replace("VYELLOW", "VYE")
            CellString = CellString.replace("NRED", "NRD")
            CellString = CellString.replace("IVORY", "IVY")
            # CellString = CellString.replace("RED", "RED")        # Color red doesn't change
            # CellString = CellString.replace("TAN", "TAN")        # Color tan doesn't change
            w_sheet.write(i, name_index, CellString)
            w_sheet.write(i, sku_index, CellString)
    wb.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\123.xls')
    return


def same_shoe_error(filename):    #  identify the same shoes' row number when they are the same catagory
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('Name')  # name_index is a integer that indicates the column number
    color_matrix = first_row_list.index('Color Matrix NEW')
    print(name_index)
    print(color_matrix)
    scale_list = []
    for i in range(r_sheet.nrows):
        # print(sheet.cell_value(i, name_index))
        target_column = r_sheet.cell_value(i, name_index)
        scale_list.append(target_column)
        # print(target_column)
    scale_list = [x for x in scale_list if ":" not in x]
    scale_list = [x for x in scale_list if "Name" not in x]
    compared_record_row_list = []
    for rowidx in range(r_sheet.nrows):
        for j in scale_list:
            if j == r_sheet.cell_value(rowidx, name_index):
                compared_record_row_list.append(rowidx)
                print(rowidx)
    print(compared_record_row_list)
    final_list = []
    for i in range(len(scale_list)):
        if i <= len(scale_list) - 2:
            if scale_list[i] in scale_list[i+1] or scale_list[i+1] in scale_list[i]:
                if r_sheet.cell_value(compared_record_row_list[i], color_matrix) == r_sheet.cell_value(compared_record_row_list[i+1], color_matrix):
                    final_list.append(compared_record_row_list[i])
                    final_list.append(compared_record_row_list[i+1])
                    print("same record, highlighted")
    print(scale_list)  # ['10002472', '10002472BLK']
    fill_color_list = [x+1 for x in final_list]
    print("final_list here represents any two records that are actually the same(by color), need to be highlighted:", fill_color_list)     # final_list here represent any two records that are actually the same, need to be highlighted
    return fill_color_list


def same_width_error(filename):    #  identify the same show row number when they are the same catagory with same width
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('Name')  # name_index is a integer that indicates the column number
    color_matrix = first_row_list.index('Width')
    print(name_index)
    print(color_matrix)
    scale_list = []
    for i in range(r_sheet.nrows):
        # print(sheet.cell_value(i, name_index))
        target_column = r_sheet.cell_value(i, name_index)
        scale_list.append(target_column)
        # print(target_column)
    scale_list = [x for x in scale_list if ":" not in x]
    scale_list = [x for x in scale_list if "Name" not in x]
    compared_record_row_list = []
    for rowidx in range(r_sheet.nrows):
        for j in scale_list:
            if j == r_sheet.cell_value(rowidx, name_index):
                compared_record_row_list.append(rowidx)
                print(rowidx)
    print(compared_record_row_list)
    final_list = []
    for i in range(len(scale_list)):
        if i <= len(scale_list) - 2:
            if scale_list[i] in scale_list[i+1] or scale_list[i+1] in scale_list[i]:
                if r_sheet.cell_value(compared_record_row_list[i], color_matrix) == r_sheet.cell_value(compared_record_row_list[i+1], color_matrix):
                    final_list.append(compared_record_row_list[i])
                    final_list.append(compared_record_row_list[i+1])
                    print("same record, highlighted")
    print(scale_list)  # ['10002472', '10002472BLK']
    fill_color_list = [x + 1 for x in final_list]
    print("final_list here represents any two records that are actually the same(by width), need to be highlighted:", fill_color_list)    # final_list here represent any two records that are actually the same, need to be highlighted
    return fill_color_list


def NewBalance_width_error(filename):    #  identify the overlap record for NewBalance shoes
    rb = open_workbook(filename)
    r_sheet = rb.sheet_by_index(0)
    #wb = copy(rb)
    #w_sheet = wb.get_sheet(0)         #################################################
    r_sheet.cell_value(0, 0)
    first_row_list = r_sheet.row_values(0)
    name_index = first_row_list.index('Name')  # name_index is a integer that indicates the column number
    width_column = first_row_list.index('Width')
    print(name_index)
    print(width_column)
    scale_list = []
    for i in range(r_sheet.nrows):
        # print(sheet.cell_value(i, name_index))
        target_column = r_sheet.cell_value(i, name_index)
        scale_list.append(target_column)
    print(len(scale_list))
    compared_record_row_list = []
    for rowidx in range(r_sheet.nrows):
        for j in scale_list:
            if j == r_sheet.cell_value(rowidx, name_index):
                compared_record_row_list.append(rowidx)
                #print(rowidx)
    print(len(compared_record_row_list))
    final_list = []
    for i in range(len(scale_list)):
        if i <= len(scale_list) - 2:
            if scale_list[i] in scale_list[i+1] or scale_list[i+1] in scale_list[i]:
                if r_sheet.cell_value(compared_record_row_list[i], width_column) == r_sheet.cell_value(
                            compared_record_row_list[i + 1], width_column) or len(
                    r_sheet.cell_value(compared_record_row_list[i], width_column)) is 0 or len(
                    r_sheet.cell_value(compared_record_row_list[i+1], width_column)) is 0:
                    final_list.append(compared_record_row_list[i])
                    final_list.append(compared_record_row_list[i + 1])
    print(r_sheet.cell_value(23,6))
    fill_color_list = [x + 1 for x in final_list]
    print("final_list here represents any two records that are actually the same, need to be highlighted:",
          fill_color_list)
    return final_list


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


#######################################################################################
# ################################Driver program#######################################
# to achieve different functionality that fit different excel format（ex: some have 3 column and some have 6 column）,
# each
replace_shoe_color('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\CustomItemSearch4Results293.xlsx')

######################################################################################
######################################################################################

# if __name__ == "__main__":


""""""""""""""""""""""

#######################################################################################
# ################################Driver program#######################################

same_shoe_error('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\complete_test.xlsx')

######################################################################################
#######################################################################################

#######################################################################################
# ################################Driver program#######################################

same_width_error('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Same Width Item.xlsx')

######################################################################################
#######################################################################################

#######################################################################################
# ################################Driver program#######################################

NewBalance_width_error('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\NewBalanceWidthError724Results672.xlsx')

######################################################################################
######################################################################################

"""""""""""""""""""""""""""""