from xlrd import open_workbook
#from xlutils.copy import copy
#import openpyxl
import xlsxwriter
#import xlwt

def same_width_error(filename):
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
    print("final_list here represents any two records that are actually the same, need to be highlighted:", fill_color_list)    # final_list here represent any two records that are actually the same, need to be highlighted
    return fill_color_list

#######################################################################################
# ################################Driver program#######################################

same_width_error('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Same Width Item.xlsx')

######################################################################################
#######################################################################################



