from xlrd import open_workbook
#from xlutils.copy import copy
#import openpyxl
#import xlsxwriter
#import xlwt

def NewBalance_width_error(filename):    #
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


#######################################################################################
# ################################Driver program#######################################

NewBalance_width_error('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\NewBalanceWidthError724Results672.xlsx')

######################################################################################
#######################################################################################



