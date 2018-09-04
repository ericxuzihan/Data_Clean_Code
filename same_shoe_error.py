from xlrd import open_workbook
#from xlutils.copy import copy
#import openpyxl
import xlsxwriter
#import xlwt

def same_shoe_error(filename):
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
    print("final_list here represents any two records that are actually the same, need to be highlighted:", fill_color_list)     # final_list here represent any two records that are actually the same, need to be highlighted
    return fill_color_list


#######################################################################################
# ################################Driver program#######################################

same_shoe_error('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\SameShoeError724Results424.xlsx')

######################################################################################
#######################################################################################

""""""""""""""""""""""
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('C:\\Users\\Eric\\Desktop\\Data_Sheet\\SameShoeError724Results424.xlsx')
pattern = xlwt.Pattern() # Create the Pattern
pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
style = xlwt.XFStyle() # Create the Pattern
style.pattern = pattern # Add Pattern to Style
worksheet.write(0, 0, 'Cell Contents', style)
workbook.save('C:\\Users\\Eric\\Desktop\\Data_Sheet\\SameShoeError724Results424.xlsx')
#data_format1 = workbook.add_format({'bg_color': '#FFC7CE'})

#for row in range(result_list):
    #print(row)
    #worksheet.write_row(row, row[0], row[0], cell_format=data_format1)

    #worksheet.write(row, 0, "Hello")
    #worksheet.write(row + 1, 0, "world")

#workbook.close()


def highlight_color(filename):
    result_list = []
    result_list = same_shoe_error(filename)
    print("result_list:", result_list)
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    data_format1 = workbook.add_format({'bg_color': '#FFC7CE'})
    for i in result_list:
        print(i)
    return
"""""""""""""""""""""""""""""


