from xlrd import open_workbook
from xlutils.copy import copy
#import xlwt

def replace_shoe_color(filename):
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
    wb.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\anotherone.xls')
    return


#######################################################################################
# ################################Driver program#######################################

replace_shoe_color('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\CustomItemSearch4Results293.xlsx')

######################################################################################
######################################################################################

# if __name__ == "__main__":