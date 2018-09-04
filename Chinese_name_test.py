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

    #list444 = []
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")

    oldsheet1 = rb.sheet_by_index(0)
    for i in range(oldsheet1.nrows):
        if 'CT A/S HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'J' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布童鞋")

        elif 'CT A/S HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                '65' in str(oldsheet1.cell(i, 1).value) and \
                len(str(oldsheet1.cell(i, 1).value)) == 7:
            sheet1.write(i, 0, "匡威查克泰勒全明星二代高帮帆布童鞋")

        elif 'CT A/S HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'F' in str(oldsheet1.cell(i, 1).value) or 'M' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'C' in str(oldsheet1.cell(i, 1).value) and '13' or '14' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮全皮运动鞋")

        elif 'CT A/S HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'C' in str(oldsheet1.cell(i, 1).value) and '15' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")


        elif str(oldsheet1.cell(i, 4).value) == 'CT A/S ZEBRA OX' and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布斑马纹童鞋")


        elif 'CT A/S LEOPARD OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                '30' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布豹纹童鞋")

        elif 'CT A/S LEOPARD OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) \
                and '10' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")


        elif 'CONVERSE HN' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'F' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威HN运动童鞋")

        elif 'CONVERSE HN' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'C' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威HN男士运动鞋")


        elif 'CT A/S OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                '3J' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布童鞋")

        elif 'CT A/S OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value) and \
                'M' in str(oldsheet1.cell(i, 1).value) and \
                len(str(oldsheet1.cell(i, 1).value)) == 5:
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")

        elif 'CT A/S OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")


        elif 'CTAS MADISON OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布童鞋")

        elif str(oldsheet1.cell(i, 4).value) == 'CT A/S SPACE HI'and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布童鞋")

        elif 'CT A/S II HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布童鞋")


        elif 'CT A/S II OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星二代低帮帆布童鞋")

        elif 'REVIVAL OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威REVIVAL 低帮女士运动鞋")

        elif 'OPTIUM OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威OPTIUM 低帮女士运动鞋")

        elif 'MT STAR 3 OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威MT STAR 低帮女士运动鞋")

        elif 'CTAS DAINTY OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星DANITY低帮帆布女鞋")

        elif 'CT AS MADISON OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星MADISON低帮帆布女鞋")

        elif 'CT A/S PATCHWORK OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")


        elif 'CT A/S DOUBLE UPPER OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")

        elif 'CT A/S DOUBLE TOUNGE OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")

        elif 'CT A/S DOTS OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")

        elif 'CRIMSON OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威CRIMSON男士运动鞋")

        elif 'EL DISTRITO' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威EL DISTRITO 男士运动鞋")

        elif 'CT A/S DETAILS HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S DOUBLE UPPER HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S LOGOS HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S GRADIATED HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S MIDSOLES HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S ROLL DOWN HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S LAYER UP HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S SPECIAL HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S SEASNL HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S ROLL DOWN PLAID HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")



        elif 'KARVE OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威KARVE 男士运动鞋")

        elif 'ESCAPE OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威ESCAPE 男士运动鞋")

        elif 'WEAPON OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威WEAPON 男士运动鞋")

        elif 'CT A/S HI LEATHER' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮全皮运动鞋")


        elif 'ALL STAR LEATHER HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮全皮运动鞋")

        elif 'T OX LEATHER' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮全皮运动鞋")

        elif 'CT STREET HIKER' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星男士高帮登山鞋")

        elif 'CT A/S SHEARLING SLIP ON LEATHER' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星男士低帮休闲鞋")

        elif 'CT A/S CAMOUFLAGE HI' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮帆布男鞋")

        elif 'CT A/S OX LEATHER' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮全皮运动鞋")

        elif 'CT A/S HI ATHLETIC LEATHER' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星高帮全皮运动鞋")



    book.save('C:\\Users\\Eric\\Desktop\\Xiao_Xu\\chinese_name_output.xls')

    return
#######################################################################################
# ################################Driver program#######################################

Xiao_Xu_test3('C:\\Users\\Eric\\Desktop\\MSS 海关备案_new.xlsx')

######################################################################################
######################################################################################


"""""""""
        elif 'CT A/S OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value)and \
                'M' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")

        elif 'CT A/S OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value)and \
                '1J' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")

        elif 'CT A/S OX' == str(oldsheet1.cell(i, 4).value) and \
                'CONVERSE' in str(oldsheet1.cell(i, 5).value)and \
                'F' in str(oldsheet1.cell(i, 1).value):
            sheet1.write(i, 0, "匡威查克泰勒全明星低帮帆布男鞋")
"""""