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
        # mens_shoe_error_list = []
        # womens_shoe_error_list = []

        list_T_Shirt = []
        list_Snapack = []
        list_Shorts = []
        list_Backpacks = []
        list_Beanie = []
        list_Belt = []
        list_Jackets = []
        list_Cap = []
        list_Bucket = []
        list_Hoodie = []
        list_Sunglasses = []
        list_Fashion_Sneakers = []
        list_Running = []
        list_Casual_Sneakers = []
        list_Shoe_Cleaner = []
        list_Skateboard = []
        list_Jeans = []
        list_Casual_Watches = []
        list_Accessories = []
        list_Rings = []
        list_Denim_Jeans = []
        list_Bracelets = []
        list_Socks = []
        list_Flat = []
        list_delete = []



        for i in range(oldsheet.nrows):
            if ' Tee' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif ' T-Shirt' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif ' T Shirt' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif ' Summit Pocket' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif ' Snapback' in str(oldsheet.cell(i, name_index).value):
                list_Snapack.append(i)
            elif ' Short' in str(oldsheet.cell(i, name_index).value):
                list_Shorts.append(i)
            elif ' Gully Pullon' in str(oldsheet.cell(i, name_index).value):
                list_Shorts.append(i)
            elif ' Shorts' in str(oldsheet.cell(i, name_index).value):
                list_Shorts.append(i)
            elif ' Backpack' in str(oldsheet.cell(i, name_index).value):
                list_Backpacks.append(i)
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
            elif ' Jkt' in str(oldsheet.cell(i, name_index).value) and '9Five' not in str(oldsheet.cell(i, name_index).value) and 'Rastaclat' not in str(oldsheet.cell(i, name_index).value):
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

            elif '9Five Eyewear' in str(oldsheet.cell(i, name_index).value):
                list_Sunglasses.append(i)

            elif 'Adidas' in str(oldsheet.cell(i, name_index).value):
                list_Fashion_Sneakers.append(i)
            elif 'Alpha Industries' in str(oldsheet.cell(i, name_index).value):
                list_Jackets.append(i)
            elif 'Article Number' in str(oldsheet.cell(i, name_index).value):
                list_Fashion_Sneakers.append(i)
            elif 'Asics' in str(oldsheet.cell(i, name_index).value):
                list_Running.append(i)
            elif 'Brooklyn Projects' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Chinatown Market' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Clear Weather' in str(oldsheet.cell(i, name_index).value):
                list_Fashion_Sneakers.append(i)
            elif 'Converse' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Sneakers.append(i)
            elif 'Crep Protect' in str(oldsheet.cell(i, name_index).value):
                list_Shoe_Cleaner.append(i)
            elif 'Crooks & Castles' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Diadora' in str(oldsheet.cell(i, name_index).value):
                list_Fashion_Sneakers.append(i)
            elif 'DVS' in str(oldsheet.cell(i, name_index).value):
                list_Skateboard.append(i)
            elif 'Embellish' in str(oldsheet.cell(i, name_index).value):
                list_Jeans.append(i)
            elif 'Flud Watch' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Watches.append(i)
            elif 'G-Shock' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Watches.append(i)
            elif 'Good Worth' in str(oldsheet.cell(i, name_index).value):
                list_Accessories.append(i)
            elif 'Han Cholo' in str(oldsheet.cell(i, name_index).value):
                list_Rings.append(i)
            elif 'I&M Jeans' in str(oldsheet.cell(i, name_index).value):
                list_Denim_Jeans.append(i)
            elif 'Jason Markk' in str(oldsheet.cell(i, name_index).value):
                list_Shoe_Cleaner.append(i)
            elif 'Komono' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Watches.append(i)
            elif 'Local Supply' in str(oldsheet.cell(i, name_index).value):
                list_Sunglasses.append(i)
            elif 'Nixon' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Watches.append(i)
            elif 'Oliver People' in str(oldsheet.cell(i, name_index).value):
                list_Sunglasses.append(i)
            elif 'Onitsuka Tiger' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Sneakers.append(i)
            elif 'Pangea' in str(oldsheet.cell(i, name_index).value):
                list_Bracelets.append(i)
            elif 'Puma' in str(oldsheet.cell(i, name_index).value):
                list_Fashion_Sneakers.append(i)
            elif 'Saucony' in str(oldsheet.cell(i, name_index).value):
                list_Running.append(i)
            elif 'Stance' in str(oldsheet.cell(i, name_index).value):
                list_Socks.append(i)
            elif 'Vans' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Sneakers.append(i)

            elif ' Flat' in str(oldsheet.cell(i, name_index).value):
                list_Flat.append(i)
            elif ' Classic Canvas' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Sneakers.append(i)
            elif ' Classic Spiced' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Sneakers.append(i)
            elif ' Crochet' in str(oldsheet.cell(i, name_index).value):
                list_Casual_Sneakers.append(i)

            elif 'Attic' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Diesel' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Guess' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Jansport' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Krew' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Mighty Healthy' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Navarre' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Orchill' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Proof Cases' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Rosewood Cutter' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Rustic Dime' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Skunk Juice' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Slvdr' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)
            elif 'Ssur' in str(oldsheet.cell(i, name_index).value):
                list_delete.append(i)

            elif 'Pink Dolphin' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Rogue Status' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Team Cozy' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Upxndr' in str(oldsheet.cell(i, name_index).value):
                list_T_Shirt.append(i)
            elif 'Rastaclat' in str(oldsheet.cell(i, name_index).value):
                list_Bracelets.append(i)
            elif 'Shwood' in str(oldsheet.cell(i, name_index).value):
                list_Sunglasses.append(i)
            elif 'Sneaker Lab' in str(oldsheet.cell(i, name_index).value):
                list_Shoe_Cleaner.append(i)


        book = xlwt.Workbook(encoding="utf-8")
        sheet1 = book.add_sheet("Sheet 1")

        for i, n in enumerate(list_T_Shirt):
            sheet1.write(n, 0, "Knits & Tees")
        for i, n in enumerate(list_Snapack):
            sheet1.write(n, 0, "Casual Hat")
        for i, n in enumerate(list_Shorts):
            sheet1.write(n, 0, "Shorts")
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
        for i, n in enumerate(list_Sunglasses):
            sheet1.write(n, 0, "Sunglasses")
        for i, n in enumerate(list_Fashion_Sneakers):
            sheet1.write(n, 0, "Fashion Sneakers")
        for i, n in enumerate(list_Running):
            sheet1.write(n, 0, "Running")
        for i, n in enumerate(list_Casual_Sneakers):
            sheet1.write(n, 0, "Casual Sneakers")
        for i, n in enumerate(list_Shoe_Cleaner):
            sheet1.write(n, 0, "Shoe Cleaner")
        for i, n in enumerate(list_Skateboard):
            sheet1.write(n, 0, "Skateboard")
        for i, n in enumerate(list_Jeans):
            sheet1.write(n, 0, "Jeans")
        for i, n in enumerate(list_Casual_Watches):
            sheet1.write(n, 0, "Casual Watches")
        for i, n in enumerate(list_Accessories):
            sheet1.write(n, 0, "Accessories")
        for i, n in enumerate(list_Rings):
            sheet1.write(n, 0, "Rings")
        for i, n in enumerate(list_Denim_Jeans):
            sheet1.write(n, 0, "Denim Jeans")
        for i, n in enumerate(list_Bracelets):
            sheet1.write(n, 0, "Bracelets")
        for i, n in enumerate(list_Socks):
            sheet1.write(n, 0, "Socks")
        for i, n in enumerate(list_Flat):
            sheet1.write(n, 0, "Flats")
        for i, n in enumerate(list_delete):
            sheet1.write(n, 0, "Delete")


    book.save('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\A1232322332test.xls')
    return

#######################################################################################
# ################################Driver program#######################################

all_exclusive_inventory('C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\All Exclusive Inventory.xlsx')

######################################################################################
######################################################################################

