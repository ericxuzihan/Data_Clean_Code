f1 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\BW_SC_GenerateInventoryReport.js')
f2 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazon.itemfieldsimport.js')
f3 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazon.itemfieldsimportbuyandwalk.js')
f4 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazon.itemfieldsimportmybay.js')
f5 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.inventoryadjustmentsimport.js')
f6 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.itemfulfillmentimport.js')
f7 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.returnreceipts.js')
f8 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.returns.js')
f9 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.salesorder.js')
f10 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.toimport.js')
f11 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.buyandwalk.amazonfba.toreceipt.js')
f12 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\Celigo.integrator.realtime.async.Export.min.js')
f13 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\deleterecords.js')
f14 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\mav_accounts.js')
f15 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\NS SS Items to link to Images.js')
f16 = open('C:\\Users\\Eric\\Desktop\\SuiteScript\\NS UE ITEM PRICING DATE WHEN RETURN AUTHORIZATIONS.js')

#'C:\\Users\\Eric\\Desktop\\Test_Data_Sheet\\Rod_test.xlsx')

for line in f1.readlines():
    #line = str(line.lower())
    line = line.strip().lower()
    #print(str(line))
    if str(line) == "buyandwalk":
        print(str(line))
    #else:
        #print(str(line))