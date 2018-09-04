from selenium import webdriver
import time
import logging

def chrome_script():

    driver = webdriver.Chrome('C:\\Users\\Eric\\Downloads\\chromedriver_win32\\chromedriver')
    driver.get("https://www.google.com/maps/dir/40.4515849,-3.6903752/41.380896,2.1228198/@40.4682789,-5.2677025,6z/am=t?hl=en")
    #driver.implicitly_wait(10)

    # <button vet="13034" jsaction="print.print" class="widget-print-button kd-button kd-button-submit"> Print </button>

    #driver.find_element_by_xpath("//button[@class='widget-print-button kd-button kd-button-submit']").click()

    time.sleep(5)

    return

chrome_script()

"""""""""""""""""""""""

    driver.find_element_by_id("ctl00_MainContent_StatusFilter").click()

    driver.find_element_by_xpath("//option[text()='- Entered in System']").click()

    driver.find_element_by_name("ctl00$MainContent$SearchButton").click()

    time.sleep(1)

    driver.find_element_by_xpath("//input[@type='checkbox']").click()

    time.sleep(0.3)

    #######################################################################
    # ################# After checked all boxes ###########################

    driver.find_element_by_id("ctl00_MainContent_selectedOrdersPanel").click()

    time.sleep(0.3)

    driver.find_element_by_xpath("//option[text()='Export To A2000']").click()

    time.sleep(0.3)

    driver.find_element_by_name("ctl00$MainContent$BatchButton").click()

    time.sleep(0.3)

    driver.find_element_by_xpath("//a[@href='../../Assets/A2K.CSV']").click()

    time.sleep(4)

    #######################################################################
    # ################# After download the A2K.CSV ########################

    driver.find_element_by_xpath("//input[@type='checkbox']").click()

    time.sleep(1)

    driver.find_element_by_id("ctl00_MainContent_selectedOrdersPanel").click()

    time.sleep(2)

    driver.find_element_by_xpath("//option[text()='Print Invoices']").click()

    time.sleep(2)

    driver.find_element_by_name("ctl00$MainContent$BatchButton").click()

    time.sleep(2)

    driver.find_element_by_name("ctl00$MainContent$Print").click()

    time.sleep(10)

    driver.forward()

    driver.find_element_by_xpath("//button[@class='print default']").click()





def Firefox_Fax_Script():
    logging.basicConfig(filename='fax_script_log.log', level=logging.INFO)
    usr = "exodec"
    psw = "3x0@dm1n"
    profile = webdriver.FirefoxProfile(
        "C:\\Users\\David\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\sjnvoxb7.default")
    driver = webdriver.Firefox(profile)
    driver.implicitly_wait(5)
    driver.get("http://faxportal.faxsipit.com/?level=agent")
    logging.info('load page successful')

    time.sleep(3)
    driver.find_element_by_id("USERLOGIN").send_keys(usr)
    driver.find_element_by_id("USERPASSWORD").send_keys(psw)
    driver.find_element_by_id("SUBMIT").click()
    logging.info('log in successful')

    time.sleep(3)
    driver.find_element_by_id("BTNAGENTREPORT").click()
    time.sleep(0.3)
    driver.find_element_by_id("CBTIMEPERIOD").click()
    time.sleep(0.3)
    driver.find_element_by_xpath("//*[@id='CBTIMEPERIOD']//following::option[@value='2']").click()
    time.sleep(0.3)
    driver.find_element_by_id("CBREPOUTPUT").click()
    time.sleep(0.3)
    driver.find_element_by_xpath("//*[@id='CBREPOUTPUT']//following::option[@value='1']").click()
    time.sleep(0.3)
    driver.find_element_by_id("BTNSHOWREPORT").click()
    time.sleep(8)
    logging.info('download successful')

    driver.quit()


Firefox_Fax_Script()
"""""""""""""""