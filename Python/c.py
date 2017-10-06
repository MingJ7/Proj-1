import selenium
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
#NOTE if IDMS runs slowly, it might be time to ask to restart the server

def correctURL(url):
    if url == ieb.current_url:
        return True
    else:
        return False

class ToLoopOut(Exception):
    def __init__(self,*args,**kwargs):
        Exception.__init__(self,*args,**kwargs)

ieb = webdriver.Ie()
ieb.get("")
while True: #THERE HAS TO BE A BETTER WAY TO DO THIS
    try:
        while not correctURL(""):
            time.sleep(0.5)
        break
    except selenium.common.exceptions.UnexpectedAlertPresentException:
        time.sleep(3)


#ieb.refresh()
ieb.find_element(By.PARTIAL_LINK_TEXT, "My Task").click()
WebDriverWait(ieb,5).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow\"]")))
#ieb.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow\"]").click()
#ieb.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown\"]/div/ul/li[3]").click()

table = ieb.find_element(By.ID, 'ctl00_ContentPlaceHolder1_grdTask_ctl00')
try:
    rows = table.find_elements(By.CLASS_NAME, "rgRow") + table.find_elements(By.CLASS_NAME, "rgAltRow") #scanning the odd rows 1st then the even rows
    print(rows)
    for row in rows:
        a = row.find_element(By.XPATH, "./td[2]")
        print(a.text)
        if a.text == "SRQ":
            a.find_element(By.CSS_SELECTOR, "a").click()
            print("clicked")
            raise ToLoopOut
#    rows = table.find_elements(By.CSS_SELECTOR, "tr[class='rgAltRow']") #scanning the even rows
#    for row in rows:
#        a = row.find_element(By.XPATH, "./td[2]")
#        print(a.text)
#        if a.text == "SRQ":
#            a.find_element(By.CSS_SELECTOR, "a").click()
#            print("clicked")
#            raise ToLoopOut
except ToLoopOut: #above will cause an error when there is no more table rows
    None
ieb.find_element(By.PARTIAL_LINK_TEXT, "View Current Access").click()
access_win = ieb.find_element(By.ID, "ctl00_ContentPlaceHolder1_UserRadTreeList")
print(access_win.text)





#ieb.close()



#<a id="ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow" style="overflow: hidden;display: block;position: relative;outline: none;">select</a>
#//*[@id="ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow"]
##ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow
#

#<li class="rcbHovered">50</li>
#//*[@id="ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown"]/div/ul/li[3]
##ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown > div > ul > li.rcbHovered
#