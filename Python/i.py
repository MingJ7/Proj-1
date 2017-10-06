import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter
import tkinter.ttk
import time
import selenium.webdriver.remote.switch_to
import selenium.webdriver.remote.webelement
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
#NOTE if IDMS runs slowly, it might be time to ask to restart the server
def correctURL(url, browser): #return true if current url matches given url
    if url == browser.current_url:
        return True
    else:
        return False
def max_page_size(browser):
    try:
        WebDriverWait(browser,4).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow\"]"))) #wait till page size dropdown appears
    except selenium.common.exceptions.TimeoutException:
        return
    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow").click()
    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown > div > ul > li:nth-child(3)").click()
def filter_for(search, browser):
    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl02_ctl03_FilterTextBox_AttestationName").send_keys(search)
    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl02_ctl03_Filter_AttestationName").click()
    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_rfltMenu_detached > ul > li:nth-child(2)").click()
def longin():
    idms.find_element(By.NAME,"txtUserName").send_keys("zainal\tdi\n")
def get_IDMS_username(browser, employID):
    browser.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_grdUser_ctl00_ctl02_ctl03_FilterTextBox_EmployeeId\"]").send_keys(employID)
    browser.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_grdUser_ctl00_ctl02_ctl03_Filter_EmployeeId\"]").click()
    browser.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_grdUser_rfltMenu_detached\"]/ul/li[a[span=\"Contains\"]]/a/span").click()
    time.sleep(1)
    table = browser.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_grdUser_ctl00\"]")
    username = table.find_element(By.XPATH,"./tbody/tr/td[a]/a").text
    return username
def add_voice_roles(roles,browser):
    browser.find_element(By.XPATH,"//input[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.inputSearch\"]").send_keys("VOICES*\t\n")
    time.sleep(1)
    for role in roles:
        table = browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.availableRolesTable-content\"]/span/table/tbody")
        try:
            table.find_element(By.XPATH,"./tr[td//div[span=\""+role+"\"]]/td[1]").click()
        except selenium.common.exceptions.NoSuchElementException:
            browser.find_element(By.XPATH,"//div[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.availableRolesTable-scrollV\"]/table/tbody/tr[3]").click()
            table = browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.availableRolesTable-content\"]/span/table/tbody")
            try:
                table.find_element(By.XPATH,"./tr[td//div[span=\""+role+"\"]]/td[1]").click()
            except selenium.common.exceptions.NoSuchElementException:
                browser.find_element(By.XPATH,"//div[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.availableRolesTable-scrollV\"]/table/tbody/tr[1]").click()
                table = browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.availableRolesTable-content\"]/span/table/tbody")
                table.find_element(By.XPATH,"./tr[td//div[span=\""+role+"\"]]/td[1]").click()
        browser.find_element(By.XPATH,"//a[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.buttonRoleAdd\"]").click()
def remove_voice_roles(roles,browser):
    browser.find_element(By.XPATH,"//input[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.InputField1\"]").send_keys("VOICES*\t\n")
    time.sleep(1)
    for role in roles:
        table = browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.assignedRolesTable-content\"]/span/table/tbody")
        try:
            table.find_element(By.XPATH,"./tr[td//div[span=\""+role+"\"]]/td[1]").click()
        except selenium.common.exceptions.NoSuchElementException:
            browser.find_element(By.XPATH,"//div[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.assignedRolesTable-scrollV\"]/table/tbody/tr[3]").click()
            table =browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.assignedRolesTable-content\"]/span/table/tbody")
            try:
                table.find_element(By.XPATH,"./tr[td//div[span=\""+role+"\"]]/td[1]").click()
            except selenium.common.exceptions.NoSuchElementException:
                browser.find_element(By.XPATH,"//div[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.assignedRolesTable-scrollV\"]/table/tbody/tr[1]").click()
                table = browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.assignedRolesTable-content\"]/span/table/tbody")
                table.find_element(By.XPATH,"./tr[td//div[span=\""+role+"\"]]/td[1]").click()
        browser.find_element(By.XPATH,"//a[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.buttonRoleRemove\"]").click()
def delete_voice(browser):
    browser.find_element(By.XPATH,"//input[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.InputField1\"]").send_keys("VOICES*\t\n")
    time.sleep(1)
    for a in range(6):
        table = browser.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.assignedRolesTable-content\"]/span/table/tbody")
        try:
            table.find_element(By.XPATH,"./tr/td[4]").click()
        except selenium.common.exceptions.NoSuchElementException:
            break
        browser.find_element(By.XPATH,"//a[@id=\"aaaaJKNEPINJAHAG.RemoveParentRolesView.buttonRoleRemove\"]").click()
taskList = []
wintelTeam = []
application_name="l"
accessList = []
requestType = ""
username = ""
#START main code here
idms = webdriver.Ie()
#go to login page
idms.get("")
longin()
while True: #THERE HAS TO BE A BETTER WAY TO DO THIS
    try:
        while not correctURL("", idms):
            time.sleep(0.5)
        break
    except selenium.common.exceptions.UnexpectedAlertPresentException:
        time.sleep(3)
idms.get("")
max_page_size(idms)
time.sleep(1) #to let the page reload with 50 tasks
table = idms.find_element(By.ID, 'ctl00_ContentPlaceHolder1_grdTask_ctl00')
cbox = table.find_element(By.XPATH, "./tbody/tr/td[2][a=\"" + application_name + "\"]/a")
cbox.click()
time.sleep(1) #to let the page load
accessList = idms.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_GenericRequestDetailsUC_AccessNameLabel\"]")
accessList = accessList.text.split(",")
requestType = idms.find_element(By.XPATH,"//*[@id=\"ctl00_ContentPlaceHolder1_GenericRequestDetailsUC_RequestTypeLabel\"]").text
employID = idms.find_element(By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_INHUserInfoUC_txtEmpID\"]").get_attribute("value")
username = idms.find_element(By.XPATH,"//input[@name=\"ctl00$ContentPlaceHolder1$INHUserInfoUC$txtUserName\"]").get_attribute("value")
requeatURL = idms.current_url
print(username)
idms.get("")
username = get_IDMS_username(idms, employID)
print(accessList)
print(requestType)
print(employID)
print(username)
idms.get(requeatURL)
#SAP part
sap = webdriver.Ie()
switcher = selenium.webdriver.remote.switch_to.SwitchTo(sap)
sap.get("")
sap.find_element(By.XPATH,"//input[@id=\"logonuidfield\"]").send_keys("PCMJ072B\tCondense555\n")
WebDriverWait(sap,4).until(EC.presence_of_element_located((By.XPATH, "//table[@id=\"level1\"]/tbody/tr/td[a=\"User Administration\"]")))
sap.find_element(By.XPATH,"//table[@id=\"level1\"]/tbody/tr/td[a=\"User Administration\"]").click()
time.sleep(1)
switcher.frame("Desktop Innerpage    ")
switcher.frame("isolatedWorkArea")
sap.find_element(By.XPATH,"//input[@title=\"Enter a search string\"]").send_keys(username+"\t\n")
time.sleep(1)
table = sap.find_element(By.XPATH,"//td[@id=\"aaaaJKNEPINJ.UserSearchResultView.userResultTable-content\"]")
table.find_element(By.XPATH,".//tbody/tr[@rr=\"1\"]/td[@cc=\"2\"]/span").click()
WebDriverWait(sap,3).until(EC.presence_of_element_located((By.XPATH,"//span[@id=\"aaaaJKNEPINJ.DisplayUserView.associatedRoles-focus\"]")))
try:
    sap.find_element(By.XPATH,"//span[@id=\"aaaaJKNEPINJ.DisplayUserView.associatedRoles-focus\"]").click()
except selenium.common.exceptions.ElementNotInteractableException:
    sap.find_element(By.XPATH,"//span[@id=\"aaaaJKNEPINJ.DisplayUserView.TabStrip-menu\"]").click()
    sap.find_element(By.XPATH,"//span[@id=\"aaaaJKNEPINJ.DisplayUserView.TabStrip:-r\"]//tbody/tr[td[@class=\"urMnuTxt\" and span=\"Assigned Roles\"]]").click()
sap.find_element(By.XPATH,"//span[a=\"Modify\"]/a").click()
WebDriverWait(sap,3).until(EC.presence_of_element_located((By.XPATH,"//input[@id=\"aaaaJKNEPINJAHAG.AssignParentRolesView.inputSearch\"]")))
if "Amendment Add" in requestType:
    add_voice_roles(accessList,sap)
elif "Amendment Remove" in requestType:
    remove_voice_roles(accessList,sap)
time.sleep(1)
sap.find_element(By.XPATH,"//a[@id=\"aaaaJKNEPINJ.ModifyUserView.save\"]").click()
idms.find_element_by_tag_name("body").send_keys(Keys.CONTROL+"p")
#sam's stuff
#NOTE TEST DATA BELOW
