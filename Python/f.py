import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter
import tkinter.ttk
import time
#NOTE if IDMS runs slowly, it might be time to ask to restart the server

def correctURL(url, browser): #return true if current url matches given url
    if url == browser.current_url:
        return True
    else:
        return False
def max_page_size(browser):
    WebDriverWait(ieb,5).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow\"]"))) #wait till page size dropdown appears
    ieb.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow").click()
    ieb.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown > div > ul > li:nth-child(3)").click()
###NOTE Depreciated
###def filter_for(search, browser):
###    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl02_ctl03_FilterTextBox_AttestationName").send_keys(search)
###    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl02_ctl03_Filter_AttestationName").click()
###    browser.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_rfltMenu_detached > ul > li:nth-child(2)").click()
###NOTE Depreciated
def longin():
    ieb.find_element(By.NAME,"txtUserName").send_keys("zainal\tdi\n")
    ###NOTE Depreciated, faster to find username field and tabbing to password field
    ###ieb.find_element(By.NAME,"txtPassword").send_keys("di\n") # "\n" works here for input
    ###ieb.find_element(By.NAME,"btnLogin").click() #Depreciated, faster to press enter key
    ###NOTE Depreciated, faster to find username field and tabbing to password field
def check_boxes(app_name, browser):
    table = browser.find_element(By.ID, 'ctl00_ContentPlaceHolder1_grdTask_ctl00')
    ###XXX Depreciated, easier to understand but slower as each row is scanned one by one
    ###rows = table.find_elements(By.XPATH, "./tbody/tr") #find all rows in the table
    ###print(len(rows))
    ###for row in rows:
    ###    cell = row.find_element(By.XPATH, "./td[3]") #for each row, get
    ###    print(cell.text)
    ###    if cell.text == "HRMS/FLINS":
    ###        row.find_element(By.XPATH, "./td[1]").click()
    ###XXX Depreciated, easier to understand but slower as each row is scanned one by one

    #NOTE DONE:1. Get only the rows with the 3rd column having an element a with text "HRMS/FLINS" inside
    #NOTE DONE:2. Within those rows, get the checkbox
    cboxes = table.find_elements(By.XPATH, "./tbody/tr[td[3][a=\"" + app_name + "\"]]/td[1]")
    for cbox in cboxes:
        cbox.click()
def fill(field, name, browser):
    if field == "assign_to":
        field = "ctl00_ContentPlaceHolder1_UserNameAutoCompleteBox_Input"
        browser.find_element(By.ID, field).send_keys(name)
        time.sleep(1.5)
        browser.find_element(By.ID, field).send_keys("\n")
    elif field == "delegate_to":
        field = "ctl00_ContentPlaceHolder1_DelegatedToAutoCompleteBox_Input"
        browser.find_element(By.ID, field).send_keys(name)
        time.sleep(1.5)
        browser.find_element(By.ID, field).send_keys("\n")
    elif field == "comments":
        field = "ctl00_ContentPlaceHolder1_txtComments"
        browser.find_element(By.ID, field).send_keys(name)
def get_IDMS_name(browser):
    return browser.find_element(By.ID, "ctl00_DIHeader1_lblProfile").text
class WindowData:
    a=None
    b=None
    def __init__(self, master, list1, list2):
        self.master = master #Save the master window as a variable

        self.e={} #have a dictionary to store the list of radio uttons
        self.string = tkinter.StringVar()
        self.string.set("elliot")
        self.label1 = tkinter.Label(master, text="Type of Tasks to Delegate:").grid(row=0)
        self.label2 = tkinter.Label(master, text="Delegate to:").grid(row=1)

        self.e[1] = tkinter.ttk.Combobox(master, width = 30, justify="left", state ="readonly", values= list2)
        self.e[1].set("Windows (ITD Only)") #set the default value for the
        for i, item  in enumerate(list1, start=2):#make the list of radio buttons
            self.e[i] = tkinter.Radiobutton(master, text=item, variable=self.string, value=item, justify="center")

        self.e[1].grid(row=0, column=1)#organising
        for i, item in enumerate(list1, start=2):
            self.e[i].grid(row=i, column=1, sticky="w")
        self.button1 = tkinter.Button(master, text='Cancel', command=quit).grid(row=len(list1)+2, column=0, sticky="w", pady=4)
        self.button2 = tkinter.Button(master, text='Ok', command=self.get_data).grid(row=len(list1)+2, column=1, sticky="w", pady=4)
    def get_task_type(self):
        return self.e[1].get()
    def get_delegate_user(self):
        return self.string.get()
    def get_data(self):
        self.a = self.get_task_type()
        self.b = self.get_delegate_user()
        self.master.destroy()#closes the window
###NOTE depreciated, No more use
###class ToLoopOut(Exception):
###    def __init__(self,*args,**kwargs):
###        Exception.__init__(self,*args,**kwargs)
###NOTE depreciated, No more use
taskList = ["s (a a)","2.f"]
wintelTeam = ["u", "o", "p"]
#START main code here
window = tkinter.Tk()
windowData = WindowData(window, wintelTeam, taskList)
window.mainloop()
application_name = windowData.a
user_to_delegate = windowData.b
ieb = webdriver.Ie()
#go to login page
ieb.get("")
longin()
while True: #THERE HAS TO BE A BETTER WAY TO DO THIS
    try:
        while not correctURL("", ieb):
            time.sleep(0.5)
        break
    except selenium.common.exceptions.UnexpectedAlertPresentException:
        time.sleep(3)
user = get_IDMS_name(ieb)
#go to bulk assign group task page
ieb.get("")
max_page_size(ieb)
time.sleep(1) #to let the page reload with 50 tasks
check_boxes(application_name, ieb)
fill("assign_to", user, ieb)
ieb.find_element(By.XPATH, "//input[@type=\"submit\" and @value=\"Save\"]").click()
#go to create delegation page
ieb.get("")
fill("delegate_to", user_to_delegate, ieb)
fill("comments", "FYA, please", ieb)
check_boxes(application_name, ieb)
ieb.find_element(By.XPATH, "//input[@type=\"submit\" and @value=\"Save\"]").click()
#ieb.close()