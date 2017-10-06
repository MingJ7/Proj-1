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
    WebDriverWait(browser,5).until(EC.presence_of_element_located((By.XPATH, "//*[@id=\"ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow\"]"))) #wait till page size dropdown appears
    ieb.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_Arrow").click()
    ieb.find_element(By.CSS_SELECTOR,"#ctl00_ContentPlaceHolder1_grdTask_ctl00_ctl03_ctl01_PageSizeComboBox_DropDown > div > ul > li:nth-child(3)").click()

def check_boxes(app_name, browser):
    table = browser.find_element(By.ID, 'ctl00_ContentPlaceHolder1_grdTask_ctl00')
    #NOTE DONE:1. Get only the rows with the 3rd column having an element a with text "HRMS/FLINS" inside
    #NOTE DONE:2. Within those rows, get the checkbox
    cboxes = table.find_elements(By.XPATH, "./tbody/tr[td[3][a=\"" + app_name + "\"]]/td[1]")
    for cbox in cboxes:
        cbox.click()
def fill(field, name, browser):
    if field == "assign_to":
        field = "ctl00_ContentPlaceHolder1_UserNameAutoCompleteBox_Input"
        browser.find_element(By.ID, field).send_keys(name)
        time.sleep(1)
        browser.find_element(By.ID, field).send_keys("\n")
    elif field == "delegate_to":
        field = "ctl00_ContentPlaceHolder1_DelegatedToAutoCompleteBox_Input"
        browser.find_element(By.ID, field).send_keys(name)
        time.sleep(1)
        browser.find_element(By.ID, field).send_keys("\n")
    elif field == "comments":
        field = "ctl00_ContentPlaceHolder1_txtComments"
        browser.find_element(By.ID, field).send_keys(name)
def get_IDMS_name(browser):
    WebDriverWait(browser,5).until(EC.presence_of_element_located((By.ID, "ctl00_DIHeader1_lblProfile"))) #wait till page loads
    return browser.find_element(By.ID, "ctl00_DIHeader1_lblProfile").text
def end_prog():
    raise SystemExit
class WindowData:
    a=None
    b=None
    def __init__(self, master, list1, list2):
        self.master = master #Save the master window as a variable

        self.radio={} #have a dictionary to store the list of radio uttons
        self.string = tkinter.StringVar()
        self.string.set(list1[0])
        self.label1 = tkinter.Label(master, text="Type of Tasks to Delegate:").grid(row=0)
        self.label2 = tkinter.Label(master, text="Delegate to:").grid(row=1)

        self.cbox = tkinter.ttk.Combobox(master, width = 30, justify="left", state ="readonly", values= list2)
        self.cbox.set("Windows (ITD Only)") #set the default value for the
        for i, item  in enumerate(list1, start=1):#make the list of radio buttons
            self.radio[i] = tkinter.Radiobutton(master, text=item, variable=self.string, value=item, justify="center")

        self.cbox.grid(row=0, column=1)#organising
        for i, item in enumerate(list1, start=1):
            self.radio[i].grid(row=i, column=1, sticky="w")
        self.button1 = tkinter.Button(master, text='Cancel', command=end_prog).grid(row=len(list1)+2, column=0, sticky="w", pady=4)
        self.button2 = tkinter.Button(master, text='Ok', command=self.get_data).grid(row=len(list1)+2, column=1, sticky="w", pady=4)
    def get_data(self):
        WindowData.a = self.cbox.get() #To save the data as data cannot be gotten once window is destroyed
        WindowData.b = self.string.get()
        self.master.destroy()#closes the window


taskList = ["1","2","3/4","5","6 7"]
wintelTeam = ["1", "2", "3"]
#START main code here
window = tkinter.Tk(className="Choose Data")
windowData = WindowData(window, wintelTeam, taskList)
window.mainloop()
application_name = WindowData.a
user_to_delegate = WindowData.b
ieb = webdriver.Ie()
#go to login page
ieb.get("")
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
fill("assign_to", user, ieb)
check_boxes(application_name, ieb)
ieb.find_element(By.XPATH, "//input[@type=\"submit\" and @value=\"Save\"]").click()
#go to create delegation page
ieb.get("")
fill("delegate_to", user_to_delegate, ieb)
fill("comments", "FYA, please", ieb)
check_boxes(application_name, ieb)
ieb.find_element(By.XPATH, "//input[@type=\"submit\" and @value=\"Save\"]").click()
print("Program has ended, you can now close this window")
#ieb.close()