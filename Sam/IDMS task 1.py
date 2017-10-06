#DONE: 1. wait for user to log in
#DONE: 2. Go to display tasks for him or her
#DONE: 3. Display pop-up box with "OKAY" button
#DONE: 4. wait for user to navigate to requests page and click the button
#XXX : 5. retrieve system ID
#XXX : 5. retrieve user ID, idms request no, generate date
#DONE: 6. start up GUI with fields, wait for user to click "Okay" button
#DONE: 7. Find a way to print
#DONE: 8. integrate "Save As" window code
#DONE: 9.Prompt the user for "Okay" again. If he or she clicks "cancel" or closes the window, the program will
#TODO: 10. Clean up code: get rid of totally general exception handling
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pywinauto
from pywinauto import Application, Desktop
import datetime as dt
import time
import re
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import platform
import os, errno

def suggest_fields(event=None):
    """Retrieve the fields to suggest to the user"""
    #try:#finding System ID from "Application Name:"
    systemIDField = ''
    systemIDField = myD.find_element_by_xpath('//tr[td[span[starts-with(.,\"Request Type\")]]]/td[3]').text
        #systemIDField = myD.find_element_by_css_selector('#ctl00_ContentPlaceHolder1_GenericRequestDetailsUC_SystemNameLabel').text
    #except selenium.common.exceptions.NoSuchElementException:
    print("Unable to find \"Application Name\" field.")

    try:#finding Username from "Username" entry box
        userNameField = ''
        userNameField = myD.find_element_by_css_selector('#ctl00_ContentPlaceHolder1_INHUserInfoUC_txtUserName').get_attribute('value')
    except selenium.common.exceptions.NoSuchElementException:
        print("Unable to find \"Username:\" entry field")

    try:
        #try to get request ID from "Exeption Remarks:"
        requestIDRawLine = myD.find_element_by_css_selector('#ctl00_ContentPlaceHolder1_ExceptionRemarksValueLabel').text
        print(requestIDRawLine)
        requestIDFoundObject = re.search(r'request (\d+)', requestIDRawLine)
        print(requestIDFoundObject)
        print(requestIDFoundObject.group(0))
        requestID = requestIDFoundObject.group(1)
    except selenium.common.exceptions.NoSuchElementException:
        print("Finding \"Request {ID}\" via Exception Remarks failed")
        try:
            #retrieve request ID by "Request ID:"
            requestID = ""
            requestID = myD.find_element_by_xpath('//tr[td[span[starts-with(.,\"Request ID\")]]]/td[3]').text
        except selenium.common.exceptions.NoSuchElementException:
            print("Finding \"Request ID\" via \"Request ID\" field  failed")

    requestID = "IDMS " + requestID
    print(requestID)
    return [systemIDField, userNameField, requestID, str(dt.date.today())]

def quit_script():
    raise SystemExit

def enter_saveasListItem(filename,timeOut=5):
    try:#find the folder on the list, and select(highlight) it.
        saveas.window(title=filename, control_type='ListItem', visible_only=False).select()
    except pywinauto.findwindows.ElementNotFoundError:#if the folder cannot be found, make one with that name, and select(highlight) it.
        saveas.window(title='Create New Folder', control_type='Button').click()
        saveas.type_keys(filename+'{ENTER}', with_spaces=True)
    #send 'ENTER' to open the folder
    saveas.type_keys('{ENTER}')

class myDriver(webdriver.Ie):
    def __init__(self):
        super().__init__()

    def ex_wait(self, waitTime, *locator):
        return WebDriverWait(self, waitTime).until(EC.presence_of_element_located(locator))

class EntryWithHelp(ttk.Frame):
    class AssignmentNotFound(Exception):
        None

    def __init__(self, master=None, preInputList=None, fieldNameToHelpMessage={"Default fieldname":"Default helpmessage for default fieldname"}):
        super().__init__(master)
        #to suggest input for the user
        self.preInputList = preInputList
        self.fNTHM = fieldNameToHelpMessage
        self.helpBoxHelpMessage = "(HelpBox:)I'm here to help!"
        self.okayHelpMessage = "Save the information and close this window."
        #Send the frame into existence
        self.pack(fill=tk.BOTH, expand=1)
        #Create widgets to put on the frame
        self.create_widgets()
        #Push the toplevel widget the widget is in to the front
        self.winfo_toplevel().attributes("-topmost", True)
        #Force focus onto the toplevel widget
        self.winfo_toplevel().focus_force()

    def create_widgets(self):
        #Create styles for the widgets
        self.fieldNameStyle = ttk.Style()
        self.helpBoxStyle = ttk.Style()
        self.fieldNameStyle.configure('fieldName.TLabel', font=('Times New Roman', 10))
        self.helpBoxStyle.configure('helpBox.TLabel', relief=tk.SUNKEN, font=('Garamond', 10))

        self.insert_rows(self)

        self.okay = ttk.Button(self, text='Okay')
        self.helpBox = ttk.Label(self, width=40, style='helpBox.TLabel')

        self.okay.pack(ipadx=1, ipady=10, padx=10, pady=10, side=tk.RIGHT)
        self.helpBox.pack(padx=13, pady=13, side=tk.LEFT, fill=tk.X, expand=1)
        #bind "binds" a key to a function
        self.okay.bind('<Button-1>', self.check_and_quit)#Button-1 is left click
        self.bind_all('<Return>', self.check_and_quit)
        self.bind_all('<FocusIn>', self.help_box)#FocusIn is when focus "starts" on the widget
        self.bind_all('<Enter>', self.help_box)

    def insert_rows(self, root):
        self.entries = {}
        for n, fieldName in enumerate(self.fNTHM):
            row = ttk.Frame(root)
            label = ttk.Label(row, style='fieldName.TLabel', text=fieldName, width=16, anchor=tk.W)
            entry = ttk.Entry(row)

            entry.insert(0, preInputList[n])

            row.pack(padx=10, pady=5, side=tk.TOP, fill=tk.X, expand=1)
            label.pack(side=tk.LEFT)
            entry.pack(side=tk.RIGHT, fill=tk.X, expand=1)

            self.entries[fieldName] = entry

    def help_box(self, event):
        #print(event)
        helpMessage = "(HelpBox:)Please input the values."#default message
        for fieldName in self.fNTHM:
            if event.widget == self.entries[fieldName]:
                helpMessage = self.fNTHM[fieldName]
        if event.widget == self.helpBox:
            helpMessage = self.helpBoxHelpMessage
        if event.widget == self.okay:
            helpMessage = self.okayHelpMessage
        self.helpBox.config(text=" "+helpMessage)

    def check_and_quit(self, event):
        try:
            for name in self.entries:
                print(self.entries[name].get())
                if not self.entries[name].get().replace(' ',''):
                    print("Idiot didn't input a field!")
                    raise self.AssignmentNotFound
        except self.AssignmentNotFound:
            tk.messagebox.showwarning("ERROR:", "Do not leave fields blank!")
        else:
            self.winfo_toplevel().destroy()

class IDMSSimpleForm(EntryWithHelp):
    def __init__(self, master=None, preInputList=None, fieldNameToHelpMessage={"Default fieldname":"Default helpmessage for default fieldname"}):
        super().__init__(master, preInputList, fieldNameToHelpMessage)
        self.dateLabel = "Processing Date:"
        self.helpBoxHelpMessage = "(HelpBox:)Click me to re-suggest fields!"
        self.add_on_to_create_widgets()

    def add_on_to_create_widgets(self):
        self.helpBox.bind('<Button-1>', self.resuggest_fields)
        #map the "X" close button to quit_script
        self.winfo_toplevel().protocol('WM_DELETE_WINDOW', quit_script)

    def resuggest_fields(self, event):
        suggestList = suggest_fields()
        for n, fieldName in enumerate(self.fNTHM):
            self.entries[fieldName].delete(0, tk.END)
            self.entries[fieldName].insert(0, suggestList[n])
    def check_and_quit(self, event):
        try:
            for name in self.entries:
                print(self.entries[name].get())
                if not self.entries[name].get().replace(' ',''):
                    print("Idiot didn't input a field!")
                    raise self.AssignmentNotFound
            self.inputtedDate = dt.date(*[int(partOfDate) for partOfDate in self.entries[self.dateLabel].get().split('-')])
            print(self.inputtedDate)
        except self.AssignmentNotFound:
            tk.messagebox.showwarning("ERROR", "Do not leave fields blank!")
        except (ValueError, TypeError):
            tk.messagebox.showwarning("ERROR", "Please input a valid date!")
        else:
            self.name_of_file = "_".join([(self.entries[name].get().lstrip(' ')) for name in self.entries if name != self.dateLabel]) + "_{:04d}-{:02d}-{:02d}".format(self.inputtedDate.year, self.inputtedDate.month, self.inputtedDate.day)
            self.winfo_toplevel().destroy()

myD = myDriver()
#Navigate to the portal
myD.get('')
##NOTE: TO DELETE LATER! type in the username and password and enter
#myD.find_element_by_css_selector('#txtUserName').send_keys('zainal\tdi\n')
##Wait for user to log into the IDMS portal
#while True:
#    try:
#        myD.find_element_by_css_selector('#txtUserName')
#        print("Still on login page")
#        time.sleep(1)
#    except selenium.common.exceptions.NoSuchElementException:
#        break
while True:
    """Wait for person to click "Print as PDF" button"""
    #This is the GUI creation
    root = tk.Tk(className=" RPA")
    #map the "X" close button to quit_script
    root.protocol('WM_DELETE_WINDOW', quit_script)
    frame = ttk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=1)
    printAsPDFButtonStyle = ttk.Style()
    quitButtonStyle = ttk.Style()
    printAsPDFButtonStyle.configure("printAsPDFButtonStyle.TButton", font=('Arial Black', 30))
    quitButtonStyle.configure("quitButtonStyle.TButton", font=('Arial Black', 30), foreground='red')
    printAsPDFButton = ttk.Button(style="printAsPDFButtonStyle.TButton", text="Print as PDF", command=frame.winfo_toplevel().destroy)
    quitButton = ttk.Button(style="quitButtonStyle.TButton", text="Quit!", command=quit_script)
    printAsPDFButton.pack(side=tk.TOP, fill=tk.BOTH, expand=1)
    quitButton.pack(side=tk.TOP, fill=tk.BOTH, expand=1)
    frame.winfo_toplevel().attributes("-topmost", True)
    frame.winfo_toplevel().focus_force()
    frame.mainloop()
    print("Script continues!")

    #DONE: add GUI here
    root = tk.Tk(className=' Input the values')
    fieldNameToHelpMessage = {
        "System ID(s):"     :"If more than 1, seperate IDs with commas.",
        "User ID(s):"       :"If more than 1, seperate IDs with commas.",
        "Source of Request:":"Eg. \"IDMS {NUMBER}\", \"Email\", etc.",
        "Processing Date:"  :"Enter the date in YYYY-MM-DD format."
    }
    preInputList = suggest_fields()
    dlg = IDMSSimpleForm(master=root, preInputList=preInputList\
                        ,fieldNameToHelpMessage=fieldNameToHelpMessage)
    #Start the GUI
    dlg.mainloop()
    print(dlg.inputtedDate)
    print(dlg.name_of_file)

    try:
        myD.find_element_by_xpath('//input[@type=\"submit\" and @value=\"Expand All\"]').click()
    except selenium.common.exceptions.NoSuchElementException:
        while True:
            try:
                myD.find_element_by_xpath('//input[@type=\"submit\" and @name[contains(.,\"GECBtnExpandColumn\")] and @class=\"rgExpand\"]').click()
            except selenium.common.exceptions.NoSuchElementException:
                break

    #Open up the "Print" window
    myD.execute_script('window.print();')
    #pywinauto for the "Print" window
    IE_DIR = r'C:\Program Files\Internet Explorer\iexplore.exe'

    """Things in the Internet Explorer "Print" window"""
    if platform.release() == '7':
        prt = Desktop(backend='win32').window(title_re='Print', class_name='#32770')
        prt.wait('ready', retry_interval=0.02)
        printerList = prt.window(title_re="FolderView", control_type="ListView", class_name_re="SysListView32")
        printerList.Select("CutePDF Writer") #the select selects an item in a ListView window
        prt.type_keys("%p")
    elif platform.release() == '10':
        ie = Application(backend='uia').connect(path=IE_DIR)
        ieTab = ie.window(title_re='.*- Internet Explorer')
        prt = ieTab.window(title='Print', class_name='#32770')#class name refers to a dialog
        prt.wait('visible', retry_interval=0.02)
        prt.window(title='CutePDF Writer', control_type='ListItem', visible_only=False).select()
        prt.type_keys("%p")

    """Things in the "Save As" window"""
    saveapp = Application(backend='uia')
    saveas = Application(backend='uia')
    print(saveapp)
    while True:
        try:
            saveapp = saveapp.connect(title='Save As')
            print("Connected.")
            saveas = saveapp.window(title='Save As')
            print("Found it!")
            break
        except:
            print("Failed to conenct to \"Save As\". Trying again!")
    print("Sees the Save As window")
    saveas.wait('enabled', retry_interval=0.02)

    #Type PDF name into the editing box
    saveas.type_keys("%n")
    saveas.type_keys(dlg.name_of_file+" -0", with_spaces=True)
    print("typed in the name of the file")

    """Start navigating the directory"""
    saveas.type_keys('%i%{DOWN}')
    print("opened?")
    try:#to try and select for the 2nd time. Known glitch.
        saveas.SaveInList1.window(title=r'',control_type='ListItem').click_input()
    except pywinauto.findwindows.ElementNotFoundError:
        print("First try failed")
        saveas.type_keys('%i%{DOWN}')
        saveas.SaveInList1.window(title=r'',control_type='ListItem').click_input()
    save_file_dir = r'N:/ITSec/Daily Requests/Robot UIA/{:04d}/{:04d}-{:02d}/{:04d}-{:02d}-{:02d}'.format(dlg.inputtedDate.year,\
                                                                                                        dlg.inputtedDate.year, dlg.inputtedDate.month,\
                                                                                                        dlg.inputtedDate.year, dlg.inputtedDate.month, dlg.inputtedDate.day)
    print(save_file_dir)
    try:#Try to make the directory if it doesn't exist. To speed up finding.
        os.makedirs(save_file_dir)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise
    for fileName in save_file_dir.split('/')[1:]:
        print(fileName)
        enter_saveasListItem(fileName)
    #Try to save the file
    saveas.type_keys("%s")
    i = 1
    while True:
        try: #in case the "Confirm Save As" window pops up, rename the file with a numerical behind.
            saveas.child_window(title='Confirm Save As', class_name='#32770').NoButton.click()
            saveas.FileNameEdit.wait('active', retry_interval=0.02)
            saveas.type_keys("%n{RIGHT}")
            saveas.type_keys("\b" * len(repr(i-1)))
            saveas.type_keys("{}".format(i))
            saveas.type_keys("%s")
            i += 1
        except Exception as e:
            print("Caught an error in \"Confirm Save As\"!",str(e))
            break

print("\033[1mALL PASS\033[0m")