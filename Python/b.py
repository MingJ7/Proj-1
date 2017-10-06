from pywinauto import application
from pywinauto import clipboard
import pywinauto
import time
import os
#app1 = application.Application(backend="uia")  NOTE: To do sam's way
app1 = application.Application()
app2 = application.Application()
app3 = application.Application()
outlook_dir = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\outlook.exe"
#Starting Outlook
try:
    app1.connect(path = outlook_dir)
    outlook = app1.window(title_re = ".* - Outlook") #connecting outlook window
except application.ProcessNotFoundError:
    app1.start(outlook_dir)
    outlookopen = app1.window(title_re = "Opening - Outlook")
    outlookopen.wait("exists")
    outlookopen.wait_not("exists", timeout = 30)
    outlook = app1.window(title_re = "Inbox - .* - Outlook") #connecting outlook window
#End outlook start
pywinauto.clipboard.EmptyClipboard()#clear keyboard
outlook.restore()
#print(pywinauto.clipboard.GetClipboardFormats())
#print(pywinauto.clipboard.GetFormatName(13))#number 7 is where clipboard data is saved in bytes(for string), this get type of data stored
#print(pywinauto.clipboard.GetData(format_id= 13))#getting info from the clipboard. number 13 is for unicode strings
outlook.type_keys("%hlc{ENTER}") #navigate to the 1st catagory
time.sleep(5)
outlook.type_keys("+{TAB}{ENTER}") #navigate into the email
time.sleep(2)
outlooktop = app1.top_window()
outlooktop.type_keys("+{TAB}" * 5)#move to the subject field
outlooktop.type_keys("^a^c")#copy the subject
subject = pywinauto.clipboard.GetData(format_id= 13)#getting info from the clipboard
#notepad.restore()#bring the notepad to the front
#notepad.edit.type_keys(subject)#patse output into notepad
print(subject)
outlooktop.type_keys("+{TAB}" * 2)#move to the date and time field
outlooktop.type_keys("^a^c")#copy the date and time
datetime = pywinauto.clipboard.GetData(format_id= 13)#getting info from the clipboard
print(datetime)
outlooktop.type_keys("%fpp")#to print(cutepdf needs to be set as default printer)
time.sleep(1)
outlooktop.close()
for i in range(0,10): # try to find save as window for 5 seconds
    try:
        a = app3.connect(title_re="Save As")
        break
    except pywinauto.findwindows.ElementNotFoundError:
        print(i)
        time.sleep(0.5)
a = a.window(title_re="Save As")
a.draw_outline()
#NOTE save as window also has hotkeys to access things
#outlooktop.type_keys("%{F4}")#close the message window
#notepad.restore()#bring the notepad to the front
#notepad.edit.type_keys(datetime)#patse output into notepad
#notepad.type_keys("{VK_APPS}") # this is  the "right click" button on the keyboard
print("PASS'")

##Don't kno what to do from here
#NOTE sam's way
#outlook.MailFoldersTree.JunkEmailTreeItem.select()
#outlook.MsoDockTopPane.RibbonToolBar.RibbonPane.pane.pane.RibbonPane.LowerRibbonPane.HomeGroup.FindGroup.draw_outline()
#outlook.MsoDockTopPane.RibbonToolBar.RibbonPane.pane.pane.RibbonPane.LowerRibbonPane.HomeGroup.FindGroup.FilterEmail.draw_outline()
#outlook.MsoDockTopPane.RibbonToolBar.RibbonPane.pane.pane.RibbonPane.LowerRibbonPane.HomeGroup.FindGroup.FilterEmail.press()
