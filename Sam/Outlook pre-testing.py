import re
import time
import pywinauto
from pywinauto.application import Application
def saveasListItem(string):
    try:
        saveas.window(title=string, control_type='ListItem', visible_only=False).select()
    except pywinauto.findwindows.ElementNotFoundError:
        saveas.window(title='Create New Folder', control_type='Button').click()
        saveas.type_keys(string+'{ENTER}', with_spaces=True)
    saveas.type_keys('{ENTER}')
#outlookPath = r'C:\Program Files (x86)\Microsoft Office\root\Office16\outlook.exe'
#try:#if application fails to connect
#    app = Application(backend='uia').connect(path=outlookPath)
#except:#startup outlook
#    print("Application failed to connect")
#    app = Application(backend='uia').start(outlookPath)
#while 1:
#    try:
#        o = app.window(title_re='.*-.*- Outlook', visible_only=False).wait('exists')
#        o = app.window(title_re='.*-.*- Outlook', visible_only=False)
#        print("1")
#        break
#    except Exception as e:
#        print("Error at finding outlook main window!", e)
#        o = app.window(title_re='.*- Outlook', visible_only=False)
#        break
#print("Start-up complete!")
#toBeSaved = o.TableViewTable.Dataitem1.wait('exists ready', retry_interval=0.02) #Dataitem1 is the first dataitem(the first mail) in TableViewTable(the mail selector pane)
def connect_to_and_start_outlook(outlookPath):
    try:#if application fails to connect
        global outlook
        outlook = outlook.connect(path=outlookPath)
    except pywinauto.application.ProcessNotFoundError:#start up outlook
        print("Failed to find any instance of Outlook")
        outlook.start(outlookPath)
    global o
    o = outlook.window(title_re='.*-.*- Outlook', visible_only =False) #this is the mail window
    while True:
        try:#make sure the outlook window is ready
            o.restore()
            o.maximize()
            o.window(title_re='AllOfToBePr0cessed.*', control_type='TreeItem').wait('ready exists', retry_interval=0.02)
            print("Outlook is ready!")
            break
        except Exception as e:
            print("Error at starting:", type(e), e)

OUTLOOK64_PATH = r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'
OUTLOOK32_PATH = r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'

"""Get an instance of Outlook running"""
outlook = Application(backend='uia')
try:
    connect_to_and_start_outlook(OUTLOOK64_PATH)
except pywinauto.application.AppStartError as e:
    print("Could not find 64 bit version of Outlook")
    connect_to_and_start_outlook(OUTLOOK32_PATH)
#Open "Color Categories" to tag a mail
o.type_keys('%hga')
colorCategories = o.window(title='Color Categories', class_name='#32770')
colorCategories.window(title='Processed', control_type='ListItem', visible_only=False).wait('exists').set_focus()
colorCategories.type_keys(' '+'{ENTER}', with_spaces=True)

####textlist = ['14/2/2139']
####date = foundDate = ""
####for text in textlist:
####    foundDate = re.search(r'(?P<day>\d{1,2})/(?P<month>\d{1,2})/(?P<year>\d{4})', text)
####    if foundDate != None:
####        print("Source text:{}".format(text))
####        print("Successful match = {}".format(foundDate))
####        date += foundDate.group('day') + "-" + foundDate.group('month') + "-" + foundDate.group('year') + " "
####print("date = \"{}\"".format(date))
####print("foundDate =", foundDate)
####saveapp = Application(backend='uia').connect(title='Save As')
#####Start "Save As" part
####saveas = saveas1 = saveapp.window(title='Save As')
####saveas.set_focus()
####print("Focus set!")
####saveas.type_keys('%n'+date)#type in date
#####Start selecting 'd&ps (\\fsp\sp\ITD) (N:)'
####saveas.type_keys('%i%{DOWN}')
##saveaschildren = saveas.descendants(control_type='ListItem')
##Counter=0
##for child in saveaschildren:
##    print(child)
##    if r'd&ps (\\fsp\sp\ITD) (N:)' in repr(child) and 'ListItem' in repr(child):
##        print("\033[96mWrapper found!\033[0m")
##        Counter += 1
##        saveas1 = child
##saveas.type_keys('%i%{DOWN}')
##saveas.type_keys('%i%{DOWN}')
##print("opened?")
##try:#to try and select for the 2nd time. Known glitch.
##    saveas.SaveInList1.window(title=r'd&ps (\\fsp\sp\ITD) (N:)',control_type='ListItem').click_input()
##except pywinauto.findwindows.ElementNotFoundError:
##    print("First try failed")
##    saveas.type_keys('%i%{DOWN}')
##    saveas.SaveInList1.window(title=r'd&ps (\\fsp\sp\ITD) (N:)',control_type='ListItem').click_input()
##saveasListItem('ITSec')
##saveasListItem('Daily Requests')
#saveasListItem('Robot UIA test')
#saveasListItem(foundDate.group(3))
#saveasListItem(foundDate.group(2))
#saveasListItem(foundDate.group(1))

print("ALL PASS!", '#Counter', '#endTime-startTime')