from pywinauto.application import Application
import datetime as dt
import pywinauto
def enter_saveasListItem(filename,timeOut=1):
    try:#find the folder on the list, and select(highlight) it.
        saveas.window(title=filename, control_type='ListItem', visible_only=False).select()
    except pywinauto.findwindows.ElementNotFoundError:#if the folder cannot be found, make one with that name, and select(highlight) it.
        saveas.window(title='Create New Folder', control_type='Button').click()
        saveas.type_keys(filename+'{ENTER}', with_spaces=True)
    #send 'ENTER' to open the folder
    saveas.type_keys('{ENTER}')

IE_DIR = r'C:\Program Files\Internet Explorer\iexplore.exe'
ie = Application(backend='uia').connect(path=IE_DIR)
ieTab = ie.window(title_re='.*- Internet Explorer')
for child in ieTab.descendants():
    print(child)
prt =ieTab.window(title='Print', class_name='#32770')#class name refers to a dialog
prt.window(title='CutePDF Writer', control_type='ListItem', visible_only=False).select()
prt.window(title='Print', control_type='Button').click()

inputtedDate = dt.date.today()
saveapp = Application(backend='uia')
saveas = Application(backend='uia')
print(saveapp)
while True:
    try:
        saveapp = saveapp.connect(title='Save As')
        print("Connected.")
        break
    except:
        print("Failed to conenct to \"Save As\". Trying again!")
saveas = saveapp.window(title='Save As')
saveas.wait('enabled', retry_interval=0.02)
print("Sees the Save As window")

#Type PDF name into the editing box
saveas.type_keys("%n")
saveas.type_keys('test this shite'+" -0", with_spaces=True)
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
enter_saveasListItem('ITSec', 5)
enter_saveasListItem('Daily Requests')
enter_saveasListItem('Robot UIA')
enter_saveasListItem("{:04d}".format(inputtedDate.year))
enter_saveasListItem("{:04d}-{:02d}".format(inputtedDate.year,inputtedDate.month,))
enter_saveasListItem("{:04d}-{:02d}-{:02d}".format(inputtedDate.year,inputtedDate.month, inputtedDate.day))

#Try to save the file
saveas.type_keys("%s")
i = 1
while True:
    try: #incase the "Confirm Save As" window pops up, rename the file with a numerical behind.
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