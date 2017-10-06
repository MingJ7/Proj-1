#TODO: Fix create_tag_rule.
#TODO: search up usage of listToExtend.extend(list) instead of +
#TODO: when creating folders under main, use return from o.texts() to get email
#TODO: when color categories has too many tags, need to scroll down before click_input() and spacebar to select tag
import pywinauto
from pywinauto import Application

def print_descendants_of(window, x=-1):
    x += 1
    print("|  " * x + str(window))
    for child in window.children():
        print_descendants_of(child)

#def print_descendants_of(window, x=-1, searchFor_re=None, listOfMatches=[]):
#    x += 1
#    print("|  "*x + str(window))
#    for child in window.children():
#     listOfMatches += print_descendants_of(child, x=x, listofMatches=listofMatches, searchFor_re=searchFor_re)
#    #if search critera is given
#    if searchFor_re is not None:#probably don't need "is not None"
#        matches = re.search(searchFor_re, str(window))#need to test this statement, I forgot exactly what it returns if nothing is inside
#        if matches is not None:
#            listOfMatches += [window]
#    return listOfMatches
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

def create_new_folder_under_main(folderName):
    o.wait('ready', retry_interval=0.02)
    o.type_keys('%on')
    createNewFolder = o.window(title='Create New Folder', class_name='#32770')
    createNewFolder.Edit.type_keys(folderName, with_spaces=True)
    createNewFolder.window(title_re='.*@.*[.][cC][oO][mM].*', control_type='TreeItem', visible_only=False, found_index=0).select()
    createNewFolder.type_keys('{ENTER}')

def create_new_tag(tagName):
    o.wait('ready', retry_interval=0.02)
    o.type_keys('%hga')
    colorCategories = o.window(title='Color Categories', class_name='#32770')
    #create a new tag
    colorCategories.type_keys('%n')
    addNewCategory = colorCategories.window(title='Add New Category', class_name='#32770')
    addNewCategory.window(class_name='RichEdit20WPT').type_keys(tagName+'{ENTER}', with_spaces=True)
    tag = colorCategories.window(title=tagName, control_type='ListItem', visible_only=False)
    #uncheck the new tag and press "ENTER" key to accept and close window
    colorCategories.wait('ready', retry_interval=0.02)
    colorCategories.type_keys(' '+'{ENTER}', with_spaces=True)

def create_tag_rule():
    o.wait('ready', retry_interval=0.02)
    o.type_keys('%hrrl')
    rulesAndAlerts = o.window(title='Rules and Alerts', class_name='#32770', found_index=0)
    #create new rule
    rulesAndAlerts.type_keys('%n')
    rulesWizard = rulesAndAlerts.window(title='Rules Wizard', class_name='#32770')
    print_descendants_of(rulesWizard)
    rulesWizard.window(title='Apply rule on messages I receive', control_type='ListItem', visible_only=False, found_index=0).click_input()
    #activate "Next >"
    rulesWizard.type_keys('%n', with_spaces=True)
    print_descendants_of(rulesWizard)
    rulesWizard.window(title='assigned to category category', control_type='CheckBox', visible_only=False, found_index=0).click()
    #focus on "category" from "assigned to >category< category", ENTER to open
    rulesWizard.type_keys('{TAB}{ENTER}')
    colorCategories = rulesWizard.window(title='Color Categories', class_name='#32770')
    colorCategories.window(title='Processed', control_type='ListItem', visible_only=False).click_input()
    colorCategories.type_keys(' '+'{ENTER}', with_spaces=True)
    #activate "Next >"
    rulesWizard.type_keys('%n')
    rulesWizard.window(title='move it to the specified folder', control_type='CheckBox', visible_only=False, found_index=0).click()
    rulesWizard.type_keys('{TAB}{DOWN}{ENTER}')
    rulesAndAlertsUnderWizard = rulesWizard.window(title='Rules and Alerts', class_name='#32770')
    rulesAndAlertsUnderWizard.window(title='AlreadyPr0cessed', control_type='TreeItem', visible_only=False).select()
    rulesAndAlertsUnderWizard.type_keys('{ENTER}')
    rulesWizard.window(title='Finish', control_type='Button').click()
    microsoftOutlook = rulesWizard.window(title='Microsoft Outlook', class_name='#32770')
    microsoftOutlook.window(title='OK', control_type='Button').click()
    #activate "Apply"
    rulesAndAlerts.type_keys('%a')
    rulesAndAlerts.close()

OUTLOOK64_PATH = r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'
OUTLOOK32_PATH = r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'
"""Get an instance of Outlook running"""
outlook = Application(backend='uia')
try:
    connect_to_and_start_outlook(OUTLOOK64_PATH)
except pywinauto.application.AppStartError as e:
    print("Could not find 64 bit version of Outlook")
    connect_to_and_start_outlook(OUTLOOK32_PATH)

print("o_window texts: ", o.texts())
#create_new_folder_under_main('AllOfToBePr0cessed')
#create_new_folder_under_main('AlreadyPr0cessed')
#create_new_tag('Processed')
#create_tag_rule()