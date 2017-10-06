#TODO: 1. find out whether taking the wrapper from descendants() function and then referring to it is faster than just finding in the window.
#TODO: 2. Make code full-proof
#TODO: Naming convention, _01 at the end if has copy for file
#TODO: when selecting tags to tag something with, open "All Categories" because tag list shows most recently used tags, not necessarily the "Processed" tag will show up
#TODO: maybe when finding the outlook.exe find it via glog module
#TODO: UNINPORTANT: Aesthetics for label on Windows7 machines (width not long enough)
#NOTE: I will need to update the other files right after updating one file...
import time
import re
from pywinauto import Application
import pywinauto
import ctypes
import datetime as dt
import tkinter as tk
from tkinter import ttk
import os, errno
#class and function definitions
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

def enter_saveasListItem(filename,timeOut=1):
    try:#find the folder on the list, and select(highlight) it.
        saveas.window(title=filename, control_type='ListItem', visible_only=False).select()
    except pywinauto.findwindows.ElementNotFoundError:#if the folder cannot be found, make one with that name, and select(highlight) it.
        saveas.window(title='Create New Folder', control_type='Button').click()
        saveas.type_keys(filename+'{ENTER}', with_spaces=True)
    #send 'ENTER' to open the folder
    saveas.type_keys('{ENTER}')


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

class SimpleForm(EntryWithHelp):
    def __init__(self, master=None, preInputList=None, fieldNameToHelpMessage={"Default fieldname":"Default helpmessage for default fieldname"}):
        self.dateLabel = "Processing Date:"
        super().__init__(master, preInputList, fieldNameToHelpMessage)

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
            print(self.name_of_file)
            self.winfo_toplevel().destroy()

#more definitions
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

"""Select 'Junk Mail' tree item"""
try:
    o.window(title_re='AllOfToBePr0cessed.*', control_type='TreeItem').wait('exists ready', retry_interval=0.02).select()
except Exception as e:
    print("Finding \"Junk Email\" failed :[", e)
    raise e
    exit()

#Process mail
try:
    while True:#To process more than 1 mail till it gives an error, whereby no more mail are visible to be processed
        """Start up the GUI to ask for information from the email"""
        #Create objects for the mainloop to loop with
        root = tk.Tk(className=' Input the values')
        fieldNameToHelpMessage = {
            "System ID(s):"     :"If more than 1, seperate IDs with commas.",
            "User ID(s):"       :"If more than 1, seperate IDs with commas.",
            "Source of Request:":"Eg. \"IDMS {NUMBER}\", \"Email\", etc.",
            "Processing Date:"  :"Enter the date in YYYY-MM-DD format."
        }
        preInputList = ["", "", "Email", str(dt.date.today())]
        dlg = SimpleForm(master=root, preInputList=preInputList\
                            ,fieldNameToHelpMessage=fieldNameToHelpMessage)
        #Start the GUI
        dlg.mainloop()
        print(dlg.inputtedDate)
        """Save as '<PDF_NAME> -0' """
        #Open "Save As" window
        o.type_keys('%fp')
        o.PrintButton.wait('exists ready', retry_interval=0.02)#wait for some part of the window to be interactable(visible)
        o.type_keys('i')
        o.CutePDFWriterListItem.select()
        try:
            o.PrintButton.click_input()
        except TypeError:
            None
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
        save_file_dir = r'/Daily Requests/Robot UIA/{:04d}/{:04d}-{:02d}/{:04d}-{:02d}-{:02d}'.format(dlg.inputtedDate.year,\
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

        """Categorise mail as 'Processed' """
        #o.type_keys('%hztg') #for when window is not maximised
        o.type_keys("%hg")
        try:
            o.Processed.click_input()
        except TypeError:
            None

        """Run the rules and get the 'Processed' tagged mail outta the current folder"""
        #o.type_keys('{VK_ESCAPE}') #for it window is not maximised
        o.type_keys('%hrrl')
        rna = o.window(title='Rules and Alerts')
        rna.RunRulesNow.click()
        rnn = o.window(title='Run Rules Now')
        rnn.ProcessedCheckBox1.click_input(double=True)
        rnn.RunNow.click()
        rnn.close()
        rna.close()
except pywinauto.findwindows.ElementAmbiguousError:
    print("There was more than one element that was matched!")
except pywinauto.findwindows.ElementNotFoundError:
    print("No element could be found!")
except pywinauto.findwindows.WindowAmbiguousError:
    print("There was more than one window that matched!")
except pywinauto.findwindows.WindowNotFoundError:
    print("No window could be found!")
except pywinauto.findbestmatch.MatchError:
    print("ALL PASS!")
