from pywinauto.application import Application
import re
from sys import getsizeof
import ctypes
from datetime import date
from tkinter import *
outlookPath = r'C:\Program Files (x86)\Microsoft Office\root\Office16\outlook.exe'
##  Styles:
##  0 : OK
##  1 : OK | Cancel
##  2 : Abort | Retry | Ignore
##  3 : Yes | No | Cancel
##  4 : Yes | No
##  5 : Retry | No
##  6 : Cancel | Try Again | Continue
#ctypes.windll.user32.MessageBoxW(0, "Messageeee", "Window Title!", 1(style))
"""Start up the GUI to ask for information from the email"""
receivedDate = None
dateLabel ='Processing Date:\n[YYYY-MM-DD]'
fields = 'System ID:', 'User ID:', 'Source of Request:', dateLabel

class AssignmentNotFound(Exception):
    None
def check_and_quit(event):
    try:
        for name in entries:
            print(entries[name].get())
            if not entries[name].get().replace(' ',''):
                print("Idiot didn't input a field!")
                raise AssignmentNotFound
        receivedDate = date(*[int(partOfDate) for partOfDate in entries[dateLabel].get().split('-')])
    except AssignmentNotFound:
        ctypes.windll.user32.MessageBoxW(0, "Do not leave fields blank!", "ERROR:", 0)
    except (ValueError, TypeError):
        ctypes.windll.user32.MessageBoxW(0, "Please input a valid date!", "ERROR:", 0)
    else:
        dlg.quit()

def insert_rows(root, fields):
    entries={}
    #making the entry fields
    for field in fields:
        row = Frame(root, relief=SUNKEN)
        label = Label(row, text=field, width=15, anchor='nw')
        entry = Entry(row)
        row.pack(pady=5, padx=5, side=TOP, fill=X, expand=1)
        label.pack(side=LEFT)
        entry.pack(side=RIGHT, padx=5, fill=X, expand=1)
        entries[field] = entry
    return entries

dlg = Tk(className='Input the values')#Dialog:dlg
dlg.geometry('500x300')

entries = insert_rows(dlg, fields)
entries[dateLabel].insert(0, str(date.today()))
okay = Button(dlg, text='Okay')
okay.pack(pady=5, side=RIGHT, fill=BOTH, expand=1)

okay.bind('<Button-1>',check_and_quit)
dlg.bind('<Return>', check_and_quit)

dlg.mainloop()

print("_".join([(entries[name].get().lstrip(' ')) for name in entries]))
