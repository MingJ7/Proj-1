from pywinauto.application import Application
import re
from sys import getsizeof
import ctypes
import datetime as dt
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox


class EntryWithHelp(ttk.Frame):
    class AssignmentNotFound(Exception):
        None

    def __init__(self, master=None, preInputList=None, fieldNameToHelpMessage={"Default fieldname":"Default helpmessage for default fieldname"}):
        super().__init__(master)
        #to suggest input for the user
        self.preInputList = preInputList
        self.fNTHM = fieldNameToHelpMessage
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
            helpMessage = "(HelpBox:)I'm here to help!"
        if event.widget == self.okay:
            helpMessage = "Save the information and close this window."
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
            self.winfo_toplevel().destroy()


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