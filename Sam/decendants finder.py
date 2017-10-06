import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pywinauto
from pywinauto.application import Application
import datetime as dt
import time
import re
import tkinter as tk
import tkinter.ttk as ttk
from selenium.webdriver.common.keys import Keys

def print_descendants_of(window, x=-1):
    x += 1
    print("|  "*x + str(window))
    for child in window.children():
        print_descendants_of(child, x)

outlookPath = r'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'

"""Get an instance of Outlook running"""
outlook = Application(backend='uia')
try:#if application fails to connect
    outlook = outlook.connect(path=outlookPath)
except pywinauto.application.ProcessNotFoundError:#start up outlook
    print("Failed to find any instance of Outlook")
    outlook.start(outlookPath)
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

print_descendants_of(o)