import selenium
from selenium import webdriver
import time
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pywinauto.application import Application
import datetime as dt
import pywinauto
import tkinter as tk
import tkinter.ttk as ttk
from tkinter.messagebox import showwarning
root = tk.Tk(className=" RPA")
frame = ttk.Frame(root)
frame.pack(fill=tk.BOTH, expand=1)
printAsPDFButtonStyle = ttk.Style()
quitButtonStyle = ttk.Style()
printAsPDFButtonStyle.configure("printAsPDFButtonStyle.TButton", font=('Arial Black', 30))
quitButtonStyle.configure("quitButtonStyle.TButton", font=('Arial Black', 30), foreground='red')
printAsPDFButton = ttk.Button(style="printAsPDFButtonStyle.TButton", text="Print as PDF", command=frame.winfo_toplevel().destroy)
quitButton = ttk.Button(style="quitButtonStyle.TButton", text="Quit!", command=lambda:(_ for _ in ()).throw(SystemExit))
printAsPDFButton.pack(side=tk.TOP, fill=tk.BOTH, expand=1)
quitButton.pack(side=tk.TOP, fill=tk.BOTH, expand=1)
#Push the toplevel widget the widget is in to the front
self.winfo_toplevel().attributes("-topmost", True)
#Force focus onto the toplevel widget
self.winfo_toplevel().focus_force()
frame.mainloop()
print("Script continues!")