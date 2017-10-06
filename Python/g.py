#import ctypes  # An included library with Python install.   
#ctypes.windll.user32.MessageBoxW(0, "Your text", "Your title", 1)


from tkinter import *
import tkinter.ttk
class App:
  def __init__(self, master):
    frame = Frame(master)
    frame.pack()
    self.button = Button(frame, 
                         text="QUIT", fg="red",
                         command=master.quit)
    self.button.pack(side=LEFT)
    self.slogan = Button(frame,
                         text="Hello",
                         command=self.write_slogan)
    self.slogan.pack(side=LEFT)
  def write_slogan(self):
    print("Tkinter is easy to use!")

root = Tk()
app = App(root)
root.mainloop()

list1=['ASDFGH','NAMES AKAK','THE GRAND STAND-OFF','Slowing Samuel\'s brain:\n\033[1m———THE DOCUMENTARY———\033[0m','Ming Jiu gets the worm: The biography']

class get_data:
    def __init__(self, master,list1):
        self.e={}
        self.v = IntVar()
        self.v.set(1)
        self.label1 = Label(master, text="Type of Tasks to Delegate:").grid(row=0)
        self.label2 = Label(master, text="Delegate to:").grid(row=1)
        
        self.e[1] = tkinter.ttk.Combobox(master,width = 30, justify="left", state ="readonly", values= list1)
        self.e[1].set("Windows (ITD Only)")
        for i, item  in enumerate(list1,start=2):
            self.e[i] = Radiobutton(master, text=item, variable=self.v, value=i, justify="center")
        
        self.e[1].grid(row=0, column=1)
        for i, item in enumerate(list1,start=2):
            self.e[i].grid(row=i, column=1, sticky="w")
        
        self.button1 = Button(master, text='Cancel', command=quit).grid(row=len(list1)+2, column=0, sticky=W, pady=4)
        self.button2 = Button(master, text='Show', command=self.show_entry_fields).grid(row=len(list1)+2, column=1, sticky=W, pady=4)
    def show_entry_fields(self):
        print("First Name: %s\nLast Name: %s" % (self.e[1].get(), self.v.get()))
master = Tk()
dd = get_data(master,list1)
master.mainloop()

"""
commands:
quit - quits the program
Tk().destroy - closes the window
Tk().quit - continues the program without closing the window
"""