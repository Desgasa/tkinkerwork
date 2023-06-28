from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from docx import Document

from database.configdatabaseserver import *

class additionalagreement():
    def __init__(self):
        global Treatyid
        global Filename
        self.root = Tk()
        self.root.geometry("300x250")
        self.root.title("KepHelper")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        Label(self.root, text = 'FileName').place(x=87,y=40)
        Filename = Entry(self.root, width = 30)
        Filename.place(x=87, y=65)
        Label(self.root,text='Treatyid').place(x=87,y=100)
        Treatyid = Entry(self.root,width=30)
        Treatyid.place(x=87, y=125)
        replacebutton = ttk.Button(self.root,text="Replace Text")
        replacebutton.pack(
            
            expand= True,
            ipadx = 2, 
            ipady = 7,
            anchor = "s",
            pady = 30
            )
        

    def on_closing(self):
        if messagebox.askokcancel("Вийти?","Ви бажаєте закрити додаток?"):
            self.root.destroy()
            
    
    def start(self):
           self.root.mainloop()
               