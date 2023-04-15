from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from KepFiles.DepTepmlUsual import * 
from KepFiles.DepTempEarDes import * 
from KepFiles.DepTem3x20 import * 
from KepFiles.DepTem3x45 import *
from KepFiles.Test import *

class Window():
    def __init__(self):
        self.startwindow = Tk()
        self.startwindow.geometry("200x250")
        self.startwindow.title("KepHelper")
        self.startwindow.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        main_frame = Frame(self.startwindow)
        main_frame.pack(fill=BOTH,expand=True)

        main_menu = Menu(self.startwindow)
        self.startwindow.configure(menu=main_menu)

        first_item = Menu(main_menu, tearoff=0)
    
        second_item = Menu(main_menu, tearoff=0)

        main_menu.add_cascade(label ="Files", menu= first_item)
        main_menu.add_cascade(label ="Помощь", menu= second_item)
        
        first_item.add_command(label = "DepositTemplateUsual",command= self.deptemplusual)
        first_item.add_command(label = "DepositTemplateEarlyDessolution",command=self.deptemped)
        first_item.add_command(label = "DepositTemplate3x20", command= self.deptem3x20)
        first_item.add_command(label = "DepositTemplate3x45",command= self.deptem3x45)
        first_item.add_command(label = "Test",command= self.test)
        first_item.add_separator()
        first_item.add_command(label = "exit", command= self.on_closing)
        

    def deptemplusual(self):
        DepTempU = DepTU()
        DepTempU.start()

    def deptemped(self):
        DepTed = DepTED()
        DepTed.start()
    
    def deptem3x20(self):
        DepT3x20 = DepTem3x20()
        DepT3x20.start()
        
    def deptem3x45(self):
        DepT3x45 = DepTem3x45()
        DepT3x45.start()

    def test(self):
        Test = test()
        Test.start()



    def on_closing(self):
        if messagebox.askokcancel("Вийти?","Ви бажаєте закрити додаток?"):
            self.startwindow.destroy()
    
    def start(self):
           self.startwindow.mainloop()   