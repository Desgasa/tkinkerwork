from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from windows.KepWindow import *
from windows.Additionalagreement import *


class Window():
    def __init__(self):
        self.startwindow = Tk()
        self.startwindow.geometry("400x400")
        self.startwindow.title("KepHelper")
        self.startwindow.protocol("WM_DELETE_WINDOW", self.on_closing)
        

        main_menu = Menu(self.startwindow)
        self.startwindow.configure(menu=main_menu)

        main_menu.add_command(label ="Депозит КЕП", command = self.KepWindow)
        main_menu.add_command(label ="Додаткова Угода", command = self.additional_agreement)
        main_menu.add_command(label="Вихід", command=self.on_closing)
    def KepWindow(self):
        KepWindows = root()
        KepWindows.start
    def additional_agreement(self):
        Additional_agreement = additionalagreement()
        Additional_agreement.start()  
    def on_closing(self):
        if messagebox.askokcancel("Вийти?","Ви бажаєте закрити додаток?"):
            self.startwindow.destroy()
    
    def start(self):
           self.startwindow.mainloop()       