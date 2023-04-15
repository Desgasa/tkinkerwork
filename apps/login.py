from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from apps.startwindow import *



class Login():
    def __init__(self):
        self.root = Tk()
        self.root.geometry("300x250")
        self.root.title("KepHelper")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        global username
        global password

        

        Label(self.root, text="UserName").place(x=40,y=40)
        Label(self.root,text='Password').place(x=40,y=100)
        username = Entry(self.root)
        username.place(x=40, y=65)
        password = Entry(self.root,show=('*'))
        password.place(x=40, y=125)
        Button(self.root,command=self.login,text="Login",height=1,width=7).place(x=170,y=160)
    def login(self):
        

    

        if(username.get().capitalize()=="Opereod" and password.get()=="123456a"):
            messagebox.showinfo("Successfully","Login is correct")
            self.root.destroy()
            myWindow2 = Window()
            myWindow2.start()
        elif(username.get()=="" and password.get()==""):
            messagebox.showerror("Error","Enter UserName and Password")
            return False
        elif(username.get() == "opereod" or password.get() != "123456a"):
            messagebox.showerror("Error","Enter correct password")
            password.delete(0,END)
            return False    
        elif(username.get() != "opereod" or password.get() != "123456a"):
            messagebox.showerror("Error","Enter correct login and password")
            username.delete(0,END)
            password.delete(0,END)
            return False  


    def on_closing(selt):
        if messagebox.askokcancel("Вийти?","Ви бажаєте закрити додаток?"):
            selt.root.destroy()
    
    
    def start(self):
           self.root.mainloop()

            