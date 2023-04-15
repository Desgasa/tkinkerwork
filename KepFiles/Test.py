from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from docx import Document

from database.configdatabase import *





def replace_text_in_paragraph(paragraph, key, value):
            if key in paragraph.text:
                inline = paragraph.runs
                for item in inline:
                    if key in item.text:
                        item.text = item.text.replace(key, value)  

class test():
    def __init__(self):
        global Treatyid

        self.test = Tk()
        self.test.geometry("250x250")
        self.test.title("KepHelper")
        self.test.protocol("WM_DELETE_WINDOW", self.on_closing)
        Label(self.test,text='Treatyid').place(x=250,y=15)
        Treatyid = Entry(self.test,width=50)
        Treatyid.place(x=50, y=50)
        Button(self.test,text="Replace Text",command=self.search,height=3,width=10).place(x= 100,y=100)

        

    def search(self):
        mycursor = dbconn.cursor()
        mycursor.execute("SELECT ClientCode.value_text, CurrentDate.value_date FROM `ClientCode`, `CurrentDate` WHERE ClientCode.Treatyid  = %s AND CurrentDate.Treatyid = %s ; ",(Treatyid.get(),Treatyid.get(),))
        result = mycursor.fetchone()




        template_file_path = ''
        output_file_path = ''
        variables = {
        "{":"",
        "}":"",
        "ClientCode": result[0],
        "CurrentDate": result[1]
        }



        template_document = Document(template_file_path)

        for variable_key, variable_value in variables.items():
            for paragraph in template_document.paragraphs:
                replace_text_in_paragraph(paragraph, variable_key, variable_value)

            for table in template_document.tables:
                for col in table.columns:
                    for cell in col.cells:
                        for paragraph in cell.paragraphs:
                            replace_text_in_paragraph(paragraph, variable_key, variable_value)

        template_document.save(output_file_path)



    def on_closing(self):
        if messagebox.askokcancel("Вийти?","Ви бажаєте закрити додаток?"):
            self.test.destroy()
            
    
    def start(self):
           self.test.mainloop()
           

         
