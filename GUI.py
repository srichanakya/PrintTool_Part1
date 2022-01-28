import os
import shutil
import time
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from openpyxl import *
from FileLogic import logic

class GraphicalInterfaceClass:



    def __init__(self, logic:logic ):
        self.Flogic = logic
        self.COUNT = 0
        self.list_of_name = ["Purchase Order", "Purchase Agreement", "Delivery Statement", "Invoice"]
        self.list_of_entities = ["AU25", "US60"]
        self.window = Tk()
        self.window.geometry("550x180")
        self.window.title(string="Print tool")

        self.heading1 = Label(text="Source Path", font=("Arial 16 bold"))
        self.heading1.config(pady=10, padx=20)
        self.heading1.grid(column=0, row=0)

        self.sourcePath = Entry(width=30)
        self.sourcePath.insert(0,"/Users/srichanakyachowdary/Desktop/SourceFolder/doc1.pdf")
        self.sourcePath.grid(column=1, row=0, columnspan=2)

        self.heading2 = Label(text="Destination Path", font=("Arial 16 bold"))
        self.heading2.config(pady=10, padx=20)
        self.heading2.grid(column=0, row=1)

        self.destinationPath = Entry(width=30)
        self.destinationPath.insert(0,"/Users/srichanakyachowdary/Desktop/TargetDirectory")
        self.destinationPath.grid(column=1, row=1, columnspan=2)

        self.heading3 = Label(text="Select File Name and Entity", font=("arial 16 bold"))
        self.heading3.grid(column=0, row=2)
        self.heading3.config(pady=10,padx=20)
        self.filenames = ttk.Combobox(value=self.list_of_name, width=15)
        self.filenames.grid(column=1, row=2)
        self.filenames.current(0)
        self.countrynames = ttk.Combobox(value=self.list_of_entities, width=5)
        self.countrynames.grid(row=2, column=2)
        self.countrynames.current(0)

        self.excelName = Entry(width=15)
        self.excelName.insert(0,"Excel Name File")
        self.excelName.grid(column=0,row=3)
        #self.excelName.config(state=DISABLED)
        self.excelName.bind('<Enter>',self.Clicked)
        self.execute_btn = Button(text="Execute", command=self.StartExecution, width=10)
        self.execute_btn.grid(column=1, row=3)

        self.quit_btn2 = Button(text="Exit", command=self.window_close, width=10)
        self.quit_btn2.grid(column=2, row=3)

        click_btn = PhotoImage(file='info.png')
        self.info = Button(command=self.Get_Info,image=click_btn,width=10,height=10)
        self.info.grid(column=2,row=4)
        #self.window.overrideredirect(1)
        self.window.resizable(0,0)
        self.window.mainloop()

    def Clicked(self,event):
        # self.excelName.delete(0, END)
        if self.COUNT == 0:
            self.excelName.config(state=DISABLED)
            messagebox.showerror(message="Sorry field has been Disabled")
            self.COUNT += 1

    def StartExecution(self):
        self.Flogic.FileExecution(self.sourcePath.get(),self.destinationPath.get(),self.filenames.get(), self.countrynames.get())

    def Get_Info(self):
        newWindow = Tk()
        newWindow.title(string="Info")
        newWindow.geometry("220x60")
        label1 = Label(newWindow,text="Version 1.0.0 ",font=("Arial 10 bold"))
        # name1=Label(newWindow,text=": Print Tool",font=("Arial 10 bold"))
        label2 = Label(newWindow,text="Developed in Python ",font=("Times New Romn" ,10 ,"bold"))
        # name2 = Label(newWindow, text=": Python", font=("Arial 10 bold"))
        label3 = Label(newWindow, text="Developed by SRI CHANAKYA YENNANA", font=("Times New Romn",10 ,"bold"))
        # name3 = Label(newWindow, text=": Sri Chanakya Yennana", font=("Arial 10 bold"))
        label1.grid(column=0,row=0)
        # name1.grid(column=1,row=0)
        label2.grid(column=0, row=1)
        # name2.grid(column=1, row=1)
        #label3.config(pady=10, padx=2)
        label3.grid(column=0, row=2)
        # name3.grid(column=1, row=0)
        newWindow.resizable(0,0)
        # newWindow.overrideredirect(1)
        newWindow.mainloop()

    def window_close(self):
        Quit = messagebox.askokcancel(title="Quit",message="Do you want to close the application")
        if Quit:
            self.window.quit()


