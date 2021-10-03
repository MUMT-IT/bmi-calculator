import math
import openpyxl
from openpyxl import Workbook
import datetime
from tkinter import*
from tkinter import messagebox
from tkinter.messagebox import askyesno

def delete():
    myentry.delete(0, 'end')
 
root = Tk()
root.geometry('180x120')
entry = IntVar() 
myentry = Entry(root, width = 20,textvariable=entry)
myentry.pack(pady = 5)
 
mybutton = Button(root, text = "Delete", command = delete)
mybutton.pack(pady = 5)

Label(root,text = entry).pack(pady=5)
root.mainloop()