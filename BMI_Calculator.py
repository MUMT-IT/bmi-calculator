import math
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook
import datetime
from tkinter import*
from tkinter.messagebox import askyesno

root = Tk()

root.title("BMI calculator")
root.geometry("500x500")
nameInput = StringVar()
weightInput = IntVar()
heightInput = IntVar()
userNameLabel = Label(text = "ชื่อผู้ใช้").grid(row=0,column=0)
userNameEntry = Entry(width = 30,textvariable=nameInput).grid(row=0,column=1)
weightLabel = Label(text = "น้ำหนัก(กก.)").grid(row=1,column=0)
weightEntry = Entry(width = 30,textvariable=weightInput).grid(row=1,column=1)
heightLabel = Label(text="ส่วนสูง(ซม.)").grid(row=2,column=0)
heightEntry = Entry(width = 30,textvariable=heightInput).grid(row=2,column=1)


#------------------------Prepare database-----------------------

filePath = "C:/Users/Peerapat/Desktop/bmi-calculator/Database.xlsx"
wb = openpyxl.load_workbook(filePath)


#---------------------------------------------------------------

#--------------Check for existing name in database--------------
def checkName() :
	sheetNames = wb.sheetnames
	existingName = 0
	name = nameInput.get()
	if len(name) == 0:
		messagebox.showerror("Error","กรุณากรอกชิ่อผู้ใช้งาน")
	else :	
		for x in range(len(sheetNames)):
			if name == sheetNames[x] :
				existingName +=1

		if existingName > 0 :
			ws = wb[name]
		else :
			confirm = askyesno(title = "Confirmation",message=f"ไม่พบชื่อผู้ใช้งาน {name} ต้องการสร้างรายชื่อใหม่หรือไม่")
			if confirm :
				ws = wb.create_sheet(name)
				ws['A1'] = name
				ws['A2'] = 'Time'
				ws['B2'] = 'Weight'
				ws['C2'] = 'Height'
				ws['D2'] = 'BMI'
				wb.save(filePath)
			else :
				pass
		return ws
#---------------------------------------------------------------

def calculateBMI():
	height = heightInput.get()
	weight = weightInput.get()

	
	ws = checkName()
	if weight == 0 :
		messagebox.showerror("Error","กรุณากรอกข้อมูลน้ำหนัก")
	if type(weight) != int :
		messagebox.showerror("Error","กรุณากรอกข้อมูลน้ำหนักเป็นตัวเลข")
	if height == 0 :
		messagebox.showerror("Error","กรุณากรอกข้อมูลความสูง")
	if type(height) != int :
		messagebox.showerror("Error","กรุณากรอกข้อมูลความสูงเป็นตัวเลข")

	if weight != 0 and type(weight) == int and height != 0 and type(height) == int:
		height2 = math.pow(height/100, 2)
		BMI = round(weight/height2,2)
		resultLabel = Label(text=f"BMI ของคุณคือ {BMI}").grid(row=4,column=0)
		maxRow = ws.max_row+1
		date = datetime.datetime.now()
		ws.cell(maxRow,1).value = date.strftime('%d/%m/%Y %H:%M')
		ws.cell(maxRow,2).value = weight
		ws.cell(maxRow,3).value = height
		ws.cell(maxRow,4).value = BMI
		wb.save(filePath)

def showData():
	ws = checkName()
	maxRow = ws.max_row+1
	#headerLabel1 = Label(text="เวลา").grid(row=4,column=1)
	#headerLabel2 = Label(text="น้ำหนัก").grid(row=4,column=2)
	#headerLabel3 = Label(text="ส่วนสูง").grid(row=4,column=3)
	#headerLabel4 = Label(text="BMI").grid(row=4,column=4)
	for row in range(2,maxRow):
		for col in range(1,5):
			dataLabel = Label(text=ws.cell(row,col).value).grid(row=row+3,column=col-1)
			print(type(dataLabel))
	clearBtn = Button(root,text="ล้างข้อมูล",command= lambda : clearData ).grid(row=3,column=2)


def clearData():
	ws = checkName()
	maxRow = ws.max_row+1
	ws.delete_rows(3,maxRow)
	wb.save(filePath)


#testLabel = Label(root,text="result").grid(row=3,column=2)
#testBtn = Button(root,text="Clear",command = lambda : testLabel.destroy()).grid(row=3,column=3)
calBtn = Button(root,text="คำนวณ BMI",command=calculateBMI).grid(row=3,column=0)
showBtn = Button(root,text="เรียกดูข้อมูล",command=showData).grid(row=3,column=1)



root.mainloop()