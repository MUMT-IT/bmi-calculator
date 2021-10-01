import math
import openpyxl
from openpyxl import Workbook
import datetime

filePath = "C:/Users/Peerapat/Desktop/bmi-calculator/Database.xlsx"

wb = openpyxl.load_workbook(filePath)


print(type(wb.sheetnames))

sheetNames = wb.sheetnames

name = input("Enter your name: ")


enteredName = 0

for x in range(len(sheetNames)):

	if name == sheetNames[x] :
		enteredName +=1

if enteredName > 0 :
	ws = wb[name]
else :
	confirm = input("Name not found. Do you want to enter new name ? (Y/N)")
	if confirm == 'Y' :
		ws = wb.create_sheet(name)
		ws['A1'] = name
		ws['A2'] = 'Time'
		ws['B2'] = 'Weight'
		ws['C2'] = 'Height'
		ws['D2'] = 'BMI'
	else :
		exit()

weight = int(input("Enter your weight(kg): "))
height = int(input("Enter your height(cm): "))
height2 = math.pow(height/100, 2)
BMI = weight/height2

maxRow = ws.max_row+1

ws.cell(maxRow,1).value = datetime.datetime.now()
ws.cell(maxRow,2).value = weight
ws.cell(maxRow,3).value = height
ws.cell(maxRow,4).value = BMI


wb.save(filePath)

print("Your BMI is :", round(BMI, 2))
