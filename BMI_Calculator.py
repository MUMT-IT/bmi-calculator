import math

weight =  int(input("Enter your weight(kg): "))
height =  int(input("Enter your height(cm): "))

height2 = math.pow(height/100,2)

BMI = weight/height2

print ("Your BMI is :", BMI)