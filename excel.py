pip install openpyxl
import openpyxl
import os
form openpyxl import Workbook

dir(openpyxl) #open the dir
wb  = Workbook()   #wb = workbook
sheet = wb.active #new sheet make comment
sheet
sheet["A1"] = "name" #what ever u enter etc...
sheet["B1"] = "age" #what ever u enter etc...
wb.save(filename = "test.xlsx" #file name enter i am enter test.xlsx u enter etc...
os.getcwd() #find the dir comment
wb.save(filename = "C:\\download\\test.xlsx") # dir path name c diver /download folder


