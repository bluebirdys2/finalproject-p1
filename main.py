#   This file calls all function for a statics problem of a car going up a hill
#
#

#   Imports the modules needed to solve the problem and openpyxl
import functions
import openpyxl as pyxl

#   Defines the Excel file in which all the given data is stored
workbook1 = pyxl.load_workbook("data.xlsx")

print("program started")
#   Calls a function that returns the drag coefficient, cross sectional area, and name of car selected by the user
functions.main(workbook1)

