#   This file calls all function for a statics problem of a car going up a hill
#
#

#   Imports the modules needed to solve the problem and openpyxl
import functions
import openpyxl as pyxl

#   Defines the Excel file in which all the given data is stored
workbook = pyxl.load_workbook("data.xlsx")
sheet = workbook["Problem Variables"]

#   Calls a function that returns the drag coefficient, cross sectional area, and name of car selected by the user
dragC,csA,name = functions.carChoice(sheet)

#   Returns the gear ratios from the excel file 
gears = functions.getGears(sheet)

#   Function that returns all other varibles in the excel for use by later functions
v,tslope,wbase,radius,rollres,hA,fdrive,teff,weight,airden,dratio,centerg = functions.getothers(sheet)

#   Function that returns angularvex and torquex
angularvex,torquex = functions.getdyno(sheet)

#   Function that calculates roadload given rollres,weight,tslope,airden,dragC,csA, and v
roadLoad = functions.calcRoadLoad(rollres,weight,tslope,airden,dragC,csA,v)

#   Function that calculates traction,torqued,acceleration, and angularve
traction,torqued,acceleration,angularve = functions.traction(v,radius,dratio,teff,fdrive,gears,angularvex,torquex,roadLoad,weight)

#   Calculates and returns hp
hp = functions.hp(angularve,torqued)

#   Outputs graphs
functions.graphs(angularve, torqued,name)