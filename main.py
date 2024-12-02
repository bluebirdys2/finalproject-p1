#   This file calls all function for a statics problem of a car going up a hill
#
#

#   Imports the modules needed to solve the problem and openpyxl
import functions
import openpyxl as pyxl

#   Defines the Excel file in which all the given data is stored
workbook = pyxl.load_workbook("data.xlsx")


#   Calls a function that returns the drag coefficient, cross sectional area, and name of car selected by the user
dragC,csA,name,sheet = functions.carChoice(workbook)

#   Returns the gear ratios from the excel file 
gears = functions.getGears(sheet)

#   Function that returns all other varibles in the excel for use by later functions
v,tslope,wbase,radius,rollres,hA,fdrive,teff,weight,airden,dratio,centerg = functions.getothers(sheet)

#   Function that returns angularvex and torquex
angularvex,torquex = functions.getdyno(sheet)

#   Function that calculates roadload given rollres,weight,tslope,airden,dragC,csA, and v
traction,torqued,acceleration, angularve, roadLoad,rearload,frontload,frontloads,rearloads=functions.doMath(rollres,weight,tslope,airden,dragC,csA,v,radius,dratio,teff,fdrive,gears,angularvex,torquex,name,wbase,centerg,hA)

#   Outputs graphs
functions.graphs(angularve, torqued,name)