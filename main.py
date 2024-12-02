#   This file calls all function for a statics problem of a car going up a hill
#
#

#   Imports the modules needed to solve the problem and openpyxl
import functions
import openpyxl as pyxl

#   Defines the Excel file in which all the given data is stored
workbook1 = pyxl.load_workbook("data.xlsx")


#   Calls a function that returns the drag coefficient, cross sectional area, and name of car selected by the user
dragC,csA,name,count,sheet = functions.carChoice(workbook1)

#   Returns the gear ratios from the excel file 
gears = functions.getGears(sheet)

#   Function that returns all other varibles in the excel for use by later functions
v,tslope,wbase,radius,rollres,hA,fdrive,teff,weight,airden,dratio,centerg = functions.getothers(sheet)

#   Function that returns angularvex and torquex
if(name!="cr-28 var"):
    dynosheet=workbook1["Dynamometer"]
else:
    dynosheet=workbook1["cr-dyno"]
angularvex,torquex = functions.getdyno(dynosheet)

#   Function that calculates roadload given rollres,weight,tslope,airden,dragC,csA, and v
traction,torqued,acceleration, angularve, roadLoad,rearload,frontload,frontloads,rearloads,te=functions.doMath(rollres,weight,tslope,airden,dragC,csA,v,radius,dratio,teff,fdrive,gears,angularvex,torquex,name,wbase,centerg,hA)

#   Outputs graphs
#functions.graphs(angularve, torqued,name)
workbook2=pyxl.load_workbook("finalPResults.xlsx")
newsheet=functions.makeres(count,sheet,dragC,csA)
functions.outputs(gears,angularve,te,acceleration,traction,roadLoad,newsheet,name,rearloads,frontloads,frontload,rearload)