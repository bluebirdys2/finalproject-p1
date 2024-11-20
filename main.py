#   Main
#
#

import functions.py
import openpyxl as pyxl

workbook= pyxl.load_workbook("data.xlsx")
sheet = workbook["Problem Variables"]

dragC,area,name = functions.carChoice(sheet)

speed,tslope,wbase,radius,rollre,hA,fdrive,teff,weight,airden,dratio,centerg = functions.getGears(sheet)

angularvex,torquex = functions.getdyno(sheet)

roadLoad = functions.calcRoadLoad(rollres,weight,tslope,airden,dragC,csA,v)

traction,torqued,acceleration = functions.traction(v,radius,dratio,teff,fdrive,gears,angularvex,torquex,roadLoad,weight)

hp = functions.hp(angularve,torque)

functions.graphs(angularve, torque,name)