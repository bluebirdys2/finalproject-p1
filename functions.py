#This module contatins all function for a statics problem of a car going up a hill
#contains input, calculation, and output functions
import numpy as np
import openpyxl as pyxl
import matplotlib.pyplot as plt
from numpy.polynomial import Polynomial as poly
workbook= pyxl.load_workbook("data.xlsx")
sheetmain = workbook["Problem Variables"]

def carChoice(sheet):
    """
    This function gives the user an option for which car they want
    and then returns the drag coefficient and area of front bumbper
    Args:
    sheet= needs a sheet to get all value
    Returns:
    dragC= Drag coefficient
    area= area of front hood
    
    
    """
    option=0 #initiales loop
    #prompts user for what car they want data on
    print("**********************************************")
    print("1. {} {}\n2. {} {}\n3. {} {}\n4.{} {}\n5. {} {}".format(sheet["B26"].value,sheet["A26"].value,sheet["B27"].value,sheet["A27"].value, sheet["B28"].value,sheet["A28"].value,sheet["B29"].value,sheet["A29"].value,sheet["B30"].value,sheet["A30"].value))
    print("**********************************************")
    while(option!=5 and option!=4 and option!=3 and option!=2 and option!=1): #loop is active while 'option' value is invalid

        option=int(input("Select the car choice options (1-5) "))
        match option:
            case 1: #Ford Escape Hybrid
                dragC=sheet["C26"].value
                area=sheet["D26"].value
                name=sheet["B26"].value
            case 2: #Toyota Camery
                dragC=sheet["C27"].value
                area=sheet["D27"].value
                name=sheet["B27"].value
            case 3: #Smart Roadster Coupe
                dragC=sheet["C28"].value
                area=sheet["D28"].value
                name=sheet["B28"].value
            case 4: #Toyota Prius
                dragC=sheet["C29"].value
                area=sheet["D29"].value
                name=sheet["B29"].value
            case 5: #Chevrolet Volt
                dragC=sheet["C30"].value
                area=sheet["D30"].value
                name=sheet["B30"].value
            
            case _:
            #HR Friendly error message
                print("ERROR! Selection is invalid\n")
    return dragC,area,name
def getGears(sheet):
    """
    This function takes in the spreadsheet and then uses the sheet to 
    get the different gear ratios, then checks to make sure all values are positive 
    if a value is negative the program prompts the user to change the value to a new value
    
    args:
    sheet: takes in an excell sheet
    returns:
    gear1: 1st gear, gear ratio
    gear2: 2nd gear, gear ratio
    gear3: 3rd gear, gear ratio
    gear4: 4th gear, gear ratio
    gear5: 5th gear, gear ratio
    gear6: 6th gear, gear ratio
    """
    #Checks to see if gear value is invalid and forces user to correct it if needed
    gears=[]
    for i in range(12,24):
        while(sheet["B{}".format(i)].value<0): 
            sheet["B{}".format(i)]=float(input("The value of your {} is negative, input a new value: ".format(sheet["A{}".format(i)])))
        gears.append(sheet["B{}".format(i)])
    gears=np.array(gears)
    return gears
def getothers(sheet):
    #Checks to see if other values, mostly constants, are invalid
    val=[]
    for i in range(12,24): 
        while(sheet["C{}".format(i)].value<0):
            sheet["C{}".format(i)]=float(input("The value of your {} is negative, input a new value: ".format(sheet["A{}".format(i)])))
 
    speed=sheet["C12"].value /3.6
    tslope=sheet["C13"].value
    wbase=sheet["C14"].value
    radius=sheet["C15"].value
    rollre=sheet["C16"].value
    hA=sheet["C17"].value
    fdrive=sheet["C18"].value
    teff=sheet["C19"].value
    weight=sheet["C20"].value
    airden=sheet["C21"].value
    dratio=sheet["C22"].value
    centerg=sheet["C23"].value
    return speed,tslope,wbase,radius,rollre,hA,fdrive,teff,weight,airden,dratio,centerg
#gets separate arrays for both 'anglarvex' and 'torquex' from the Dyno
def getdyno(sheet): 
    angularvex=sheet["A4":"A15"].value
    torquex=sheet["B4":"B15"].value
    return angularvex,torquex
#road load calculations
def calcRoadLoad(rollres,weight,tslope,airden,dragC,csA,v): 
    airres=.5*airden*dragC*csA*v**2
    rRoll=rollres*weight*np.cos((np.atan(tslope/100)))
    roadLoad=rRoll+weight*np.sin((np.atan(tslope/100)))+airres
    return roadLoad

#Calculates traction
def traction(v,radius,dratio,teff,fdrive,gears,angularvex,torquex,roadLoad,weight):
    
    angularve=gears*dratio*v/radius
    finaldrivE=fdrive/100
    fx=poly.fit(angularvex,torquex,2)
    te=fx(angularve)
    torqued=finaldrivE*dratio*(teff/100)*gears*te
    traction=torqued/radius
    acceleration=(((-1*(traction-roadLoad))*9.81)/weight)
    return traction,torqued,acceleration
#Calculates front and rear axle load
def loads(weight,tslope,wbase,centerg):
    frontload=(weight*np.cos(tslope)*centerg)/wbase
    rearload=(weight*np.cos(tslope))-frontload
    return rearload,frontload
#Calculates horsepower
def hp(angularve,torque):
    hp=torque*angularve/5252
    return hp
#Graph configurations    
def graphs(angularve, torque,name):
    fx=poly.fit(angularve,torque,2)
    x=np.linspace(min(angularve),max(angularve),6000)
    fy=fx(x)
    plt.subplot(2,2,1)
    plt.plot(angularve,torque,"ro")
    plt.plot(x,fy,"r-")
    plt.title("Experimental dyno data and fit")
    plt.ylabel("engine torque(n m)")
    plt.xlabel("angular velocity(rpm)")
    plt.legend([name])
