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
    count=0
    option=0
    print("**********************************************")
    print("1. {} {}\n2. {} {}\n3. {} {}\n4.{} {}\n5. {} {}".format(sheet["B26"].value,sheet["A26"].value,sheet["B27"].value,sheet["A27"].value, sheet["B28"].value,sheet["A28"].value,sheet["B29"].value,sheet["A29"].value,sheet["B30"].value,sheet["A30"].value))
    print("**********************************************")
    while(option!= 6 and option!=5 and option!=4 and option!=3 and option!=2 and option!=1):

        option=int(input("Select the car choice options (1-5) "))
        match option:
            case 1:
                dragC=sheet["C26"].value
                area=sheet["D26"].value
                name=sheet["B26"].value
                count+=1

            case 2:
                dragC=sheet["C27"].value
                area=sheet["D27"].value
                name=sheet["B27"].value
                count+=1
            case 3:
                dragC=sheet["C28"].value
                area=sheet["D28"].value
                name=sheet["B28"].value
                count+=1
            case 4:
                dragC=sheet["C29"].value
                area=sheet["D29"].value
                name=sheet["B29"].value
                count+=1
            case 5:
                dragC=sheet["C30"].value
                area=sheet["D30"].value
                name=sheet["B30"].value
                count+=1
            
            case 6: 
                dragC=sheet["C31"].value
                area=sheet["D31"].value
                name=sheet["B31"].value
                count+=1

            case _:
            #HR Friendly error message
                print("ERROR! Selection is invalid\n")
    return dragC,area,name,count
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
    gears=[]
    for i in range(12,24):
        while(sheet["B{}".format(i)].value<0):
            sheet["B{}".format(i)]=float(input("The value of your {} is negative, input a new value: ".format(sheet["A{}".format(i)])))
        gears.append(sheet["B{}".format(i)])
    gears=np.array(gears)
    return gears
def getothers(sheet):
    
    for i in range(12,24):
        while(sheet["C{}".format(i)].value<0):
            sheet["C{}".format(i)]=float(input("The value of your {} is negative, input a new value: ".format(sheet["A{}".format(i)])))
 
    speed=sheet["C12"].value /3.6
    tslope=sheet["C13"].value/100
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
def getdyno(sheet):
    angularvex=sheet["A4":"A15"].value
    torquex=sheet["B4":"B15"].value
    return angularvex,torquex
"""def calcRoadLoad(rollres,weight,tslope,airden,dragC,csA,v):
    airres=.5*airden*dragC*csA*v**2
    rRoll=rollres*weight*np.cos((np.atan(tslope/100)))
    roadLoad=rRoll+weight*np.sin((np.atan(tslope/100)))+airres
    return roadLoad
"""
def doMath(rollres,weight,tslope,airden,dragC,csA,v,radius,dratio,teff,fdrive,gears,angularvex,torquex,roadLoad):
    airres=.5*airden*dragC*csA*v**2
    rRoll=rollres*weight*np.cos((np.atan(tslope)))
    roadLoad=rRoll+weight*np.sin((np.atan(tslope)))+airres



    angularve=gears*dratio*v/radius
    finaldrivE=fdrive/100
    fx=poly.fit(angularvex,torquex,4)
    te=fx(angularve)
    torqued=finaldrivE*dratio*(teff/100)*gears*te
    traction=torqued/radius
    acceleration=((((traction-roadLoad))*9.81)/weight)
    return traction,torqued,acceleration, angularve, roadLoad

def loads(weight,tslope,wbase,centerg,airres,hA):
    Wperp=weight*np.cos(tslope)
    
    
    frontload=(airres*0.6)+(weight/9.81)*hA-(wbase-centerg)*Wperp+weight*np.sin(tslope)*hA
    rearload=(airres*0.6)+(weight/9.81)*hA-(centerg)*Wperp+weight*np.sin(tslope)*hA

    frontloads=(Wperp*(wbase-centerg)-weight*np.sin(tslope)*hA)/(wbase)
    rearloads=(Wperp*centerg+weight*np.sin(tslope)*hA)/wbase
    #frontloads=(weight*np.cos(tslope)*centerg)/wbase
    #rearloads=(weight*np.cos(tslope))-frontload

    return rearload,frontload,frontloads,rearloads

def hp(angularve,torque,tslope,centerg,wbase,weight):
    frontload=(weight*np.cos(tslope)*centerg)/wbase
    rearload=(weight*np.cos(tslope))-frontload
    return rearload,frontload

def calchp(angularve,torque):
    hp=torque*angularve/5252
    return hp
    
def graphs(angularvex, torquex,name):
    fx=poly.fit(angularvex,torquex,2)
    x=np.linspace(min(angularvex),max(angularvex),6000)
    fy=fx(x)
    plt.subplot(2,2,1)
    plt.plot(angularvex,torquex,"ro")
    plt.plot(x,fy,"r-")
    plt.title("Experimental dyno data and fit")
    plt.ylabel("engine torque(n m)")
    plt.xlabel("angular velocity(rpm)")
    plt.legend([name])
    #



def makeres(count,ogsheet,dragC,area):
    workbook=pyxl.load_workbook("All_Results.xlsx")
    newsheet=workbook["Results_{}".format(count)]
    newsheet["4A"]="Gears"
    newsheet["4B"]="Ratio"
    newsheet["4C"]="WR(N)"
    newsheet["4D"]="WF(N)"
    newsheet["4E"]="WRstat(N)"
    newsheet["4F"]="WFstat(N)"
    newsheet["4G"]="a(m/s^2)"
    newsheet["4H"]="P(N)"
    newsheet["4I"]="R(N)"
    newsheet["4J"]="HP(hp)"
    newsheet["4K"]="Te(Nm)"
    newsheet["4L"]="omega_e(RPM)"

    newsheet["12A"]="DATA"
    for i in range(12,24):
        newsheet["A{}".format(i+12)]=ogsheet["B{}".format(i)].value
        newsheet["B{}".format(i+12)]=ogsheet["C{}".format(i)].value
    newsheet["26A"]="CD"
    newsheet["26B"]=dragC
    newsheet["27A"]="Area"
    newsheet["27B"]=area
    return newsheet

def outputs(gears,angularve,te,accel,trac,roadload,newsheet,name,rearloadS,frontLoadS):
    for i in range(0,6):
        hp=calchp(angularve[i],te[i])
        if(angularve[i]>7000 and name!="CR-28"):
            accel[i]=None
            trac[i]=None
            roadload=None
            te=None
            angularve=None
            hp=None
            rearloadS=None
            frontLoadS=None
        newsheet["B{}".format(i+5)]=gears[i]
        newsheet["E{}".format(i+5)]=rearloadS
        newsheet["F{}".format(i+5)]=frontLoadS
        newsheet["G{}".format(i+5)]=accel[i]
        newsheet["H{}".format(i+5)]=trac[i]
        newsheet["I{}".format(i+5)]=roadload
        newsheet["J{}".format(i+5)]=hp
        newsheet["K{}".format(i+5)]=te[i]
        newsheet["L{}".format(i+5)]=angularve[i]