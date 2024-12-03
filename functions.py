#This module contatins all function for a statics problem of a car going up a hill
#contains input, calculation, and output functions
import numpy as np
import openpyxl as pyxl
import matplotlib.pyplot as plt
from numpy.polynomial import Polynomial as poly
workbook= pyxl.load_workbook("data.xlsx")


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
    for i in range(4,10):
        while(sheet["B{}".format(i)].value<0):
            sheet["B{}".format(i)]=float(input("The value of your {} is negative, input a new value: ".format(sheet["A{}".format(i)])))
        gears.append(sheet["B{}".format(i)].value)
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
    rollres=sheet["C16"].value
    hA=sheet["C17"].value
    fdrive=sheet["C18"].value
    teff=sheet["C19"].value
    weight=sheet["C20"].value
    airden=sheet["C21"].value
    dratio=sheet["C22"].value
    centerg=sheet["C23"].value
    return speed,tslope,wbase,radius,rollres,hA,fdrive,teff,weight,airden,dratio,centerg

def getdyno(sheet):
    angularvex=[]
    torquex=[]
    for i in range(4,16):
        angularvex.append(sheet["A{}".format(i)].value)
        
        torquex.append(sheet["B{}".format(i)].value)
    angularvex=np.array(angularvex)
    torquex=np.array(torquex)
    return angularvex,torquex
"""def calcRoadLoad(rollres,weight,tslope,airden,dragC,csA,v):
    airres=.5*airden*dragC*csA*v**2
    rRoll=rollres*weight*np.cos((np.atan(tslope/100)))
    roadLoad=rRoll+weight*np.sin((np.atan(tslope/100)))+airres
    return roadLoad
"""
def doMath(rollres,weight,tslope,airden,dragC,csA,v,radius,dratio,teff,fdrive,gears,angularvex,torquex,name,wbase,centerg,hA):
    airres=.5*airden*dragC*csA*v**2
    rRoll=rollres*weight*np.cos((np.atan(tslope)))
    roadLoad=rRoll+weight*np.sin((np.atan(tslope)))+airres



    angularve=gears*dratio*v/radius
    finaldrivE=fdrive/100
    if(name!="CR-28"):
        fx=poly.fit(angularvex,torquex,4)
    else:
        fx=poly.fit(angularvex,torquex,5)
    te=fx(angularve)
    torqued=finaldrivE*dratio*(teff/100)*gears*te
    traction=torqued/radius
    acceleration=((((traction-roadLoad))*9.81)/weight)
    Wperp=weight*np.cos(tslope)
    
    
    frontload=((airres*hA)+(weight/9.81)*hA*acceleration-(wbase-centerg)*Wperp+weight*np.sin(tslope)*hA)/wbase
    rearload=((airres*hA)+(weight/9.81)*hA*acceleration+(centerg)*Wperp+weight*np.sin(tslope)*hA)/wbase

    frontloads=(Wperp*(wbase-centerg))/(wbase)
    rearloads=(Wperp*centerg)/wbase
    #frontloads=(weight*np.cos(tslope)*centerg)/wbase
    #rearloads=(weight*np.cos(tslope))-frontload


    return traction,torqued,acceleration, angularve, roadLoad,rearload,frontload,frontloads,rearloads,te



def calchp(angularve,torque):
    hp=torque*angularve*0.7375621493/5252
    return hp
    
def graphs(angularvex, torquex,name):
    #dyno graph
    if(name!="cr-28"):
        fx=poly.fit(angularvex,torquex,4)
    else:
        fx=poly.fit(angularvex,torquex,5)
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
    try:
        newsheet=workbook["Results_{}".format(count)]
    except KeyError:
        newsheet=workbook.create_sheet("Results_{}".format(count))
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

def outputs(gears,angularve,te,accel,trac,roadload,newsheet,name,rearloadS,frontLoadS,frontLoad,rearLoad):
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
            rearLoad[i]=None
            frontLoad[i]=None
        newsheet["B{}".format(i+5)]=gears[i]
        newsheet["E{}".format(i+5)]=rearloadS
        newsheet["F{}".format(i+5)]=frontLoadS
        newsheet["G{}".format(i+5)]=accel[i]
        newsheet["H{}".format(i+5)]=trac[i]
        newsheet["I{}".format(i+5)]=roadload
        newsheet["J{}".format(i+5)]=hp
        newsheet["K{}".format(i+5)]=te[i]
        newsheet["L{}".format(i+5)]=angularve[i]
        newsheet["C{}".format(i+5)]=rearLoad[i]
        newsheet["D{}".format(i+5)]=frontLoad[i]


def carChoice(workbook):
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
    sheet1=workbook["Problem Variables"]
    sheet2=workbook["cr-28 var"]
    print("**********************************************")
    print("1. {} {}\n2. {} {}\n3. {} {}\n4.{} {}\n5. {} {}\n6. {} {}\n7.Exit".format(sheet1["B26"].value,sheet1["A26"].value,sheet1["B27"].value,sheet1["A27"].value, sheet1["B28"].value,sheet1["A28"].value,sheet1["B29"].value,sheet1["A29"].value,sheet1["B30"].value,sheet1["A30"].value,sheet2["B26"].value,sheet2["A26"].value))
    print("**********************************************")
    while(option!=7):
        while(option!=7 and option!= 6 and option!=5 and option!=4 and option!=3 and option!=2 and option!=1):
        
            option=int(input("Select the car choice options (1-6) "))
            match option:
                case 1:
                    dragC=sheet1["C26"].value
                    area=sheet1["D26"].value
                    name=sheet1["A26"].value
                    count+=1
                    sheet=sheet1

                case 2:
                    dragC=sheet1["C27"].value
                    area=sheet1["D27"].value
                    name=sheet1["A27"].value
                    count+=1
                    sheet=sheet1
                case 3:
                    dragC=sheet1["C28"].value
                    area=sheet1["D28"].value
                    name=sheet1["A28"].value
                    count+=1
                    sheet=sheet1
                case 4:
                    dragC=sheet1["C29"].value
                    area=sheet1["D29"].value
                    name=sheet1["A29"].value
                    count+=1
                    sheet=sheet1
                case 5:
                    dragC=sheet1["C30"].value
                    area=sheet1["D30"].value
                    name=sheet1["A30"].value
                    count+=1
                    sheet=sheet1
                case 6: 
                
                    dragC=sheet2["C26"].value
                    area=sheet2["D26"].value
                    name=sheet2["A26"].value
                    count+=1
                    sheet=sheet2
                case 7:
                    print("Have a good day")
                case _:
                #HR Friendly error message
                    print("ERROR! Selection is invalid\n")
        if(option!=7):
            gears = getGears(sheet)

            #   Function that returns all other varibles in the excel for use by later functions
            v,tslope,wbase,radius,rollres,hA,fdrive,teff,weight,airden,dratio,centerg = getothers(sheet)

            # Function that returns angularvex and torquex
            if(name!="cr-28"):
                dynosheet=workbook["Dynamometer"]
            else:
                dynosheet=workbook["cr-dyno"]
            angularvex,torquex = getdyno(dynosheet)

            #   Function that calculates roadload given rollres,weight,tslope,airden,dragC,csA, and v
            traction,torqued,acceleration, angularve, roadLoad,rearload,frontload,frontloads,rearloads,te=doMath(rollres,weight,tslope,airden,dragC,area,v,radius,dratio,teff,fdrive,gears,angularvex,torquex,name,wbase,centerg,hA)

            #   Outputs graphs
            #functions.graphs(angularve, torqued,name)

            newsheet=makeres(count,sheet,dragC,area)    
            outputs(gears,angularve,te,acceleration,traction,roadLoad,newsheet,name,rearloads,frontloads,frontload,rearload)
