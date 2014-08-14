import win32com.client #Allow interface with outlook
import time
import datetime
import selenium #Web control
#from selenium import webdriver #Allows selenium to drive broswers
#from selenium.webdriver.common.keys import Keys #Allows selenium to enter keystrokes
import getpass #allows password input securly in cmd
import numpy as np
from tkinter import * #Gui
from win32api import *

print(GetDomainName())
##Experiment with tkinter
#l = tkinter.Label(text = "See me?")
#l.pack()
#l.mainloop()

#master = Tk()

#f = Frame(master, height=32, width=32)
##f.pack_propagate(0) # don't shrink
#f.pack()
#def callback():
#    print("Click!")

#b = Button(master, text="Button", command=callback)
##b = Button(master, text="longtext", anchor=W, justify=LEFT, padx=2) #Button formatting
#b.config(padx = 20) #Horizontal padding between text and border. padx is horiz, pady is vert padding
#b.pack()
#b.mainloop()

#Checkboxes

#var = IntVar()

#c = Checkbutton(master, text="Expand", variable=var)
#c.pack()

#mainloop()

#Radiobuttons
#v = IntVar()

#Radiobutton(master, text="One", variable=v, value=1).pack(anchor=W)
#Radiobutton(master, text="Two", variable=v, value=2).pack(anchor=W)

#mainloop()

outlook = win32com.client.Dispatch("Outlook.Application") 
namespace = outlook.GetNamespace("MAPI") 
appointments = namespace.GetDefaultFolder(9).Items 
print("Day =")
print(datetime.datetime.today().weekday()) #Day of the week (Monday == 0)
#Calculate the day of the week and count days over Sunday
if datetime.datetime.today().weekday() == 6:
    dayOverSun = 0
else:
    dayOverSun = datetime.datetime.today().weekday() + 1

# Restrict to items in the next (days =) days
#dayOverSun is used to bring the start of the calender reading back until Sunday
begin = datetime.date.today() + datetime.timedelta(days = -dayOverSun) #set start for calender reading period
end = begin + datetime.timedelta(days = 7); #set end date for calender reading period
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
restrictedItems = appointments.Restrict(restriction)


# Iterate through restricted AppointmentItems and print them
o = 0
w = 0
offTime = np.zeros((100, 8)) #[0]startMonth, [1]startDay, [2]startHour, [3]startMinute, [4]endMonth, [5]endDay, [6]endHour, [7]endMinute
workTime = np.zeros((100, 8))#[0]startMonth, [1]startDay, [2]startHour, [3]startMinute, [4]endMonth, [5]endDay, [6]endHour, [7]endMinute
for appointmentItem in restrictedItems:
    ##Prints the subject, start time, and end time in columns
    #print("Subject: {0}, Start: {1}, End: {2}".format(
    #      appointmentItem.Subject, appointmentItem.Start., 
    #      appointmentItem.End.))

        
    #If subject of appointment is at work
    if(("Walkin" in appointmentItem.Subject) or ("INC0" in appointmentItem.Subject) or ("ITASK0" in appointmentItem.Subject) or ("Jack Checks" in appointmentItem.Subject)):
        workTime[w][0] = appointmentItem.Start.month
        workTime[w][1] = appointmentItem.Start.day
        workTime[w][2] = appointmentItem.Start.hour
        workTime[w][3] = appointmentItem.Start.minute
        workTime[w][4] = appointmentItem.End.month
        workTime[w][5] = appointmentItem.End.day
        workTime[w][6] = appointmentItem.End.hour
        workTime[w][7] = appointmentItem.End.minute
        w = w + 1

    #If the subject is off or a class of some type or other event
    else:
        offTime[o][0] = appointmentItem.Start.month
        offTime[o][1] = appointmentItem.Start.day
        offTime[o][2] = appointmentItem.Start.hour
        offTime[o][3] = appointmentItem.Start.minute
        offTime[o][4] = appointmentItem.End.month
        offTime[o][5] = appointmentItem.End.day
        offTime[o][6] = appointmentItem.End.hour
        offTime[o][7] = appointmentItem.End.minute
        o = o + 1

workTimeCut = np.zeros((w, 8)) #Create new 2d array the size of the input data
workTimeCut[:][:] = workTime[:w][:] #Cut and paste old array into new array
offTimeCut = np.zeros((o, 8)) #Create new 2d array the size of the input data
offTimeCut[:][:] = offTime[:o][:] #Cut and paste old array into new array
workTimeCut = workTimeCut[workTimeCut[:,1].argsort()] #Sort work time by start date
offTimeCut = offTimeCut[offTimeCut[:,1].argsort()] #Sort off time by start date
print("Off Times")
print(offTime)
print("Work Times")
print(workTime)
print("Off Times sorted")
print(offTimeCut)
print("Work Times sorted")
print(workTimeCut)

#TIME TO FIGURE BLANK TIME MOTHERFUCKER
date = begin + datetime.timedelta(days = 1) #Set date to Monday of whatever week it is
blankTime = np.zeros((100, 8)) #Create new blankTime 2d array 100 long by 8 wide
o = 0 #Count number of blankTime events
for i in range(0, 1): #5Iterate From monday through friday, loop runs on friday on last loop then iterates to saturday and quits
    print("Loop number is", i)
    print("Loop day is", date.day)
    j = 32 #Increments of 15 minutes. Each j is 15 minutes. This sets j to 8:00 AM
    counting = False #Records if blankTime is being counted currently
    while (j <= 68): #While j is less than 5:00 PM
        print("Hour is", int(j/4))
        print("Minute is", (j%4)*15)
        #Find blank time by running through day in 15 minute block and comparing with offTime and workTime to see if current 15 minute block is laready occupied
        blank = True
        for k in range(0, len(offTimeCut)): #Search through offTimeCut list in search of time that is blank
            if (offTimeCut[k][1] == date.day): #Start time is the same as current day being evaluated
                if ((((offTimeCut[k][2])*4) + (int((offTimeCut[k][3])/4))) <= j):
                    if ((((offTimeCut[k][6])*4) + (int((offTimeCut[k][7])/4))) > j):
                            blank = False
            if (workTimeCut[k][1] == date.day): #Start time is the same as current day being evaluated
                if ((((workTimeCut[k][2])*4) + (int((workTimeCut[k][3])/4))) <= j):
                    if ((((workTimeCut[k][6])*4) + (int((workTimeCut[k][7])/4))) > j):
                            blank = False


        print("blankTime", blank) #Test

        if ((blank == True) and (counting == False)): #If current time is beginning of blank time
            startTime = j #Record the start of this free time
            counting = True #Record that free time is being counted already
        if ((blank == False) and (counting == True)): #If current time is end of blank time
            #Record blankTime
            blankTime[o][0] = date.month #Start month
            blankTime[o][1] = date.day #Start day
            blankTime[o][2] = int(startTime/4) #Start hour
            blankTime[o][3] = (startTime%4)*15 #Start minute
            blankTime[o][4] = date.month #End month
            blankTime[o][5] = date.day #End day
            blankTime[o][6] = int(j/4) #End hour
            blankTime[o][7] = (j%4)*15 #End minute
            o = o + 1 #Iterate event counter
            counting = False
        j = j + 1
    date = date + datetime.timedelta(days = 1) #Iterate through days

blankTimeCut = np.zeros((o, 8)) #Create new 2d array the size of the input data
blankTimeCut[:][:] = blankTime[:o][:] #Cut and paste old array into new array

print("Day after loop", date.day)

#[0]startMonth, [1]startDay, [2]startHour, [3]startMinute, [4]endMonth, [5]endDay, [6]endHour, [7]endMinute
hoursWorked = 0
#Add hours worked
for i in range(0, w):
    hoursWorked = hoursWorked + (workTimeCut[i][6] - workTimeCut[i][2])
    if (workTimeCut[i][2] == 30): #If start time is half hour, subtract time
        hoursWorked = hoursWorked-.5
    if (workTimeCut[i][6] == 30): #If end time is half hour, add time
        hoursWorked = hoursWorked+.5
print("Hours worked is", hoursWorked)
print("Should be 30.5 with current blank time")
print("Blank time cut", blankTimeCut)


##Experiment w/ selenium
#browser = webdriver.Firefox() #Open browser
#print('Opening browser')
#browser.get('https://portal.prod.cu.edu/MyCUInfoFedAuthLogin.html') #Go to a web page
#assert 'myCUinfo' in browser.title #Assert that page has loaded
#elem = browser.find_element_by_id('username')  # Find the username
#elem.send_keys(input('username: ')) # + Keys.RETURN presses return
#elem = browser.find_element_by_id('password')  # Find the password
#elem.send_keys(getpass.getpass()) #Uses send keys to type password

#browser.find_element_by_id('submit').click()
##elem.click()
#browser.get('https://portal.prod.cu.edu/psp/epprod/UCB2/ENTP/h/?tab=CU_STAFF')
#assert 'MyCUinfo' in browser.title #Assert that page has loaded
#browser.get('https://portal.prod.cu.edu/psp/epprod_newwin/EMPLOYEE/ENTP/s/WEBLIB_CU_S_MEN.ISCRIPT2.FieldFormula.IScript_leave?PORTALPARAM_PTCNAV=CU_LEAVE_TG&EOPP.SCNode=ENTP&EOPP.SCPortal=UCB2&EOPP.SCName=ADMN_CU_QUICK_LINKS&EOPP.SCLabel=CU%20Quick%20Links&EOPP.SCPTcname=&FolderPath=PORTAL_ROOT_OBJECT.PORTAL_BASE_DATA.CO_NAVIGATION_COLLECTIONS.ADMN_CU_QUICK_LINKS.ADMN_S201308091032101805630266&IsFolder=false')
#browser.set_window_size(1000, 700)




