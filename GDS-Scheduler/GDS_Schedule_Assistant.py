#Code by Mitch Zinser
import win32com.client #Allow interface with outlook
import time
import datetime
import numpy as np
#from win32api import *

#How to get other calenders
#recipient = namespace.createRecipient("Adam Page (OIT)")
#resolved = recipient.Resolve()
#sharedCalender = namespace.GetSharedDefaultFolder(recipient, 9)
#appointments = sharedCalender.Items

#Start working with outlook calender
outlook = win32com.client.Dispatch("Outlook.Application") 
namespace = outlook.GetNamespace("MAPI")
appointments = namespace.GetDefaultFolder(9).Items 
print("Current Day =", datetime.datetime.today().weekday(), "(Where Monday = 0)") #Day of the week (Monday = 0)
#print(datetime.datetime.today().weekday()) #Day of the week (Monday == 0)
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
print("[0]startMonth, [1]startDay, [2]startHour, [3]startMinute, [4]endMonth, [5]endDay, [6]endHour, [7]endMinute")
#print("Off Times")
#print(offTime)
#print("Work Times")
#print(workTime)
print("Off Times sorted")
print(offTimeCut)
print("Work Times sorted")
print(workTimeCut)
