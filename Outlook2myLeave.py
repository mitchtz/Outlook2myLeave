import win32com.client
import time
import datetime

#myOlApp = CreateObject("Outlook.Application")
#myItem = myOlApp.CreateItem(olAppointmentItem)




outlook = win32com.client.Dispatch("Outlook.Application") 
namespace = outlook.GetNamespace("MAPI") 
appointments = namespace.GetDefaultFolder(9).Items 

print(datetime.datetime.today().weekday()) #Day of the week (Monday == 0)


# Restrict to items in the next 30 days
begin = datetime.date.today()
end = begin + datetime.timedelta(days = 30);
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
restrictedItems = appointments.Restrict(restriction)


# Iterate through restricted AppointmentItems and print them
for appointmentItem in restrictedItems:
    print("Subject: {0}, Start: {1}, End: {2}".format(
          appointmentItem.Subject, appointmentItem.Start, 
          appointmentItem.End))




##print appointments.count 
#x = 4 # This is a number for one of the calendar entries 
#print(appointments[x])
#print(appointments[x].start)
#print(appointments[x].end)
#print(appointments[x].RecurrenceState) 
#print(appointments[x].Subject)
#print(appointments[x].IsRecurring)
