import win32com.client
import time
import datetime
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import getpass

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
#for appointmentItem in restrictedItems:

#    #Prints the subject, start time, and end time in columns
#    #print("Subject: {0}, Start: {1}, End: {2}".format(
#    #      appointmentItem.Subject, appointmentItem.Start, 
#    #      appointmentItem.End))
#    if("Off" in appointmentItem.Subject):
#        print("Subject: {0}, Start: {1}, End: {2}".format(
#           appointmentItem.Subject, appointmentItem.Start, 
#           appointmentItem.End))


#Experiment w/ selenium

browser = webdriver.Firefox() #Open browser
print('Opening browser')
browser.get('https://portal.prod.cu.edu/MyCUInfoFedAuthLogin.html') #Go to a web page
assert 'myCUinfo' in browser.title #Assert that page has loaded
elem = browser.find_element_by_id('username')  # Find the username
elem.send_keys(input('username: ')) # + Keys.RETURN presses return
elem = browser.find_element_by_id('password')  # Find the password
elem.send_keys(getpass.getpass()) #Uses send keys to type password

browser.find_element_by_id('submit').click()
#elem.click()
browser.get('https://portal.prod.cu.edu/psp/epprod/UCB2/ENTP/h/?tab=CU_STAFF')
assert 'MyCUinfo' in browser.title #Assert that page has loaded
browser.get('https://portal.prod.cu.edu/psp/epprod_newwin/EMPLOYEE/ENTP/s/WEBLIB_CU_S_MEN.ISCRIPT2.FieldFormula.IScript_leave?PORTALPARAM_PTCNAV=CU_LEAVE_TG&EOPP.SCNode=ENTP&EOPP.SCPortal=UCB2&EOPP.SCName=ADMN_CU_QUICK_LINKS&EOPP.SCLabel=CU%20Quick%20Links&EOPP.SCPTcname=&FolderPath=PORTAL_ROOT_OBJECT.PORTAL_BASE_DATA.CO_NAVIGATION_COLLECTIONS.ADMN_CU_QUICK_LINKS.ADMN_S201308091032101805630266&IsFolder=false')
















