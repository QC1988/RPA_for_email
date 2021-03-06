import win32com.client
import win32api
import pythoncom
import re


class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            
            subject = mail.Subject
            print(subject)
            try:
                # Taking all the "BLAHBLAH" which is enclosed by two "%". 
                command = re.search(r"%qi%", subject).group(1)

                print(command) # Or whatever code you wish to execute.
            except:
                pass

outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
#and then an infinit loop that waits from events.
# win32api.PostQuitMessage()
pythoncom.PumpMessages() 