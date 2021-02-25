# coding: utf-8
import win32api
import pythoncom
import re
import win32com.client as win32
import warnings
import sys
import os
import time
import uuid

_VERSION_ = 'v0.1.0'

def show_version():
    print("=" * 30)     
    print("Send_email_with_python.{}".format(_VERSION_).center(20))
    print("=" * 30)
show_version()

def send_email(fn):
    warnings.filterwarnings('ignore')
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)


    """Set parameters"""

    # account_send_mail   = 'sekikishuo@gmail.com' 
    To_list             = ['sekikishuo@gmail.com','qiqichang1@163.com'] 
    Cc_list             = ['sekikishuo@gmail.com']
    Subject             = 'NEP inspection data'
    HTMLBody            = '''
                            <H2>Hello, This is a response mail.</H2>
                            Hello Guys. 
                        '''
    attachment_path_filename_1 = 'C:\\Users\\Qichang Ql\\Desktop\\Demo\\' + fn
    # attachment_path_filename_2 = ""
    # gmail = outlook.Session.Accounts.Item(2)

    """End for set parameters"""



    # Recipient
    To_str = ''
    for t in To_list:
        To_str = To_str + ';' + t
    mail.To = To_str
    # CC
    Cc_str = ''
    for c in Cc_list:
        Cc_str = Cc_str + ';' + t
    mail.CC = Cc_str
    # Subject
    mail.Subject = Subject
    # Body
    mail.BodyFormat = 2  # 2: Html format
    mail.HTMLBody = HTMLBody
    # Attachments
    mail.Attachments.Add(attachment_path_filename_1)
    # mail.Attachments.Add(attachment_path_filename_2)

    # Send mail
    Send_mail = outlook.Session.Accounts.Item(2)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, Send_mail))
    # mail.Send()

    # Display
    mail.Display()
    # if __name__ == '__main__':
    #     sys.exit(main())



def download_attachments():
    file_name =''
    date_time_stamp = time.strftime("%Y%m%d-%H%M%S")
    #set custom working directory
    os.chdir('C:\\Users\\Qichang Ql\\Desktop\\Demo')
    print(os.getcwd())
    namespace = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    account = namespace.Folders['sekikishuo@gmail.com']
    # main_inbox = outlook.GetDefaultFolder(3)
    # subfolder = main_inbox.Folders.Item("NEP")
    folder = account.Folders.Item("[Gmail]")
    subfolder = folder.Folders.Item('NEP')
    print(subfolder.Name)

    subfolderitems = subfolder.Items
    message = subfolderitems.GetFirst()
    attachment_name = 'RMA.xlsx'

    #Loop to pick messages that are unread
    for message in subfolderitems:
            if message.UnRead == True:
                    print("New Mail Found... Downloading Attachment...")
                    #Loop to check if the attachment name is the same
                    i = 1
                    # num_attach = len([x for x in message.Attachments])
                    # print(num_attach)
                    for attachments in message.Attachments:
                        if attachments.FileName == attachment_name:
                            print(message.Subject)
                            print(attachments.FileName)
                            print(i)
                            # attachment = attachments.Item(i)
                            #Saves to the attachment to the working directory
                            file_uuid = str(uuid.uuid4())
                            attachments.SaveAsFile(os.getcwd() + '\\' + date_time_stamp + file_uuid + attachment_name  )
                            print(os.getcwd() + '\\' + date_time_stamp + file_uuid + attachment_name  )

                            # print (attachments)
                            time.sleep(1)
                            print("Successfully!")
                            i = i + 1
                            # break
                        else:
                            print(message.Subject)
                            print(attachments.FileName)
                        #Go to next unread messages if any
                    time.sleep(1)
                    message.UnRead = False
                    time.sleep(1)
                    message = subfolderitems.GetNext()
            else:
                    print ("Checking...")
    file_name = date_time_stamp + file_uuid + attachment_name
    return file_name


class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # RecrivedItemIDs is a collection of mail IDs separated by a ",".
        # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            subject = mail.Subject
            print(subject)
            if subject in 'NEP':
                # try:
                # Taking all the "BLAHBLAH" which is enclosed by two "%". 
                # command = re.search(r"%NEP%", subject).group(1)
                time.sleep(1)
                fn = download_attachments()
                print(fn)
                time.sleep(4)
                send_email(fn)
                # print(command) # Or whatever code you wish to execute.
                # except:
                #     pass

outlook = win32.DispatchWithEvents("Outlook.Application", Handler_Class)
#and then an infinit loop that waits from events.
# win32api.PostQuitMessage()
pythoncom.PumpMessages() 
