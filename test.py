import win32com.client as win32
import os
import time

import uuid
file_uuid = str(uuid.uuid4())


date_time_stamp = time.strftime("%Y%m%d-%H%M%S")
#set custom working directory
os.chdir('C:\\Users\\Qichang Ql\\Desktop')
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
                # i = 1
                num_attach = len([x for x in message.Attachments])
                print(num_attach)
                for x in range(1, num_attach+1):
                    attachments = message.Attachments

                    if attachments.Item(x).FileName == attachment_name:
                        print(message.Subject)
                        print(attachments.Item(x).FileName)
                        print(file_uuid)
                        # attachment = attachments.Item(i)
                        #Saves to the attachment to the working directory 
                        attachment = attachments.Item(x)
                        attachment.SaveAsFile(os.getcwd() + '\\' + date_time_stamp + file_uuid + attachment_name  )
                        # print (attachments)
                        time.sleep(3)
                        print("Successfully!")
                        # i = i + 1
                        # break
                    else:
                        print(message.Subject)
                        print(attachments.FileName)
                    #Go to next unread messages if any
                message.UnRead = False
                message = subfolderitems.GetNext()
        else:
                print ("Checking...")