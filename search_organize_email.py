import win32com.client as win32
import os


get_path = 'C:\\Users\\Qichang Ql\\Desktop'


outlook = win32.Dispatch('Outlook.Application')


namespace = outlook.GetNamespace('MAPI')
account = namespace.Folders['sekikishuo@gmail.com']
Gmail = account.Folders['[Gmail]']
jyuyo = Gmail.Folders['ゴミ箱']
# print(inbox.Items.count)
yt_email = [mail for mail in jyuyo.Items if mail.SenderEmailAddress.endswith('gmail.com')]
for mail in yt_email:
    print(mail)
    attachments = mail.Attachments
    num_attach = len([x for x in attachments])
    print(num_attach)
    for x in range(1, num_attach+1):
        print(x)
        attachment = attachments.Item(x)
        attachment.SaveAsFile(os.path.join(get_path, attachment.FileName))
        print(attachment, "saved!")

