# coding: utf-8
import win32com.client as win32
import pythoncom
import warnings
import sys

_VERSION_ = 'v0.1.0'

def show_version():
    print("=" * 30)     
    print("Send_email_with_python.{}".format(_VERSION_).center(20))
    print("=" * 30)
show_version()

warnings.filterwarnings('ignore')
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)




"""Set parameters"""

account_send_mail   = 'sekikishuo@gmail.com' 
To_list             = ['sekikishuo@gmail.com','qiqichang1@163.com'] 
Cc_list             = ['sekikishuo@gmail.com']
Subject             = 'NEP inspection data'
HTMLBody            = '''
                        <H2>Hello, This is a test mail.</H2>
                        Hello Guys. 
                     '''
attachment_path_filename_1 = "C:\\Users\\Qichang Ql\\Desktop\\テスト　FOR　PYTHON\\RPA_for_test\\RMA.xlsx"
# attachment_path_filename_2 = ""
# gmail = outlook.Session.Accounts.Item(2)
# gmail = outlook.Session.Accounts[2]

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
mail.Cc = Cc_str
# Subject
mail.Subject = Subject
# Body
mail.BodyFormat = 2  # 2: Html format
mail.HTMLBody = HTMLBody
# Attachments
mail.Attachments.Add(attachment_path_filename_1)
# mail.Attachments.Add(attachment_path_filename_2)

# Send mail
Send_mail = outlook.Session.Accounts[account_send_mail]
mail._oleobj_.Invoke(*(64209, 0, 8, 0, Send_mail))
# mail.Send()

# Display
mail.Display()


# if __name__ == '__main__':
#     sys.exit(main())