import win32com.client as win32


def send_mail():
  outlook_app = win32.Dispatch('Outlook.Application')
  # choose sender account
  send_account = None
  print("11111111111111111")
  for account in outlook_app.Session.Accounts:
    print("2222222222222222")
    if account.DisplayName == 'sekikishuo@gmail.com':
        print("-------------")  #account.DisplayName
        send_account = account
        break
  mail_item = outlook_app.CreateItem(0)  # 0: olMailItem
  # mail_item.SendUsingAccount = send_account not working
  # the following statement performs the function instead
  mail_item._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))

#   mail._oleobj_.Invoke(*(64209, 0, 8, 0, gmail))
#   mail_item.Recipients.Add('sekikishuo@gmail.com')
  mail_item.To = 'sekikishuo@gmail.com'
  mail_item.Subject = 'Test sending using gmail account'
  mail_item.BodyFormat = 2  # 2: Html format
  mail_item.HTMLBody = '''
    <H2>Hello, This is a test mail.</H2>
    Hello Guys. 
    '''
  mail_item.Send()
#   mail_item.Display()


send_mail()
# if __name__ == '__main__':
#   send_mail()