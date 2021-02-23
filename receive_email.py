import win32com.client as win32
import xlrd
import re
import time

outlook = win32.Dispatch('outlook.application')

send_account = None
mail = outlook.CreateItem(0)
outlook.MailItem.SendUsingAccount = ["sekikishuo@gmail.com"]
# 遍历所有的账户信息进行筛选
# for account in outlook.Session.Accounts:
#     # 选择要使用的邮箱账户
#     if account.DisplayName == "***@gmail.com":
#        # 赋值发件账户
#        send_account = account
#        break
# mail = outlook.CreateItem(0)
# 设置邮件的发件账户
# mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
# 接下来操作同 1.3


mail.GetInspector # 这里很关键，有了这代码，下面才能获取到outlook默认签名
receivers = ['sekikishuo@gmail.com']
mail.To = receivers[0]
mail.Subject = 'test1'
#print(mail.HTMLBody) #这里打印的就是签名，调用了mail.GetInspector之后，HTMLBody就会自动变为签名，需要添加正文的话，把正文加进去就好了

# bodystart = re.search("<body.*?>", mail.HTMLBody) # 找到签名里面的body头，签名是html格式的
# mail.HTMLBody = re.sub(bodystart.group(), bodystart.group() + "THE text", mail.HTMLBody) # 在签名里的body头后面插入正文


# workbook = xlrd.open_workbook("C:\\Users\\Qichang Ql\\Desktop\\テスト　FOR　PYTHON\\RPA_for_test\\RMA.xlsx")
# mySheet = workbook.sheet_by_index(0)
# nrows = mySheet.nrows
# content = []
# for i in range(nrows):
#     ss = mySheet.row_values(i)
#     content.append(ss)
#     print(content)
#     Truecontent = str(content)
# mail.Body = Truecontent
mail.Body = "这里是邮件正文" #Body和HTMLBody只用一个

mail.Attachments.Add("C:\\Users\\Qichang Ql\\Desktop\\テスト　FOR　PYTHON\\RPA_for_test\\RMA.xlsx")
time.sleep(1)

mail.Send()