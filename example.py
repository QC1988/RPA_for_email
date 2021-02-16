Python 调用outlook发送邮件

使用模块：win32com
1. 模块安装
pip install pypiwin32
1
2. 模块使用
import win32com
# 调用outlook application
outlook = win32com.client.Dispatch('outlook.application')
1
2
3
3. 发送邮件
# 创建一个item
mail =  outlook.CreateItem(0)
# 接收人
mail.To =  "***@outlook.com;***@outlook.com"
# 抄送人
mail.CC =  "***@outlook.com;***@outlook.com"
# 主题
mail.Subject = "这里是一个邮件的主题"
# Body
mail.Body = "这里是一个邮件的主要内容"
# 添加附件
mail.Attachments.Add("这里是要添加附件的位置")
# 可添加多个附件
mail.Attachments.Add("这里是要添加附件的位置")
# 最后发送邮件
mail.Send()
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
4. 当outlook中有多个账号登陆时，选择某个特定的账号进行邮件的发送
# 发件账户
send_account = None

# 遍历所有的账户信息进行筛选
for account in outlook.Session.Accounts:
    # 选择要使用的邮箱账户
    if account.DisplayName == "***@outlook.com":
       # 赋值发件账户
       send_account = account
       break
mail = outlook.CreateItem(0)
# 设置邮件的发件账户
mail._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
# 接下来操作同 1.3
