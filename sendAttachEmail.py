import smtplib

from email.message import EmailMessage

# # 1、设置邮箱域名、发件人邮箱、邮箱授权码、收件人邮箱
mail_sender = "xxx@163.com"
# 邮箱授权码,注意这里不是邮箱密码
mail_license = "xxxxxx"
# 收件人邮箱，可以为多个收件人
mail_receivers = ["xxxx@163.com"]
# SMTP服务器,这里使用163邮箱
smtp_server = "smtp.163.com"

def send_email_with_excel(receivers, subject, content, excel_filename):
    msg = EmailMessage()
    # 1. 设置title和内容
    msg['Subject'] = subject
    msg['From'] = mail_sender
    msg['To'] = receivers
    msg.set_content(content)
    # 2. 添加附件
    with open(excel_filename, "rb") as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=excel_filename)

    # 3. 准备发送邮件
    server = smtplib.SMTP(smtp_server, 25)
    server.set_debuglevel(1)
    server.login(mail_sender, mail_license)
    server.send_message(msg)

send_email_with_excel(mail_receivers, "I'm subject", "I'm content", "table.xlsx")