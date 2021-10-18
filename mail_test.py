import smtplib
from email.mime.text import MIMEText
from email.header import Header

SERVER = "143.166.224.194"
sender = "gabriel_zhang@dell.com"
#mail_host = 'smtp.dell.com'
#mail_user = '18298268658'
#mail_pass = 'cat123'
#sender = '18298268658@163.com'
receivers = ['gabriel_zhang@dell.com','leo_wu1@dell.com']
cc=['cici_chen2@dell.com','gabriel_zhang@dell.com']
subject = "(Test)APJ MFG Auto On-Hand Recon Data Extraction Result"
body='''<html><h1>2021-10-08 Auto Recon Result</h1><p style="color:red">C6 file not generate!</p><p>APCC recon 
result captured！</p><p>C2 recon result captured！</p><p>C4 recon result captured！</p><p>ICC recon result 
captured！</p><p>Today's result uploaded to database!</p></br></br></br><h5>Gross $ Accracy rate:</h5><table 
border="1"><thead><tr><th>CC6</th><th>APCC</th><th>CC2</th><th>CC4</th><th>ICC</th></tr></thead><tbody><tr><td>99
.9900%</td><td>98.5300%</td><td style="color:red">95.4200%</td><td>99.5300%</td><td 
style="color:red">97.5300%</td></tr></tbody></table></html><a 
href="https://dscpowerbi.dell.com/reports/powerbi/DSC/Inventory/PROD/MFG_AutoRecon_DashBoard">PBI Link</a> '''

message = MIMEText(body, 'html', 'utf-8')
message['From'] = Header('MFG_Daily_Auto_Recon', 'utf-8')
message['To'] = Header(';'.join(receivers), 'utf-8')
message['Cc'] = Header(';'.join(cc), 'utf-8')
message['Subject'] = Header(subject, 'utf-8')


try:
    smtpObj = smtplib.SMTP(SERVER)
    smtpObj.sendmail(sender, receivers+cc, message.as_string())
    print("邮件发送成功")
except smtplib.SMTPException:
    print("Error: 无法发送邮件")