import datetime
import smtplib
from email.mime.text import MIMEText
from email.header import Header


def send_result(result, to_address, cc_addres, accuracy):
    SERVER = "143.166.224.194"
    sender = "gabriel_zhang@dell.com"
    subject = "(Test)APJ MFG Auto On-Hand Recon Data Extraction Result"
    body = '<h1>%s Auto Recon Result</h1>' % today1
    for item in result:
        temp = ""
        if "captured" in item or "database" in item:
            temp = "<p>" + item + "</p>"
        else:
            temp = '<p style="color:red">' + item + "</p>"
        body += temp
    body += '</br></br></br>'
    body += '<h5>Gross $ Accracy rate:</h5>'
    final_site = ''
    final_rate = ''
    for site, rate in accuracy.items():
        temp_site = '<th>%s</th>' % site
        final_site += temp_site
        if rate < 0.98:
            rate = format(rate, '.4%')
            temp_rate = '<td style="color:red">%s</td>' % rate
        else:
            rate = format(rate, '.4%')
            temp_rate = '<td>%s</td>' % rate
        final_rate += temp_rate
    final_site = '<table border="1"><thead><tr>%s</tr></thead>' % final_site
    final_rate = '<tbody><tr>%s</tr></tbody></table>' % final_rate
    body += final_site
    body += final_rate
    body='<html>%s</html>' % body

    body+='<a href="https://dscpowerbi.dell.com/reports/powerbi/DSC/Inventory/PROD/MFG_AutoRecon_DashBoard">PBI Link</a>'

    message = MIMEText(body, 'html', 'utf-8')
    message['From'] = Header('No-Reply MFG_Daily_Auto_Recon', 'utf-8')
    message['To'] = Header(';'.join(to_address), 'utf-8')
    message['Cc'] = Header(';'.join(cc_address), 'utf-8')
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtp_mail = smtplib.SMTP(SERVER)
        smtp_mail.sendmail(sender, to_address+cc_address, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException:
        print("邮件发送失败")

result1= {'CC6':0.9999,'APCC':0.9853,'CC2':0.9542,'CC4':0.9953,'ICC':0.9753}
today1 = datetime.datetime.now().strftime('%Y-%m-%d')
result=['C6 file not generate!','APCC recon result captured！','C2 recon result captured！','C4 recon result captured！','ICC recon result captured！',"Today's result uploaded to database!"]

to_address = ['gabriel_zhang@dell.com','leo_wu1@dell.com']
cc_address = ['cici_chen2@dell.com','gabriel_zhang@dell.com']
send_result(result, to_address, cc_address, result1)

