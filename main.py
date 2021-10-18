# import pymssql #引入pymssql模块
# 最近版本09/23， 新增outlook邮件功能~

import pandas as pd
import glob
import os
import datetime
import logging
from sqlalchemy import create_engine
import smtplib
from email.mime.text import MIMEText
from email.header import Header
#import win32com.client as win32

### 如果跑当日的需要更改三个地方~~

strFileRoot = r"\\AMERBIZPRDMP01.amer.dell.com\SCG_36158prodMP\FACTORY_RECON"
today = datetime.datetime.now().strftime('%Y_%m_%d')
#today = '2021_09_25'
apcc_files = glob.glob(os.path.join(strFileRoot, "APC\RECON\stockStatusReconReport_" + today + "*.xlsx"))
c2_files = glob.glob(os.path.join(strFileRoot, "CC2\RECON\stockStatusReconReport_" + today + "*.xlsx"))
c4_files = glob.glob(os.path.join(strFileRoot, "CC4\RECON\stockStatusReconReport_" + today + "*.xlsx"))
c6_files = glob.glob(os.path.join(strFileRoot, "CC6\RECON\stockStatusReconReport_" + today + "*.xlsx"))
icc_files = glob.glob(os.path.join(strFileRoot, "ICCX\RECON\stockStatusReconReport_*" + today + "*.xlsx"))
today1 = datetime.datetime.now().strftime('%Y-%m-%d')

# today1='2021-09-25'


def read_icc(dataframe, result, accuracy):
    try:
        if len(icc_files) > 0:
            for file in icc_files:
                icc_frame = pd.read_excel(file)
                icc_frame = icc_frame[icc_frame['Bulk Expense'] == 'N']
                icc_frame['Total_Amount'] = icc_frame['Standard Cost'] * icc_frame['Prism-OH QTY']  # 增加总金额
                Total_Amount = sum(icc_frame['Total_Amount'])  # 获得总金额
                Issue_Part_Gross_Amt = icc_frame[icc_frame['Net Variance Qty'] != 0]['Gross Variance Values'].sum()
                Issue_Part_Net_Amt = icc_frame[icc_frame['Net Variance Qty'] != 0]['Net Variance Values'].sum()
                Total_Line_Qty = icc_frame['Inventory Org'].count()  # 获得行数
                Issue_Line_Qty = icc_frame[icc_frame['Net Variance Qty'] != 0]['Inventory Org'].count()  # 获取非0行数
                Issue_Line_Percentage = Issue_Line_Qty / Total_Line_Qty
                Gross_Amount_Accuracy = 1 - Issue_Part_Gross_Amt / Total_Amount
                Net_Amount_Accuracy = 1 - abs(Issue_Part_Net_Amt / Total_Amount)
                record_row = []
                Recon_Site = 'ICC'
                # Recon_date=datetime.datetime.now().strftime('%Y-%m-%d')
                Recon_date = today1
                record_row.append(Recon_Site)
                record_row.append(Recon_date)
                record_row.append(Total_Amount)
                record_row.append(Issue_Part_Gross_Amt)
                record_row.append(Issue_Part_Net_Amt)
                record_row.append(Total_Line_Qty)
                record_row.append(Issue_Line_Qty)
                record_row.append(Issue_Line_Percentage)
                record_row.append(Gross_Amount_Accuracy)
                record_row.append(Net_Amount_Accuracy)
                # 将记录添加进总的dataframe
                dataframe.loc[-1] = record_row
                dataframe.index = dataframe.index + 1
                dataframe = dataframe.sort_index()
                accuracy[Recon_Site] = Gross_Amount_Accuracy
                logger.info("ICC recon result captured！")
                result.append('ICC recon result captured！')
        else:
            logger.info("ICC file not generate!")
            result.append('ICC file not generate!')
            Recon_Site = 'ICC'
            # Recon_date = datetime.datetime.now().strftime('%Y-%m-%d')
            Recon_date = today1
            record_row = [Recon_Site, Recon_date, 0, 0, 0, 0, 0, 0, 0, 0]
            dataframe.loc[-1] = record_row
            dataframe.index = dataframe.index + 1
            dataframe = dataframe.sort_index()
            accuracy[Recon_Site] = 0
    except:
        logger.error("异常信息", exc_info=True)


def read_c4(dataframe, result, accuracy):
    try:
        if len(c4_files) > 0:
            for file in c4_files:
                c4_frame = pd.read_excel(file)
                c4_frame = c4_frame[c4_frame['Bulk Expense'] == 'N']
                c4_frame['Total_Amount'] = c4_frame['Standard Cost'] * c4_frame['Prism-OH QTY']  # 增加总金额
                Total_Amount = sum(c4_frame['Total_Amount'])  # 获得总金额
                Issue_Part_Gross_Amt = c4_frame[c4_frame['Net Variance Qty'] != 0]['Gross Variance Values'].sum()
                Issue_Part_Net_Amt = c4_frame[c4_frame['Net Variance Qty'] != 0]['Net Variance Values'].sum()
                Total_Line_Qty = c4_frame['Inventory Org'].count()  # 获得行数
                Issue_Line_Qty = c4_frame[c4_frame['Net Variance Qty'] != 0]['Inventory Org'].count()  # 获取非0行数
                Issue_Line_Percentage = Issue_Line_Qty / Total_Line_Qty
                Gross_Amount_Accuracy = 1 - Issue_Part_Gross_Amt / Total_Amount
                Net_Amount_Accuracy = 1 - abs(Issue_Part_Net_Amt / Total_Amount)
                record_row = []
                Recon_Site = 'CC4'
                # Recon_date=datetime.datetime.now().strftime('%Y-%m-%d')
                Recon_date = today1
                record_row.append(Recon_Site)
                record_row.append(Recon_date)
                record_row.append(Total_Amount)
                record_row.append(Issue_Part_Gross_Amt)
                record_row.append(Issue_Part_Net_Amt)
                record_row.append(Total_Line_Qty)
                record_row.append(Issue_Line_Qty)
                record_row.append(Issue_Line_Percentage)
                record_row.append(Gross_Amount_Accuracy)
                record_row.append(Net_Amount_Accuracy)
                # 将记录添加进总的dataframe
                dataframe.loc[-1] = record_row
                dataframe.index = dataframe.index + 1
                dataframe = dataframe.sort_index()
                accuracy[Recon_Site] = Gross_Amount_Accuracy
                logger.info("C4 recon result captured！")
                result.append('C4 recon result captured！')
        else:
            logger.info("C4 file not generate!")
            result.append('C4 file not generate!')
            Recon_Site = 'CC4'
            # Recon_date = datetime.datetime.now().strftime('%Y-%m-%d')
            Recon_date = today1
            record_row = [Recon_Site, Recon_date, 0, 0, 0, 0, 0, 0, 0, 0]
            dataframe.loc[-1] = record_row
            dataframe.index = dataframe.index + 1
            dataframe = dataframe.sort_index()
            accuracy[Recon_Site] = 0
    except:
        logger.error("异常信息", exc_info=True)


def read_c2(dataframe, result, accuracy):
    try:
        if len(c2_files) > 0:
            for file in c2_files:
                c2_frame = pd.read_excel(file)
                c2_frame = c2_frame[c2_frame['Bulk Expense'] == 'N']
                # c2_frame['Recon_Date']='2021-09-09'  #增加日期
                # c2_frame[['Con','Inventory Org']]=c2_frame[['Inventory Org','Con']]#交换列的值,列名不变
                # c2_frame.loc[c2_frame['Con']=='CC6']='CCC6'  #替换列的值
                c2_frame['Total_Amount'] = c2_frame['Standard Cost'] * c2_frame['Prism-OH QTY']  # 增加总金额
                # c2_frame.drop(['Con', 'Operating Unit', 'Item Description', 'Bulk Expense', 'AX-OH QTY', 'ISEMC', 'Return On ASN','Consigned'],axis=1,inplace=True)  #删除多余列
                Total_Amount = sum(c2_frame['Total_Amount'])  # 获得总金额
                # print("总金额："+str(round(Total_Amount)))
                Issue_Part_Gross_Amt = c2_frame[c2_frame['Net Variance Qty'] != 0]['Gross Variance Values'].sum()
                # print("Gross Var："+str(round(Issue_Part_Gross_Amt)))
                Issue_Part_Net_Amt = c2_frame[c2_frame['Net Variance Qty'] != 0]['Net Variance Values'].sum()
                # print("Net Var：" + str(round(Issue_Part_Net_Amt)))
                Total_Line_Qty = c2_frame['Inventory Org'].count()  # 获得行数
                # print("总行数：" + str(Total_Line_Qty))
                Issue_Line_Qty = c2_frame[c2_frame['Net Variance Qty'] != 0]['Inventory Org'].count()  # 获取非0行数
                # print("差异行数：" + str(Issue_Line_Qty))
                Issue_Line_Percentage = Issue_Line_Qty / Total_Line_Qty
                # print("差异行数百分比：" + str(round(Issue_Line_Percentage,6)))
                Gross_Amount_Accuracy = 1 - Issue_Part_Gross_Amt / Total_Amount
                # print("gross准确率：" + str(round(Gross_Amount_Accuracy,6)))
                Net_Amount_Accuracy = 1 - abs(Issue_Part_Net_Amt / Total_Amount)
                # print("Net准确率：" + str(round(Net_Amount_Accuracy,6)))
                record_row = []
                Recon_Site = 'CC2'
                # Recon_date=datetime.datetime.now().strftime('%Y-%m-%d')
                Recon_date = today1
                record_row.append(Recon_Site)
                record_row.append(Recon_date)
                record_row.append(Total_Amount)
                record_row.append(Issue_Part_Gross_Amt)
                record_row.append(Issue_Part_Net_Amt)
                record_row.append(Total_Line_Qty)
                record_row.append(Issue_Line_Qty)
                record_row.append(Issue_Line_Percentage)
                record_row.append(Gross_Amount_Accuracy)
                record_row.append(Net_Amount_Accuracy)
                # print(record_row)
                # 将记录添加进总的dataframe
                dataframe.loc[-1] = record_row
                dataframe.index = dataframe.index + 1
                dataframe = dataframe.sort_index()
                accuracy[Recon_Site] = Gross_Amount_Accuracy
                logger.info("C2 recon result captured！")
                result.append('C2 recon result captured！')
            # return dataframe
        else:
            logger.info("C2 file not generate!")
            result.append('C2 file not generate!')
            Recon_Site = 'CC2'
            # Recon_date = datetime.datetime.now().strftime('%Y-%m-%d')
            Recon_date = today1
            record_row = [Recon_Site, Recon_date, 0, 0, 0, 0, 0, 0, 0, 0]
            dataframe.loc[-1] = record_row
            dataframe.index = dataframe.index + 1
            dataframe = dataframe.sort_index()
            accuracy[Recon_Site] = 0
    except:
        logger.error("异常信息", exc_info=True)


def read_apcc(dataframe, result, accuracy):
    try:
        if len(apcc_files) > 0:
            for file in apcc_files:
                apcc_frame = pd.read_excel(file)
                apcc_frame = apcc_frame[apcc_frame['Bulk Expense'] == 'N']
                apcc_frame['Total_Amount'] = apcc_frame['Standard Cost'] * apcc_frame['Prism-OH QTY']  # 增加总金额
                Total_Amount = sum(apcc_frame['Total_Amount'])  # 获得总金额
                Issue_Part_Gross_Amt = apcc_frame[apcc_frame['Net Variance Qty'] != 0]['Gross Variance Values'].sum()
                Issue_Part_Net_Amt = apcc_frame[apcc_frame['Net Variance Qty'] != 0]['Net Variance Values'].sum()
                Total_Line_Qty = apcc_frame['Inventory Org'].count()  # 获得行数
                Issue_Line_Qty = apcc_frame[apcc_frame['Net Variance Qty'] != 0]['Inventory Org'].count()  # 获取非0行数
                Issue_Line_Percentage = Issue_Line_Qty / Total_Line_Qty
                Gross_Amount_Accuracy = 1 - Issue_Part_Gross_Amt / Total_Amount
                Net_Amount_Accuracy = 1 - abs(Issue_Part_Net_Amt / Total_Amount)
                record_row = []
                Recon_Site = 'APCC'
                # Recon_date=datetime.datetime.now().strftime('%Y-%m-%d')
                Recon_date = today1
                record_row.append(Recon_Site)
                record_row.append(Recon_date)
                record_row.append(Total_Amount)
                record_row.append(Issue_Part_Gross_Amt)
                record_row.append(Issue_Part_Net_Amt)
                record_row.append(Total_Line_Qty)
                record_row.append(Issue_Line_Qty)
                record_row.append(Issue_Line_Percentage)
                record_row.append(Gross_Amount_Accuracy)
                record_row.append(Net_Amount_Accuracy)
                # 将记录添加进总的dataframe
                dataframe.loc[-1] = record_row
                dataframe.index = dataframe.index + 1
                dataframe = dataframe.sort_index()
                accuracy[Recon_Site] = Gross_Amount_Accuracy
                logger.info("APCC recon result captured！")
                result.append('APCC recon result captured！')
        else:
            logger.info("APCC file not generate!")
            result.append('APCC file not generate!')
            Recon_Site = 'APCC'
            # Recon_date = datetime.datetime.now().strftime('%Y-%m-%d')
            Recon_date = today1
            record_row = [Recon_Site, Recon_date, 0, 0, 0, 0, 0, 0, 0, 0]
            dataframe.loc[-1] = record_row
            dataframe.index = dataframe.index + 1
            dataframe = dataframe.sort_index()
            accuracy[Recon_Site] = 0
    except:
        logger.error("异常信息", exc_info=True)


def read_c6(dataframe, result, accuracy):
    try:
        if len(c6_files) > 0:
            for file in c6_files:
                c6_frame = pd.read_excel(file)
                c6_frame = c6_frame[c6_frame['Bulk Expense'] == 'N']
                # c6_frame['Recon_Date']='2021-09-09'  #增加日期
                # c6_frame[['Con','Inventory Org']]=c6_frame[['Inventory Org','Con']]#交换列的值,列名不变
                # c6_frame.loc[c6_frame['Con']=='CC6']='CCC6'  #替换列的值
                c6_frame['Total_Amount'] = c6_frame['Standard Cost'] * c6_frame['Prism-OH QTY']  # 增加总金额
                # c6_frame.drop(['Con', 'Operating Unit', 'Item Description', 'Bulk Expense', 'AX-OH QTY', 'ISEMC', 'Return On ASN','Consigned'],axis=1,inplace=True)  #删除多余列
                Total_Amount = sum(c6_frame['Total_Amount'])  # 获得总金额
                # print("总金额："+str(round(Total_Amount)))
                Issue_Part_Gross_Amt = c6_frame[c6_frame['Net Variance Qty'] != 0]['Gross Variance Values'].sum()
                # print("Gross Var："+str(round(Issue_Part_Gross_Amt)))
                Issue_Part_Net_Amt = c6_frame[c6_frame['Net Variance Qty'] != 0]['Net Variance Values'].sum()
                # print("Net Var：" + str(round(Issue_Part_Net_Amt)))
                Total_Line_Qty = c6_frame['Inventory Org'].count()  # 获得行数
                # print("总行数：" + str(Total_Line_Qty))
                Issue_Line_Qty = c6_frame[c6_frame['Net Variance Qty'] != 0]['Inventory Org'].count()  # 获取非0行数
                # print("差异行数：" + str(Issue_Line_Qty))
                Issue_Line_Percentage = Issue_Line_Qty / Total_Line_Qty
                # print("差异行数百分比：" + str(round(Issue_Line_Percentage,6)))
                Gross_Amount_Accuracy = 1 - Issue_Part_Gross_Amt / Total_Amount
                # print("gross准确率：" + str(round(Gross_Amount_Accuracy,6)))
                Net_Amount_Accuracy = 1 - abs(Issue_Part_Net_Amt / Total_Amount)
                # print("Net准确率：" + str(round(Net_Amount_Accuracy,6)))
                record_row = []
                Recon_Site = 'CC6'
                # Recon_date=datetime.datetime.now().strftime('%Y-%m-%d')
                Recon_date = today1
                record_row.append(Recon_Site)
                record_row.append(Recon_date)
                record_row.append(Total_Amount)
                record_row.append(Issue_Part_Gross_Amt)
                record_row.append(Issue_Part_Net_Amt)
                record_row.append(Total_Line_Qty)
                record_row.append(Issue_Line_Qty)
                record_row.append(Issue_Line_Percentage)
                record_row.append(Gross_Amount_Accuracy)
                record_row.append(Net_Amount_Accuracy)
                # print(record_row)
                # 将记录添加进总的dataframe
                dataframe.loc[-1] = record_row
                dataframe.index = dataframe.index + 1
                dataframe = dataframe.sort_index()
                accuracy[Recon_Site] = Gross_Amount_Accuracy
                logger.info("C6 recon result captured！")
                result.append('C6 recon result captured！')
        else:
            logger.info("C6 file not generate!")
            result.append('C6 file not generate!')
            Recon_Site = 'CC6'
            # Recon_date = datetime.datetime.now().strftime('%Y-%m-%d')
            Recon_date = today1
            record_row = [Recon_Site, Recon_date, 0, 0, 0, 0, 0, 0, 0, 0]
            dataframe.loc[-1] = record_row
            dataframe.index = dataframe.index + 1
            dataframe = dataframe.sort_index()
            accuracy[Recon_Site] = 0
    except:
        logger.error("异常信息", exc_info=True)


'''
def conn():
    connect = pymssql.connect('ctupwcc6role2', 'Inv', 'f$msFT7_&#$!', 'Inv') #服务器名,账户,密码,数据库名
    if connect:
        print("连接成功!")
    return connect
'''
'''
def create_eng():
    enging=create_engine('mssql+pymssql://Inv:f$msFT7_&#$!@ctupwcc6role2/Inv')
    return enging
'''

'''
#调用pywin32, 但是task scheduler不能用
def send_result(result, to_address, cc_addres, accuracy):
    outlook = win32.Dispatch("Outlook.Application")
    # 创建一个邮件对象
    mail = outlook.CreateItem(0)
    # 对邮件的各个属性进行赋值
    # mail.To = "test@outlook.com;test2@outlook.com;test3@outlook.com;"  # 收件邮箱列表,如多人用;隔开
    mail.To = to_address
    # mail.CC = 'test@outlook.com'  # 抄送邮箱列表
    mail.CC = cc_addres
    # mail.BCC = "test@outlook.com" # 密抄邮箱列表，谨慎使用
    mail.Subject = "(Test)APJ MFG Auto On-Hand Recon Data Extraction Result"
    mail.BodyFormat = 2  # 2: Html format
    # mail.Body = "邮件正文"  # 如不需要HTML格式使用
    body = '<h1>%s Auto Recon Result(BE Excluded)</h1>' % today1

    for item in result:
        temp = ""
        if "captured" in item or "database" in item:
            temp = "<p>" + item + "</p>"
        else:
            temp = '<p style="color:red">' + item + "</p>"
        body += temp
    body += '</br></br></br>'
    body+='<h5>Gross $ Accracy rate:</h5>'

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
    body = '<html>%s</html>' % body
    #print(body)
    mail.HTMLBody = body
    # mail.Attachments.Add(attachment_url)
    # 添加多个附件
    # mail.Attachments.Add("附件1绝对路径")
    # mail.Attachments.Add("附件2绝对路径")...
    # 邮件发送
    mail.Send()

'''

#采用smtp模式
def send_result(result, to_address, cc_address, accuracy):
    SERVER = "143.166.224.194"
    sender = "gabriel_zhang@dell.com"
    subject = "(Test)APJ MFG Auto On-Hand Recon Data Extraction Result"


    body = '<h1>%s Auto Recon Result(BE Excluded)</h1>' % today1

    for item in result:
        temp = ""
        if "captured" in item or "database" in item:
            temp = "<p>" + item + "</p>"
        else:
            temp = '<p style="color:red">' + item + "</p>"
        body += temp
    body += '</br></br></br>'
    body+='<h5>Gross $ Accracy rate:</h5>'

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
    body = '<html>%s</html>' % body
    body += '<a href="https://dscpowerbi.dell.com/reports/powerbi/DSC/Inventory/PROD/MFG_AutoRecon_DashBoard">PBI Link</a>'
    body+='<p style="color:green">If any question, can talk to Gabriel!</p>'
    message = MIMEText(body, 'html', 'utf-8')
    message['From'] = Header('MFG_Daily_Auto_Recon', 'utf-8')
    message['To'] = Header(';'.join(to_address), 'utf-8')
    message['Cc'] = Header(';'.join(cc_address), 'utf-8')
    message['Subject'] = Header(subject, 'utf-8')
    try:
        smtp_mail = smtplib.SMTP(SERVER)
        smtp_mail.sendmail(sender, to_address+cc_address, message.as_string())
        logger.info('Mail already send out!')
    except smtplib.SMTPException:
        logger.error("异常信息", exc_info=True)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # 打印日志设置
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    # LOG_FORMAT = "%(asctime)s - %(levelname)s - %(filename)s, line:%(lineno)d - %(message)s" #带有文件名、行号
    DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
    # 创建Logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    # 输出到文件
    today = datetime.datetime.now().strftime('%Y_%m_%d')
    #today = '2021_09_25'
    file_name = 'Log/Job_Log_{0}.log'.format(today)
    file_handler = logging.FileHandler(file_name, mode='a', encoding='utf-8')
    # 输出到控制台
    stream_handler = logging.StreamHandler()

    # 错误日志单独输出到一个文件
    # error_handler = logging.FileHandler('error.log', mode='a', encoding='utf-8')
    # error_handler.setLevel(logging.ERROR)
    # 注意这里，错误日志只记录ERROR级别的日志
    # 将所有的处理器加入到logger中
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    # logger.addHandler(error_handler)
    formatter = logging.Formatter(fmt=LOG_FORMAT, datefmt=DATE_FORMAT)
    # 设置格式化
    file_handler.setFormatter(formatter)
    stream_handler.setFormatter(formatter)
    # error_handler.setFormatter(formatter)
    # 打印日志
    # logger.debug('debug 信息')
    # logger.info('info 信息')
    # logger.warning('warn 信息')
    # logger.error('error 信息')
    # logger.critical('critical 信息')
    #list_all_site = pd.DataFrame(
    #    columns=['Recon_Site', 'Recon_date', 'Total_Amount', 'Issue_Part_Gross_Amt', 'Issue_Part_Net_Amt',
    #             'Total_Line_Qty', 'Issue_Line_Qty', 'Issue_Line_Percentage', 'Gross_Amount_Accuracy',
    #            'Net_Amount_Accuracy'])

    list_all_site = pd.DataFrame(
        columns=['recon_site', 'recon_date', 'total_amount', 'issue_part_gross_amt', 'issue_part_net_amt',
                 'total_line_qty', 'issue_line_qty', 'issue_line_percentage', 'gross_amount_accuracy',
                'net_amount_accuracy'])
    result = []
    gross_accuracy = {}
    read_c6(list_all_site, result, gross_accuracy)
    #read_apcc(list_all_site, result, gross_accuracy)
    # list_all=pd.concat([list_all_site,read_c2(list_all_site)],axis=0) #如果一天返回多个文件，需要抓每条文件的数据，就要用这个~~
    #read_c2(list_all_site, result, gross_accuracy)
    #read_c4(list_all_site, result, gross_accuracy)
    #read_icc(list_all_site, result, gross_accuracy)
    list_all_site.fillna(value=0, inplace=True)
    try:


        #conn = create_engine('mssql+pymssql://Inv:f$msFT7_&#$!@ctupwcc6role2/Inv')
        conn = create_engine(
            'postgresql+psycopg2://' + 'gabriel_zhang' + ':' + 'dc37d0dd' + '@ddlgpmprd11.us.dell.com' + ':' + str(
                6420) + '/' + 'gp_ns_ddl_prod')
        list_all_site.to_sql(name='mfg_recon_data', con=conn, if_exists='append', index=False,schema='ws_go_invn')
        logger.info("Today's result uploaded to database!")
        result.append("Today's result uploaded to database!")
        final_result = list(set(result))
        #to_address = "gabriel_zhang@dell.com;Mars_Liu@Dell.com;Muthukannan_P@DELL.com;Jayaraj_Raj@Dell.com;Caihua_Ke@Dell.com;Liuqiang_Wang@Dell.com"
        to_address =["gabriel_zhang@dell.com"]
        cc_address =["gabriel_zhang@dell.com"]
        #cc_address = "Lay_Ching_Tan@dell.com;Lant_Zhou@Dell.com"
        send_result(result, to_address, cc_address, gross_accuracy)

    except:
        logger.error("异常信息", exc_info=True)
