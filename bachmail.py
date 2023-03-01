#!/usr/local/bin/python3
# -*- coding: UTF-8 -*-
# @Time    : 2023/03/01
# @Author  : Chen Jiang
# @Mail    : chenjiang@microshield.com.cn

import time
import xlrd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header

# Open the Excel file and select the first sheet
workbook = xlrd.open_workbook('/Users/chenjiang/Desktop/SE2023H1-SE调薪计划-V2.2-20230228.xlsx')
#workbook = xlrd.open_workbook('test.xlsx')
worksheet = workbook.sheet_by_index(0)
sender_email='chenjiang@microshield.com.cn'

# Loop through each row and get the email address
for row in range(2, worksheet.nrows):
    name = worksheet.cell_value(row, 0)
    receiver_email = str(worksheet.cell_value(row, 7)).replace(u'\xa0','')
    sem_value = worksheet.cell_value(row, 8)
    cc_email=[]
    cc_email.extend(sem_value.split(','))
    rm_email = worksheet.cell_value(row, 9)
    cc=[receiver_email]
    if cc_email != '':
        cc.extend(cc_email)
    if rm_email != '':
        cc.append(rm_email)
    #cc.append('chenjiang@msmail.com.cn')
    cc.append('ken.chen@microshield.com.cn')
    currpay=worksheet.cell_value(row, 1)
    willpay=worksheet.cell_value(row, 5)
    rate= "{}%".format(round((willpay-currpay)*100/currpay,1))

    # Define the email message
    subject = '%s 2023H1调薪' %(name)
    body='Hi! %s\n\n感谢你过去一年为麦弗瑞公司所做出的贡献.  根据你的杰出表现, 从2023年1月1日起，你每月基本薪水将由原来的%s调整到%s (增长%s). 具体调整将反映到2月份实发的1月工资里.  如有疑问，可跟财务部门尚姐或我及时沟通。\n\n新的一年里，希望你继续提高业务能力, 努力服务好客户, 跟麦弗瑞一起成长.\n\nBR!\n\nChen Jiang\n\nMicroshield Technology Co., Ltd\n\n北京市海淀区西三环北路50号豪柏大厦C2座18-19层 100048\n\n(86)10-88518768\n\n(86)18612696123' %(name,currpay,willpay,rate)
    message = MIMEMultipart()


    # Add the message body to the message object
    message.attach(MIMEText(body, 'plain', 'utf-8'))

    # Set the email headers
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject
    message['Cc'] = ','.join(cc)
    print(message)

    # Log in to the SMTP server and send the email
    try:
        with smtplib.SMTP('smtp.qiye.aliyun.com', 25) as server:
            server.login(sender_email, 'C5h2e1n5')
            server.sendmail(sender_email, cc, message.as_string())
            server.quit()
            print("Email sent successfully!")
            time.sleep(20)
    except smtplib.SMTPException as error:
       print(error)
        