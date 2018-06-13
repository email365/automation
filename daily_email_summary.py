import schedule
import time

import pandas as pd
import numpy as np
from datetime import datetime
from functools import reduce
from datetime import date
from os import listdir
from os.path import isfile, join
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys

import email_process_data


def get_addresses(page):
    mail_file = pd.read_excel('emails.xlsx', sheet_name = '{}'.format(page))
    to_addresses1 = []
    for i in mail_file['email']:
        to_addresses1.append(i)
    return to_addresses1


def send_email_on_time(address_list):
    path = r'C:/Users/Administrator/Desktop/自动群发邮件/'
    for i in range(len(address_list)):
        msg = MIMEMultipart('related')
        msg['From'] = 'huanglh@hst.com'
        msg['To'] =  address_list[i]
        file_name = '销售业绩汇总表.xlsx'
        msg['Subject'] = '{} 销售体系新增业绩名细汇总'.format(datetime.now().strftime('%Y-%m-%d'))

        body = "大家好，今日销售体系新增业绩名细汇总已发布，详见附件。如有疑问，请与我联系，谢谢！"

        msg.attach(MIMEText(body, 'plain'))

        att = MIMEText(open(os.path.join(path, file_name), 'rb').read(), 'base64', 'gbk')
        att["Content-Type"] = 'application/octet-stream'
        att.add_header('Content-Disposition', 'attachment', filename=('gbk', '', file_name))
        msg.attach(att)

        smtp = smtplib.SMTP()
        smtp.connect('smtp.exmail.qq.com', '25')
        smtp.login('huanglh@hst.com', 'Hlh121119.')
        smtp.sendmail(msg['From'], msg['To'], msg.as_string())

test = "test"
real = "Sheet2"

def job(t):
    email_process_data.get_汇总表()
    send_email_on_time(get_addresses(real))

schedule.every().day.at("18:25").do(job,'sending emails')

while True:
    schedule.run_pending()
    time.sleep(1)