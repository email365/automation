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

import process_data

def get_个人业绩csv():
    currentMonth = datetime.now().month
    current_month_file_path = 'Y:/款项登记/18年回款登记/2018-{}月回款/2018-{}月回款.xlsx'.format(currentMonth, currentMonth)
    df_currentMonth = pd.read_excel(current_month_file_path, sheet_name='2018年{}月'.format(currentMonth))
    df_currentMonth['收款日期'] = pd.to_datetime(df_currentMonth['收款日期'], format='%Y-%m-%d')
    df_previous = pd.read_excel('Y:/款项登记/好视通客户成交管理-总表 .xlsx', sheet_name='2018年')
    # budget
    budget = pd.read_excel('Y:/款项登记/18年回款登记/业绩目标.xlsx', sheet_name=0)

    # 离职人员名单
    exit_employees = pd.read_excel('Y:/款项登记/18年回款登记/2018年入职登记表.xlsx', sheet_name='离职登记')
    exit_employees = exit_employees[['离职日期','部门','姓名']]
    exit_employees.离职日期 = pd.to_datetime(exit_employees['离职日期'], format='%Y-%m-%d')
    exit_employees = exit_employees[(exit_employees['离职日期'].dt.month == currentMonth-1)]

    exit_employee_department = set(exit_employees.部门.unique())
    exit_employee_name = set(exit_employees.姓名.unique())

    # df_currentMonth = pd.read_excel(current_month_file_path, sheet_name='2018年{}月'.format(currentMonth))
    当月业绩 = df_currentMonth[['销售人员', '净现金业绩', '所属部门', '新单数']].groupby(['所属部门', '销售人员']).sum()
    当月业绩 = 当月业绩.reset_index()
    当月业绩 = 当月业绩.rename(index=str, columns={'净现金业绩': '{}月净现金业绩'.format(currentMonth)})

    # 获取：个人排名_前一个月份业绩 列表
    # 获取：个人排名_前一个月份业绩 列表
    i = 1
    df_list = []
    while i < currentMonth:
        previous_month_file_path = 'Y:/款项登记/18年回款登记/2018-{}月回款/2018-{}月回款.xlsx'.format(currentMonth - i,
                                                                                       currentMonth - i)
        last_month_df = pd.read_excel(previous_month_file_path, sheet_name='2018年{}月'.format(currentMonth - i))
        last_month_df = last_month_df[['所属部门', '销售人员', '净现金业绩']].groupby(['所属部门', '销售人员']).sum()
        last_month_df = last_month_df.reset_index()
        last_month_df = last_month_df.rename(index=str, columns={'净现金业绩': '{}月净现金业绩'.format(currentMonth - i)})
        #     last_month_df = last_month_df.reset_index()
        df_list.append(last_month_df)
        i += 1

    前单月业绩合并表 = reduce(lambda left, right: pd.merge(left, right, left_on=['所属部门', '销售人员'],
                                                   right_on=['所属部门', '销售人员'], how='outer'), df_list)
    个人业绩YTD = pd.merge(当月业绩, 前单月业绩合并表, left_on=['所属部门', '销售人员'],
                       right_on=['所属部门', '销售人员'], how='outer')
    个人业绩YTD['1-{}月净现金业绩'.format(currentMonth)] = 个人业绩YTD.sum(numeric_only=True, axis=1)
    个人业绩YTD = 个人业绩YTD.reset_index()
    个人业绩YTD = 个人业绩YTD.round(2)
    # 个人业绩YTD.iloc[:, 2:] = 个人业绩YTD.iloc[:, 2:].astype(float)
    个人业绩YTD = 个人业绩YTD.fillna(0)
    个人业绩YTD = 个人业绩YTD.drop(columns=['index'])
    个人业绩YTD = 个人业绩YTD.sort_values(['所属部门','{}月净现金业绩'.format(currentMonth)], ascending=[True,False])
    个人业绩YTD = 个人业绩YTD[~个人业绩YTD['销售人员'].isin(exit_employee_name) & ~个人业绩YTD['销售人员'].isin(exit_employee_department)]

#     个人业绩名细及排名 = get_个人业绩YTD()
    个人业绩YTD.index = np.arange(1, len(个人业绩YTD) + 1)
    #个人业绩YTD.to_csv('个人业绩名细及排名.csv',encoding = 'gbk',index = True)
    writer = pd.ExcelWriter('个人业绩名细及排名.xlsx')
    个人业绩YTD.to_excel(writer,'Sheet1')
    writer.save()


办事处月度完成率 = process_data.get_办事处汇总表()
#办事处月度完成率.to_csv('办事处月度完成率.xlsx',encoding = 'gbk',index = True)
writer1 = pd.ExcelWriter('办事处月度完成率.xlsx')
办事处月度完成率.to_excel(writer1,'Sheet1')
writer1.save()



def get_addresses():
    mail_file = pd.read_excel('emails.xlsx')
    to_addresses1 = []
    for i in mail_file.email:
        to_addresses1.append(i)
    return to_addresses1
    
get_addresses()


def send_email_on_time(address_list):
    path = r'C:/Users/Administrator/Desktop/自动群发邮件/'
    for i in range(len(address_list)):
        msg = MIMEMultipart('related')
        msg['From'] = 'huanglh@hst.com'
        msg['To'] =  address_list[i]
        file_name = '个人业绩名细及排名.xlsx'
        # msg['Subject'] = os.path.basename(file_name)
        msg['Subject'] = '{} 销售体系业绩播报'.format(datetime.now().strftime('%Y-%m-%d'))

        body = "大家好，今日销售体系业绩播报已发布，详见附件。如有疑问，请与我联系，谢谢！"

        msg.attach(MIMEText(body, 'plain'))

        att = MIMEText(open(os.path.join(path, file_name), 'rb').read(), 'base64', 'gbk')
        att["Content-Type"] = 'application/octet-stream'
        att.add_header('Content-Disposition', 'attachment', filename=('gbk', '', file_name))
        msg.attach(att)

        smtp = smtplib.SMTP()
        smtp.connect('smtp.exmail.qq.com', '25')
        smtp.login('huanglh@hst.com', 'Hlh121119.')
        smtp.sendmail(msg['From'], msg['To'], msg.as_string())

def send_email_on_time1(address_list):
    path = r'C:/Users/Administrator/Desktop/自动群发邮件/'
    for i in range(len(address_list)):
        msg = MIMEMultipart('related')
        msg['From'] = 'huanglh@hst.com'
        msg['To'] =  address_list[i]
        file_name = '办事处月度完成率.xlsx'
        # msg['Subject'] = os.path.basename(file_name)
        msg['Subject'] = '{} 销售体系月度完成率'.format(datetime.now().strftime('%Y-%m-%d'))

        body = "大家好，今日销售体系月度完成率播报已发布，详见附件。如有疑问，请与我联系，谢谢！"

        msg.attach(MIMEText(body, 'plain'))

        att = MIMEText(open(os.path.join(path, file_name), 'rb').read(), 'base64', 'gbk')
        att["Content-Type"] = 'application/octet-stream'
        att.add_header('Content-Disposition', 'attachment', filename=('gbk', '', file_name))
        msg.attach(att)

        smtp = smtplib.SMTP()
        smtp.connect('smtp.exmail.qq.com', '***')
        smtp.login('***@***.com', '***.')
        smtp.sendmail(msg['From'], msg['To'], msg.as_string())

def job(t):
    get_个人业绩csv()
    send_email_on_time(get_addresses())
    send_email_on_time1(get_addresses())

schedule.every().day.at("18:20").do(job,'sending emails')

while True:
    schedule.run_pending()
    time.sleep(1)
