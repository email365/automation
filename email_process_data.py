import pandas as pd
from datetime import datetime
from functools import reduce
currentMonth = datetime.now().month
from datetime import date
lastyear_lastday = date(date.today().year-1, 12, 31)
today = date.today().isoformat()


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


# helper function: return report columns
def make_column(targetColumn, columnName, data):
    df_target = data[['所属部门', targetColumn]].groupby('所属部门').sum()/10000
    df_target = df_target.reset_index()
    df_target = df_target.rename(index=str, columns={targetColumn: columnName})
    df_target = df_target.round(2)
    return df_target


def get_汇总表():
    # 本月收款
    本月收款 = make_column('本期收款', '本月收款', df_currentMonth)
    # 软件金额
    软件金额 = make_column('净现金业绩', '软件金额', df_currentMonth)
    # 单数_新签代理
    单数_新签代理 = df_currentMonth[['所属部门', '新单数']].groupby('所属部门').sum()
    单数_新签代理 = 单数_新签代理.reset_index()
    # 当日业绩
    df_currentMonth['收款日期'] = pd.to_datetime(df_currentMonth['收款日期'], format='%Y-%m-%d')
    df_today = df_currentMonth.loc[(df_currentMonth.收款日期 == today)]
    当日业绩 = make_column('净现金业绩', '当日业绩', df_today)

    currentMonth_column_name = '{}月'.format(currentMonth)
    budget_currentMonth = budget[['部门', currentMonth_column_name, '2018年预测']]
    budget_currentMonth = budget_currentMonth.rename(index=str,
                                                     columns={'部门': '所属部门', budget_currentMonth.columns[1]: '月度任务'})
    # 之前月份累计收款
    之前月总收款 = make_column('本期收款', '之前月总收款', df_previous)
    # 之前月总业绩
    之前月总业绩 = make_column('净现金业绩', '之前月总业绩', df_previous)
    # 合并表
    dfs = [软件金额, budget_currentMonth, 单数_新签代理, 本月收款, 当日业绩, 之前月总收款, 之前月总业绩]
    合并表 = reduce(lambda left, right: pd.merge(left, right, on='所属部门', how='left'), dfs)
    合并表 = 合并表.fillna(0)

    # 年度累计收款
    合并表['年度累计收款'] = 合并表['本月收款'] + 合并表['之前月总收款']
    # 年度累计业绩
    合并表['年度累计业绩'] = 合并表['软件金额'] + 合并表['之前月总业绩']
    汇总 = 合并表.round(2)
    汇总 = 汇总.append(汇总.sum(numeric_only=True), ignore_index=True)
    汇总.iloc[-1, 0] = '汇总'
    汇总['月度完成率(%)'] = 汇总['软件金额'].div(汇总['月度任务']).multiply(100).round(2)
    汇总['年度达成率(%)'] = 汇总['年度累计业绩'].div(汇总['2018年预测']).multiply(100).round(2)
    汇总 = 汇总.round(2)
    汇总 = 汇总.astype(str)
    汇总 = 汇总.replace('inf',0)
    汇总 = 汇总.replace('nan', 0)
    汇总 = 汇总.replace('0.0', 0)
    汇总['月度完成率(%)'] = 汇总['月度完成率(%)'].astype(float)
    #汇总.to_csv('销售业绩汇总表.csv',encoding = 'gbk',index = False)
    writer = pd.ExcelWriter('销售业绩汇总表.xlsx')
    汇总.to_excel(writer,'Sheet1')
    writer.save()