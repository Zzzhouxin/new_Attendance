"""
============================
# -*- coding: utf-8 -*-
# @Time    : 2024/6/5 22:02
# @Author  : zhouxin
# @FileName: main.py
# @Software: PyCharm
===========================
"""
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment
import pandas as pd
import logging

# 配置logging模块
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class Employee:
    """
    员工基类
    """

    def __init__(self, equipment_id=None, employee_id=None, name=None, department=None):
        self.Equipment_ID = equipment_id  # 设备工号
        self.Employee_ID_number = employee_id  # 员工工号
        self.Employee_name = name  # 姓名
        self.Employee_department = department  # 部门
        self.Attendance_Record = {}  # 考勤记录
        self.Attendance_Result = {
            'Numbers_of_late_arrivals': 0,  # 迟到次数
            "Numbers_of_early_departures": 0,  # 早退次数
            "Numbers_of_not_clocked_in_at_noon": 0,  # 中午未打卡次数
            "Number_of_absences": 0,  # 缺勤次数
            "Duration_of_study_in_the_laboratory": 0,  # 在实验室学习时长
        }  # 考情统计结果


class EmployeeManager:
    """
    员工管理类
    """

    def __init__(self):
        self.employees = {}

    def add_employee(self, employee):
        """添加员工"""
        self.employees[employee.Employee_name] = employee


def load_data():
    """
    读取本地的打卡记录文件数据
    :return:
    """
    # 初始化员工管理类
    manager = EmployeeManager()

    file_path = "考勤记录5月-李志老师.xlsx"
    attendance_data = pd.read_excel(file_path)
    for index, row in attendance_data.iterrows():
        attendance_data_dict = row.to_dict()

        if attendance_data_dict['姓名'] not in manager.employees:
            new_employee = Employee()
            new_employee.Equipment_ID = attendance_data_dict['设备工号']
            new_employee.Employee_ID_number = attendance_data_dict['员工工号']
            new_employee.Employee_name = attendance_data_dict['姓名']
            new_employee.Employee_department = attendance_data_dict['部门']

            manager.add_employee(new_employee)

            # 添加打卡记录
            attendance_time_dict = {}
            data_times = attendance_data_dict['考勤时间'].split("  ")
            # for data_time in data_times:
            attendance_time_dict[attendance_data_dict['考勤日期']] = data_times

            new_employee.Attendance_Record.update(attendance_time_dict)

        else:
            # 添加打卡记录
            attendance_time_dict = {}
            data_times = attendance_data_dict['考勤时间'].split("  ")
            attendance_time_dict[attendance_data_dict['考勤日期']] = data_times
            manager.employees[attendance_data_dict['姓名']].Attendance_Record.update(attendance_time_dict)

    logging.info("考勤数据加载完成")

    return manager


def load_work_dates_from_excel():
    """
    从Excel文件中读取需要上班的日期。

    参数:
    file_path (str): Excel文件的路径

    返回:
    list: 包含需要上班的日期的列表，日期格式为YYYY-MM-DD
    """
    # 读取Excel文件
    df = pd.read_excel("上班时间.xlsx", header=None)

    # 需要上班的日期在第一列
    work_dates = df.iloc[:, 0].tolist()

    # 确保日期格式为字符串 YYYY-MM-DD
    work_dates = [date.strftime('%Y-%m-%d') if not pd.isna(date) else '' for date in
                  pd.to_datetime(work_dates, errors='coerce')]
    work_dates = [date for date in work_dates if date]

    return work_dates


def bool_in_time_duration(time_record, start_time, end_time):
    """
    判断time_record中有没有记录在起止时间内
    :param time_record:
    :param start_time:
    :param end_time:
    :return:
    """
    start_time = datetime.strptime(start_time, '%H:%M:%S').time()
    end_time = datetime.strptime(end_time, '%H:%M:%S').time()

    filtered_times = []
    for time_str in time_record:
        time_obj = datetime.strptime(time_str, '%H:%M:%S').time()
        if start_time <= time_obj <= end_time:
            filtered_times.append(time_str)

    if len(filtered_times):
        return True
    else:
        return False


def Check_in_status_in_the_morning(employee, work_day):
    """
    分析7-10点间的打卡记录情况
    :param employee:
    :param work_day:
    :return:
    """
    time_record = employee.Attendance_Record[work_day]

    if bool_in_time_duration(time_record=time_record, start_time='07:00:00', end_time='09:00:00'):
        # 早上正常打卡
        pass

    elif bool_in_time_duration(time_record=time_record, start_time='09:00:00', end_time='10:00:00'):
        # 早上打卡迟到
        employee.Attendance_Result['Numbers_of_late_arrivals'] += 1

    else:
        # 缺勤
        employee.Attendance_Result['Number_of_absences'] += 1


def Check_in_status_in_the_moon(employee, work_day):
    time_record = employee.Attendance_Record[work_day]

    if bool_in_time_duration(time_record=time_record, start_time='12:50:00', end_time='13:30:00'):
        # 中午正常打卡
        pass

    elif bool_in_time_duration(time_record=time_record, start_time='13:30:00', end_time='14:30:00'):
        # 中午打卡迟到
        employee.Attendance_Result['Numbers_of_late_arrivals'] += 1

    else:
        # 中午未打卡
        employee.Attendance_Result['Numbers_of_not_clocked_in_at_noon'] += 1


def Check_in_status_in_the_night(employee, work_day):
    time_record = employee.Attendance_Record[work_day]

    if bool_in_time_duration(time_record=time_record, start_time='18:00:00', end_time='23:59:59'):
        # 晚上正常打卡
        pass

    elif bool_in_time_duration(time_record=time_record, start_time='17:00:00', end_time='18:00:00'):
        # 晚上早退
        employee.Attendance_Result['Numbers_of_early_departures'] += 1

    else:
        # 晚上早退
        employee.Attendance_Result['Number_of_absences'] += 1


def get_Duration_of_study(employee, work_day):
    time_record = employee.Attendance_Record[work_day]

    if len(time_record) < 2:
        pass
    else:
        # 将时间字符串转换为datetime对象
        start_time = datetime.strptime(time_record[0], '%H:%M:%S')
        end_time = datetime.strptime(time_record[-1], '%H:%M:%S')

        # 计算时间差
        time_diff = end_time - start_time

        # 将时间差转换为小时数，保留两位小数
        # 如果上午和下午都有打卡记录，时间-2，如果没有则不变
        time_record = employee.Attendance_Record[work_day]
        if (bool_in_time_duration(time_record=time_record, start_time='12:50:00', end_time='23:59:59') and
            bool_in_time_duration(time_record=time_record, start_time='07:00:00', end_time='11:00:00')):
            hours_diff = round(time_diff.total_seconds() / 3600, 2) - 2
        else:
            hours_diff = round(time_diff.total_seconds() / 3600, 2)

        # 如果只有一次打卡记录 不能为负值
        if hours_diff < 0:
            hours_diff = 0

        employee.Attendance_Result['Duration_of_study_in_the_laboratory'] += hours_diff
        employee.Attendance_Result['Duration_of_study_in_the_laboratory'] = round(
            employee.Attendance_Result['Duration_of_study_in_the_laboratory'], 2)


def get_output_result(employees, work_days):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    header = ["姓名", "迟到次数", "早退次数", "中午未打卡次数", "缺勤次数", "在实验室学习时长"]

    sheet.append(header)

    alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 遍历第一行的所有单元格并应用居中对齐
    for cell in sheet[1]:
        cell.alignment = alignment

    for stu in Employees.employees:
        attendance_result = Employees.employees[stu].Attendance_Result
        sheet.append([
            stu,
            attendance_result['Numbers_of_late_arrivals'],
            attendance_result['Numbers_of_early_departures'],
            attendance_result['Numbers_of_not_clocked_in_at_noon'],
            attendance_result['Number_of_absences'],
            attendance_result['Duration_of_study_in_the_laboratory']
        ])

    result_filename = "统计结果.xlsx"
    workbook.save(result_filename)


if __name__ == "__main__":

    Employees = load_data()

    Work_days = load_work_dates_from_excel()

    # 一天一天的检查打卡记录
    for work_day in Work_days:

        for employee in Employees.employees:

            stu = Employees.employees[employee]

            if work_day not in stu.Attendance_Record:
                stu.Attendance_Result['Number_of_absences'] += 2
                stu.Attendance_Result['Numbers_of_not_clocked_in_at_noon'] += 1

            else:
                Check_in_status_in_the_morning(employee=stu, work_day=work_day)

                Check_in_status_in_the_moon(employee=stu, work_day=work_day)

                Check_in_status_in_the_night(employee=stu, work_day=work_day)

                get_Duration_of_study(stu, work_day)

    get_output_result(Employees, Work_days)

    logging.info("统计结果输出完成")