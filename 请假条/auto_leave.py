# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-19 12:18:45
LastEditTime: 2021-09-24 13:00:42
LastEditors: Lumen
Description:
FilePath: \auto_office\请假条\auto_leave.py
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
"""

import os
from math import ceil  # 向上取整

import pandas as pd
from docxtpl import DocxTemplate


def excel_to_excel(old_excel: list,
                   temp_path='./模板/temp') -> list:
    """将excel表格转换成适合使用的新excel表格

    Args:
        old_excel (list): 初始统计表格，应将所有信息放置在同一工作表中
        temp_path (str, optional): 生成excel表格保存路径. Defaults to './模板/temp'.

    Returns:
        list: 生成的excel表格列表
    """
    temp_excel_list = get_excel_list(temp_path)
    if temp_excel_list is None:
        print('无临时文件')
    else:
        for excel in temp_excel_list:  # 删除上次运行时生成的临时excel文件
            os.remove(temp_path + excel)

    if not os.path.exists(temp_path):
        os.makedirs(temp_path)
    frame = pd.read_excel(old_excel)  # 载入需要转换的excel表格

    frame['年级'] = frame['专业班级'].str[2:4]  # 切分班级列，方便按要求排序
    frame['年级'] = frame['年级'].map(lambda x: int(x))

    frame['个人班级'] = frame['专业班级'].str[4:]
    frame['个人班级'] = frame['个人班级'].map(lambda x: int(x))

    frame['专业'] = frame['专业班级'].str[:2]

    frame = frame.sort_values(by=['年级', '专业', '个人班级'], ascending=True)  # 排序

    frame['时间段'] = frame.apply(get_time_quantum, axis=1)  # 根据时间段赋值

    time_college_grouping = frame.groupby([frame['时间'], frame['学院']])  # 按照时间和学院进行分组

    time_college_grouping_list = []  # 创建新的分组表

    for i in time_college_grouping:  # 向分组表添加新分组
        time_college_grouping_list.append(i)
    # 根据长度分组
    for i in range(len(time_college_grouping_list)):  # 创建临时excel表，并且设置表格居中
        df = pd.DataFrame(time_college_grouping_list[i][1])
        df = df.loc[:, ~df.columns.str.contains('Unnamed')]  # 去除unnamed列
        name = str(time_college_grouping_list[i][0][1])
        max_raw = df.shape[0]
        block = ceil(max_raw / 18)  # 向上取整
        print(max_raw, block)

        for x in range(block):
            if x == block-1:
                new_df = df[x*18:max_raw]
                print(new_df)
                writer = pd.ExcelWriter(f'./模板/temp/{name}-{i}.{x+1}.xlsx', engine='xlsxwriter')  # 居中保存进excel
                new_df = new_df.style.set_properties(**{'text-align': "center"})
                new_df.to_excel(writer, sheet_name='Sheet1')
                writer.save()
            else:
                new_df = df[x*18:(x+1)*18]
                print(new_df)
                writer = pd.ExcelWriter(f'./模板/temp/{name}-{i}.{x+1}.xlsx', engine='xlsxwriter')  # 居中保存进excel
                new_df = new_df.style.set_properties(**{'text-align': "center"})
                new_df.to_excel(writer, sheet_name='Sheet1')
                writer.save()

    new_excel_list = get_excel_list("./模板/temp")  # 生成的临时excel文件名列表

    return new_excel_list


def get_time_quantum(frame: str) -> str:
    """根据表格内的请假时间来判断请假时间段

    Args:
        frame (str): 请假时间

    Returns:
        str: 时间段
    """
    if frame["时间"] == '半天（8:00-12:00）':
        return "上半天"
    elif frame["时间"] == '半天（14:00-17:50）':
        return "下半天"
    elif frame["时间"] == '一天（8:00-17:50）':
        return "白天"
    elif frame["时间"] == '晚上（19:00-21:00）':
        return "晚上"
    elif frame["时间"] == '一天（8:00-21:00）':
        return "全天"
    else:
        return "未知"


def excel_to_word(excel_name: str,
                  the_people_name: str,
                  the_date1: str,
                  the_thing: str,
                  the_date2: str,
                  the_n: int,
                  root: str = '.\\'):
    """将符合要求的excel文件转换成模板word文件

    Args:
        excel_name (str): 需要转换的excel
        the_people_name (str): 人员类型
        the_date1 (str): 活动日期
        the_thing (str): 活动事项
        the_date2 (str): 请假条制作日期
        the_n (int): 避免重复，给定不重复数字
        root (str, optional): 保存路径. Defaults to '.\'.
    """
    sheet = pd.read_excel(excel_name)
    name_list = []  # 姓名列表
    class_list = []  # 班级列表

    college_name = sheet['学院'][0]
    time = sheet['时间'][0]
    time_quantum = sheet['时间段'][0]
    peoples_name = the_people_name
    date1 = the_date1
    thing = the_thing
    date2 = the_date2
    number = the_n

    tpl = DocxTemplate('.\\模板\\请假条程序套用模板.docx')
    name_list = list(sheet['姓名'])
    class_list = list(sheet['专业班级'])

    for i in range(len(name_list)):  # 两个字的姓名与三个字姓名对齐
        if len(name_list[i]) == 2:
            name_list[i] = name_list[i][0] + "  " + name_list[i][-1]

    if len(name_list) < 18:  # 填充空白
        for i in range(18 - len(name_list)):
            name_list.append('')

    if len(class_list) < 18:  # 填充空白
        for i in range(18 - len(class_list)):
            class_list.append('')

    context = {
        'college_name': college_name,
        'peoples_name': peoples_name,
        'date1': date1,
        'thing': thing,
        'time': time,
        'date2': date2,
    }

    for i in range(1, 19):
        context['cell{}1'.format(i)] = class_list[i-1]
        context['cell{}2'.format(i)] = name_list[i-1]

    if not os.path.exists(root + thing + '请假条'):
        os.makedirs(root + thing + '请假条')

    tpl.render(context=context)

    if time_quantum == '未知':
        tpl.save(root + thing + '请假条' + '\\' + college_name + thing +
                 '请假条' + '-' + str(number + 1) + '.docx')
    else:
        tpl.save(root + thing + '请假条' + '\\' + college_name + thing +
                 '请假条' + time_quantum + '-' + str(number + 1) + '.docx')


def get_excel_list(path: str) -> list:
    """获取路径下的excel文件

    Args:
        path (str): 路径

    Returns:
        list: 路径下的excel列表
    """
    excel_lists = []

    for i in os.listdir(path):
        if str(i).endswith('.xlsx'):
            excel_lists.append(i)
    return excel_lists
