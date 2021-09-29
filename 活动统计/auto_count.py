# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-18 19:50:15
LastEditTime: 2021-09-29 14:34:43
LastEditors: Lumen
Description:
FilePath: \auto_office\活动统计\auto_count.py
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
"""
import os
from math import ceil  # 向上取整

import pandas as pd
from docxtpl import DocxTemplate


def activity_score(sheet: str,
                   date: str,
                   activity: list,
                   root: str = None) -> None:
    """统计参与活动生成活动分证明

    Args:
        sheet (str): 统计表路径，Excel表应为各部门总统计表
        date (str): 制表时间
        activity (list): 活动列表，属于活动分范畴的活动
        root (str, optional): 输出文件路径. Defaults to None.
    """
    sheet_names = list(pd.read_excel(sheet, sheet_name=None).keys())
    print(sheet_names)
    for sheet_name in sheet_names:
        print(sheet_name)
        data = pd.read_excel(sheet, sheet_name=sheet_name)
        personal_information = ['学院', '专业班级', '姓名', '学号']
        # 为本年度需要参与活动分证明的活动

        columns = personal_information + activity

        data = pd.DataFrame(data, columns=columns)
        the_name = '奖学金证明'
        if not os.path.exists(root + '.\\' + the_name + '.\\' + sheet_name):
            os.makedirs(root + '.\\' + the_name + '.\\' + sheet_name)

        for x in range(data.shape[0]):
            tpl = DocxTemplate('.\\模板\\奖学金活动证明模板.docx')
            person = []  # 个人信息
            things = []  # 个人活动事项

            for y in range(4, data.shape[1]):
                # print(x, y, data.iloc[x, y]) # 定位出错点
                if data.iloc[x, y] > 0:
                    things.append([data.columns[y], int(data.iloc[x, y])])

            # if len(things) == 0: # 去除未参加活动人员信息
            #     continue
            person = [data.iloc[x, 0], data.iloc[x, 1], data.iloc[x, 2], data.iloc[x, 3]]
            # 判断姓名位数
            if type(person[2]) is not float and 1 < len(person[2]) < 3:
                person[2] = list(person[2])[0] + '  ' + list(person[2])[-1]
            print(person, things, sep='\n')
            # 用空白填充模板多余部分，不可省略
            if len(things) < 20:
                for i in range(20 - len(things)):
                    things.append(['', ''])
            context = {
                'college_name': person[0],
                'class': person[1],
                'name': person[2],
                'p_id': person[3],
                'date': date
            }
            for i in range(1, 20):
                context[f'c{i}1'] = things[i - 1][0]
                if things[i - 1][1] != '':
                    times = things[i - 1][1]
                    if times == 2:
                        cn_times = '两次'
                    elif 10 < times < 20:
                        cn_times = an_2_cn(str(times))[1:] + '次'
                    else:
                        cn_times = an_2_cn(str(times)) + '次'
                    context[f'c{i}2'] = cn_times
                else:
                    context[f'c{i}2'] = ''

            tpl.render(context=context)
            tpl.save(root + '.\\' + the_name + '.\\' + sheet_name + '\\' +
                     f'{person[2]}' + the_name + '.docx')

        print('--' * 20)


def second_class_score(sheet: str,
                       date: str,
                       activity: dict,
                       root: str = None)  -> None:
    """统计参与活动生成第二课堂分证明

    Args:
        sheet (str): 统计表路径，Excel表应为各部门总统计表
        date (str): 制表时间
        activity (dict): 活动字典，键为属于第二课堂分范畴的活动名称，值为对应活动单次参加所对应的天数
        root (str, optional): 输出文件路径. Defaults to None.
    """
    sheet_names = list(pd.read_excel(sheet, sheet_name=None).keys())
    print(sheet_names)
    for sheet_name in sheet_names:
        print(sheet_name)
        data = pd.read_excel(sheet, sheet_name=sheet_name)
        personal_information = ['学院', '专业班级', '姓名', '学号']
        # 为本年度需要参与第二课堂证明的活动
        activitys = list(activity.keys())
        times = list(activity.values())
        columns = personal_information + activitys

        data = pd.DataFrame(data, columns=columns)
        the_name = '第二课堂证明'
        if not os.path.exists(root + '.\\' + the_name + '.\\' + sheet_name):
            os.makedirs(root + '.\\' + the_name + '.\\' + sheet_name)

        for x in range(data.shape[0]):
            tpl = DocxTemplate('.\\模板\\第二课堂活动证明模板.docx')
            person = []  # 个人信息
            things = []  # 个人活动事项

            for y in range(4, data.shape[1]):
                # print(x, y, data.iloc[x, y]) # 定位出错点
                if data.iloc[x, y] > 0:
                    things.append([data.columns[y], int(data.iloc[x, y])])

            if len(things) == 0:  # 去除未参加活动人员信息
                continue

            person = [data.iloc[x, 0], data.iloc[x, 1], data.iloc[x, 2], data.iloc[x, 3]]
            # 判断姓名位数
            if type(person[2]) is not float and 1 < len(person[2]) < 3:
                person[2] = list(person[2])[0] + '  ' + list(person[2])[-1]
            print(person, things, sep='\n')
            if len(things) < 20:
                for i in range(20 - len(things)):
                    things.append(['', ''])
            context = {
                'college_name': person[0],
                'class': person[1],
                'name': person[2],
                'p_id': person[3],
                'date': date
            }
            for i in range(1, 20):
                context[f'c{i}1'] = things[i - 1][0]
                if things[i - 1][1] != '':
                    day = ceil(things[i - 1][1] * times[i - 1])
                    if day == 2:
                        cn_day = '两天'
                    elif 10 < day < 20:
                        cn_day = an_2_cn(str(day))[1:] + '天'
                    else:
                        cn_day = an_2_cn(str(day)) + '天'
                    context[f'c{i}2'] = cn_day
                else:
                    context[f'c{i}2'] = ''

            tpl.render(context=context)
            tpl.save(root + '.\\' + the_name + '.\\' + sheet_name + '\\' +
                     f'{person[2]}' + the_name + '.docx')

        print('--' * 20)


def an_2_cn(num_str: str) -> str:
    """将数字转换成中文数字，最多四位数字
    Args:
        num_str (str): 输入字符串数字

    Returns:
        str: 返回中文数字
    """
    result = ""
    han_list = ["零" , "一" , "二" , "三" , "四" , "五" , "六" , "七" , "八" , "九"]
    unit_list = ["", "", "十" , "百" , "千"]
    num_len = len(num_str)
    for i in range(num_len):
        num = int(num_str[i])
        if i!=num_len-1:
            if num!=0:
                result=result+han_list[num]+unit_list[num_len-i]
            else:
                if result[-1]=='零':
                    continue
                else:
                    result=result+'零'
        else:
            if num!=0:
                result += han_list[num]
    return result
