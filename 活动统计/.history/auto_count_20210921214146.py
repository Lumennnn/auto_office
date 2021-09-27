# -*- coding: utf-8 -*-
'''
Author: Lumen
Date: 2021-09-18 16:43:51
LastEditTime: 2021-09-21 21:41:46
LastEditors: Lumen
Description:
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
'''
import os
import pandas as pd
from docxtpl import DocxTemplate


def activity_score(sheet: str, date: str, activity: list, root: str = None):
    """统计参与活动生成活动分证明

    Args:
        sheet (str): 统计表路径
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
            # if len(data.iloc[x, 2]) == 2: # 两个字的姓名与三个字姓名对齐
            #     data.iloc[x, 2] = data.iloc[x, 2][0] + "  " + data.iloc[x, 2][-1]
            # if len(things) == 0: # 去除未参加活动人员信息
            #     continue
            person = [
                data.iloc[x, 0], data.iloc[x, 1], data.iloc[x, 2], data.iloc[x,
                                                                             3]
            ]
            print(person)
            print(things)

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
                    if times == 1:
                        context[f'c{i}2'] = '一次'
                    elif times == 2:
                        context[f'c{i}2'] = '两次'
                    elif times == 3:
                        context[f'c{i}2'] = '三次'
                    elif times == 4:
                        context[f'c{i}2'] = '四次'
                    elif times == 5:
                        context[f'c{i}2'] = '五次'
                    elif times == 6:
                        context[f'c{i}2'] = '六次'
                    elif times == 7:
                        context[f'c{i}2'] = '七次'
                else:
                    context[f'c{i}2'] = ''

            tpl.render(context=context)
            tpl.save(root + '.\\' + the_name + '.\\' + sheet_name + '\\' +
                     f'{person[2]}' + the_name + '.docx')

        print('--' * 20)


def second_class_score(sheet: str, date: str, activity: dict, root=None):
    """统计参与活动生成第二课堂分证明

    Args:
        sheet (str): 统计表路径
        date (str): 制表时间
        activity (dict): 活动字典，键为属于第二课堂分范畴的活动名称，值为对应活动单次对应的天数
        root ([type], optional): [description]. Defaults to None.
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
            # if len(data.iloc[x, 2]) == 2: # 两个字的姓名与三个字姓名对齐
            #     data.iloc[x, 2] = data.iloc[x, 2][0] + "  " + data.iloc[x, 2][-1]
            if len(things) == 0:  # 去除未参加活动人员信息
                continue
            person = [
                data.iloc[x, 0], data.iloc[x, 1], data.iloc[x, 2], data.iloc[x,
                                                                             3]
            ]
            print(person)
            print(things)

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
                    day = things[i - 1][1] * times[i - 1]
                    if day == 0.5:
                        context[f'c{i}2'] = '半天'
                    elif day == 1:
                        context[f'c{i}2'] = '一天'
                    elif day == 1.5:
                        context[f'c{i}2'] = '两天'
                    elif day == 2:
                        context[f'c{i}2'] = '两天'
                    elif day == 2.5:
                        context[f'c{i}2'] = '三天'
                    elif day == 3:
                        context[f'c{i}2'] = '三天'
                    elif day == 3.5:
                        context[f'c{i}2'] = '四天'
                    elif day == 4:
                        context[f'c{i}2'] = '四天'
                    elif day == 4.5:
                        context[f'c{i}2'] = '五天'
                    elif day == 5:
                        context[f'c{i}2'] = '五天'
                    elif day == 5.5:
                        context[f'c{i}2'] = '六天'
                    elif day == 6:
                        context[f'c{i}2'] = '六天'
                else:
                    context[f'c{i}2'] = ''

            tpl.render(context=context)
            tpl.save(root + '.\\' + the_name + '.\\' + sheet_name + '\\' +
                     f'{person[2]}' + the_name + '.docx')

        print('--' * 20)