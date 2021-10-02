# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-18 19:50:15
LastEditTime: 2021-10-02 23:18:19
LastEditors: Lumen
Description:
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
"""

import os
from typing import List, NoReturn, Dict
from math import ceil  # 向上取整

import pandas as pd
from docxtpl import DocxTemplate
from pandas.core.frame import DataFrame


def excel_to_excel(old_excel: str,
                   temp_path: str='./模板/temp') -> List[str]:
    """将excel表格转换成适合使用的新excel表格

    Args:
        old_excel (str): 初始统计表格，应将所有信息放置在同一工作表中
        temp_path (str, optional): 生成的中间excel表格保存路径. Defaults to './模板/temp'.

    Returns:
        list: 生成的excel表格列表
    """
    if not os.path.exists(temp_path):
        os.makedirs(temp_path)

    frame: DataFrame = pd.read_excel(old_excel)  # 载入需要转换的excel表格

    frame['年级'] = frame['专业班级'].str[2:4]  # 切分班级列，方便按要求排序
    frame['年级'] = frame['年级'].map(lambda x: int(x))

    frame['个人班级'] = frame['专业班级'].str[4:]
    frame['个人班级'] = frame['个人班级'].map(lambda x: int(x))

    frame['专业'] = frame['专业班级'].str[:2]

    frame = frame.sort_values(by=['年级', '专业', '个人班级'], ascending=True)  # 排序

    college_grouping: DataFrame = frame.groupby([frame['学院']])  # 按照时间和学院进行分组
    college_grouping_list: List[DataFrame] = []  # 创建新的分组表

    for i in college_grouping:  # 向分组表添加新分组
        college_grouping_list.append(i)
    # 根据长度分组
    for i in range(len(college_grouping_list)):  # 创建临时excel表，并且设置表格居中
        df = pd.DataFrame(college_grouping_list[i][1])
        df = df.loc[:, ~df.columns.str.contains('Unnamed')]  # 去除unnamed列
        name: str = str(college_grouping_list[i][0])
        max_raw: int = df.shape[0]
        block: int = ceil(max_raw / 18)  # 向上取整
        # print(max_raw, block)

        for x in range(block):
            if x == block-1:
                new_df = df[x*18:max_raw]
                #print(new_df)
                writer = pd.ExcelWriter(f'./模板/temp/{name}-{i}.{x+1}.xlsx', engine='xlsxwriter')  # 居中保存进excel
                new_df = new_df.style.set_properties(**{'text-align': "center"})
                new_df.to_excel(writer, sheet_name='Sheet1')
                writer.save()
            else:
                new_df = df[x*18:(x+1)*18]
                #print(new_df)
                writer = pd.ExcelWriter(f'./模板/temp/{name}-{i}.{x+1}.xlsx', engine='xlsxwriter')  # 居中保存进excel
                new_df = new_df.style.set_properties(**{'text-align': "center"})
                new_df.to_excel(writer, sheet_name='Sheet1')
                writer.save()

    new_excel_list: List[str] = get_excel_list("./模板/temp")  # 生成的临时excel文件名列表

    return new_excel_list


def excel_to_word(excel_name: str,
                  the_thing: str,
                  the_date: str,
                  the_n: int,
                  template: str,
                  root: str = '.\\') -> NoReturn:
    """将符合要求的excel文件转换成模板word文件

    Args:
        excel_name (str): 需要转换的excel
        the_thing (str): 活动事项
        the_date (str): 请假条制作日期
        the_n (int): 避免重复，给定不重复数字
        template (str): 模板路径
        root (str, optional): 保存路径. Defaults to '.\'.
    """
    if not os.path.exists(root + the_thing):
        os.makedirs(root + the_thing)

    sheet: DataFrame = pd.read_excel(excel_name)
    name_list: List[str] = []  # 姓名列表
    class_list: List[str] = []  # 班级列表

    college_name: List[str] = list(sheet['学院'])[0]

    tpl: DocxTemplate = DocxTemplate(template)
    name_list: List[str] = list(sheet['姓名'])
    class_list: List[str] = list(sheet['专业班级'])

    for i in range(len(name_list)):  # 两个字的姓名与三个字姓名对齐
        if len(name_list[i]) == 2:
            name_list[i] = name_list[i][0] + "  " + name_list[i][-1]

    if len(name_list) < 18:  # 填充空白
        for i in range(18 - len(name_list)):
            name_list.append('')

    if len(class_list) < 18:  # 填充空白
        for i in range(18 - len(class_list)):
            class_list.append('')

    context: Dict[str, str] = {
        'college_name': college_name,
        'date': the_date,
    }
    # 用空白填充模板多余部分，不可省略
    for i in range(1, 19):
        context['cell{}1'.format(i)] = class_list[i-1]
        context['cell{}2'.format(i)] = name_list[i-1]


    tpl.render(context=context)
    tpl.save(root + the_thing + '\\' + college_name + the_thing + '-' + str(the_n + 1) + '.docx')


def get_excel_list(path: str) -> List[str]:
    """获取路径下的excel文件

    Args:
        path (str): 路径

    Returns:
        list: 路径下的excel列表
    """
    excel_lists: List[str] = []

    for i in os.listdir(path):
        if str(i).endswith('.xlsx'):
            excel_lists.append(i)

    return excel_lists

