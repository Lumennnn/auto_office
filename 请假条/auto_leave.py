# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-19 12:18:45
LastEditTime: 2021-10-27 18:28:25
LastEditors: Lumen
Description:
👻👻👻👻👻👻👻👻👻👻👻👻
"""

import os
import sys
from math import ceil  # 向上取整
from typing import Dict, List, NoReturn

import pandas as pd
from docxtpl import DocxTemplate
from loguru import logger
from pandas.core.frame import DataFrame

logger.add("log.log", retention="30 days")


@logger.catch
def excel_to_excel(old_excel: str, temp_path: str = "./模板/temp/") -> List[str]:
    """将excel表格转换成适合使用的新excel表格

    Args:
        old_excel (str): 初始统计表格，应将所有信息放置在同一工作表中
        temp_path (str, optional): 生成excel表格保存路径. Defaults to './模板/temp'.

    Returns:
        List[str]: 生成的excel表格列表
    """
    if not os.path.exists(temp_path):
        os.makedirs(temp_path)

    temp_excel_list: List[str] = get_excel_list(temp_path)

    if temp_excel_list is None:
        print("无临时文件")
    else:
        for excel in temp_excel_list:  # 删除上次运行时生成的临时excel文件
            os.remove(temp_path + excel)

    frame: DataFrame = pd.read_excel(old_excel)  # 载入需要转换的excel表格

    check_excel(frame=frame)  # 检查表格是否合适

    frame["年级"] = frame["专业班级"].str[2:4]  # 切分班级列，方便按要求排序
    frame["年级"] = frame["年级"].map(lambda x: int(x))

    frame["个人班级"] = frame["专业班级"].str[4:]
    frame["个人班级"] = frame["个人班级"].map(lambda x: int(x))

    frame["专业"] = frame["专业班级"].str[:2]

    frame = frame.sort_values(by=["年级", "专业", "个人班级"], ascending=True)  # 排序

    frame["时间段"] = frame.apply(get_time_quantum, axis=1)  # 根据时间段赋值

    time_college_grouping: DataFrame = frame.groupby(
        [frame["时间"], frame["学院"]]
    )  # 按照时间和学院进行分组

    time_college_grouping_list: List[DataFrame] = []  # 创建新的分组表

    for i in time_college_grouping:  # 向分组表添加新分组
        time_college_grouping_list.append(i)
    # 根据长度分组
    for i in range(len(time_college_grouping_list)):  # 创建临时excel表，并且设置表格居中
        df: DataFrame = pd.DataFrame(time_college_grouping_list[i][1])
        df = df.loc[:, ~df.columns.str.contains("Unnamed")]  # 去除unnamed列
        name: str = str(time_college_grouping_list[i][0][1])
        max_raw: int = df.shape[0]
        block: int = ceil(max_raw / 18)  # 向上取整
        print(max_raw, block)

        for x in range(block):
            if x == block - 1:
                new_df: DataFrame = df[x * 18 : max_raw]
                print(new_df)
                writer = pd.ExcelWriter(
                    f"./模板/temp/{name}-{i}.{x+1}.xlsx", engine="xlsxwriter"
                )  # 居中保存进excel
                new_df = new_df.style.set_properties(**{"text-align": "center"})
                new_df.to_excel(writer, sheet_name="Sheet1")
                writer.save()
            else:
                new_df: DataFrame = df[x * 18 : (x + 1) * 18]
                print(new_df)
                writer = pd.ExcelWriter(
                    f"./模板/temp/{name}-{i}.{x+1}.xlsx", engine="xlsxwriter"
                )  # 居中保存进excel
                new_df = new_df.style.set_properties(**{"text-align": "center"})
                new_df.to_excel(writer, sheet_name="Sheet1")
                writer.save()

    new_excel_list: List[str] = get_excel_list("./模板/temp")  # 生成的临时excel文件名列表
    print("生成的Excel文件列表：\n", new_excel_list)

    return new_excel_list


@logger.catch
def get_time_quantum(frame: DataFrame) -> str:
    """根据表格内的请假时间来判断请假时间段

    Args:
        frame (DataFrame): 请假时间

    Returns:
        str: 时间段
    """
    if frame["时间"] == "半天（8:00-12:00）":
        return "上半天"
    elif frame["时间"] == "半天（14:00-17:50）":
        return "下半天"
    elif frame["时间"] == "一天（8:00-17:50）":
        return "白天"
    elif frame["时间"] == "晚上（19:00-21:00）":
        return "晚上"
    elif frame["时间"] == "一天（8:00-21:00）":
        return "全天"
    else:
        return "未知"


@logger.catch
def excel_to_word(
    excel_name: str,
    the_people_name: str,
    the_date1: str,
    the_thing: str,
    the_date2: str,
    the_n: int,
    root: str = "",
) -> NoReturn:
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
    sheet: DataFrame = pd.read_excel(excel_name)
    name_list: List[str] = []  # 姓名列表
    class_list: List[str] = []  # 班级列表

    college_name: str = sheet["学院"][0]
    time: str = sheet["时间"][0]
    time_quantum: str = sheet["时间段"][0]
    peoples_name: str = the_people_name
    date1: str = the_date1
    thing: str = the_thing
    date2: str = the_date2
    number: int = the_n

    tpl = DocxTemplate("./模板/请假条程序模板.docx")
    name_list: List[str] = list(sheet["姓名"])
    class_list: List[str] = list(sheet["专业班级"])

    for i in range(len(name_list)):  # 两个字的姓名与三个字姓名对齐
        if len(name_list[i]) == 2:
            name_list[i] = name_list[i][0] + "  " + name_list[i][-1]

    if len(name_list) < 18:  # 填充空白
        for i in range(18 - len(name_list)):
            name_list.append("")

    if len(class_list) < 18:  # 填充空白
        for i in range(18 - len(class_list)):
            class_list.append("")

    context: Dict[str, str] = {
        "college_name": college_name,
        "peoples_name": peoples_name,
        "date1": date1,
        "thing": thing,
        "time": time,
        "date2": date2,
    }

    for i in range(1, 19):
        context["cell{}1".format(i)] = class_list[i - 1]
        context["cell{}2".format(i)] = name_list[i - 1]

    if not os.path.exists(root + thing + "请假条"):
        os.makedirs(root + thing + "请假条")

    tpl.render(context=context)

    if time_quantum == "未知":
        tpl.save(
            root
            + thing
            + "请假条"
            + "\\"
            + college_name
            + thing
            + "请假条"
            + "-"
            + str(number + 1)
            + ".docx"
        )
    else:
        tpl.save(
            root
            + thing
            + "请假条"
            + "\\"
            + college_name
            + thing
            + "请假条"
            + time_quantum
            + "-"
            + str(number + 1)
            + ".docx"
        )


@logger.catch
def get_excel_list(path: str) -> List[str]:
    """获取路径下的excel文件

    Args:
        path (str): 路径

    Returns:
        List[str]: 路径下的excel列表
    """
    excel_lists: List[str] = []

    for i in os.listdir(path):
        if str(i).endswith(".xlsx"):
            excel_lists.append(i)
    return excel_lists


@logger.catch
def check_excel(frame: DataFrame) -> NoReturn:
    """检查列表是否符合规范

    Args:
        frame (DataFrame): 传入DataFrame格式表格
    """
    df_columns: set = set(frame)
    right_columns: set = set(["学院", "专业班级", "姓名", "时间"])
    if not right_columns.issubset(df_columns):
        logger.warning("检查列名是否符合规范")
        # print("检查列名是否符合规范")
        sys.exit()

    right_time: set = set(["（", "）"])
    times: List[str] = list(frame["时间"])
    for index, time in enumerate(times):
        if not right_time.issubset(set(time)):
            logger.warning(f"检查时间格式是否符合规范(使用中文括号)->行号:{index + 2}")
            # print(f"检查时间格式是否符合规范(使用中文括号)->行号:{index + 2}")
            sys.exit()

    class_names: List[str] = list(frame["专业班级"])
    for index, class_name in enumerate(class_names):
        if len(class_name) > 6:
            logger.warning(f"检查专业班级是否符合规范(超出长度限制)->行号:{index + 2}")
            # print(f"检查专业班级是否符合规范(超出长度限制)->行号:{index + 2}")
            sys.exit()

    names: List[str] = list(frame["姓名"])
    for index, name in enumerate(names):
        if len(name) > 5:
            logger.warning(f"检查姓名长度是否符合规范(超出长度限制)->行号:{index + 2}")
            # print(f"检查姓名长度是否符合规范(超出长度限制)->行号:{index + 2}")
            sys.exit()
