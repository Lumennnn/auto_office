# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-19 12:18:45
LastEditTime: 2021-11-10 15:51:46
LastEditors: Lumen
Description: 活动请假条制作

👻👻👻👻👻👻👻👻👻👻👻👻👻👻
"""
import os
from collections import Counter
from math import ceil  # 向上取整
from typing import Dict, List, NoReturn

import pandas as pd
from docxtpl import DocxTemplate
from loguru import logger
from pandas.core.frame import DataFrame
from pathos.pools import ProcessPool as Pool

logger.add("runing.log", retention="30 days", enqueue=True)


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


def get_excel_list(path: str) -> List[str]:
    """获取路径下的所有excel文件名称

    Args:
        path (str): 查找路径

    Returns:
        List[str]: 路径下的所有excel文件名称
    """
    excel_lists: List[str] = []

    for i in os.listdir(path):
        if str(i).endswith(".xlsx"):
            excel_lists.append(i)

    return excel_lists


def check_data_frame_by_column(frame: DataFrame, type: str) -> bool:
    """检查DataFrame的指定列是否符合规范

    Args:
        frame (DataFrame): 传入的DataFrame表格
        type (str): 检查类型

    Returns:
        bool: 此类型是否正确
    """
    is_right: bool = True

    if type == "columns":
        df_columns: set = set(frame)
        right_columns: set = set(["学院", "专业班级", "姓名", "时间"])
        if not right_columns.issubset(df_columns):
            logger.error("检查列名是否符合规范")
            is_right = False
    elif type == "times":
        right_time: set = set(["（", "）"])
        times: List[str] = list(frame["时间"])
        for index, time in enumerate(times):
            if not right_time.issubset(set(time)):
                logger.error(f"检查时间格式是否符合规范(使用中文括号)->行号:{index + 2}")
                is_right = False
        time_types = dict(Counter(times))
        if len(time_types) > (0.5 * len(times)):
            logger.error(f"检查时间格式是否符合规范(时间段范围出错)->行号:{index + 2}")
            is_right = False
    elif type == "class_name":
        class_names: List[str] = list(frame["专业班级"])
        for index, class_name in enumerate(class_names):
            if len(class_name) != 6:
                logger.error(f"检查专业班级是否符合规范(不符合长度限制)->行号:{index + 2}")
                is_right = False
    elif type == "names":
        names: List[str] = list(frame["姓名"])
        for index, name in enumerate(names):
            if len(name) > 5 or len(name) < 2:
                logger.error(f"检查姓名长度是否符合规范(不符合长度限制)->行号:{index + 2}")
                is_right = False

    return is_right


def check_data_frame(data_frame: DataFrame):
    """检查传入的DataFrame是否符合规范

    Args:
        data_frame (DataFrame): 传入的DataFrame

    Returns:
        [type]: 表格是否符合规范
    """
    is_all_right = True
    p = Pool()
    type_list = ["columns", "times", "class_name", "names"]
    data_list = [data_frame for _ in range(len(type_list))]

    is_right_list = p.amap(check_data_frame_by_column, data_list, type_list).get()

    # 接下来还要使用进程池，因此不进行关闭操作
    # 执行完close后不会有新的进程加入到pool,join函数等待所有子进程结束
    # p.close()
    # p.join()

    for i in is_right_list:
        if i == False:
            is_all_right = False

    return is_all_right


def split_data_frame(frame: DataFrame) -> List[DataFrame]:
    """将传入的DataFrame按格式切分成所需类型和大小的DataFrame列表

    Args:
        frame (DataFrame): 传入需要切分的DataFrame

    Returns:
        List[DataFrame]: 返回一个包含DataFrame的列表
    """
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

    time_college_groups: List[DataFrame] = []  # 创建新的分组表

    for i in time_college_grouping:  # 向分组表添加新分组
        time_college_groups.append(i)

    spilt_data_frame_group: List[DataFrame] = []
    # 根据长度切分成合适长度（最大长度18）的DataFrame表格
    for i in range(len(time_college_groups)):
        df: DataFrame = pd.DataFrame(time_college_groups[i][1])
        df = df.loc[:, ~df.columns.str.contains("Unnamed")]  # 去除unnamed列
        max_raw: int = df.shape[0]
        block: int = ceil(max_raw / 18)  # 向上取整

        for x in range(block):
            if x == block - 1:
                new_df: DataFrame = df[x * 18 : max_raw]
                spilt_data_frame_group.append(new_df)
            else:
                new_df: DataFrame = df[x * 18 : (x + 1) * 18]
                spilt_data_frame_group.append(new_df)

    return spilt_data_frame_group


def data_frame_to_word(
    data_frame: DataFrame,
    the_people_type: str,
    the_date1: str,
    the_thing: str,
    the_date2: str,
    the_n: int,
    root: str = "",
) -> NoReturn:
    """将DataFrame填充进Word模板中

    Args:
        data_frame (DataFrame): 传入DataFrame
        the_people_type (str): 人员类型
        the_date1 (str): 请假日期
        the_thing (str): 活动事项
        the_date2 (str): 批假日期
        the_n (int): 一个不重复数字
        root (str, optional): 输出路径，默认为当前文件夹. Defaults to "".

    Returns:
        NoReturn: [description]
    """
    name_list: List[str] = []  # 姓名列表
    class_list: List[str] = []  # 班级列表

    college_name: str = list(data_frame["学院"])[0]
    time: str = list(data_frame["时间"])[0]
    time_quantum: str = list(data_frame["时间段"])[0]

    tpl = DocxTemplate("./模板/请假条程序模板.docx")
    name_list: List[str] = list(data_frame["姓名"])
    class_list: List[str] = list(data_frame["专业班级"])

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
        "peoples_name": the_people_type,
        "date1": the_date1,
        "thing": the_thing,
        "time": time,
        "date2": the_date2,
    }

    for i in range(1, 19):
        context["cell{}1".format(i)] = class_list[i - 1]
        context["cell{}2".format(i)] = name_list[i - 1]

    if not os.path.exists(root + the_thing + "请假条"):
        os.makedirs(root + the_thing + "请假条")

    # 填充模板
    tpl.render(context=context)

    if time_quantum == "未知":
        tpl.save(
            root
            + the_thing
            + "请假条"
            + "\\"
            + college_name
            + the_thing
            + "请假条"
            + "-"
            + str(the_n + 1)
            + ".docx"
        )
    else:
        tpl.save(
            root
            + the_thing
            + "请假条"
            + "\\"
            + college_name
            + the_thing
            + "请假条"
            + time_quantum
            + "-"
            + str(the_n + 1)
            + ".docx"
        )


def data_frame_to_final_word(
    data_frame: DataFrame,
    the_people_type: str,
    the_date1: str,
    the_thing: str,
    the_date2: str,
    root: str = "",
) -> NoReturn:
    """将DataFrame转化为Word文件

    Args:
        data_frame (DataFrame): 传入DataFrame
        the_people_type (str): 人员类型
        the_date1 (str): 请假日期
        the_thing (str): 活动事项
        the_date2 (str): 批假日期
        root (str, optional): 输出路径，默认为当前文件夹. Defaults to "".

    Returns:
        NoReturn: [description]
    """
    p = Pool()

    spilt_data_frame_group = split_data_frame(data_frame)
    the_people_type_list = [the_people_type for _ in range(len(spilt_data_frame_group))]
    the_date1_list = [the_date1 for _ in range(len(spilt_data_frame_group))]
    the_thing_list = [the_thing for _ in range(len(spilt_data_frame_group))]
    the_date2_list = [the_date2 for _ in range(len(spilt_data_frame_group))]
    the_n_list = [i for i in range(len(spilt_data_frame_group))]
    root_list = [root for _ in range(len(spilt_data_frame_group))]

    # 加入进程池（不可以有print输出）
    p.amap(
        data_frame_to_word,
        spilt_data_frame_group,
        the_people_type_list,
        the_date1_list,
        the_thing_list,
        the_date2_list,
        the_n_list,
        root_list,
    )

    p.close()
    p.join()
    p.clear()
