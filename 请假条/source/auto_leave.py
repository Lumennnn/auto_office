# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-19 12:18:45
LastEditTime: 2022-03-25 17:18:42
LastEditors: Lumen
Description: 活动请假条制作小程序

👻👻👻👻👻👻👻👻👻👻👻👻👻👻
"""
import os
from collections import Counter
from math import ceil
from typing import Dict, List
from multiprocessing import Pool

import pandas as pd
from docxtpl import DocxTemplate
from loguru import logger
from pandas.core.frame import DataFrame


logger.add("runing.log", retention="30 days", enqueue=True)


def rm_all_space_and_special_character(sth: str) -> str:
    """去除多余空格和特殊字符,只保留中文字符

    Args:
        sth (str): 传入的字符串

    Returns:
        str: 去除后的字符串
    """
    new_str = ""
    for s in sth:
        if "\u4e00" <= s <= "\u9fa5":
            new_str += s

    return new_str


def preprocess_excel(excel_path: str) -> DataFrame:
    """对excel文件进行预处理

    Args:
        excel_path (str): 文件路径

    Returns:
        DataFrame: _description_
    """
    frame = pd.read_excel(excel_path)
    # 去除空行
    frame.dropna(how="all", inplace=True)
    # 填充空值
    frame.fillna(value="空", inplace=True)
    # 去除列名中的多余空格和非中文字符
    old_columns = frame.columns
    new_columns = []
    for col in old_columns:
        new_columns.append(rm_all_space_and_special_character(col))
    frame.columns = new_columns
    # 去除姓名列的多余空格和非中文字符
    frame["姓名"] = frame["姓名"].transform(rm_all_space_and_special_character)

    return frame


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


def check_data_frame_by_column(frame: DataFrame, col_type: str) -> bool:
    """检查DataFrame的指定列是否符合规范

    Args:
        frame (DataFrame): 传入的DataFrame表格
        col_type (str): 检查类型

    Returns:
        bool: 此类型是否正确
    """
    is_right: bool = True

    if col_type == "times":
        right_time: set = set(["（", "）"])
        time_list: List[str] = list(frame["时间"])
        for index, time in enumerate(time_list):
            if len(time) < 2 or not right_time.issubset(set(time)):
                logger.error(f"检查时间格式是否符合规范(应使用中文括号)->行号:{index + 2}")
                is_right = False
        time_types = dict(Counter(time_list))
        if len(time_types) > (0.5 * len(time_list)):
            logger.error(f"检查时间格式是否符合规范(时间段范围数量出错)")
            is_right = False
    elif col_type == "class_name":
        class_names: List[str] = list(frame["专业班级"])
        for index, name in enumerate(class_names):
            if len(name) != 6:
                logger.error(f"检查专业班级是否符合规范(不符合长度限制)->行号:{index + 2}")
                is_right = False
    elif col_type == "names":
        names: List[str] = list(frame["姓名"])
        for index, name in enumerate(names):
            if len(name) > 5 or len(name) < 2:
                logger.error(f"检查姓名长度是否符合规范(不符合长度限制)->行号:{index + 2}")
                is_right = False

    return is_right


def check_data_frame(data_frame: DataFrame, is_multiprocess: bool = True):
    """检查传入的DataFrame是否符合规范

    Args:
        data_frame (DataFrame): 传入的DataFrame
        is_multiprocess (bool)=True: 是否启动多进程

    Returns:
        [type]: 表格是否符合规范
    """

    # 检查列名称
    df_columns: set = set(data_frame)
    right_columns: set = set(["学院", "专业班级", "姓名", "时间"])
    if not right_columns.issubset(df_columns):
        logger.error("检查列名是否符合规范")
        return False

    type_list = ["times", "class_name", "names"]
    is_right_list: list[bool] = []

    if is_multiprocess:
        p = Pool()
        data_list = [data_frame for _ in range(len(type_list))]
        zip_args = list(zip(data_list, type_list))

        is_right_list = p.starmap(check_data_frame_by_column, zip_args)

        # # 执行完close后不会有新的进程加入到pool,join函数等待所有子进程结束
        p.close()
        p.join()
    else:
        for _type in type_list:
            is_right_list.append(check_data_frame_by_column(data_frame, _type))

    return all(is_right_list)


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

    time_college_grouping = frame.groupby(["时间", "学院"])  # 按照时间和学院进行分组

    time_college_groups = []  # 创建新的分组表

    for _ in time_college_grouping:  # 向分组表添加新分组
        time_college_groups.append(_)

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
) -> None:
    """将DataFrame填充进Word模板并保存Word文件

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

    tpl = DocxTemplate("./source/请假条程序模板-2.docx")
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


def data_frame_to_words(
    data_frame: DataFrame,
    the_people_type: str,
    the_date1: str,
    the_thing: str,
    the_date2: str,
    root: str = "",
) -> None:
    """批量将DataFrame转化为Word文件

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

    spilt_data_frame_group = split_data_frame(data_frame)

    for i in range(len(spilt_data_frame_group)):
        data_frame_to_word(
            spilt_data_frame_group[i],
            the_people_type,
            the_date1,
            the_thing,
            the_date2,
            i,
            root,
        )
