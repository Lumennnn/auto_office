# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-18 19:50:15
LastEditTime: 2022-03-23 15:40:02
LastEditors: Lumen
Description:
ð±âðð±âðð±âðð±âðð±âðð±âðð±âðð±âðð±âðð±âð
"""
import os
from typing import List, Dict, NoReturn
from math import ceil  # åä¸åæ´

import pandas as pd
from docxtpl import DocxTemplate
from pandas.core.frame import DataFrame


def activity_score(
    sheet: str, date: str, activity: List[str], root: str = ""
) -> NoReturn:
    """ç»è®¡åä¸æ´»å¨çææ´»å¨åè¯æ

    Args:
        sheet (str): ç»è®¡è¡¨è·¯å¾ï¼Excelè¡¨åºä¸ºåé¨é¨æ»ç»è®¡è¡¨
        date (str): å¶è¡¨æ¶é´
        activity (list): æ´»å¨åè¡¨ï¼å±äºæ´»å¨åèç´çæ´»å¨
        root (str, optional): è¾åºæä»¶è·¯å¾. Defaults to None.
    """
    sheet_names: List[str] = list(pd.read_excel(sheet, sheet_name=None).keys())
    print(sheet_names)
    for sheet_name in sheet_names:
        sheet_name: str
        print(sheet_name)
        data: DataFrame = pd.read_excel(sheet, sheet_name=sheet_name)
        personal_information: List[str] = ["å­¦é¢", "ä¸ä¸ç­çº§", "å§å", "å­¦å·"]
        # ä¸ºæ¬å¹´åº¦éè¦åä¸æ´»å¨åè¯æçæ´»å¨

        columns: List[str] = personal_information + activity

        data = pd.DataFrame(data, columns=columns)
        the_name: str = "å¥å­¦éè¯æ"
        if not os.path.exists(root + ".\\" + the_name + ".\\" + sheet_name):
            os.makedirs(root + ".\\" + the_name + ".\\" + sheet_name)

        for x in range(data.shape[0]):
            tpl = DocxTemplate(".\\source\\å¥å­¦éæ´»å¨è¯ææ¨¡æ¿.docx")
            person: List[str] = []  # ä¸ªäººä¿¡æ¯
            things: List[List[str, int]] = []  # ä¸ªäººæ´»å¨äºé¡¹

            for y in range(4, data.shape[1]):
                # print(x, y, data.iloc[x, y]) # å®ä½åºéç¹
                if data.iloc[x, y] > 0:
                    things.append([data.columns[y], int(data.iloc[x, y])])

            # if len(things) == 0: # å»é¤æªåå æ´»å¨äººåä¿¡æ¯
            #     continue
            person = [
                data.iloc[x, 0],
                data.iloc[x, 1],
                data.iloc[x, 2],
                data.iloc[x, 3],
            ]
            # å¤æ­å§åä½æ°
            if (type(person[2]) is not float) and (1 < len(person[2]) < 3):
                person[2] = list(person[2])[0] + "  " + list(person[2])[-1]
            print(person, things, sep="\n")
            # ç¨ç©ºç½å¡«åæ¨¡æ¿å¤ä½é¨åï¼ä¸å¯çç¥
            if len(things) < 20:
                for i in range(20 - len(things)):
                    things.append(["", ""])
            context: Dict[str, str] = {
                "college_name": person[0],
                "class": person[1],
                "name": person[2],
                "p_id": person[3],
                "date": date,
            }
            for i in range(1, 20):
                context[f"c{i}1"] = things[i - 1][0]
                if things[i - 1][1] != "":
                    times = things[i - 1][1]
                    if times == 2:
                        cn_times = "ä¸¤æ¬¡"
                    elif 10 < times < 20:
                        cn_times = an_2_cn(str(times))[1:] + "æ¬¡"
                    else:
                        cn_times = an_2_cn(str(times)) + "æ¬¡"
                    context[f"c{i}2"] = cn_times
                else:
                    context[f"c{i}2"] = ""

            tpl.render(context=context)
            tpl.save(
                root
                + ".\\"
                + the_name
                + ".\\"
                + sheet_name
                + "\\"
                + f"{person[2]}"
                + the_name
                + ".docx"
            )

        print("--" * 20)


def second_class_score(
    sheet: str, date: str, activity: Dict[str, int], root: str = ""
) -> NoReturn:
    """ç»è®¡åä¸æ´»å¨çæç¬¬äºè¯¾å åè¯æ

    Args:
        sheet (str): ç»è®¡è¡¨è·¯å¾ï¼Excelè¡¨åºä¸ºåé¨é¨æ»ç»è®¡è¡¨
        date (str): å¶è¡¨æ¶é´
        activity (dict): æ´»å¨å­å¸ï¼é®ä¸ºå±äºç¬¬äºè¯¾å åèç´çæ´»å¨åç§°ï¼å¼ä¸ºå¯¹åºæ´»å¨åæ¬¡åå æå¯¹åºçå¤©æ°
        root (str, optional): è¾åºæä»¶è·¯å¾. Defaults to None.
    """
    sheet_names: List[str] = list(pd.read_excel(sheet, sheet_name=None).keys())
    print(sheet_names)
    for sheet_name in sheet_names:
        sheet_name: str
        print(sheet_name)
        data = pd.read_excel(sheet, sheet_name=sheet_name)
        personal_information: List[str] = ["å­¦é¢", "ä¸ä¸ç­çº§", "å§å", "å­¦å·"]
        # ä¸ºæ¬å¹´åº¦éè¦åä¸ç¬¬äºè¯¾å è¯æçæ´»å¨
        activitys: List[str] = list(activity.keys())
        times: List[int] = list(activity.values())
        columns: List[str] = personal_information + activitys

        data = pd.DataFrame(data, columns=columns)
        the_name: str = "ç¬¬äºè¯¾å è¯æ"
        if not os.path.exists(root + ".\\" + the_name + ".\\" + sheet_name):
            os.makedirs(root + ".\\" + the_name + ".\\" + sheet_name)

        for x in range(data.shape[0]):
            tpl: DocxTemplate = DocxTemplate(".\\source\\ç¬¬äºè¯¾å æ´»å¨è¯ææ¨¡æ¿.docx")
            person = []  # ä¸ªäººä¿¡æ¯
            things = []  # ä¸ªäººæ´»å¨äºé¡¹

            for y in range(4, data.shape[1]):
                # print(x, y, data.iloc[x, y]) # å®ä½åºéç¹
                if data.iloc[x, y] > 0:
                    things.append([data.columns[y], int(data.iloc[x, y])])

            if len(things) == 0:  # å»é¤æªåå æ´»å¨äººåä¿¡æ¯
                continue

            person = [
                data.iloc[x, 0],
                data.iloc[x, 1],
                data.iloc[x, 2],
                data.iloc[x, 3],
            ]
            # å¤æ­å§åä½æ°
            if type(person[2]) is not float and 1 < len(person[2]) < 3:
                person[2] = list(person[2])[0] + "  " + list(person[2])[-1]
            print(person, things, sep="\n")
            if len(things) < 20:
                for i in range(20 - len(things)):
                    things.append(["", ""])
            context = {
                "college_name": person[0],
                "class": person[1],
                "name": person[2],
                "p_id": person[3],
                "date": date,
            }
            for i in range(1, 20):
                context[f"c{i}1"] = things[i - 1][0]
                if things[i - 1][1] != "":
                    day = ceil(things[i - 1][1] * times[i - 1])
                    if day == 2:
                        cn_day = "ä¸¤å¤©"
                    elif 10 < day < 20:
                        cn_day = an_2_cn(str(day))[1:] + "å¤©"
                    else:
                        cn_day = an_2_cn(str(day)) + "å¤©"
                    context[f"c{i}2"] = cn_day
                else:
                    context[f"c{i}2"] = ""

            tpl.render(context=context)
            tpl.save(
                root
                + ".\\"
                + the_name
                + ".\\"
                + sheet_name
                + "\\"
                + f"{person[2]}"
                + the_name
                + ".docx"
            )

        print("--" * 20)


def an_2_cn(num_str: str) -> str:
    """å°æ°å­è½¬æ¢æä¸­ææ°å­ï¼æå¤åä½æ°å­
    Args:
        num_str (str): è¾å¥å­ç¬¦ä¸²æ°å­

    Returns:
        str: è¿åä¸­ææ°å­
    """
    result: str = ""
    han_list: List[str] = ["é¶", "ä¸", "äº", "ä¸", "å", "äº", "å­", "ä¸", "å«", "ä¹"]
    unit_list: List[str] = ["", "", "å", "ç¾", "å"]
    num_len: int = len(num_str)
    for i in range(num_len):
        num = int(num_str[i])
        if i != num_len - 1:
            if num != 0:
                result = result + han_list[num] + unit_list[num_len - i]
            else:
                if result[-1] == "é¶":
                    continue
                else:
                    result = result + "é¶"
        else:
            if num != 0:
                result += han_list[num]
    return result
