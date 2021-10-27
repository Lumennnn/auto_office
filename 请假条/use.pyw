"""
Author: Lumen
Date: 2021-09-19 12:18:45
LastEditTime: 2021-10-27 18:29:01
LastEditors: Lumen
Description:
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
"""

import sys

import pandas as pd
from loguru import logger
from pywebio.input import *
from pywebio.output import *

import auto_leave as al


def check_people(people: str) -> str:
    """人员类型检查

    Args:
        people (str): 人员类型

    Returns:
        str: 不在范围内则返回提示
    """
    if people not in ["志愿者", "干部", "干事"]:
        return "确保人员类型在所提供范围内"


def check_none(the: str) -> str:
    """检查是否为空

    Args:
        the (str): 检查字段

    Returns:
        str: 为空则返回提示
    """
    if the is None or "":
        return "值不能为空"


if __name__ == "__main__":
    img1 = open(".\\模板\\bangonshi.jpg", "rb").read()
    img2 = open(".\\模板\\school.png", "rb").read()
    put_image(src=img1, width="770px", height="720px")
    put_markdown(r"""### 使用时注意事项：""")
    put_text("1.确保选择的excel文件内容为以下格式")
    put_table(
        [
            ["部门", "学院", "专业班级", "姓名", "时间"],
            ["办公室", "统计学院", "数据1903", "XXX", "晚上（19:00-21:00）"],
        ]
    )

    put_text("2.确保输入的时间段为以下格式")
    put_table(
        [
            ["序号", "时间段"],
            ["①", "半天（8:00-12:00）"],
            ["②", "半天（14:00-17:50）"],
            ["③", "一天（8:00-17:50）"],
            ["④", "一天（8:00-21:00）"],
            ["⑤", "晚上（19:00-21:00）"],
        ]
    )
    put_text("3.长安校区共有下列学院及专业")
    put_image(src=img2, width="2000px")
    put_text("4.确保输入内容的正确性")
    put_text("------------这是分割线------------")

    excel_list = al.get_excel_list(".")

    excel = radio("选择当前目录下要转换的文件（仅限后缀名为.xlsx的文件）", excel_list)
    excel: str = str(excel)
    print("选择的Excel文件：", excel)
    frame = pd.read_excel(excel, "Sheet1").head(10)
    put_code(frame, language="Python")
    put_text("以上为所选文件前10行信息，确认格式正确后按确认按钮继续")
    confirm = actions("确认继续?", ["继续", "取消"], help_text="请再次确认文件格式正确")
    put_text("------------这还是分割线------------")

    if confirm == "取消":
        sys.exit()
    else:
        get_input = input_group(
            "请假条信息",
            [
                input(
                    "请输入活动参与人（志愿者/干部/干事）",
                    name="people_name",
                    type=TEXT,
                    validate=check_people,
                ),
                input(
                    "请输入活动日期，格式为：2021年5月1日",
                    name="date1",
                    type=TEXT,
                    validate=check_none,
                ),
                input("请输入活动名称", name="thing", type=TEXT, validate=check_none),
                input(
                    "请输入落款日期，格式为：二〇二一年五月一日",
                    name="date2",
                    type=TEXT,
                    validate=check_none,
                ),
            ],
        )
        put_text("这是进度条🗡")
        try:
            put_processbar("bar")
            print("运行中.......")
            to_write_excel = al.excel_to_excel(excel)
            excel_list_len = len(to_write_excel) - 1
            for n, excel in enumerate(to_write_excel):
                set_processbar("bar", n / excel_list_len)
                print(f"\n进度：{n+1}/{excel_list_len+1}")
                al.excel_to_word(
                    "./模板/temp/" + excel,
                    the_people_name=get_input["people_name"],
                    the_date1=get_input["date1"],
                    the_thing=get_input["thing"],
                    the_date2=get_input["date2"],
                    the_n=n,
                )
        except (ValueError, AttributeError, NameError, TypeError) as e:
            logger.exception("出错了！")
            put_text("------------这又是分割线------------")
            put_text("出了一点点点点点小问题！")
        else:
            put_text("------------这又是分割线------------")
            put_markdown(r"""### 程序运行成功，请在程序所在目录查看""")
            print("\n程序运行成功，请在程序所在目录查看")
        finally:
            sys.exit()
