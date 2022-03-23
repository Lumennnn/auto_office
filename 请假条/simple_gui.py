# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2022-03-22 15:45:44
LastEditTime: 2022-03-23 15:36:48
LastEditors: Lumen
Description:
👻👻👻👻👻👻👻👻
"""
import asyncio

import PySimpleGUI as sg

import request_package.auto_leave as al


sg.theme("DefaultNoMoreNagging")


layout2 = [
    [
        sg.Text("活动参与人:", font="雅黑"),
        sg.Combo(
            ["志愿者", "干部", "干事"],
            default_value="志愿者",
            readonly=True,
            size=(6, 1),
            key="-People_name-",
            font="雅黑",
        ),
    ],
    [
        sg.Text("活动日期:", tooltip="示例：2021年5月1日", font="雅黑"),
        sg.Input(key="-Date1-", tooltip="示例：2021年5月1日", font="雅黑"),
    ],
    [sg.Text("活动名称:", font="雅黑"), sg.Input(key="-Thing-", font="雅黑")],
    [
        sg.Text("落款日期:", tooltip="示例：二〇二一年五月一日", font="雅黑"),
        sg.Input(key="-Date2-", tooltip="示例：二〇二一年五月一日", font="雅黑"),
    ],
    [
        sg.Image(
            "./source/赞.png",
            size=(500, 500),
            subsample=5,
            enable_events=True,
            key="-确认-",
            tooltip="输入完成后才可点击确认",
        )
    ],
]


layout = [
    [
        sg.Image(
            "./source/bangonshi.png",
            size=(640, 601),
            subsample=2,
            enable_events=True,
            key="-Windows1-",
        ),
        sg.Frame(
            layout=layout2, title="输入信息", visible=False, key="-Windows2-", font="雅黑"
        ),
    ]
]


window = sg.Window("青年志愿者联合会", layout, icon="./source/会徽.ico", size=(350, 350))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == "-Windows1-":
        file_path = sg.popup_get_file(
            "选择excel文件：",
            title=" ",  # 标题栏
            default_extension=".xlsx",  # 如果没有后缀名则添加
            multiple_files=False,  # 多文件选取
            grab_anywhere=True,  # 允许拖动
            keep_on_top=True,  # 始终在最上层
            initial_folder=".",  # 最开始在当前文件夹寻找
            font="雅黑",
        )
        # 选择文件后进行检查
        if file_path != None:
            frame = al.preprocess_excel(file_path)

            if not al.check_data_frame(frame, is_multiprocess=False):  # 检查表格是否合适
                sg.popup_no_buttons(
                    "表格中存在格式错误，请检查日志查看具体错误",
                    title=" ",
                    text_color="red",
                    auto_close=True,
                    auto_close_duration=3,
                    font="雅黑",
                )
                break
        window["-Windows1-"].update(visible=False)
        window["-Windows2-"].update(visible=True)

    if values["-Date1-"] and values["-Date2-"] and values["-Thing-"]:
        if event == "-确认-":
            asyncio.run(
                al.data_frame_to_final_word(
                    data_frame=frame,
                    the_people_type=values["-People_name-"],
                    the_date1=values["-Date1-"],
                    the_thing=values["-Thing-"],
                    the_date2=values["-Date2-"],
                )
            )
        break

window.close()
