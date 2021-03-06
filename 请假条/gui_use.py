# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2022-03-22 15:45:44
LastEditTime: 2022-04-10 17:59:52
LastEditors: Lumen
Description:
ð»ð»ð»ð»ð»ð»ð»ð»
"""
import PySimpleGUI as sg

import source.auto_leave as al


sg.theme("DefaultNoMoreNagging")


layout2 = [
    [
        sg.Text("æ´»å¨åä¸äºº:", font="éé»"),
        sg.Combo(
            ["å¿æ¿è", "å¹²é¨", "å¹²äº"],
            default_value="å¿æ¿è",
            readonly=True,
            size=(6, 1),
            key="-PeopleName-",
            font="éé»",
        ),
    ],
    [
        sg.Text("æ´»å¨æ¥æ:", tooltip="ç¤ºä¾ï¼2021å¹´5æ1æ¥", font="éé»"),
        sg.Input(key="-Date1-", tooltip="ç¤ºä¾ï¼2021å¹´5æ1æ¥", font="éé»"),
    ],
    [sg.Text("æ´»å¨åç§°:", font="éé»"), sg.Input(key="-Thing-", font="éé»")],
    [
        sg.Text("è½æ¬¾æ¥æ:", tooltip="ç¤ºä¾ï¼äºãäºä¸å¹´äºæä¸æ¥", font="éé»"),
        sg.Input(key="-Date2-", tooltip="ç¤ºä¾ï¼äºãäºä¸å¹´äºæä¸æ¥", font="éé»"),
    ],
    [
        sg.Image(
            filename="./source/èµ.png",
            size=(500, 500),
            subsample=5,
            enable_events=True,
            key="-ç¡®è®¤-",
            tooltip="è¾å¥å®æåæå¯ç¹å»ç¡®è®¤",
        )
    ],
]


layout = [
    [
        sg.Image(
            filename="./source/bangonshi.png",
            size=(640, 601),
            subsample=2,
            enable_events=True,
            key="-Windows1-",
        ),
        sg.Frame(
            layout=layout2, title="è¾å¥ä¿¡æ¯", visible=False, key="-Windows2-", font="éé»"
        ),
    ]
]


window = sg.Window("éå¹´å¿æ¿èèåä¼", layout, icon="./source/ä¼å¾½.ico", size=(350, 350))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == "-Windows1-":
        file_path = sg.popup_get_file(
            "éæ©excelæä»¶ï¼",
            title=" ",  # æ é¢æ 
            default_extension=".xlsx",  # å¦ææ²¡æåç¼ååæ·»å 
            multiple_files=False,  # å¤æä»¶éå
            grab_anywhere=True,  # åè®¸æå¨
            keep_on_top=True,  # å§ç»å¨æä¸å±
            initial_folder=".",  # æå¼å§å¨å½åæä»¶å¤¹å¯»æ¾
            font="éé»",
        )
        # éæ©æä»¶åè¿è¡æ£æ¥
        if file_path != None:
            frame = al.preprocess_excel(file_path)

            if not al.check_data_frame(frame, is_multiprocess=False):  # æ£æ¥è¡¨æ ¼æ¯å¦åé
                sg.popup_no_buttons(
                    "è¡¨æ ¼ä¸­å­å¨æ ¼å¼éè¯¯ï¼è¯·æ£æ¥æ¥å¿æ¥çå·ä½éè¯¯",
                    title=" ",
                    text_color="red",
                    auto_close=True,
                    auto_close_duration=3,
                    font="éé»",
                )
                break
        window["-Windows1-"].update(visible=False)
        window["-Windows2-"].update(visible=True)

    if values["-Date1-"] and values["-Date2-"] and values["-Thing-"]:
        if event == "-ç¡®è®¤-":
            al.data_frame_to_words(
                data_frame=frame,
                the_people_type=values["-PeopleName-"],
                the_date1=values["-Date1-"],
                the_thing=values["-Thing-"],
                the_date2=values["-Date2-"],
            )
        break

window.close()
