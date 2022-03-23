# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-18 19:50:15
LastEditTime: 2022-03-23 15:37:28
LastEditors: Lumen
Description:
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
"""


import os

import auto_count_single as acs

temp_excel_list = acs.get_excel_list("./source/temp")

for excel in temp_excel_list:  # 删除上次运行时生成的临时excel文件
    os.remove("./source/temp/" + excel)

excel_list = acs.excel_to_excel("迎新人员信息.xlsx")

# print(excel_list)
for n, excel in enumerate(excel_list):
    print(n, excel)
    acs.excel_to_word(
        "./source/temp/" + excel,
        the_thing="迎新志愿者",
        the_date="二〇二一年九月二十日",
        the_n=n,
        template=".\\source\\迎新志愿活动模板.docx",
    )
