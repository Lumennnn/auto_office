# -*- coding: utf-8 -*-
'''
Author: Lumen
Date: 2021-09-18 19:50:15
LastEditTime: 2021-09-21 21:22:13
LastEditors: Lumen
Description:
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
'''

import auto_count_single as al
import pandas as pd
import os

temp_excel_list = al.get_excel_list('./模板/temp')

for excel in temp_excel_list:  # 删除上次运行时生成的临时excel文件
    os.remove('./模板/temp/' + excel)

excel_list = al.excel_to_excel('迎新人员信息.xlsx')

print(excel_list)
for n, excel in enumerate(excel_list):
    print(n, excel)
    al.excel_to_word('./模板/temp/' + excel, the_thing='迎新志愿者',
                    the_date='二〇二一年九月二十日', the_n=n, template='.\\模板\\迎新志愿活动.docx')