# -*- coding: utf-8 -*-


import auto_count_single as acs
import os

temp_excel_list = acs.get_excel_list('./模板/temp')

for excel in temp_excel_list:  # 删除上次运行时生成的临时excel文件
    os.remove('./模板/temp/' + excel)

excel_list = acs.excel_to_excel('迎新人员信息.xlsx')

# print(excel_list)
for n, excel in enumerate(excel_list):
    print(n, excel)
    acs.excel_to_word('./模板/temp/' + excel,
                    the_thing='迎新志愿者',
                    the_date='二〇二一年九月二十日',
                    the_n=n,
                    template='.\\模板\\迎新志愿活动.docx')