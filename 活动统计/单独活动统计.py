# -*- coding: utf-8 -*-
"""
Author: Lumen
Date: 2021-09-18 19:50:15
LastEditTime: 2022-03-23 15:37:28
LastEditors: Lumen
Description:
π±βππ±βππ±βππ±βππ±βππ±βππ±βππ±βππ±βππ±βπ
"""


import os

import auto_count_single as acs

temp_excel_list = acs.get_excel_list("./source/temp")

for excel in temp_excel_list:  # ε ι€δΈζ¬‘θΏθ‘ζΆηζηδΈ΄ζΆexcelζδ»Ά
    os.remove("./source/temp/" + excel)

excel_list = acs.excel_to_excel("θΏζ°δΊΊεδΏ‘ζ―.xlsx")

# print(excel_list)
for n, excel in enumerate(excel_list):
    print(n, excel)
    acs.excel_to_word(
        "./source/temp/" + excel,
        the_thing="θΏζ°εΏζΏθ",
        the_date="δΊγδΊδΈεΉ΄δΉζδΊεζ₯",
        the_n=n,
        template=".\\source\\θΏζ°εΏζΏζ΄»ε¨ζ¨‘ζΏ.docx",
    )
