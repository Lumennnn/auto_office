# -*- coding: utf-8 -*-
'''
Author: Lumen
Date: 2021-09-18 18:15:43
LastEditTime: 2021-09-19 23:08:32
LastEditors: Lumen
Description:
🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍🐱‍🏍
'''
import auto_count as ac
import time

if __name__ == '__main__':
    date = '二〇二一年九月十六日'
    sheet1 = '长安活动统计表.xlsx'
    sheet2 = '翠雁活动统计表.xlsx'

    start = time.time()
    activity_dict = {'拉拉手': 0.5, '双选会': 2, '爱在心中': 0.5, '爱心义卖': 1, '亲情陪伴': 0.5, '垃圾分类': 0.5, '绿植领养': 0.5, '光盘行动': 0.5,
                    '红色行动': 1, '联谊晚会': 1, '免费午餐': 0.5, '总结大会': 0.5, '陕广大家帮': 0.5, '流浪狗之家': 0.5, '植树节活动': 1, '校区交流会': 0.5,
                    '素拓志愿者': 1, '三农文化周': 0.5, '体检志愿者': 1, '地铁志愿者': 0.5, '团青科晚会': 1, '血车十四分队': 1, '防疫系列活动': 2,
                    '国际志愿者日': 0.5,  '新老生交流会': 0.5, '让爱不再流浪': 0.5, '卫生巾互助盒': 0.5, '女青系列活动': 0.5, '十周年系列活动': 0.5,
                    '博物馆志愿服务': 2, '十四运系列活动': 0.5, '助残月系列活动': 0.5, '三下乡系列活动': 0.5, '预防肺结核宣传': 0.5, '急救知识宣传活动': 0.5,
                    '防艾知识系列活动': 0.5, '趣味运动会志愿者': 2}
    # 第二课堂证明按天进行计算
    ac.second_class_score(sheet=sheet1, date=date, activity=activity_dict, root='长安')
    ac.second_class_score(sheet=sheet2, date=date, activity=activity_dict, root='翠雁')

    activity_list = ['拉拉手', '双选会', '爱心义卖', '爱在心中', '亲情陪伴', '清扫校园', '点亮心心', '绿植领养',
                    '爱心收集', '红色行动', '垃圾分类', '免费午餐','陕广大家帮', '流浪狗之家', '体检志愿者',
                    '植树节活动', '卫生巾互助盒', '血车十四分队', '防疫系列活动','十四运系列活动',
                    '博物馆志愿服务', '助残月系列活动', '防艾知识系列互动', '急救知识宣传活动']
    # 活动分证明按次进行计算
    ac.activity_score(sheet=sheet1, date=date, activity=activity_list, root='长安')
    ac.activity_score(sheet=sheet2, date=date, activity=activity_list, root='翠雁')

    end = time.time()
    print(end - start)