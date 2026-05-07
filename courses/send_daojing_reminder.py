#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
道德经学习提醒发送脚本
根据日期自动发送当天学习内容
"""

from datetime import date

# 进度映射：Day N = 4月26日+(N-18)天 = 对应章节
# Day 18 = 4月26日 = 58-60章
# Day 19 = 4月27日 = 61-63章
# Day 20 = 4月28日 = 64-66章
# Day 21 = 4月29日 = 67-69章
# Day 22 = 4月30日 = 70-72章
# Day 23 = 5月1日 = 73-75章

DAY_MAP = {
    18: (58, 60, "4月26日"),
    19: (61, 63, "4月27日"),
    20: (64, 66, "4月28日"),
    21: (67, 69, "4月29日"),
    22: (70, 72, "4月30日"),
    23: (73, 75, "5月1日"),
    24: (76, 78, "5月2日"),
    25: (79, 81, "5月3日"),
}

def get_current_day():
    base = date(2026, 4, 26)
    today = date.today()
    return 18 + (today - base).days

def get_message():
    day = get_current_day()
    if day in DAY_MAP:
        cs, ce, date_str = DAY_MAP[day]
        return f"""三先生，道德经学习时间到！🐦

今天是第{day}天（{date_str}），学习第{cs}-{ce}章。

请开始学习！📚"""
    else:
        return f"""三先生，道德经学习时间到！🐦

今天是第{day}天，请确认章节后开始学习。📚"""

if __name__ == '__main__':
    print(get_message())
