#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英语学习每日生成并发送 - 直接发飞书
"""

import os
import shutil
from datetime import date

COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/英语学习今日内容.docx'

FILES = {
    1: (f'{COURSES_DIR}/英语学习_L1-5_AesopFables.docx', 'L1-5 初级'),
    2: (f'{COURSES_DIR}/英语学习_L6-10_AesopFables.docx', 'L6-10 中级'),
    3: (f'{COURSES_DIR}/英语学习_L11-15_AesopFables.docx', 'L11-15 中级'),
    4: (f'{COURSES_DIR}/英语学习_L16-20_AesopFables.docx', 'L16-20 进阶'),
    5: (f'{COURSES_DIR}/英语学习_L21-25_AesopFables.docx', 'L21-25 进阶'),
    6: (f'{COURSES_DIR}/英语学习_L26-30_AesopFables.docx', 'L26-30 高级'),
    7: (f'{COURSES_DIR}/英语学习_L31-35_AesopFables.docx', 'L31-35 高级'),
}

def get_day_index():
    START = date(2026, 4, 27)
    today = date.today()
    return (today - START).days + 1

def main():
    day = get_day_index()
    print(f"今天是学习第 {day} 天")

    if day in FILES:
        src, desc = FILES[day]
    else:
        # 超出7天的循环到第7天
        cycle_day = ((day - 1) % 7) + 1
        src, desc = FILES[cycle_day]
        print(f"循环到第 {cycle_day} 天 (共7个级别)")

    if not os.path.exists(src):
        print(f"错误：找不到 {src}")
        return

    shutil.copy2(src, OUTBOUND)
    print(f"✅ 已复制: {desc}")
    print(f"📤 文件路径: {OUTBOUND}")
    print(f"📤 请发送附件: {OUTBOUND} 给三先生")

if __name__ == '__main__':
    main()