#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英语学习每日智能发送脚本 v3
根据进度按顺序发送L1-15 Aesop Fables内容
从Day2开始发送L6-10（三先生确认的中级内容）
"""

import os
import shutil
from datetime import date

# ===== 课程文件 =====
COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/英语学习今日内容.docx'

# 文件列表（按顺序）
FILES = {
    1: (f'{COURSES_DIR}/英语学习_L1-5_AesopFables.docx', 'L1-5 初级'),
    2: (f'{COURSES_DIR}/英语学习_L6-10_AesopFables.docx', 'L6-10 中级'),
    3: (f'{COURSES_DIR}/英语学习_L11-15_AesopFables.docx', 'L11-15 中级'),
}

def get_day_index():
    """计算学习进度（第几天）- 从2026-04-27开始"""
    START = date(2026, 4, 27)
    today = date.today()
    return (today - START).days + 1

def main():
    day = get_day_index()
    print(f"今天是学习第 {day} 天")

    # 分配内容：L1-5是Day1，L6-10是Day2，L11-15是Day3+
    if day == 1:
        src, desc = FILES[1]
    elif day == 2:
        src, desc = FILES[2]
    else:
        src, desc = FILES[3]

    if not os.path.exists(src):
        print(f"错误：找不到 {src}")
        return

    shutil.copy2(src, OUTBOUND)
    print(f"✅ 已复制: {desc}")
    print(f"📤 发送文件: {OUTBOUND}")

if __name__ == '__main__':
    main()