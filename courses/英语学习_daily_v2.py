#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英语学习每日智能发送脚本 v2
根据进度按顺序发送L1-10 Aesop Fables内容
"""

import os
import shutil
from datetime import date

# ===== 课程文件 =====
COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/英语学习今日内容.docx'

# 已有文件
L1_5 = f'{COURSES_DIR}/英语学习_L1-5_AesopFables.docx'
L6_10 = f'{COURSES_DIR}/英语学习_L6-10_AesopFables.docx'

# 进度：学完L1-5和L6-10后，继续其他中级材料
# 当前进度：2026-04-27是第1天

def get_day_index():
    """计算学习进度（第几天）"""
    # 以2026-04-27为第1天
    START = date(2026, 4, 27)
    today = date.today()
    return (today - START).days + 1

def main():
    day = get_day_index()
    print(f"今天是学习第 {day} 天")

    if not os.path.exists(L1_5):
        print(f"错误：找不到 {L1_5}")
        return

    if not os.path.exists(L6_10):
        print(f"错误：找不到 {L6_10}")
        return

    # 分配内容
    # L1-5: 第1-5天
    # L6-10: 第6-10天
    # 第11天起：继续中级材料（后续扩展）

    if day <= 5:
        src = L1_5
        lesson_num = day
        fname = '英语学习_L1-5_AesopFables'
        desc = f'Lessons 1-5（第{day}课进行中）'
    elif day <= 10:
        src = L6_10
        lesson_num = day - 5
        fname = '英语学习_L6-10_AesopFables'
        desc = f'Lessons 6-10（第{lesson_num}课进行中）'
    else:
        # 后续扩展：中级英语材料
        print("L1-10已学完，等待后续内容...")
        src = L6_10
        fname = '英语学习_L6-10_AesopFables'
        desc = 'Lessons 6-10（复习）'

    # 复制到outbound
    shutil.copy2(src, OUTBOUND)
    print(f"✅ 已复制: {fname}")
    print(f"📝 {desc}")
    print(f"📤 发送文件: {OUTBOUND}")

if __name__ == '__main__':
    main()
