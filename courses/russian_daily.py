#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
俄语每日课件发送脚本
根据日期自动计算进度，发送对应天的课件
"""
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, '/opt/homebrew/lib/node_modules/openclaw/node_modules')

LESSON_DIR = "/Users/mac/.openclaw/workspace/courses/俄语课件"
START_DATE = datetime(2026, 6, 11)  # 6月11日开始

def get_today_lesson():
    """根据日期计算今天的课件"""
    today = datetime.now()
    days_passed = (today - START_DATE).days
    
    # 当前进度：1=Day1已发送，正在学习Day2
    current_day = days_passed + 2  # Day1已发送，从Day2开始
    
    if current_day > 84:
        return None, "已完成全部84课学习！"
    
    lesson_file = os.path.join(LESSON_DIR, "俄语_Day%d_L1.pptx" % current_day)
    
    if not os.path.exists(lesson_file):
        return None, "Day%d课件尚未生成" % current_day
    
    return lesson_file, "Day%d" % current_day

if __name__ == "__main__":
    lesson_path, day_info = get_today_lesson()
    print("今日课件:", day_info)
    if lesson_path:
        print("路径:", lesson_path)
    else:
        print(day_info)
