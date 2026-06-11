#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
俄语7天速成每日发送脚本
根据日期自动发送对应天的课件
"""
import os
import sys
from datetime import datetime

LESSON_DIR = "/Users/mac/.openclaw/workspace/courses/俄语课件"
START_DATE = datetime(2026, 6, 11)  # Day 1 = 6月11日

def get_today_lesson():
    """根据日期计算今天的课件"""
    today = datetime.now()
    days_passed = (today - START_DATE).days
    current_day = (days_passed % 7) + 1  # 1-7循环
    
    lesson_file = os.path.join(LESSON_DIR, "俄语速成_Day%d.pptx" % current_day)
    
    if not os.path.exists(lesson_file):
        return None, "Day%d课件不存在" % current_day
    
    return lesson_file, current_day

if __name__ == "__main__":
    lesson_path, day = get_today_lesson()
    print("今日课件: Day %d" % day)
    print("路径:", lesson_path)
