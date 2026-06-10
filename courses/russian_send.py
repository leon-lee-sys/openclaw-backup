#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
俄语每日课件发送脚本
根据日期自动计算进度，发送对应天的课件
"""
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

LESSON_DIR = "/Users/mac/.openclaw/workspace/courses/俄语课件"
START_DATE = datetime(2026, 6, 10)  # Day1 = 6月10日

def get_today_lesson():
    """根据日期计算今天的课件"""
    today = datetime.now()
    days_passed = (today - START_DATE).days
    
    current_day = days_passed + 1  # Day1=6月10日
    
    if current_day < 1:
        return None, "学习尚未开始"
    if current_day > 84:
        return None, "已完成全部84课学习！"
    
    lesson_file = os.path.join(LESSON_DIR, "俄语_Day%d_L1.pptx" % current_day)
    
    if not os.path.exists(lesson_file):
        return None, "Day%d课件尚未生成" % current_day
    
    return lesson_file, current_day

def get_lesson_title(day_num):
    """根据天数返回课件标题"""
    titles = {
        1: "名词第1格和第2格",
        2: "名词第3格（给谁）",
        3: "名词第4格（做什么）",
        4: "名词第5格（用什么）",
        5: "名词第6格（关于什么）",
        6: "形容词性数配合（一）",
        7: "形容词性数配合（二）",
    }
    if day_num <= 30:
        return titles.get(day_num, "名词各格学习")
    elif day_num <= 60:
        return "动词学习"
    else:
        return "实用口语"

if __name__ == "__main__":
    lesson_path, day_info = get_today_lesson()
    
    if lesson_path is None:
        print(day_info)
        sys.exit(1)
    
    print("今日课件:", day_info)
    print("路径:", lesson_path)
    print("标题:", get_lesson_title(day_info))
