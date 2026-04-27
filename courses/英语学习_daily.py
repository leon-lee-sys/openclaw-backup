#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英语学习每日智能发送脚本
根据星期几自动选择不同的学习内容
"""

import os
import sys
from datetime import date

# ===== 每周内容安排（初级→中级，5天循环） =====
# Week 1: Beginner Basics (5 days)
# 周一: W1D1 - Daily Routine (日常生活)
# 周二: W1D2 - Simple Conversations (简单对话)
# 周三: W1D3 - At the Store (商店场景)
# 周四: W1D4 - Asking for Directions (问路)
# 周五: W1D5 - Making Appointments (预约)
# 然后 Week 2, Week 3...

# 计算今天是第几天（从2026-04-08算起）
START_DATE = date(2026, 4, 8)  # 周三
TODAY = date.today()
DAYS_SINCE_START = (TODAY - START_DATE).days

# 确定今天是周几 (0=周一, 1=周二, ..., 6=周日)
WEEKDAY = TODAY.weekday()

# 确定当前是第几周的第几天
DAY_OF_CYCLE = DAYS_SINCE_START % 35  # 7天循环
WEEK_NUM = (DAYS_SINCE_START // 35) + 1
DAY_IN_WEEK = DAY_OF_CYCLE % 5  # 0-4 对应周一到周五

# 学习内容列表（初级5天主题）
TOPICS = [
    ('W1D1', 'Daily Routine（日常生活）'),
    ('W1D2', 'Simple Conversations（简单对话）'),
    ('W1D3', 'At the Store（商店场景）'),
    ('W1D4', 'Asking for Directions（问路）'),
    ('W1D5', 'Making Appointments（预约）'),
]

# 第二周
TOPICS_W2 = [
    ('W2D1', 'At the Restaurant（餐厅）'),
    ('W2D2', 'Talking about Work（谈工作）'),
    ('W2D3', 'Weather & Seasons（天气与季节）'),
    ('W2D4', 'Hobbies & Interests（爱好）'),
    ('W2D5', 'Family & Friends（家庭与朋友）'),
]

# 第三周（Presentation Skills）
TOPICS_W3 = [
    ('W3D1', 'How to Start a Presentation（演讲开场）'),
    ('W3D2', 'Presenting Data（数据展示）'),
    ('W3D3', 'Handling Questions（应对提问）'),
    ('W3D4', 'Closing a Presentation（演讲结尾）'),
    ('W3D5', 'Business Meeting（商务会议）'),
]

ALL_TOPICS = [TOPICS, TOPICS_W2, TOPICS_W3]
WEEK_IDX = (WEEK_NUM - 1) % 3
CURRENT_TOPICS = ALL_TOPICS[WEEK_IDX]

topic_id, topic_name = CURRENT_TOPICS[DAY_IN_WEEK]

print(f"今日日期: {TODAY} ({['周一','周二','周三','周四','周五','周六','周日'][WEEKDAY]})")
print(f"当前周期: Week {WEEK_NUM}, Day {DAY_IN_WEEK+1}")
print(f"学习内容: {topic_id} - {topic_name}")

# 检查是否有对应的脚本
SCRIPT_PATH = f'/Users/mac/.openclaw/workspace/courses/英语学习_{topic_id}.py'
DOCX_PATH = f'/Users/mac/.openclaw/workspace/courses/英语学习_{topic_id}_docx.py'

# 尝试生成文档
generated = False

# 先检查是否有专门的生成脚本
if os.path.exists(DOCX_PATH):
    print(f"找到生成脚本: {DOCX_PATH}")
    result = os.system(f'/tmp/docx_venv2/bin/python "{DOCX_PATH}"')
    if result == 0:
        generated = True

# 如果没有专门的脚本，尝试通用脚本
if not generated:
    gen_script = f'/Users/mac/.openclaw/workspace/courses/英语学习_gen.py'
    if os.path.exists(gen_script):
        # 修改通用脚本生成当天内容
        print(f"使用通用生成器...")
        # 这里可以添加通用生成逻辑
    else:
        print(f"警告: 没有找到 {topic_id} 的生成脚本")

# 复制到outbound
OUTBOUND = '/Users/mac/.openclaw/media/outbound/英语学习_today.docx'

# 查找可能的docx文件
possible_paths = [
    f'/Users/mac/.openclaw/workspace/courses/英语学习_{topic_id}.docx',
    f'/Users/mac/.openclaw/workspace/courses/英语学习_{topic_id}_Beginner.docx',
]

for pp in possible_paths:
    if os.path.exists(pp):
        os.system(f'cp "{pp}" "{OUTBOUND}"')
        print(f"已复制到: {OUTBOUND}")
        print(f"原文件: {pp}")
        break
else:
    print(f"错误: 找不到 {topic_id} 的Word文档")
    sys.exit(1)

print(f"\n✅ 请发送附件: {OUTBOUND}")
print(f"📝 今日学习: {topic_id} - {topic_name}")
