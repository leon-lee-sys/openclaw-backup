#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日晚8点统一提醒脚本
动态计算明天日期，发送对应日程
"""

import os
import re
from datetime import date, timedelta

TOMORROW_EVENTS_FILE = '/Users/mac/.openclaw/workspace/schedule/tomorrow-reminders.md'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/tomorrow_reminder.txt'

def get_tomorrow_date():
    return date.today() + timedelta(days=1)

def parse_schedule_file():
    """解析日程文件，找到明天的日程"""
    tomorrow = get_tomorrow_date()
    # 文件格式: ## 5月7日（周四）日程
    target_day = tomorrow.day

    if not os.path.exists(TOMORROW_EVENTS_FILE):
        return None

    with open(TOMORROW_EVENTS_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # 分割各日程块
    blocks = re.split(r'\n---\n', content)

    for block in blocks:
        block = block.strip()
        if not block:
            continue
        # 检查是否以"## 5月X日（周Y）日程"开头
        # 匹配: 5月7日 -> 月=5, 日=7
        match = re.match(r'## 5月(\d+)日（([^）]+)）日程', block)
        if match:
            day = int(match.group(1))
            if day == target_day:
                # 提取日程内容（去掉标题行）
                lines = block.split('\n')
                body_lines = []
                in_body = False
                for line in lines:
                    if line.startswith('## '):
                        in_body = True
                        continue
                    if in_body:
                        body_lines.append(line)
                body = '\n'.join(body_lines).strip()
                header = f"5月{day}日（{match.group(2)}）日程"
                if body:
                    return header, body
                return header, "无其他特殊安排"

    return None

def main():
    result = parse_schedule_file()

    if result is None:
        reminder_text = f"三先生，明天没有特殊日程安排 🐦\n\n祝您工作顺利！"
    else:
        header, body = result
        reminder_text = f"三先生，明天日程提醒 📅\n\n## {header}\n\n{body}\n\n请提前做好准备，准时参加！"

    # 保存到outbound
    with open(OUTBOUND, 'w', encoding='utf-8') as f:
        f.write(reminder_text)

    print(reminder_text)
    print(f"\n✅ 提醒文本已保存到: {OUTBOUND}")

if __name__ == '__main__':
    main()