#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日晚8点统一提醒脚本
读取日程文件，发送次日提醒
"""

import os
from datetime import date, timedelta

TOMORROW_EVENTS_FILE = '/Users/mac/.openclaw/workspace/schedule/tomorrow-reminders.md'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/tomorrow_reminder.txt'

def main():
    # 读取日程文件
    if not os.path.exists(TOMORROW_EVENTS_FILE):
        reminder_text = f"三先生，明天没有特殊日程安排 🐦\n\n祝您工作顺利！"
    else:
        with open(TOMORROW_EVENTS_FILE, 'r', encoding='utf-8') as f:
            content = f.read().strip()

        if content:
            reminder_text = f"三先生，明天日程提醒 📅\n\n{content}\n\n请提前做好准备，准时参加！"
        else:
            reminder_text = f"三先生，明天没有特殊日程安排 🐦\n\n祝您工作顺利！"

    # 保存到outbound
    with open(OUTBOUND, 'w', encoding='utf-8') as f:
        f.write(reminder_text)

    print(reminder_text)
    print(f"\n✅ 提醒文本已保存到: {OUTBOUND}")
    print("请用 message 工具发送给三先生")

if __name__ == '__main__':
    main()
