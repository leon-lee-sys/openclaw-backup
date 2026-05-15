#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英语学习每日生成并发送 - v6
规则：每天只发一次（10点），15点是检查如果10点已发就跳过
"""

import os
import shutil
from datetime import date

COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/英语学习今日内容.docx'
TRACK_FILE = os.path.join(COURSES_DIR, '.learning_progress.txt')
DATE_FILE = os.path.join(COURSES_DIR, '.last_sent_date.txt')

FILES = [
    (f'{COURSES_DIR}/英语学习_L1-5_AesopFables.docx', 'L1-5 初级'),
    (f'{COURSES_DIR}/英语学习_L6-10_AesopFables.docx', 'L6-10 中级'),
    (f'{COURSES_DIR}/英语学习_L11-15_AesopFables.docx', 'L11-15 中级'),
    (f'{COURSES_DIR}/英语学习_L16-20_AesopFables.docx', 'L16-20 进阶'),
    (f'{COURSES_DIR}/英语学习_L21-25_AesopFables.docx', 'L21-25 进阶'),
    (f'{COURSES_DIR}/英语学习_L26-30_AesopFables.docx', 'L26-30 高级'),
    (f'{COURSES_DIR}/英语学习_L31-35_AesopFables.docx', 'L31-35 高级'),
]

def get_progress():
    if os.path.exists(TRACK_FILE):
        with open(TRACK_FILE, 'r') as f:
            content = f.read().strip()
            if content:
                return int(content)
    return 0

def save_progress(idx):
    with open(TRACK_FILE, 'w') as f:
        f.write(str(idx))

def main():
    today_str = date.today().isoformat()
    
    # 检查今天是否已经发过
    if os.path.exists(DATE_FILE):
        with open(DATE_FILE, 'r') as f:
            last_date = f.read().strip()
        if last_date == today_str:
            print(f"⚠️ 今天已发送过，跳过（避免重复）")
            print(f"📅 上次发送日期: {last_date}")
            return
    
    # 今天还没发，获取当前进度并发送
    current_idx = get_progress()
    src, desc = FILES[current_idx]
    
    if not os.path.exists(src):
        print(f"错误：找不到 {src}")
        return
    
    shutil.copy2(src, OUTBOUND)
    print(f"✅ 已复制: {desc}")
    print(f"📤 文件路径: {OUTBOUND}")
    
    # 推进到下一个级别（为明天准备）
    next_idx = (current_idx + 1) % len(FILES)
    save_progress(next_idx)
    
    # 记录今天日期
    with open(DATE_FILE, 'w') as f:
        f.write(today_str)
    
    print(f"📝 进度已更新: {desc} → 下次发送 {FILES[next_idx][1]}")

if __name__ == '__main__':
    main()