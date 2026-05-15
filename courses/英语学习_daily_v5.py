#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
英语学习每日生成并发送 - 直接发飞书
带进度跟踪，避免重复发送
"""

import os
import shutil
from datetime import date

COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/英语学习今日内容.docx'
TRACK_FILE = os.path.join(COURSES_DIR, '.learning_progress.txt')

# 7个级别循环
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
    """获取当前进度索引（0-6），如果没有记录则从0开始"""
    if os.path.exists(TRACK_FILE):
        with open(TRACK_FILE, 'r') as f:
            content = f.read().strip()
            if content:
                return int(content)
    return 0

def save_progress(idx):
    """保存当前进度索引"""
    with open(TRACK_FILE, 'w') as f:
        f.write(str(idx))

def main():
    # 获取当前进度
    current_idx = get_progress()
    
    # 读取日期检查是否需要推进到下一天
    date_file = os.path.join(COURSES_DIR, '.last_sent_date.txt')
    today_str = date.today().isoformat()
    
    # 检查今天是否已经发过
    if os.path.exists(date_file):
        with open(date_file, 'r') as f:
            last_date = f.read().strip()
        if last_date == today_str:
            # 今天已经发过，推进到下一个文件
            current_idx = (current_idx + 1) % len(FILES)
    else:
        # 第一次运行，从0开始（实际是下一个待发文件）
        pass
    
    # 使用当前索引的文件
    src, desc = FILES[current_idx]
    
    # 更新进度
    next_idx = (current_idx + 1) % len(FILES)
    save_progress(next_idx)
    
    # 记录今天日期
    with open(date_file, 'w') as f:
        f.write(today_str)
    
    if not os.path.exists(src):
        print(f"错误：找不到 {src}")
        return

    shutil.copy2(src, OUTBOUND)
    print(f"✅ 已复制: {desc}")
    print(f"📤 文件路径: {OUTBOUND}")
    print(f"📤 请发送附件: {OUTBOUND} 给三先生")

if __name__ == '__main__':
    main()