#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
道德经每日智能发送脚本 v3
跨天进度跟踪，避免重复发送
"""

import os
import shutil
from datetime import date

COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND_DIR = '/Users/mac/.openclaw/media/outbound'

# 章节区间映射（每天发送3章）
CHAPTER_RANGES = [
    (1, 3), (4, 6), (7, 9), (10, 12), (13, 15), (16, 18), (19, 21),
    (22, 24), (25, 27), (28, 30), (31, 33), (34, 36), (37, 39), (40, 42),
    (43, 45), (46, 48), (49, 51), (52, 54), (55, 57), (58, 60), (61, 63),
    (64, 66), (67, 69), (70, 72), (73, 75), (76, 78), (79, 81), (82, 84),
    (85, 87), (88, 90), (91, 93), (94, 96), (97, 99), (100, 102)
]

def get_current_chapter_index():
    """获取当前应该发送的章节索引（0开始）"""
    # 检查是否已有进度记录
    progress_file = os.path.join(COURSES_DIR, '.dao_progress.txt')
    date_file = os.path.join(COURSES_DIR, '.dao_last_date.txt')
    today_str = date.today().isoformat()
    
    # 检查今天是否已经发过
    if os.path.exists(date_file):
        with open(date_file, 'r') as f:
            last_date = f.read().strip()
        if last_date == today_str:
            # 今天已经发过，读取当前进度
            if os.path.exists(progress_file):
                with open(progress_file, 'r') as f:
                    content = f.read().strip()
                if content:
                    idx = int(content)
                    return idx
    
    # 今天还没发，推进到下一个
    if os.path.exists(progress_file):
        with open(progress_file, 'r') as f:
            content = f.read().strip()
        if content:
            current_idx = int(content)
        else:
            current_idx = 0
    else:
        current_idx = 0
    
    # 推进到下一个（循环）
    next_idx = (current_idx + 1) % len(CHAPTER_RANGES)
    return next_idx

def save_progress(idx):
    progress_file = os.path.join(COURSES_DIR, '.dao_progress.txt')
    with open(progress_file, 'w') as f:
        f.write(str(idx))

def save_date():
    date_file = os.path.join(COURSES_DIR, '.dao_last_date.txt')
    with open(date_file, 'w') as f:
        f.write(date.today().isoformat())

def get_ppt_filename(chapters):
    """根据章节查找对应的PPT文件名"""
    start, end = chapters
    # 扫描courses目录找匹配文件
    for f in os.listdir(COURSES_DIR):
        if f.startswith('道德经') and f.endswith('.pptx'):
            # 提取文件名中的章节范围
            # 例如: 道德经_第24天.pptx 或 道德经_第58-60章.pptx
            if str(start) in f:
                return os.path.join(COURSES_DIR, f)
    return None

def main():
    # 获取当前应该发送的索引
    current_idx = get_current_chapter_index()
    chapters = CHAPTER_RANGES[current_idx]
    
    # 查找PPT文件
    ppt_name = get_ppt_filename(chapters)
    
    # 保存下一个进度
    next_idx = (current_idx + 1) % len(CHAPTER_RANGES)
    save_progress(next_idx)
    save_date()
    
    print(f"今日应发章节: {chapters[0]}-{chapters[1]}章")
    print(f"文件索引: {current_idx}")
    print(f"下次进度: {next_idx}")
    
    if ppt_name:
        print(f"PPT文件: {ppt_name}")
        out_file = os.path.join(OUTBOUND_DIR, '道德经_今日学习.pptx')
        shutil.copy2(ppt_name, out_file)
        print(f"✅ 已复制到: {out_file}")
    else:
        print(f"⚠️ 未找到第{chapters[0]}-{chapters[1]}章对应PPT")
        # 列出可用文件
        print("可用PPT文件:")
        for f in sorted(os.listdir(COURSES_DIR)):
            if f.startswith('道德经') and f.endswith('.pptx'):
                print(f"  {f}")

if __name__ == '__main__':
    main()