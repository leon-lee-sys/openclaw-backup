#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
道德经每日智能发送脚本 v2
直接读取进度追踪表，自动发送对应天的课件
"""
import os
import sys
import subprocess
from datetime import date

# ========== 配置 ==========
COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
OUTBOUND_DIR = '/Users/mac/.openclaw/media/outbound'

# 三先生确认的进度映射（文件第N天 → 对应章节）
# 第18天文件 = 58-60章(4月26日发)
# 第19天文件 = 61-63章(4月27日发) ← 我刚生成
# 第20天文件 = 64-66章(4月28日发)
DAY_CHAPTER_MAP = {
    18: (58, 60),  # 4月26日
    19: (61, 63),  # 4月27日
    20: (64, 66),  # 4月28日
    21: (67, 69),  # 4月29日
    22: (70, 72),  # 4月30日
    23: (73, 75),  # 5月1日
    24: (76, 78),  # 5月2日
}

def get_current_day():
    """根据今天日期计算应该发第几天（文件）"""
    # 三先生确认: 4月26日发第18天(58-60章)
    # 4月27日(今天)发第19天(61-63章)
    # 4月28日发第20天(64-66章)
    # ...
    # 对应关系: Day N 文件 = 3*(N-18+18-1)+1+6? 太复杂
    # 简单: 文件第N天的章节 = 3*(N-18+18)+52章?
    
    # 直接用已知映射
    base_day = 18  # 4月26日 = 第18天(58-60章)
    base_date = date(2026, 4, 26)
    today = date.today()
    days_diff = (today - base_date).days
    current_day = base_day + days_diff
    return current_day

def get_chapters(day):
    if day in DAY_CHAPTER_MAP:
        return DAY_CHAPTER_MAP[day]
    # 通用公式
    cs = 3 * (day - 18 + 18 - 1) + 1 + 6
    ce = cs + 2
    return cs, ce

def find_pptx(day):
    """找第N天PPT"""
    filename = f'道德经_第{day}天.pptx'
    path = os.path.join(COURSES_DIR, filename)
    if os.path.exists(path):
        return path, filename
    return None, filename

def copy_to_outbound(src):
    """复制到outbound目录"""
    dst = os.path.join(OUTBOUND_DIR, '道德经_daily.pptx')
    os.system(f'cp "{src}" "{dst}"')
    return dst

def main():
    today = date.today()
    day = get_current_day()
    cs, ce = get_chapters(day)
    
    print(f"今日日期: {today}")
    print(f"今日课件: 第{day}天文件 (第{cs}-{ce}章)")
    
    pptx_path, filename = find_pptx(day)
    
    if pptx_path is None:
        print(f"错误: 文件不存在 {filename}")
        print(f"需要生成第{day}天({cs}-{ce}章)的课件")
        sys.exit(1)
    
    out_path = copy_to_outbound(pptx_path)
    print(f"成功: 已复制到 {out_path}")
    print(f"请用 message 工具发送: {out_path}")
    print(f"附件标题: 道德经_第{day}天.pptx (第{cs}-{ce}章)")

if __name__ == '__main__':
    main()
