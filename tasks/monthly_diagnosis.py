#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每月自动配置诊断与优化脚本
每月1号晚上10点运行
"""

import os
import subprocess
from datetime import datetime

COURSES_DIR = '/Users/mac/.openclaw/workspace/courses'
SCHEDULE_FILE = '/Users/mac/.openclaw/workspace/schedule/tomorrow-reminders.md'
CRON_STATE_FILE = '/Users/mac/.openclaw/workspace/memory/cron-diagnostics.md'

def check_crons():
    """检查所有cron状态"""
    result = subprocess.run(['openclaw', 'cron', 'list'], 
                          capture_output=True, text=True, stderr=subprocess.STDOUT)
    return result.stdout

def check_learning_files():
    """检查学习文件完整性"""
    issues = []
    
    # 英语学习文件
    for level in [1, 6, 11, 16, 21, 26, 31]:
        filename = f'英语学习_L{level}-{level+4}_AesopFables.docx'
        path = os.path.join(COURSES_DIR, filename)
        if not os.path.exists(path):
            issues.append(f"缺失: {filename}")
    
    return issues

def generate_report():
    """生成诊断报告"""
    lines = []
    lines.append(f"# 配置诊断报告 - {datetime.now().strftime('%Y年%m月%d日 %H:%M')}")
    lines.append("")
    
    # 1. Cron状态
    lines.append("## Cron状态检查")
    cron_output = check_crons()
    if 'error' in cron_output or 'fail' in cron_output:
        lines.append("⚠️ 发现error/fail状态的cron，需要处理")
    else:
        lines.append("✅ 所有cron正常")
    lines.append("")
    
    # 2. 学习文件完整性
    lines.append("## 学习文件完整性")
    file_issues = check_learning_files()
    if file_issues:
        for issue in file_issues:
            lines.append(f"❌ {issue}")
    else:
        lines.append("✅ 所有学习文件完整")
    lines.append("")
    
    lines.append("---")
    lines.append("诊断完成。如有问题请手动处理。")
    
    return '\n'.join(lines)

def main():
    report = generate_report()
    print(report)
    
    # 保存报告
    with open(CRON_STATE_FILE, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n✅ 报告已保存: {CRON_STATE_FILE}")

if __name__ == '__main__':
    main()
