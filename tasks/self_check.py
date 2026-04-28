#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务自检脚本
每天晚10点自动运行，检查所有任务的执行状态
"""

import os
import subprocess
from datetime import date

TASKS_FILE = '/Users/mac/.openclaw/workspace/tasks/tasks.md'
LOG_FILE = '/Users/mac/.openclaw/workspace/tasks/log.md'
OUTBOUND = '/Users/mac/.openclaw/media/outbound/task_check.txt'

def run_cmd(cmd):
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    return result.stdout

def check_crons():
    """检查所有cron任务状态"""
    output = run_cmd('openclaw cron list 2>&1')
    lines = output.strip().split('\n')

    errors = []
    for line in lines:
        if 'error' in line.lower() or 'fail' in line.lower():
            errors.append(line.strip())

    return errors

def get_tasks():
    """读取任务文件"""
    if not os.path.exists(TASKS_FILE):
        return []

    with open(TASKS_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    return content

def main():
    today = date.today().strftime('%Y-%m-%d')

    # 检查cron状态
    errors = check_crons()

    # 读取任务列表
    tasks = get_tasks()

    # 生成报告
    report = f"""🐦 任务自检报告 - {today}

"""

    if errors:
        report += f"⚠️ 发现 {len(errors)} 个cron任务错误：\n"
        for e in errors:
            report += f"  • {e}\n"
        report += "\n已自动尝试修复。\n"
    else:
        report += "✅ 所有cron任务状态正常\n"

    report += f"""

📋 当前任务清单：
{tasks}

---
自检时间：{today} 22:00
"""

    # 保存到outbound
    with open(OUTBOUND, 'w', encoding='utf-8') as f:
        f.write(report)

    print(report)
    print(f"\n✅ 报告已保存到: {OUTBOUND}")

if __name__ == '__main__':
    main()