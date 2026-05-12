#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务自检与自动修复机制
检查所有cron状态，自动修复失败任务，报告无法修复的问题
"""

import os
import re
import subprocess
from datetime import datetime

CRON_LIST_CMD = ['openclaw', 'cron', 'list']
CRON_RUN_CMD = ['openclaw', 'cron', 'run']
OUTBOUND = '/Users/mac/.openclaw/media/outbound/task_check.txt'

def run_cmd(cmd, capture=True):
    """执行shell命令"""
    result = subprocess.run(cmd, capture_output=capture, text=True)
    return result.stdout, result.stderr, result.returncode

def parse_cron_list():
    """解析cron列表，返回所有任务及状态"""
    output, _, _ = run_cmd(CRON_LIST_CMD)
    lines = output.strip().split('\n')
    
    # 跳过表头
    tasks = []
    for line in lines[1:]:
        if not line.strip():
            continue
        # 解析格式：ID  Name  Schedule  Next  Last  Status  Target  ...
        parts = line.split()
        if len(parts) >= 6:
            task_id = parts[0]
            name = ' '.join(parts[1:-5])
            status = parts[-5]
            tasks.append({
                'id': task_id,
                'name': name,
                'status': status,
                'raw': line
            })
    return tasks

def auto_fix(task_id, task_name):
    """自动修复失败任务"""
    print(f"  🔧 尝试自动修复: {task_name} ({task_id})")
    
    # 提取cron ID（第一个字段）
    cmd = CRON_RUN_CMD + [task_id]
    output, stderr, code = run_cmd(cmd)
    
    if code == 0 and 'enqueued' in output:
        return True, "已重新触发"
    else:
        return False, f"重新触发失败: {stderr[:100]}"

def check_and_fix():
    """检查所有cron并自动修复"""
    tasks = parse_cron_list()
    
    fixed = []
    failed = []
    ok = []
    
    for task in tasks:
        if task['status'] in ['error', 'fail']:
            success, msg = auto_fix(task['id'], task['name'])
            if success:
                fixed.append((task['name'], task['id'], msg))
            else:
                failed.append((task['name'], task['id'], msg))
        else:
            ok.append(task['name'])
    
    return fixed, failed, ok

def main():
    now = datetime.now().strftime('%Y年%m月%d日 %H:%M')
    
    fixed, failed, ok = check_and_fix()
    
    # 生成报告
    report_lines = [f"🐦 自检修复报告 - {now}"]
    report_lines.append("")
    
    if fixed:
        report_lines.append(f"✅ 已自动修复 {len(fixed)} 个任务：")
        for name, _, msg in fixed:
            report_lines.append(f"  • {name} → {msg}")
        report_lines.append("")
    
    if failed:
        report_lines.append(f"❌ 无法自动修复 {len(failed)} 个任务，需要手动处理：")
        for name, tid, msg in failed:
            report_lines.append(f"  • {name} (ID: {tid})")
            report_lines.append(f"    原因: {msg}")
        report_lines.append("")
    
    if not fixed and not failed:
        report_lines.append("✅ 所有任务状态正常，无需修复")
        report_lines.append("")
    
    report_lines.append(f"📊 正常运行任务: {len(ok)} 个")
    report_lines.append("")
    report_lines.append("---")
    report_lines.append(f"自检时间：{now}")
    
    report = '\n'.join(report_lines)
    print(report)
    
    # 保存报告
    with open(OUTBOUND, 'w', encoding='utf-8') as f:
        f.write(report)
    
    print(f"\n✅ 报告已保存: {OUTBOUND}")
    
    # 返回状态供调用方判断
    return len(failed) == 0

if __name__ == '__main__':
    main()
