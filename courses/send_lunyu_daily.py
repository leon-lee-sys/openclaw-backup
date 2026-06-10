#!/usr/bin/env python3
"""
论语每日课件发送脚本
每天11:45发送论语课件给三先生
"""
import os
import glob

COURSE_DIR = "/Users/mac/.openclaw/workspace/courses/论语"
SEND_LOG = "/Users/mac/.openclaw/workspace/courses/.lunyu_sent_log.txt"
USER_ID = "ou_128ad31d43d38fb3bb5f252161fd0a5e"

def get_next_pptx():
    """获取下一份未发送的课件"""
    # 读取已发送记录
    sent = set()
    if os.path.exists(SEND_LOG):
        with open(SEND_LOG, 'r') as f:
            sent = set(line.strip() for line in f if line.strip())
    
    # 查找所有课件
    all_ppts = sorted(glob.glob(os.path.join(COURSE_DIR, "论语_第*.pptx")))
    for ppt in all_ppts:
        filename = os.path.basename(ppt)
        if filename not in sent:
            return ppt, filename
    return None, None

def main():
    pptx_path, filename = get_next_pptx()
    if not pptx_path:
        print("No more slides to send")
        return
    
    print(f"Sending: {filename}")
    
    # 记录已发送
    with open(SEND_LOG, 'a') as f:
        f.write(filename + '\n')
    
    # 注意：实际发送由cron任务调用message工具完成
    # 这里只负责准备文件路径
    print(f"Ready to send: {pptx_path}")

if __name__ == "__main__":
    main()
