---
name: memory-backup
description: 自动备份小燕子和三先生所有对话历史、记忆文件、工作区配置的技能。当三先生说"备份"、"记忆备份"、"自动备份"、"聊天记录备份"、"配置备份"时触发。也用于每日自动定时备份（cron每天晚10点执行）。确保三先生的每次对话、每个决策、每条重要信息都不丢失。
---

# memory-backup 自动备份技能

## 功能概述

小燕子的记忆和对话历史是最宝贵的资产。本技能确保所有对话永不丢失，即使 session 被截断、模型重启，重要信息依然可以找回。

## 核心原则

**对话即记忆** —— 每一次对话结束都应该留下痕迹，而不是消失在虚空中。

## 备份内容

1. **memory/YYYY-MM-DD.md** — 当天日记（每日新建，自动追加）
2. **MEMORY.md** — 长期记忆核心文件（重要决策即刻更新）
3. **SOUL.md / USER.md / AGENTS.md** — 身份和配置
4. **工作区配置** — auth-profiles.json、openclaw.json
5. **cron 任务记录** — 每天的健康检查结果

## 触发场景

| 触发词 | 动作 |
|--------|------|
| "备份" | 执行一次完整备份（tar + git commit） |
| "记忆备份" | 同上 |
| "自动备份" | 检查 cron 是否存在，不存在则创建 |
| "今天发生什么" | 读取当天 memory/YYYY-MM-DD.md |
| "我记得之前..." | 先查 memory，不走 session 历史 |

## 工作流程

### 每日自动备份（cron: 22:00）
1. 执行 `backup.sh` 生成 tar 备份
2. 执行 `dump-conversations.sh` 转存当天对话到 memory 文件
3. 检查 cron 运行结果，失败则记录并重试
4. Git add + commit + push 到远程仓库

### 重要对话即刻记录
当三先生说：
- 重要信息（API key、电话、地址）
- 决策（"我们决定用XX方案"）
- 变更（"从今天开始..."）

→ 立即写入 MEMORY.md，不要等当天备份

### 备份脚本位置
```
~/.openclaw/workspace/skills/memory-backup/scripts/
├── backup.sh              # 完整备份（tar.gz）
├── dump-conversations.sh  # 对话历史转存
└── README.md              # 本文件
```

## 备份验证

每次备份后验证：
1. Git status 确认提交成功
2. 检查 memory 文件行数（不是空的）
3. 确认 cron 任务状态（"ok"）

## 常见问题

**Q: session 历史被截断了怎么办？**
A: 不依赖 session 历史，永远从 memory/YYYY-MM-DD.md 读取

**Q: 如何查找很久以前的对话？**
A: `grep -r "关键词" memory/` 全局搜索

**Q: 备份失败怎么办？**
A: 记录到 memory/cron-failures.md，人工触发重试
