# HEARTBEAT.md

## cron任务健康检查
每6小时检查一次cron任务状态，发现失败立即记录并尝试补救

### 检查命令
```bash
openclaw cron list 2>&1 | grep -E "error|fail"
```

### 失败处理
1. 记录失败任务到 memory/cron-failures.md
2. 尝试用 `openclaw cron run <job-id>` 重新触发
3. 如果重新触发成功，更新状态
4. 如果持续失败，发通知给用户

### 记录文件
- 失败记录：memory/cron-failures.md
- 每日检查：memory/YYYY-MM-DD.md
