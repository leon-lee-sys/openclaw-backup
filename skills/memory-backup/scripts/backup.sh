#!/bin/bash
# OpenClaw 对话历史自动备份脚本
# 用途：每天自动备份所有对话历史、记忆文件、工作区配置

OPENCLAW_DIR="$HOME/.openclaw"
WORKSPACE_DIR="$HOME/.openclaw/workspace"
BACKUP_DIR="$HOME/openclaw-backups"
TIMESTAMP=$(date +%Y%m%d_%H%M%S)

# 创建备份目录
mkdir -p "$BACKUP_DIR"

# 备份文件名
BACKUP_FILE="openclaw_backup_${TIMESTAMP}.tar.gz"

echo "=== OpenClaw 对话备份开始 ==="
echo "时间：$(date)"

# 备份关键文件
tar -czf "$BACKUP_DIR/$BACKUP_FILE" \
    "$WORKSPACE_DIR/memory/" \
    "$WORKSPACE_DIR/MEMORY.md" \
    "$WORKSPACE_DIR/SOUL.md" \
    "$WORKSPACE_DIR/USER.md" \
    "$OPENCLAW_DIR/agents/main/agent/auth-profiles.json" \
    "$OPENCLAW_DIR/openclaw.json" \
    2>/dev/null

if [ $? -eq 0 ]; then
    echo "✅ 备份成功：$BACKUP_FILE"
    
    # 保留最近10个备份
    cd "$BACKUP_DIR" || exit
    ls -t openclaw_backup_*.tar.gz 2>/dev/null | tail -n +11 | xargs -r rm 2>/dev/null
    echo "✅ 旧备份清理完成（保留最近10个）"
else
    echo "❌ 备份失败！"
    exit 1
fi

echo "=== 备份完成 ==="
