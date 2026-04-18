#!/bin/bash
# OpenClaw 对话历史转存脚本
# 用途：从 OpenClaw 导出当天对话历史，写入 memory/YYYY-MM-DD.md

OPENCLAW_CMD="/opt/homebrew/bin/openclaw"
WORKSPACE_DIR="$HOME/.openclaw/workspace"
TODAY=$(date +%Y-%m-%d)
MEMORY_FILE="$WORKSPACE_DIR/memory/$TODAY.md"

echo "=== 对话历史备份 ==="
echo "日期：$TODAY"

# 检查 openclaw 命令是否存在
if [ ! -f "$OPENCLAW_CMD" ]; then
    echo "❌ openclaw 命令不存在，跳过"
    exit 0
fi

# 使用 openclaw sessions list 获取会话列表
SESSION_LIST=$($OPENCLAW_CMD sessions list --limit 20 2>/dev/null)

if [ $? -eq 0 ] && [ -n "$SESSION_LIST" ]; then
    # 格式化对话历史
    echo "" >> "$MEMORY_FILE"
    echo "---" >> "$MEMORY_FILE"
    echo "## 对话历史备份 - $TODAY $(date +%H:%M:%S)" >> "$MEMORY_FILE"
    echo "" >> "$MEMORY_FILE"
    echo '```' >> "$MEMORY_FILE"
    echo "$SESSION_LIST" >> "$MEMORY_FILE"
    echo '```' >> "$MEMORY_FILE"
    echo "✅ 对话历史已追加到 $MEMORY_FILE"
else
    echo "⚠️ 无对话历史可备份"
fi

echo "=== 完成 ==="
